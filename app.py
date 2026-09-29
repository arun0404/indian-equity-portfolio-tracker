"""
Angel One Portfolio Manager
===========================

Streamlit dashboard that ingests Angel One holdings and trade-history exports
(.xlsx / .csv), enriches every position with live NSE/BSE prices from Yahoo
Finance and reports portfolio-level return metrics (P&L, XIRR, CAGR).

Run with::

    pip install -r requirements.txt
    streamlit run app.py
"""

from __future__ import annotations

import csv
import io
import math
import re
import warnings
from collections import deque
from collections.abc import Mapping, Sequence
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from scipy import optimize

try:
    import yfinance as yf
except ImportError:  # pragma: no cover - surfaced as a UI warning instead
    yf = None


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

IST = timezone(timedelta(hours=5, minutes=30))  # India has no DST
DAYS_PER_YEAR = 365.0
HEADER_SCAN_ROWS = 50  # how deep to look for the real header below metadata
PRICE_CACHE_TTL_SECONDS = 300
TOP_N_MOVERS = 3
ALLOCATION_SLICES = 5  # named donut slices; the remainder folds into "Others"

INR_FORMAT = "₹{:,.2f}"
PCT_FORMAT = "{:+.2f}%"

# Chart colours: validated categorical palette (light / dark steps), a
# neutral "Others" slice, and fixed status colours for gain / loss.
CATEGORICAL_LIGHT = ("#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4")
CATEGORICAL_DARK = ("#3987e5", "#d95926", "#199e70", "#c98500", "#d55181")
OTHERS_COLOR = "#898781"
GAIN_COLOR = "#0ca30c"
LOSS_COLOR = "#d03b3b"
SURFACE_LIGHT = "#ffffff"
SURFACE_DARK = "#0e1117"

UNCATEGORIZED = "Uncategorized"

# Market-cap buckets, in INR crore (1 crore = 1e7).
LARGE_CAP_MIN_CR = 20_000.0
MID_CAP_MIN_CR = 5_000.0

# Capital-gains rules for Indian listed equity (FY2024-25 onward): short-term
# is a flat rate with no exemption; long-term has a per-financial-year
# exemption before the concessional rate applies.
LTCG_HOLDING_DAYS = 365
STCG_TAX_RATE = 0.20
LTCG_TAX_RATE = 0.125
LTCG_EXEMPTION = 125_000.0

# HHI concentration bands (0-10,000 scale: sum of squared percentage shares) -
# the standard US DOJ/FTC merger-guideline thresholds, repurposed here to
# flag sector concentration risk rather than market concentration.
HHI_MODERATE_THRESHOLD = 1_500.0
HHI_HIGH_THRESHOLD = 2_500.0

# Excel stores dates as days since 1899-12-30 (accounting for the 1900 bug).
EXCEL_EPOCH = pd.Timestamp("1899-12-30")
EXCEL_MAX_SERIAL = 2_958_465  # 9999-12-31

_ISO_DATE = re.compile(r"^\d{4}-\d{1,2}-\d{1,2}")
_EXCHANGE_PREFIX = re.compile(r"^(NSE|BSE)\s*[:_]\s*(.+)$")
_SERIES_SUFFIX = re.compile(r"[\s-]+(EQ|BE|BZ|BL|SM|ST|IL)$")
_TOTAL_ROW = r"(?i)^(grand |sub |net )?total\b"
_ISIN_RE = r"^[A-Z]{2}[A-Z0-9]{9}[0-9]$"
_HEADER_NOISE_TOKENS = frozenset({"rs", "inr"})
_NON_ALNUM = re.compile(r"[^A-Z0-9]")

# Static overrides for names Yahoo Finance's own ISIN/text search can't
# resolve, or that a broker export spells out instead of using the trading
# symbol. Keyed loosely - matched after stripping everything but letters and
# digits (see `_condense`), so "Kotak Gold ETF", "KOTAK GOLD ETF" and
# "KOTAKGOLDETF" are all the same key. Consulted first, in `to_yahoo_ticker`;
# ISIN-based resolution (`resolve_isin_tickers`, in `resolve_prices`) still
# runs afterward and overrides this when it finds a better match, so this
# map only has the final say when ISIN search comes back empty - as it does
# for Kotak Gold ETF's ISIN (confirmed: Yahoo indexes it by name, not ISIN).
ANGEL_ONE_SYMBOL_MAP: dict[str, str] = {
    "INFOSYS": "INFY.NS",
    "INFOSYS LIMITED": "INFY.NS",
    "INFOSYS - EQ": "INFY.NS",
    "INFY": "INFY.NS",
    "KOTAKGOLD": "GOLD1.NS",
    "KOTAKGOLDETF": "GOLD1.NS",
    "KOTAK GOLD ETF": "GOLD1.NS",
    "KOTAK GOLD": "GOLD1.NS",
    "GOLD1": "GOLD1.NS",
    "GOLDBEES": "GOLDBEES.NS",
    "NIFTYBEES": "NIFTYBEES.NS",
}


def _condense(text: str) -> str:
    """Strip everything but letters/digits and upper-case, for loose symbol matching."""
    return _NON_ALNUM.sub("", text.upper())


_SYMBOL_MAP_BY_CONDENSED = {_condense(k): v for k, v in ANGEL_ONE_SYMBOL_MAP.items()}


class ParseError(ValueError):
    """Raised when an uploaded file cannot be interpreted as the expected report."""


# ---------------------------------------------------------------------------
# Column schemas (alias lists are in priority order, already normalised)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class TableSchema:
    """Alias specification for one export layout.

    Attributes:
        name: Human-readable layout name used in error messages.
        aliases: Canonical field -> accepted normalised header spellings.
        required: Groups of fields; at least one field of every group must be
            present for a header row to qualify.
    """

    name: str
    aliases: Mapping[str, tuple[str, ...]]
    required: tuple[tuple[str, ...], ...]


SYMBOL_ALIASES = (
    "trading symbol", "tradingsymbol", "symbol", "stock symbol", "scrip",
    "scrip name", "scrip symbol", "stock", "stock name", "instrument",
    "instrument name", "security", "security name", "company",
    "company name", "name",
)
EXCHANGE_ALIASES = ("exchange", "exch", "exchange segment", "segment")

HOLDINGS_SCHEMA = TableSchema(
    name="Holdings report",
    aliases={
        "symbol": SYMBOL_ALIASES,
        "quantity": (
            "quantity", "qty", "total qty", "total quantity", "net qty",
            "net quantity", "holding qty", "holdings qty", "holding quantity",
            "balance qty", "no of shares", "shares", "units", "available qty",
            "free qty",
        ),
        "average_price": (
            "average price", "avg price", "average cost", "avg cost",
            "average cost price", "avg cost price", "average buy price",
            "avg buy price", "buy avg price", "buy average", "buy avg",
            "avg rate", "buy price", "cost price", "purchase price",
            "avg trading price", "average trading price", "avg cost per share",
            "average cost per share",
        ),
        "ltp": (
            "ltp", "last traded price", "last price", "current price",
            "current market price", "cmp", "market price", "close price",
            "closing price",
        ),
        "exchange": EXCHANGE_ALIASES,
        "isin": ("isin", "isin code", "isin no", "isin number"),
    },
    required=(("symbol",), ("quantity",), ("average_price",)),
)

TRANSACTION_SCHEMA = TableSchema(
    name="Trade book (one row per buy/sell)",
    aliases={
        "date": (
            "trade date", "date", "order date", "transaction date", "txn date",
            "execution date", "trade date time", "order date time",
            "date time", "datetime", "timestamp",
        ),
        "symbol": SYMBOL_ALIASES,
        "type": (
            "buy sell", "buy or sell", "b s", "type", "trade type",
            "transaction type", "txn type", "side", "action", "transaction",
        ),
        "quantity": (
            "traded qty", "trade qty", "filled qty", "executed qty",
            "quantity", "qty", "total qty", "net qty",
        ),
        "price": (
            "trade price", "traded price", "execution price",
            "avg traded price", "average traded price", "avg price",
            "average price", "price", "rate", "trade rate", "net rate",
            "net price",
        ),
        "value": (
            "trade value", "traded value", "value", "amount", "net amount",
            "total value", "turnover",
        ),
        "exchange": EXCHANGE_ALIASES,
    },
    required=(("date",), ("symbol",), ("type",), ("quantity",), ("price", "value")),
)

PAIRED_PNL_SCHEMA = TableSchema(
    name="P&L statement (buy and sell legs on one row)",
    aliases={
        "symbol": SYMBOL_ALIASES,
        "quantity": ("quantity", "qty", "total qty", "net qty"),
        "buy_date": (
            "buy date", "purchase date", "date of purchase", "buy trade date",
            "entry date", "acquisition date",
        ),
        "buy_qty": ("buy qty", "buy quantity", "purchase qty", "purchase quantity"),
        "buy_price": (
            "buy price", "buy rate", "buy avg price", "buy average price",
            "avg buy price", "average buy price", "buy avg", "purchase price",
            "purchase rate", "entry price",
        ),
        "buy_value": (
            "buy value", "buy amount", "purchase value", "purchase amount",
            "total buy value", "buy turnover",
        ),
        "sell_date": (
            "sell date", "sale date", "date of sale", "sell trade date", "exit date",
        ),
        "sell_qty": ("sell qty", "sell quantity", "sale qty", "sale quantity"),
        "sell_price": (
            "sell price", "sell rate", "sell avg price", "sell average price",
            "avg sell price", "average sell price", "sell avg", "sale price",
            "sale rate", "exit price",
        ),
        "sell_value": (
            "sell value", "sell amount", "sale value", "sale amount",
            "total sell value", "sell turnover", "sale consideration",
        ),
        "exchange": EXCHANGE_ALIASES,
    },
    required=(("symbol",), ("buy_date",), ("quantity", "buy_qty"), ("buy_price", "buy_value")),
)

MUTUAL_FUND_SCHEMA = TableSchema(
    name="Mutual Fund holdings",
    aliases={
        "fund_name": (
            "fund name", "scheme name", "scheme", "fund", "mf name",
            "mutual fund", "mutual fund name",
        ),
        "isin": ("isin", "isin code", "isin no", "isin number"),
        "units": ("units", "unit", "qty", "quantity", "total units", "balance units"),
        "average_nav": (
            "average nav", "avg nav", "purchase nav", "avg cost", "average cost",
            "avg purchase nav", "cost nav",
        ),
        # "NAV as on <date>" normalises to "nav as on" - see _strip_trailing_date.
        "current_nav": ("nav as on", "current nav", "latest nav", "nav", "cmp"),
        "invested_value": (
            "invested value", "investment value", "amount invested",
            "purchase value", "invested amount", "cost value",
        ),
        # "Market Value as on <date>" normalises to "market value as on".
        "current_value": (
            "market value as on", "market value", "current value", "present value",
        ),
        "gain_loss": (
            "overall gain loss", "gain loss", "profit loss", "p and l",
            "unrealised gain loss", "unrealized gain loss",
        ),
        "xirr_pct": ("xirr", "xirr pct", "annualised return", "annualized return"),
    },
    required=(
        ("fund_name",), ("units",),
        ("average_nav", "invested_value"),
        ("current_nav", "current_value"),
    ),
)

HOLDINGS_COLUMNS = ["symbol", "ticker", "quantity", "average_price", "ltp", "isin"]
TRADE_COLUMNS = ["date", "symbol", "ticker", "type", "quantity", "price", "cash_flow"]
MUTUAL_FUND_COLUMNS = [
    "fund_name", "isin", "units", "average_nav", "current_nav",
    "invested_value", "current_value", "xirr_pct",
]


# ---------------------------------------------------------------------------
# Low-level cleaning helpers
# ---------------------------------------------------------------------------


def _strip_trailing_date(tokens: list[str]) -> list[str]:
    """Drop a trailing run of numeric tokens if it looks like an embedded date.

    Some Angel One reports bake the download date into a header, e.g. ``"NAV
    as on 2026-09-08"`` or ``"Market Value as on 2026-09-08"`` - a fresh
    download shifts the date, so a literal alias would only ever match one
    day. Requires at least one non-numeric token before the run (never
    strips a wholly-numeric header) and at least one 4-digit token in it (a
    year), so ordinary headers with a trailing number are untouched.
    """
    idx = len(tokens)
    while idx > 0 and tokens[idx - 1].isdigit():
        idx -= 1
    trailing = tokens[idx:]
    if idx > 0 and any(len(t) == 4 for t in trailing):
        return tokens[:idx]
    return tokens


def normalize_header(value: object) -> str:
    """Canonicalise a header cell to lowercase alphanumeric tokens.

    Examples: ``"Avg. Price (₹)"`` -> ``"avg price"``, ``"Buy/Sell"`` ->
    ``"buy sell"``, ``"LTP (Rs.)"`` -> ``"ltp"``, ``"NAV as on 2026-09-08"``
    -> ``"nav as on"`` (see :func:`_strip_trailing_date`).
    """
    text = clean_text(value).lower().replace("&", " and ")
    tokens = re.findall(r"[a-z0-9]+", text)
    tokens = [t for t in tokens if t not in _HEADER_NOISE_TOKENS]
    return " ".join(_strip_trailing_date(tokens))


def clean_text(value: object) -> str:
    """Return a stripped string for any cell value ("" for blanks / NaN)."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and value.is_integer():
        return str(int(value))  # BSE scrip codes read from Excel as floats
    return str(value).strip()


def to_number(series: pd.Series) -> pd.Series:
    """Coerce broker-formatted numbers to floats.

    Handles thousands separators, currency symbols, percent signs, unicode
    minus and accounting negatives such as ``(1,234.50)``. Anything that still
    cannot be parsed becomes NaN.
    """
    if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
        return series.astype(float)
    text = series.map(clean_text).str.replace("−", "-", regex=False)
    negative = text.str.fullmatch(r"\(.*\)")
    cleaned = text.str.replace(r"(?i)rs\.?|inr|[₹,\s%()]", "", regex=True)
    numbers = pd.to_numeric(cleaned, errors="coerce").astype(float)
    return numbers.where(~negative, -numbers.abs())


def _parse_one_date(value: object) -> pd.Timestamp:
    """Parse a single date cell (datetime, Excel serial, or text) to a naive date."""
    if value is None or isinstance(value, bool):
        return pd.NaT
    try:
        if isinstance(value, (pd.Timestamp, datetime, date, np.datetime64)):
            ts = pd.Timestamp(value)
        elif isinstance(value, (int, float, np.integer, np.floating)):
            if not np.isfinite(value) or not 1 <= value <= EXCEL_MAX_SERIAL:
                return pd.NaT
            ts = EXCEL_EPOCH + pd.to_timedelta(float(value), unit="D")
        else:
            text = str(value).strip()
            if not text:
                return pd.NaT
            if re.fullmatch(r"\d{8}", text):  # 20240115
                ts = pd.to_datetime(text, format="%Y%m%d", errors="coerce")
            else:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    # Indian exports are day-first; ISO strings must not be.
                    ts = pd.to_datetime(
                        text, dayfirst=not _ISO_DATE.match(text), errors="coerce"
                    )
    except (ValueError, TypeError, OverflowError):
        return pd.NaT
    if pd.isna(ts):
        return pd.NaT
    if ts.tzinfo is not None:
        ts = ts.tz_convert(IST).tz_localize(None)
    return ts.normalize()


def parse_dates(series: pd.Series) -> pd.Series:
    """Parse a column of heterogeneous date cells into ``datetime64`` (NaT if invalid)."""
    return pd.to_datetime(series.map(_parse_one_date), errors="coerce")


def normalize_side(series: pd.Series) -> pd.Series:
    """Map free-text transaction types (B, Buy, Purchase, S, Sold...) to BUY/SELL/""."""
    text = series.map(clean_text).str.upper()
    side = np.select(
        [
            text.str.match(r"^(B|BUY|BOUGHT|PURCHASE|P)\b"),
            text.str.match(r"^(S|SELL|SOLD|SALE)\b"),
        ],
        ["BUY", "SELL"],
        default="",
    )
    return pd.Series(side, index=series.index)


def to_yahoo_ticker(raw_symbol: str, exchange: str = "") -> tuple[str, str]:
    """Convert an Angel One symbol to ``(display_symbol, yahoo_ticker)``.

    ``RELIANCE`` -> ``RELIANCE.NS``; ``RELIANCE-EQ`` -> ``RELIANCE.NS``;
    ``NSE:INFY`` -> ``INFY.NS``; BSE rows and numeric BSE scrip codes -> ``.BO``.
    Symbols that already carry a ``.NS`` / ``.BO`` suffix keep it. Checked
    first against :data:`ANGEL_ONE_SYMBOL_MAP` for names that don't convert
    to a real ticker by this heuristic alone (spelled-out company names,
    ETFs Yahoo's ISIN search can't find) - skipped when the row explicitly
    says BSE, since the map's NSE tickers would otherwise silently override
    a real exchange hint.
    """
    symbol = raw_symbol.strip().upper()
    exchange = exchange.strip().upper()
    if "BSE" not in exchange:
        mapped = _SYMBOL_MAP_BY_CONDENSED.get(_condense(symbol))
        if mapped:
            return mapped.rsplit(".", 1)[0], mapped

    prefixed = _EXCHANGE_PREFIX.match(symbol)
    if prefixed:
        exchange, symbol = prefixed.group(1), prefixed.group(2)
    if symbol.endswith((".NS", ".BO")):
        symbol, suffix = symbol[:-3], symbol[-3:]
    else:
        suffix = ".BO" if ("BSE" in exchange or symbol.isdigit()) else ".NS"
    symbol = _SERIES_SUFFIX.sub("", symbol).replace(" ", "")
    return symbol, f"{symbol}{suffix}"


def _add_tickers(frame: pd.DataFrame) -> pd.DataFrame:
    """Replace ``raw_symbol`` / ``exchange`` with normalised ``symbol`` and ``ticker``."""
    pairs = [
        to_yahoo_ticker(sym, exch)
        for sym, exch in zip(frame["raw_symbol"], frame["exchange"], strict=True)
    ]
    out = frame.drop(columns=["raw_symbol", "exchange"])
    out["symbol"] = [p[0] for p in pairs]
    out["ticker"] = [p[1] for p in pairs]
    return out


# ---------------------------------------------------------------------------
# File reading & header detection
# ---------------------------------------------------------------------------


def _read_csv_grid(content: bytes) -> pd.DataFrame:
    """Read a CSV into a header-less grid, tolerating ragged metadata rows."""
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = content.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    try:
        delimiter = csv.Sniffer().sniff(text[:20_000], delimiters=",;\t|").delimiter
    except csv.Error:
        delimiter = ","
    rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    width = max((len(r) for r in rows), default=0)
    if width == 0:
        raise ParseError("The CSV file is empty.")
    grid = [[cell.strip() or None for cell in r] + [None] * (width - len(r)) for r in rows]
    return pd.DataFrame(grid, dtype=object)


def read_sheets(content: bytes, filename: str) -> dict[str, pd.DataFrame]:
    """Load every sheet of an ``.xlsx`` / ``.csv`` upload as a raw header-less grid."""
    if not content:
        raise ParseError(f"'{filename}' is empty.")
    suffix = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    if suffix in {"xlsx", "xlsm"}:
        try:
            sheets = pd.read_excel(
                io.BytesIO(content), sheet_name=None, header=None, engine="openpyxl"
            )
        except Exception as exc:  # openpyxl raises a zoo of exception types
            raise ParseError(f"Could not open '{filename}' as an Excel workbook ({exc}).") from exc
        sheets = {name: df for name, df in sheets.items() if not df.dropna(how="all").empty}
        if not sheets:
            raise ParseError(f"'{filename}' contains no data.")
        return sheets
    if suffix in {"csv", "txt"}:
        return {"CSV": _read_csv_grid(content)}
    raise ParseError(f"Unsupported file type '.{suffix}'. Upload an .xlsx or .csv export.")


@dataclass(frozen=True)
class ExtractedTable:
    """A report table located inside a raw grid, with canonical column names."""

    schema: TableSchema
    frame: pd.DataFrame
    sheet: str
    header_row: int


def _match_header(cells: Sequence[str], schema: TableSchema) -> dict[str, int]:
    """Map schema fields to column positions for one candidate header row."""
    first_position: dict[str, int] = {}
    for idx, cell in enumerate(cells):
        if cell:
            first_position.setdefault(cell, idx)
    positions: dict[str, int] = {}
    for field_name, aliases in schema.aliases.items():
        for alias in aliases:
            if alias in first_position:
                positions[field_name] = first_position[alias]
                break
    return positions


def _missing_groups(positions: Mapping[str, int], schema: TableSchema) -> list[tuple[str, ...]]:
    """Return the required field groups that a header mapping fails to satisfy."""
    return [group for group in schema.required if not any(f in positions for f in group)]


def extract_table(
    sheets: Mapping[str, pd.DataFrame], schemas: Sequence[TableSchema]
) -> ExtractedTable:
    """Find the best header row across sheets/layouts and slice out the data below it.

    Every one of the first ``HEADER_SCAN_ROWS`` rows of every sheet is scored
    against every schema; the row that satisfies all required groups with the
    most recognised columns wins (earliest row on ties). This skips the
    client-name / date-range / disclaimer rows that brokers prepend.

    Raises:
        ParseError: If no row satisfies any schema. The message lists what
            was missing so the user can fix the file.
    """
    best: tuple[int, TableSchema, str, int, dict[str, int]] | None = None
    nearest: tuple[int, TableSchema, list[tuple[str, ...]], list[str]] | None = None

    for sheet_name, raw in sheets.items():
        for row_idx in range(min(len(raw), HEADER_SCAN_ROWS)):
            cells = [normalize_header(v) for v in raw.iloc[row_idx].tolist()]
            if not any(cells):
                continue
            for schema in schemas:
                positions = _match_header(cells, schema)
                if not positions:
                    continue
                missing = _missing_groups(positions, schema)
                score = len(positions)
                if not missing:
                    if best is None or score > best[0]:
                        best = (score, schema, sheet_name, row_idx, positions)
                elif nearest is None or score > nearest[0]:
                    nearest = (score, schema, missing, [c for c in cells if c])

    if best is None:
        expected = " / ".join(s.name for s in schemas)
        if nearest is None:
            raise ParseError(
                f"No recognisable column headers found (expected a {expected})."
            )
        _, schema, missing, seen = nearest
        wanted = "; ".join(
            " or ".join(f"'{schema.aliases[f][0]}'" for f in group) for group in missing
        )
        raise ParseError(
            f"Closest match was a {schema.name}, but it is missing: {wanted}. "
            f"Columns found: {', '.join(seen[:15])}."
        )

    _, schema, sheet_name, row_idx, positions = best
    body = sheets[sheet_name].iloc[row_idx + 1 :]
    frame = pd.DataFrame(
        {name: body.iloc[:, pos].to_numpy() for name, pos in positions.items()}
    ).dropna(how="all")
    return ExtractedTable(schema=schema, frame=frame, sheet=sheet_name, header_row=row_idx)


# ---------------------------------------------------------------------------
# Holdings & trade-history parsing
# ---------------------------------------------------------------------------


def clean_holdings(frame: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Standardise a raw holdings table to ``HOLDINGS_COLUMNS``.

    Drops footer/total rows and closed positions, merges duplicate tickers
    (e.g. pledged and free lots on separate rows) with a quantity-weighted
    average price, and derives Yahoo Finance tickers.

    Returns:
        The cleaned holdings and a list of human-readable data-quality notes.

    Raises:
        ParseError: If no valid position survives cleaning.
    """
    notes: list[str] = []
    isin = frame["isin"].map(clean_text).str.upper() if "isin" in frame else ""
    df = pd.DataFrame(
        {
            "raw_symbol": frame["symbol"].map(clean_text),
            "quantity": to_number(frame["quantity"]),
            "average_price": to_number(frame["average_price"]),
            "ltp": to_number(frame["ltp"]) if "ltp" in frame else np.nan,
            "exchange": frame["exchange"].map(clean_text) if "exchange" in frame else "",
            "isin": isin,
        }
    )
    df["isin"] = df["isin"].where(df["isin"].str.match(_ISIN_RE, na=False), "")
    df = df[(df["raw_symbol"] != "") & ~df["raw_symbol"].str.match(_TOTAL_ROW)]
    # Text-only rows (disclaimers, section titles) are dropped without a note.
    df = df[df[["quantity", "average_price", "ltp"]].notna().any(axis=1)]

    unreadable = df["quantity"].isna()
    if unreadable.any():
        notes.append(f"Skipped {int(unreadable.sum())} row(s) without a numeric quantity.")
    df = df[~unreadable]

    closed = df["quantity"] <= 0
    if closed.any():
        notes.append(f"Skipped {int(closed.sum())} closed / zero-quantity position(s).")
    df = df[~closed]

    bad_cost = df["average_price"].isna() | (df["average_price"] < 0)
    if bad_cost.any():
        names = ", ".join(df.loc[bad_cost, "raw_symbol"].head(10))
        notes.append(f"Skipped position(s) with a missing/invalid average price: {names}.")
    df = df[~bad_cost]

    df["ltp"] = df["ltp"].where(df["ltp"] > 0)
    df = _add_tickers(df)
    if df.empty:
        raise ParseError("No valid positions (symbol + quantity + average price) were found.")

    # Same stock on several rows (pledged/free lots, NSE and BSE lines): one position.
    duplicated = df["symbol"].duplicated()
    if duplicated.any():
        names = ", ".join(sorted(df.loc[duplicated, "symbol"].unique()))
        notes.append(f"Merged duplicate rows for: {names} (quantity-weighted average price).")
    df["cost"] = df["quantity"] * df["average_price"]
    merged = df.groupby("symbol", sort=False, as_index=False).agg(
        ticker=("ticker", "first"),
        quantity=("quantity", "sum"),
        cost=("cost", "sum"),
        ltp=("ltp", "first"),
        isin=("isin", "first"),
    )
    merged["average_price"] = merged["cost"] / merged["quantity"]
    return merged[HOLDINGS_COLUMNS].reset_index(drop=True), notes


def _unpivot_paired_pnl(frame: pd.DataFrame) -> pd.DataFrame:
    """Split a P&L statement (buy + sell legs per row) into one row per leg."""
    blank = pd.Series(np.nan, index=frame.index, dtype=object)
    common_qty = frame.get("quantity", blank)

    legs = []
    for side in ("buy", "sell"):
        if f"{side}_date" not in frame:
            continue
        side_qty = frame.get(f"{side}_qty", blank)
        leg = pd.DataFrame(
            {
                "date": frame[f"{side}_date"],
                "symbol": frame["symbol"],
                "type": side.upper(),
                "quantity": side_qty.where(side_qty.notna(), common_qty),
                "price": frame.get(f"{side}_price", blank),
                "value": frame.get(f"{side}_value", blank),
                "exchange": frame.get("exchange", pd.Series("", index=frame.index)),
            }
        )
        # Open positions have no sell leg - drop those silently.
        legs.append(leg[leg["date"].map(clean_text) != ""])
    return pd.concat(legs, ignore_index=True)


def clean_trades(frame: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Standardise a trade log to ``TRADE_COLUMNS`` with signed cash flows.

    Cash-flow convention (investor's perspective): BUY = negative (cash out),
    SELL = positive (cash in). Missing prices are derived from trade value /
    quantity when a value column exists.

    Returns:
        Trades sorted by date plus human-readable data-quality notes.

    Raises:
        ParseError: If no valid trade survives cleaning.
    """
    notes: list[str] = []
    quantity = to_number(frame["quantity"]).abs()
    price = to_number(frame["price"]) if "price" in frame else pd.Series(np.nan, index=frame.index)
    if "value" in frame:
        derived = to_number(frame["value"]).abs() / quantity.replace(0, np.nan)
        price = price.where(price > 0, derived)
    df = pd.DataFrame(
        {
            "date": parse_dates(frame["date"]),
            "raw_symbol": frame["symbol"].map(clean_text),
            "type": normalize_side(frame["type"]),
            "quantity": quantity,
            "price": price,
            "exchange": frame["exchange"].map(clean_text) if "exchange" in frame else "",
        }
    )
    df = df[(df["raw_symbol"] != "") & ~df["raw_symbol"].str.match(_TOTAL_ROW)]
    # Text-only rows (disclaimers, section titles) are dropped without a note.
    df = df[df["date"].notna() | df["quantity"].notna() | df["price"].notna()]

    checks = [
        (df["date"].isna(), "an unreadable date"),
        (df["type"] == "", "an unrecognised BUY/SELL type"),
        (~(df["quantity"] > 0), "a missing/zero quantity"),
        (~(df["price"] > 0), "a missing/zero price"),
    ]
    invalid = pd.Series(False, index=df.index)
    for mask, reason in checks:
        new = mask & ~invalid
        if new.any():
            notes.append(f"Skipped {int(new.sum())} row(s) with {reason}.")
        invalid |= mask
    df = df[~invalid]
    if df.empty:
        raise ParseError("No valid trades (date + symbol + BUY/SELL + quantity + price) found.")

    df = _add_tickers(df)
    direction = np.where(df["type"] == "BUY", -1.0, 1.0)
    df["cash_flow"] = direction * df["quantity"] * df["price"]
    return df[TRADE_COLUMNS].sort_values("date", kind="stable").reset_index(drop=True), notes


@st.cache_data(show_spinner=False)
def parse_holdings_file(content: bytes, filename: str) -> tuple[pd.DataFrame, list[str]]:
    """Parse an Angel One holdings export (cached on file content)."""
    table = extract_table(read_sheets(content, filename), [HOLDINGS_SCHEMA])
    return clean_holdings(table.frame)


@st.cache_data(show_spinner=False)
def parse_trades_file(content: bytes, filename: str) -> tuple[pd.DataFrame, list[str]]:
    """Parse an Angel One trade book or P&L statement (cached on file content)."""
    table = extract_table(
        read_sheets(content, filename), [TRANSACTION_SCHEMA, PAIRED_PNL_SCHEMA]
    )
    frame = table.frame
    if table.schema is PAIRED_PNL_SCHEMA:
        frame = _unpivot_paired_pnl(frame)
    trades, notes = clean_trades(frame)
    return trades, [f"Detected layout: {table.schema.name}.", *notes]


def clean_mutual_funds(frame: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Standardise a raw mutual-fund holdings table to ``MUTUAL_FUND_COLUMNS``.

    Whichever of average/current NAV vs. invested/current value is missing
    per row is cross-derived from the other (plus units). Angel One's own
    per-fund XIRR is kept as-is - a fund's cash-flow history isn't available
    from a holdings snapshot, so it can't be recomputed the way equity XIRR
    is from the trade book.

    Returns:
        The cleaned funds and a list of human-readable data-quality notes.

    Raises:
        ParseError: If no valid fund survives cleaning.
    """
    notes: list[str] = []
    isin = frame["isin"].map(clean_text).str.upper() if "isin" in frame else ""
    df = pd.DataFrame(
        {
            "fund_name": frame["fund_name"].map(clean_text),
            "isin": isin,
            "units": to_number(frame["units"]),
            "average_nav": to_number(frame["average_nav"]) if "average_nav" in frame else np.nan,
            "current_nav": to_number(frame["current_nav"]) if "current_nav" in frame else np.nan,
            "invested_value": (
                to_number(frame["invested_value"]) if "invested_value" in frame else np.nan
            ),
            "current_value": (
                to_number(frame["current_value"]) if "current_value" in frame else np.nan
            ),
            "xirr_pct": to_number(frame["xirr_pct"]) if "xirr_pct" in frame else np.nan,
        }
    )
    df["isin"] = df["isin"].where(df["isin"].str.match(_ISIN_RE, na=False), "")
    df = df[(df["fund_name"] != "") & ~df["fund_name"].str.match(_TOTAL_ROW)]
    # Text-only rows (disclaimers, section titles) are dropped without a note.
    df = df[df[["units", "average_nav", "invested_value"]].notna().any(axis=1)]

    unreadable = df["units"].isna()
    if unreadable.any():
        notes.append(f"Skipped {int(unreadable.sum())} row(s) without numeric units.")
    df = df[~unreadable]

    zero_units = df["units"] <= 0
    if zero_units.any():
        notes.append(f"Skipped {int(zero_units.sum())} closed / zero-unit fund(s).")
    df = df[~zero_units]

    # Cross-derive whichever of NAV / value is missing, per row.
    df["invested_value"] = df["invested_value"].where(
        df["invested_value"].notna(), df["units"] * df["average_nav"]
    )
    df["average_nav"] = df["average_nav"].where(
        df["average_nav"].notna(), df["invested_value"] / df["units"]
    )
    df["current_value"] = df["current_value"].where(
        df["current_value"].notna(), df["units"] * df["current_nav"]
    )
    df["current_nav"] = df["current_nav"].where(
        df["current_nav"].notna(), df["current_value"] / df["units"]
    )

    bad_cost = df["average_nav"].isna() | (df["average_nav"] < 0)
    if bad_cost.any():
        names = ", ".join(df.loc[bad_cost, "fund_name"].head(10))
        notes.append(f"Skipped fund(s) with a missing/invalid average NAV: {names}.")
    df = df[~bad_cost]
    if df.empty:
        raise ParseError("No valid mutual fund holdings (name + units + NAV/value) were found.")

    duplicated = df["fund_name"].duplicated()
    if duplicated.any():
        names = ", ".join(sorted(df.loc[duplicated, "fund_name"].unique()))
        notes.append(f"Merged duplicate rows for: {names} (units-weighted average NAV).")
    # Sum the file's own invested_value (not units * average_nav): the NAV
    # column is rounded to 2dp for display, so recomputing from it would
    # lose precision the file's own invested_value doesn't have.
    merged = df.groupby("fund_name", sort=False, as_index=False).agg(
        isin=("isin", "first"),
        units=("units", "sum"),
        invested_value=("invested_value", "sum"),
        current_nav=("current_nav", "first"),
        current_value=("current_value", "sum"),
        xirr_pct=("xirr_pct", "mean"),
    )
    merged["average_nav"] = merged["invested_value"] / merged["units"]
    return merged[MUTUAL_FUND_COLUMNS].reset_index(drop=True), notes


@st.cache_data(show_spinner=False)
def parse_mutual_funds_file(content: bytes, filename: str) -> tuple[pd.DataFrame, list[str]]:
    """Parse the Mutual Fund sheet from an Angel One Portfolio export, if present.

    Angel One's fuller "Portfolio" export (as opposed to the plain Holdings
    CSV) is a multi-sheet workbook with separate Equity and Mutual Fund
    tables; this looks for the latter in the same uploaded file. Returns an
    empty result (no error, no notes) when the file has no mutual fund table
    at all, since that's the normal case for an equity-only holdings file.
    """
    try:
        table = extract_table(read_sheets(content, filename), [MUTUAL_FUND_SCHEMA])
    except ParseError:
        return pd.DataFrame(columns=MUTUAL_FUND_COLUMNS), []
    return clean_mutual_funds(table.frame)


# ---------------------------------------------------------------------------
# Demo data
# ---------------------------------------------------------------------------

DEMO_UNIVERSE: dict[str, float] = {  # symbol -> indicative price (synthetic)
    "RELIANCE": 1_380.0,
    "TCS": 3_450.0,
    "HDFCBANK": 1_650.0,
    "INFY": 1_520.0,
    "ICICIBANK": 1_290.0,
    "ITC": 415.0,
    "SBIN": 810.0,
    "BHARTIARTL": 1_850.0,
    "LT": 3_600.0,
    "SUNPHARMA": 1_700.0,
    "M&M": 3_100.0,
    "ASIANPAINT": 2_400.0,
}


@st.cache_data(show_spinner=False)
def generate_demo_data(as_of: date, seed: int = 42) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build a self-consistent synthetic portfolio (holdings + trade history).

    Holdings are derived from the generated trades with average-cost
    accounting, so quantities reconcile exactly and XIRR/CAGR are meaningful.
    One fully exited position exercises realised (SELL) cash flows.
    """
    rng = np.random.default_rng(seed)
    today = pd.Timestamp(as_of)
    records: list[dict[str, object]] = []

    for symbol, ref_price in DEMO_UNIVERSE.items():
        ages = np.sort(rng.integers(45, 1_100, size=int(rng.integers(1, 5))))[::-1]
        bought = 0
        for age in ages:
            qty = int(rng.integers(5, 60))
            bought += qty
            records.append({
                "date": today - pd.Timedelta(days=int(age)), "symbol": symbol,
                "type": "BUY", "quantity": qty,
                "price": round(ref_price * rng.uniform(0.6, 1.15), 2),
            })
        if rng.random() < 0.35 and bought > 3:
            records.append({
                "date": today - pd.Timedelta(days=int(rng.integers(5, int(ages[-1])))),
                "symbol": symbol, "type": "SELL", "quantity": bought // 3,
                "price": round(ref_price * rng.uniform(0.85, 1.2), 2),
            })
    records += [
        {"date": today - pd.Timedelta(days=950), "symbol": "WIPRO", "type": "BUY",
         "quantity": 120, "price": 230.0},
        {"date": today - pd.Timedelta(days=300), "symbol": "WIPRO", "type": "SELL",
         "quantity": 120, "price": 275.0},
    ]

    trades = pd.DataFrame(records).sort_values("date", kind="stable").reset_index(drop=True)
    trades["ticker"] = [to_yahoo_ticker(s)[1] for s in trades["symbol"]]
    trades["cash_flow"] = (
        np.where(trades["type"] == "BUY", -1.0, 1.0) * trades["quantity"] * trades["price"]
    )

    holdings: list[dict[str, object]] = []
    for symbol, lots in trades.groupby("symbol", sort=False):
        qty, cost = 0.0, 0.0
        for side, q, p in zip(lots["type"], lots["quantity"], lots["price"], strict=True):
            if side == "BUY":
                qty, cost = qty + q, cost + q * p
            else:  # average-cost method: selling does not change the average
                cost -= cost / qty * q
                qty -= q
        if qty > 0:
            holdings.append({
                "symbol": symbol, "ticker": lots["ticker"].iloc[0], "quantity": qty,
                "average_price": cost / qty, "isin": "",
                "ltp": round(DEMO_UNIVERSE[symbol] * rng.uniform(0.82, 1.25), 2),
            })
    return pd.DataFrame(holdings)[HOLDINGS_COLUMNS], trades[TRADE_COLUMNS]


# Real fund names/ISINs with an indicative starting NAV, purely for demo mode.
DEMO_MUTUAL_FUNDS: tuple[dict[str, object], ...] = (
    {"fund_name": "Kotak Small Cap Fund", "isin": "INF174K01KT2",
     "average_nav": 210.0, "xirr_pct": 24.5},
    {"fund_name": "Parag Parikh Flexi Cap Fund", "isin": "INF879O01027",
     "average_nav": 55.0, "xirr_pct": 19.2},
    {"fund_name": "Nippon India Nifty Smallcap 250 Index Fund", "isin": "INF204KB15W0",
     "average_nav": 24.0, "xirr_pct": 15.8},
)


@st.cache_data(show_spinner=False)
def generate_demo_mutual_funds(seed: int = 43) -> pd.DataFrame:
    """Build a small synthetic mutual fund portfolio for demo mode.

    Independent of :func:`generate_demo_data`'s equity holdings/trades - a
    holdings snapshot alone (real or demo) never carries per-fund cash-flow
    history, so there is nothing to keep "consistent" here the way equity
    XIRR is consistent with its trade book; the XIRR figures are simply
    plausible fixed values, exactly as Angel One would report them.
    """
    rng = np.random.default_rng(seed)
    rows = []
    for fund in DEMO_MUTUAL_FUNDS:
        units = round(float(rng.uniform(50, 400)), 3)
        current_nav = round(fund["average_nav"] * rng.uniform(1.05, 1.45), 2)
        rows.append({
            "fund_name": fund["fund_name"], "isin": fund["isin"], "units": units,
            "average_nav": fund["average_nav"], "current_nav": current_nav,
            "invested_value": round(units * fund["average_nav"], 2),
            "current_value": round(units * current_nav, 2),
            "xirr_pct": fund["xirr_pct"],
        })
    return pd.DataFrame(rows)[MUTUAL_FUND_COLUMNS]


# ---------------------------------------------------------------------------
# Market data
# ---------------------------------------------------------------------------


def _download_last_close(tickers: Sequence[str]) -> dict[str, float]:
    """Bulk-download recent daily bars and return the latest close per ticker."""
    if not tickers:
        return {}
    data = yf.download(
        list(tickers), period="5d", interval="1d", group_by="column",
        auto_adjust=False, progress=False, threads=True,
    )
    if data is None or data.empty:
        return {}
    if isinstance(data.columns, pd.MultiIndex):
        if "Close" not in data.columns.get_level_values(0):
            return {}
        closes = data["Close"]
    elif "Close" in data.columns:
        closes = data[["Close"]].set_axis([tickers[0]], axis=1)
    else:
        return {}
    if isinstance(closes, pd.Series):
        closes = closes.to_frame(tickers[0])

    prices: dict[str, float] = {}
    for ticker in closes.columns:
        series = pd.to_numeric(closes[ticker], errors="coerce").dropna()
        series = series[series > 0]
        if not series.empty:
            prices[str(ticker)] = float(series.iloc[-1])
    return prices


def _alternate_exchange(ticker: str) -> str | None:
    """Swap ``.NS`` <-> ``.BO`` for a retry when one listing has no data."""
    if ticker.endswith(".NS"):
        return ticker[:-3] + ".BO"
    if ticker.endswith(".BO"):
        return ticker[:-3] + ".NS"
    return None


@st.cache_data(ttl=PRICE_CACHE_TTL_SECONDS, show_spinner=False)
def fetch_live_prices(tickers: tuple[str, ...]) -> tuple[dict[str, float], str, str | None]:
    """Fetch latest prices for ``tickers`` from Yahoo Finance (cached for 5 minutes).

    Tickers with no NSE data are retried on BSE (and vice versa). Failures are
    cached too, so an outage or rate limit is not hammered on every rerun; the
    sidebar refresh button clears the cache.

    Returns:
        ``(prices, fetched_at, error)`` where ``prices`` maps ticker -> price,
        ``fetched_at`` is an IST timestamp string and ``error`` is ``None`` on
        success or a message when nothing could be fetched.
    """
    fetched_at = datetime.now(IST).strftime("%d %b %Y, %H:%M:%S IST")
    if yf is None:
        return {}, fetched_at, "yfinance is not installed (pip install yfinance)."
    try:
        prices = _download_last_close(tickers)
        retry = {
            alt: original
            for original in tickers
            if original not in prices and (alt := _alternate_exchange(original))
        }
        if retry:
            for alt, price in _download_last_close(list(retry)).items():
                prices[retry[alt]] = price
    except Exception as exc:  # network errors, rate limits, schema changes
        return {}, fetched_at, f"Yahoo Finance request failed: {exc}"
    if not prices:
        return {}, fetched_at, "Yahoo Finance returned no prices (offline or rate-limited?)."
    return prices, fetched_at, None


def _search_isin(isin: str) -> tuple[str, str | None]:
    """Look up one ISIN via Yahoo's search endpoint; prefer an NSE listing."""
    try:
        quotes = yf.Search(isin, max_results=8).quotes
    except Exception:
        return isin, None
    symbols = (q.get("symbol", "") for q in quotes)
    candidates = [s for s in symbols if s.endswith((".NS", ".BO"))]
    nse = next((c for c in candidates if c.endswith(".NS")), None)
    return isin, nse or (candidates[0] if candidates else None)


@st.cache_data(ttl=86_400, show_spinner=False)
def resolve_isin_tickers(isins: tuple[str, ...]) -> dict[str, str]:
    """Resolve ISINs to Yahoo Finance tickers via Yahoo's own search.

    Angel One's full portfolio export (as opposed to its plain holdings CSV)
    lists a truncated internal scrip name instead of the NSE trading symbol
    (e.g. "Reliance Industr" rather than ``RELIANCE``), so guessing a ticker
    from that text is unreliable. ISIN is a globally unique identifier that
    Yahoo's own search indexes directly, so it resolves correctly where the
    name-based guess would not. Cached for a day since the mapping is static;
    lookups run in parallel and any that fail are silently skipped (the
    name-based ticker guess is used instead).
    """
    if yf is None or not isins:
        return {}
    resolved: dict[str, str] = {}
    with ThreadPoolExecutor(max_workers=8) as pool:
        for isin, ticker in pool.map(_search_isin, isins):
            if ticker:
                resolved[isin] = ticker
    return resolved


def enrich_holdings(holdings: pd.DataFrame, live_prices: Mapping[str, float]) -> pd.DataFrame:
    """Attach prices and compute per-position valuation and P&L.

    Price precedence: live Yahoo price -> LTP from the uploaded file -> average
    cost (flagged via ``price_source`` so the UI can warn about it).
    Percentages are NaN where the invested value is zero (e.g. bonus shares).
    """
    df = holdings.copy()
    live = df["ticker"].map(live_prices).astype(float)
    df["price_source"] = np.select(
        [live.notna(), df["ltp"].notna()], ["Live", "File"], default="Cost"
    )
    df["ltp"] = live.fillna(df["ltp"]).fillna(df["average_price"])

    df["invested_value"] = df["quantity"] * df["average_price"]
    df["current_value"] = df["quantity"] * df["ltp"]
    df["unrealized_pnl"] = df["current_value"] - df["invested_value"]
    invested = df["invested_value"].where(df["invested_value"] > 0)
    df["unrealized_pnl_pct"] = df["unrealized_pnl"] / invested * 100
    total = df["current_value"].sum()
    df["weight_pct"] = df["current_value"] / total * 100 if total > 0 else np.nan
    return df


AMFI_SCHEME_LIST_URL = "https://api.mfapi.in/mf"
AMFI_SCHEME_NAV_URL = "https://api.mfapi.in/mf/{code}"


@st.cache_data(ttl=86_400, show_spinner=False)
def fetch_amfi_scheme_isin_map() -> dict[str, int]:
    """ISIN -> AMFI scheme code, from mfapi.in's full scheme list.

    Cached a day: this master list changes rarely (new scheme launches), and
    NAVs themselves only update once per business day, so there is nothing
    to gain from fetching it more often. A request failure yields an empty
    map rather than raising - live NAV lookup then simply finds nothing and
    every fund falls back to the NAV already in the uploaded file.
    """
    try:
        response = requests.get(AMFI_SCHEME_LIST_URL, timeout=15)
        response.raise_for_status()
        schemes = response.json()
    except Exception:
        return {}
    mapping: dict[str, int] = {}
    for scheme in schemes:
        code = scheme.get("schemeCode")
        if not code:
            continue
        for isin in (scheme.get("isinGrowth"), scheme.get("isinDivReinvestment")):
            if isin:
                mapping.setdefault(isin, code)
    return mapping


def _fetch_one_amfi_nav(scheme_code: int) -> tuple[int, float | None]:
    """Latest NAV for one AMFI scheme code; ``None`` on any failure."""
    try:
        response = requests.get(AMFI_SCHEME_NAV_URL.format(code=scheme_code), timeout=15)
        response.raise_for_status()
        nav = float(response.json()["data"][0]["nav"])
    except Exception:
        return scheme_code, None
    return scheme_code, (nav if nav > 0 else None)


@st.cache_data(ttl=PRICE_CACHE_TTL_SECONDS, show_spinner=False)
def fetch_live_mf_navs(isins: tuple[str, ...]) -> dict[str, float]:
    """Latest NAV per ISIN from AMFI (via mfapi.in), threaded and cached 5 minutes.

    Mirrors :func:`fetch_live_prices` for equities: a per-ISIN failure (fund
    not found, request error) is dropped silently rather than raised, so one
    bad ISIN never blocks the rest of the batch.
    """
    if not isins:
        return {}
    isin_to_code = fetch_amfi_scheme_isin_map()
    codes = {isin: isin_to_code[isin] for isin in isins if isin in isin_to_code}
    if not codes:
        return {}
    navs_by_code: dict[int, float] = {}
    with ThreadPoolExecutor(max_workers=8) as pool:
        for code, nav in pool.map(_fetch_one_amfi_nav, set(codes.values())):
            if nav is not None:
                navs_by_code[code] = nav
    return {isin: navs_by_code[code] for isin, code in codes.items() if code in navs_by_code}


def enrich_mutual_funds(
    mutual_funds: pd.DataFrame, live_navs: Mapping[str, float] | None = None
) -> pd.DataFrame:
    """Compute per-fund P&L and portfolio weight (the mutual-fund analogue of
    :func:`enrich_holdings`).

    NAV precedence: live AMFI NAV (by ISIN, via :func:`fetch_live_mf_navs`) ->
    the NAV already in the uploaded file, flagged via ``nav_source`` the same
    way equities flag ``price_source``. ``live_navs`` defaults to empty (file
    NAV only) so callers that don't care about live data can omit it.
    """
    df = mutual_funds.copy()
    if df.empty:
        df["nav_source"] = pd.Series(dtype=object)
        return df
    live = df["isin"].map(live_navs or {}).astype(float)
    df["nav_source"] = np.where(live.notna(), "Live", "File")
    df["current_nav"] = live.fillna(df["current_nav"])
    df["current_value"] = df["units"] * df["current_nav"]

    df["unrealized_pnl"] = df["current_value"] - df["invested_value"]
    invested = df["invested_value"].where(df["invested_value"] > 0)
    df["unrealized_pnl_pct"] = df["unrealized_pnl"] / invested * 100
    total = df["current_value"].sum()
    df["weight_pct"] = df["current_value"] / total * 100 if total > 0 else np.nan
    return df


# ---------------------------------------------------------------------------
# Sector & market-cap enrichment
# ---------------------------------------------------------------------------

# Fallback sector map for common NSE large/mid-caps, used when Yahoo's
# ``.info`` payload has no sector (frequent for ETFs and thinly covered
# names) or the lookup fails outright. Keyed by display symbol, not ticker,
# so it applies regardless of NSE/BSE listing.
SECTOR_FALLBACK: dict[str, str] = {
    "RELIANCE": "Energy", "ONGC": "Energy", "GAIL": "Energy", "IOC": "Energy",
    "BPCL": "Energy", "COALINDIA": "Energy", "ADANIENT": "Energy",
    "TCS": "Technology", "INFY": "Technology", "WIPRO": "Technology",
    "HCLTECH": "Technology", "TECHM": "Technology", "LTIM": "Technology",
    "HDFCBANK": "Financial Services", "ICICIBANK": "Financial Services",
    "SBIN": "Financial Services", "KOTAKBANK": "Financial Services",
    "AXISBANK": "Financial Services", "INDUSINDBK": "Financial Services",
    "FEDERALBNK": "Financial Services", "IDFCFIRSTB": "Financial Services",
    "BAJFINANCE": "Financial Services", "BAJAJFINSV": "Financial Services",
    "HDFCLIFE": "Financial Services", "SBILIFE": "Financial Services",
    "ITC": "Consumer Defensive", "HINDUNILVR": "Consumer Defensive",
    "NESTLEIND": "Consumer Defensive", "BRITANNIA": "Consumer Defensive",
    "DABUR": "Consumer Defensive", "GODREJCP": "Consumer Defensive",
    "COLPAL": "Consumer Defensive", "TATACONSUM": "Consumer Defensive",
    "MARUTI": "Consumer Cyclical", "TATAMOTORS": "Consumer Cyclical",
    "M&M": "Consumer Cyclical", "EICHERMOT": "Consumer Cyclical",
    "HEROMOTOCO": "Consumer Cyclical", "BAJAJ-AUTO": "Consumer Cyclical",
    "TITAN": "Consumer Cyclical", "INDHOTEL": "Consumer Cyclical",
    "SUNPHARMA": "Healthcare", "CIPLA": "Healthcare", "DRREDDY": "Healthcare",
    "DIVISLAB": "Healthcare", "APOLLOHOSP": "Healthcare", "ZYDUSLIFE": "Healthcare",
    "NATCOPHARM": "Healthcare",
    "TATASTEEL": "Basic Materials", "JSWSTEEL": "Basic Materials",
    "HINDALCO": "Basic Materials", "VEDL": "Basic Materials", "SAIL": "Basic Materials",
    "NMDC": "Basic Materials", "ACC": "Basic Materials", "AMBUJACEM": "Basic Materials",
    "ULTRACEMCO": "Basic Materials", "GRASIM": "Basic Materials",
    "PIDILITIND": "Basic Materials", "UPL": "Basic Materials", "TATACHEM": "Basic Materials",
    "LT": "Industrials", "SIEMENS": "Industrials", "ABB": "Industrials",
    "BEL": "Industrials", "HAL": "Industrials", "BHEL": "Industrials",
    "CONCOR": "Industrials", "NCC": "Industrials", "ADANIPORTS": "Industrials",
    "NTPC": "Utilities", "POWERGRID": "Utilities",
    "BHARTIARTL": "Communication Services",
    "DLF": "Real Estate",
}


@st.cache_data(ttl=86_400, show_spinner=False)
def fetch_sector_and_marketcap(
    tickers: tuple[str, ...]
) -> dict[str, tuple[str | None, float | None]]:
    """Fetch ``(sector, marketCap in INR)`` per ticker from Yahoo Finance.

    NSE/BSE market caps come back already denominated in INR, so no currency
    conversion is needed. Cached for a day and run in parallel since
    ``yf.Ticker(...).get_info()`` is one request per ticker; any failure
    (network, delisted, ETF with no sector) yields ``(None, None)`` for that
    ticker rather than raising, so one bad symbol never breaks the batch.
    """
    if yf is None or not tickers:
        return {}

    def _one(ticker: str) -> tuple[str, tuple[str | None, float | None]]:
        try:
            info = yf.Ticker(ticker).get_info()
        except Exception:
            return ticker, (None, None)
        return ticker, (info.get("sector"), info.get("marketCap"))

    result: dict[str, tuple[str | None, float | None]] = {}
    with ThreadPoolExecutor(max_workers=8) as pool:
        for ticker, value in pool.map(_one, tickers):
            result[ticker] = value
    return result


def classify_market_cap(market_cap_inr: float | None) -> str:
    """Bucket an INR market cap into Large / Mid / Small Cap (₹ crore terms)."""
    if market_cap_inr is None or not np.isfinite(market_cap_inr) or market_cap_inr <= 0:
        return UNCATEGORIZED
    cap_cr = market_cap_inr / 1e7
    if cap_cr > LARGE_CAP_MIN_CR:
        return "Large Cap"
    if cap_cr >= MID_CAP_MIN_CR:
        return "Mid Cap"
    return "Small Cap"


def enrich_sector_and_cap(holdings: pd.DataFrame) -> pd.DataFrame:
    """Attach ``sector``, ``market_cap`` and ``market_cap_bucket`` columns.

    Sector precedence: Yahoo Finance -> :data:`SECTOR_FALLBACK` (by display
    symbol) -> ``"Uncategorized"``. Market cap is Yahoo-only (no local
    fallback exists); a missing value also buckets to ``"Uncategorized"``.
    """
    df = holdings.copy()
    tickers = tuple(sorted(df["ticker"].unique()))
    with st.spinner("Fetching sector and market-cap data..."):
        info = fetch_sector_and_marketcap(tickers)

    def _sector(row: pd.Series) -> str:
        sector, _ = info.get(row["ticker"], (None, None))
        return sector or SECTOR_FALLBACK.get(row["symbol"], UNCATEGORIZED)

    df["sector"] = df.apply(_sector, axis=1)
    df["market_cap"] = df["ticker"].map(lambda t: info.get(t, (None, None))[1])
    df["market_cap_bucket"] = df["market_cap"].map(classify_market_cap)
    return df


# ---------------------------------------------------------------------------
# Dividend & corporate-actions data
# ---------------------------------------------------------------------------

DIVIDEND_LOOKBACK_YEARS = 1  # trailing-twelve-months window for yield/income projections


def _fetch_one_dividend_history(ticker: str) -> tuple[str, pd.Series]:
    """Full per-share dividend history for one ticker; an empty Series on any failure.

    Yahoo returns the same empty (not missing/error) result for a stock that
    has simply never paid a dividend as for a lookup failure - there is no
    way to tell the two apart from this call alone, so both are treated the
    same way: no dividend rows, no exception.
    """
    try:
        dividends = yf.Ticker(ticker).dividends
    except Exception:
        return ticker, pd.Series(dtype=float)
    if dividends is None or dividends.empty:
        return ticker, pd.Series(dtype=float)
    dividends = dividends.copy()
    # Yahoo's index is tz-aware (Asia/Kolkata); the rest of the app is naive.
    dividends.index = pd.DatetimeIndex(dividends.index).tz_localize(None)
    return ticker, dividends


@st.cache_data(ttl=86_400, show_spinner=False)
def fetch_dividend_history(tickers: tuple[str, ...]) -> pd.DataFrame:
    """Historical per-share dividend payouts for every ticker, from Yahoo Finance.

    Cached a day (a dividend calendar changes rarely) and threaded like
    :func:`fetch_sector_and_marketcap`; a ticker with no history (an ETF, a
    non-payer) or a fetch failure simply contributes no rows, never raises.

    Returns:
        One row per historical payout across all tickers: ``ticker``,
        ``ex_dividend_date``, ``dividend_per_share``. Empty (with these
        columns) if nothing was found for any ticker.
    """
    columns = ["ticker", "ex_dividend_date", "dividend_per_share"]
    if yf is None or not tickers:
        return pd.DataFrame(columns=columns)
    frames: list[pd.DataFrame] = []
    with ThreadPoolExecutor(max_workers=8) as pool:
        for ticker, dividends in pool.map(_fetch_one_dividend_history, tickers):
            if dividends.empty:
                continue
            frames.append(pd.DataFrame({
                "ticker": ticker, "ex_dividend_date": dividends.index,
                "dividend_per_share": dividends.to_numpy(dtype=float),
            }))
    if not frames:
        return pd.DataFrame(columns=columns)
    return pd.concat(frames, ignore_index=True)


def enrich_dividend_metrics(
    holdings: pd.DataFrame, dividend_history: pd.DataFrame, as_of: pd.Timestamp
) -> pd.DataFrame:
    """Attach trailing-twelve-month dividend/share, forward yield and estimated annual
    income to holdings - needs only current holdings, not trade history (unlike realised
    dividends, which need a purchase date - see :func:`compute_realized_dividend_events`).

    A stock with nothing paid in the trailing year gets ``0``, not ``NaN``: "no dividend"
    is a real, valid answer here, not missing data.
    """
    df = holdings.copy()
    if dividend_history.empty:
        ttm = pd.Series(0.0, index=df.index)
    else:
        window_start = as_of - pd.DateOffset(years=DIVIDEND_LOOKBACK_YEARS)
        recent = dividend_history[
            dividend_history["ex_dividend_date"].gt(window_start)
            & dividend_history["ex_dividend_date"].le(as_of)
        ]
        ttm_by_ticker = recent.groupby("ticker")["dividend_per_share"].sum()
        ttm = df["ticker"].map(ttm_by_ticker).fillna(0.0)
    df["ttm_dividend_per_share"] = ttm
    df["dividend_yield_pct"] = np.where(df["ltp"] > 0, ttm / df["ltp"] * 100, np.nan)
    df["estimated_annual_dividend"] = df["quantity"] * ttm
    return df


# ---------------------------------------------------------------------------
# Financial metrics
# ---------------------------------------------------------------------------


def compute_cagr(
    current_value: float, invested_value: float, start: pd.Timestamp | None, end: pd.Timestamp
) -> float | None:
    """CAGR = (current / invested) ** (1 / years) - 1, as a fraction.

    Returns ``None`` when undefined: no start date, non-positive holding
    period or cost basis, or a numerically meaningless result (overflow).
    """
    if start is None or pd.isna(start) or not np.isfinite(invested_value):
        return None
    years = (end - start).days / DAYS_PER_YEAR
    if years <= 0 or invested_value <= 0 or not np.isfinite(current_value) or current_value < 0:
        return None
    if current_value == 0:
        return -1.0
    try:
        cagr = math.exp(math.log(current_value / invested_value) / years) - 1.0
    except (OverflowError, ValueError):
        return None
    return cagr if math.isfinite(cagr) else None


def xnpv(rate: float, amounts: np.ndarray, years: np.ndarray) -> float:
    """Net present value of dated cash flows; ``years`` are offsets from the first flow."""
    if rate <= -1.0:
        return float("nan")
    with np.errstate(over="ignore", invalid="ignore"):
        return float(np.sum(amounts * np.exp(-years * np.log1p(rate))))


def _xnpv_derivative(rate: float, amounts: np.ndarray, years: np.ndarray) -> float:
    """First derivative of :func:`xnpv` with respect to ``rate``."""
    if rate <= -1.0:
        return float("nan")
    with np.errstate(over="ignore", invalid="ignore"):
        return float(np.sum(-years * amounts * np.exp(-(years + 1.0) * np.log1p(rate))))


# Bracketing grid for the root search: dense near typical returns, sparse in the tails.
_XIRR_GRID = (
    -0.9999, -0.999, -0.99, -0.95, -0.9, -0.8, -0.6, -0.4, -0.2, -0.1, 0.0,
    0.05, 0.1, 0.2, 0.35, 0.5, 0.75, 1.0, 1.5, 2.0, 3.0, 5.0, 10.0, 25.0,
    100.0, 1_000.0, 10_000.0,
)


def solve_xirr(dates: Sequence[object], amounts: Sequence[float]) -> float | None:
    """Solve ``sum(C_i / (1 + r) ** ((d_i - d_1) / 365)) = 0`` for ``r``.

    Same-day flows are netted first. The solver scans a rate grid for sign
    changes and refines each bracket with Brent's method (guaranteed
    convergence); if no bracket exists it falls back to Newton-Raphson from
    several seeds. When multiple roots exist the one closest to 10 % wins,
    matching spreadsheet XIRR behaviour for conventional flows.

    Returns:
        The annualised rate as a fraction (0.12 == 12 %), or ``None`` if the
        flows lack a sign change or no root could be found.
    """
    flows = pd.DataFrame({"date": pd.to_datetime(list(dates)), "amount": list(amounts)})
    flows["amount"] = pd.to_numeric(flows["amount"], errors="coerce")
    flows = flows.dropna().groupby("date", as_index=False)["amount"].sum()
    scale = flows["amount"].abs().max() if not flows.empty else 0.0
    flows = flows[flows["amount"].abs() > 1e-9 * max(scale, 1.0)].sort_values("date")
    cash = flows["amount"].to_numpy(dtype=float)
    if len(cash) < 2 or not (cash > 0).any() or not (cash < 0).any():
        return None
    years = (flows["date"] - flows["date"].iloc[0]).dt.days.to_numpy(dtype=float) / DAYS_PER_YEAR
    if years[-1] <= 0:
        return None

    def npv(rate: float) -> float:
        return xnpv(rate, cash, years)

    roots: list[float] = []
    values = [npv(r) for r in _XIRR_GRID]
    brackets = zip(_XIRR_GRID[:-1], values[:-1], _XIRR_GRID[1:], values[1:], strict=True)
    for lo, f_lo, hi, f_hi in brackets:
        if not (np.isfinite(f_lo) and np.isfinite(f_hi)):
            continue
        if f_lo == 0.0:
            roots.append(lo)
        elif f_lo * f_hi < 0:
            try:
                roots.append(optimize.brentq(npv, lo, hi, xtol=1e-12, maxiter=500))
            except (ValueError, RuntimeError):
                continue
    if roots:
        return float(min(roots, key=lambda r: abs(r - 0.1)))

    tolerance = 1e-6 * float(np.abs(cash).sum())
    for guess in (0.1, 0.0, -0.5, 1.0, 10.0):
        try:
            rate = optimize.newton(
                npv, guess, fprime=lambda r: _xnpv_derivative(r, cash, years),
                tol=1e-10, maxiter=200,
            )
        except (RuntimeError, OverflowError, ZeroDivisionError, FloatingPointError):
            continue
        if np.isfinite(rate) and rate > -1.0 and abs(npv(rate)) <= tolerance:
            return float(rate)
    return None


@dataclass(frozen=True)
class PortfolioMetrics:
    """Portfolio-level headline numbers (percentages are in percent, not fractions)."""

    total_invested: float
    total_current: float
    net_pnl: float
    net_pnl_pct: float | None
    xirr_pct: float | None
    xirr_method: str  # "xirr" | "fallback" | "unavailable"
    cagr_pct: float | None
    cagr_start: pd.Timestamp | None


def first_purchase_date(trades: pd.DataFrame | None, symbols: Sequence[str]) -> pd.Timestamp | None:
    """Earliest BUY among currently held symbols (falls back to any BUY)."""
    if trades is None or trades.empty:
        return None
    buys = trades[trades["type"] == "BUY"]
    held = buys[buys["symbol"].isin(symbols)]
    source = held if not held.empty else buys
    return source["date"].min() if not source.empty else None


def compute_portfolio_metrics(
    holdings: pd.DataFrame, trades: pd.DataFrame | None, as_of: pd.Timestamp
) -> PortfolioMetrics:
    """Aggregate P&L, XIRR and CAGR for the enriched holdings.

    XIRR uses every trade cash flow plus a terminal inflow equal to today's
    market value. If the solver cannot converge it falls back to the simple
    portfolio return (``xirr_method == "fallback"``).
    """
    invested = float(holdings["invested_value"].sum())
    current = float(holdings["current_value"].sum())
    pnl = current - invested
    pnl_pct = pnl / invested * 100 if invested > 0 else None

    xirr_pct, method = None, "unavailable"
    if trades is not None and not trades.empty:
        try:
            rate = solve_xirr(
                [*trades["date"], as_of], [*trades["cash_flow"], current]
            )
        except (ValueError, TypeError, ArithmeticError):
            rate = None
        xirr_pct, method = (rate * 100, "xirr") if rate is not None else (pnl_pct, "fallback")

    start = first_purchase_date(trades, holdings["symbol"].tolist())
    cagr = compute_cagr(current, invested, start, as_of)
    return PortfolioMetrics(
        total_invested=invested,
        total_current=current,
        net_pnl=pnl,
        net_pnl_pct=pnl_pct,
        xirr_pct=xirr_pct,
        xirr_method=method,
        cagr_pct=cagr * 100 if cagr is not None else None,
        cagr_start=start,
    )


BENCHMARK_TICKER = "^NSEI"  # NIFTY 50


@st.cache_data(ttl=PRICE_CACHE_TTL_SECONDS, show_spinner=False)
def fetch_benchmark_series(start: pd.Timestamp, end: pd.Timestamp) -> pd.Series:
    """Daily close prices for the benchmark index from ``start`` to ``end`` (inclusive).

    Cached like :func:`fetch_live_prices`; an empty result on any failure
    (network, no data for the range) rather than raising, so a benchmark
    outage degrades to "no alpha available", not a crashed page.
    """
    if yf is None:
        return pd.Series(dtype=float)
    try:
        data = yf.download(
            BENCHMARK_TICKER, start=start, end=end + pd.Timedelta(days=1),
            interval="1d", auto_adjust=False, progress=False,
        )
    except Exception:
        return pd.Series(dtype=float)
    if data is None or data.empty or "Close" not in data.columns.get_level_values(0):
        return pd.Series(dtype=float)
    closes = data["Close"]
    if isinstance(closes, pd.DataFrame):
        closes = closes.iloc[:, 0]
    return pd.to_numeric(closes, errors="coerce").dropna()


def _price_on_or_before(series: pd.Series, when: pd.Timestamp) -> float | None:
    """Latest available price at or before ``when`` (handles weekends/holidays)."""
    eligible = series[series.index <= when]
    return float(eligible.iloc[-1]) if not eligible.empty else None


@dataclass(frozen=True)
class BenchmarkComparison:
    """Portfolio XIRR vs. an identical-cash-flow investment in the benchmark index."""

    benchmark_name: str
    benchmark_xirr_pct: float
    alpha_pct: float  # portfolio XIRR - benchmark XIRR


def compute_benchmark_alpha(
    trades: pd.DataFrame, as_of: pd.Timestamp, portfolio_xirr_pct: float
) -> BenchmarkComparison | None:
    """Alpha = portfolio XIRR minus a benchmark XIRR built from the *same* cash flows.

    For a fair comparison, every trade's rupee amount is hypothetically
    invested in (a BUY) or withdrawn from (a SELL) the benchmark index on
    that same date, at that day's close (or the last trading day at/before
    it, for a weekend/holiday). The resulting benchmark units, valued at the
    latest available close, become the terminal value for a ``solve_xirr``
    call on the same dated cash flows - giving a benchmark XIRR directly
    comparable to the portfolio's, rather than mixing a CAGR-basis benchmark
    return with an XIRR-basis portfolio return.

    Returns ``None`` when there isn't enough data for a fair comparison
    (benchmark data unavailable, no valid trades, or the equivalent
    benchmark investment nets to zero or negative units).
    """
    series = fetch_benchmark_series(trades["date"].min(), as_of)
    if series.empty:
        return None

    units = 0.0
    for row in trades.itertuples():
        price = _price_on_or_before(series, row.date)
        if price is None or price <= 0:
            continue
        amount = row.quantity * row.price
        units += amount / price if row.type == "BUY" else -amount / price
    if units <= 0:
        return None

    benchmark_value = units * float(series.iloc[-1])
    rate = solve_xirr([*trades["date"], as_of], [*trades["cash_flow"], benchmark_value])
    if rate is None:
        return None
    benchmark_xirr_pct = rate * 100
    return BenchmarkComparison(
        benchmark_name="NIFTY 50",
        benchmark_xirr_pct=benchmark_xirr_pct,
        alpha_pct=portfolio_xirr_pct - benchmark_xirr_pct,
    )


@dataclass(frozen=True)
class MutualFundMetrics:
    """Portfolio-level mutual fund numbers (percentages are in percent, not fractions)."""

    total_invested: float
    total_current: float
    net_pnl: float
    net_pnl_pct: float | None
    weighted_xirr_pct: float | None


def compute_mutual_fund_metrics(mutual_funds: pd.DataFrame) -> MutualFundMetrics:
    """Aggregate invested/current value, P&L and an invested-value-weighted average XIRR.

    Angel One reports XIRR per fund, not for the portfolio as a whole, and a
    true portfolio-level figure needs each fund's full cash-flow history -
    not available from a holdings snapshot. The weighted average here (funds
    with no reported XIRR excluded) is a reasonable approximation, not an
    exact one; see the equivalent equity caveat for ``xirr_method == "fallback"``.
    """
    invested = float(mutual_funds["invested_value"].sum())
    current = float(mutual_funds["current_value"].sum())
    pnl = current - invested
    pnl_pct = pnl / invested * 100 if invested > 0 else None

    weighted_xirr = None
    weights = mutual_funds["invested_value"].where(mutual_funds["xirr_pct"].notna())
    if weights.notna().any() and weights.sum() > 0:
        weighted_xirr = float((mutual_funds["xirr_pct"] * weights).sum() / weights.sum())
    return MutualFundMetrics(
        total_invested=invested, total_current=current, net_pnl=pnl,
        net_pnl_pct=pnl_pct, weighted_xirr_pct=weighted_xirr,
    )


def top_movers(holdings: pd.DataFrame, n: int = TOP_N_MOVERS) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Top ``n`` gainers and bottom ``n`` laggards by ``unrealized_pnl_pct``.

    With fewer than ``2 * n`` ranked positions the list is split so that no
    stock appears on both sides. Laggards are ordered worst first.
    """
    ranked = holdings.dropna(subset=["unrealized_pnl_pct"]).sort_values(
        "unrealized_pnl_pct", ascending=False
    )
    n_top = min(n, math.ceil(len(ranked) / 2))
    n_bottom = min(n, len(ranked) - n_top)
    return ranked.head(n_top), ranked.tail(n_bottom).iloc[::-1]


def reconcile_positions(holdings: pd.DataFrame, trades: pd.DataFrame) -> pd.DataFrame:
    """Compare held quantities with the net quantities implied by the trade log.

    A mismatch means the log does not fully explain the holdings (older
    trades, bonus/split shares, transfers), which distorts XIRR and CAGR.
    Keyed by symbol because the same stock may be traded on NSE and BSE.
    """
    signed = trades["quantity"].where(trades["type"] == "BUY", -trades["quantity"])
    traded = signed.groupby(trades["symbol"]).sum().rename("traded")
    held = holdings.groupby("symbol")["quantity"].sum().rename("held")
    combined = pd.concat([held, traded], axis=1).fillna(0.0)
    mismatch = combined[~np.isclose(combined["held"], combined["traded"], atol=1e-6)]
    return mismatch.rename_axis("symbol").reset_index()


# ---------------------------------------------------------------------------
# Portfolio breakdowns (sector, market cap, tax)
# ---------------------------------------------------------------------------


def compute_hhi(weight_pct: pd.Series) -> float:
    """Herfindahl-Hirschman Index (0-10,000 scale) from percentage weights.

    Sum of squared percentage shares: one 100% holding scores 10,000; N
    equally-weighted holdings score ``10,000 / N``.
    """
    return float((weight_pct.fillna(0.0) ** 2).sum())


def hhi_risk_level(hhi: float) -> str:
    """Classify concentration risk using the standard DOJ/FTC HHI bands."""
    if hhi < HHI_MODERATE_THRESHOLD:
        return "Low Concentration"
    if hhi < HHI_HIGH_THRESHOLD:
        return "Moderate Concentration"
    return "High Concentration"


def group_breakdown(holdings: pd.DataFrame, by: str) -> pd.DataFrame:
    """Aggregate invested/current value, P&L and weight by ``by`` (a grouping column).

    Used for both the sector and market-cap breakdowns - same shape, so the
    same table/chart-building code renders either one.
    """
    total_current = holdings["current_value"].sum()
    grouped = holdings.groupby(by, sort=False, as_index=False).agg(
        invested_value=("invested_value", "sum"),
        current_value=("current_value", "sum"),
        holdings_count=("symbol", "count"),
    )
    grouped["unrealized_pnl"] = grouped["current_value"] - grouped["invested_value"]
    invested = grouped["invested_value"].where(grouped["invested_value"] > 0)
    grouped["unrealized_pnl_pct"] = grouped["unrealized_pnl"] / invested * 100
    grouped["weight_pct"] = (
        grouped["current_value"] / total_current * 100 if total_current > 0 else np.nan
    )
    return grouped.sort_values("current_value", ascending=False).reset_index(drop=True)


def compute_sector_drilldown(holdings: pd.DataFrame, sector: str) -> pd.DataFrame:
    """Per-stock detail for one sector, with two distinct weight columns.

    A stock's ``pct_of_sector`` (share of that sector's own value) and
    ``pct_of_portfolio`` (share of the whole portfolio) tell different
    stories - a stock can dominate a small sector while being a minor
    portfolio position, or vice versa - so both are kept, not just one.
    """
    total_portfolio_value = float(holdings["current_value"].sum())
    subset = holdings[holdings["sector"] == sector].copy()
    total_sector_value = float(subset["current_value"].sum())
    subset["pct_of_portfolio"] = (
        subset["current_value"] / total_portfolio_value * 100
        if total_portfolio_value > 0 else np.nan
    )
    subset["pct_of_sector"] = (
        subset["current_value"] / total_sector_value * 100 if total_sector_value > 0 else np.nan
    )
    return subset.sort_values("current_value", ascending=False).reset_index(drop=True)


def split_gainers_losers(holdings: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Split holdings into profitable (P&L > 0) and loss-making (P&L < 0) positions."""
    profitable = holdings[holdings["unrealized_pnl"] > 0].sort_values(
        "unrealized_pnl", ascending=False
    )
    lossmaking = holdings[holdings["unrealized_pnl"] < 0].sort_values("unrealized_pnl")
    return profitable, lossmaking


def compute_tax_lots(trades: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    """Match sells against buys FIFO per symbol and return the still-held lots.

    Each row is one remaining purchase lot: ``symbol``, ``purchase_date``,
    ``quantity``, ``price`` (cost) and, once dated, ``holding_days``/``term``
    as of ``as_of``. A symbol whose sells exceed its recorded buys
    (incomplete trade history, or a short sale) simply runs out of lots to
    consume - it never goes negative or raises.
    """
    rows: list[dict[str, object]] = []
    for symbol, group in trades.groupby("symbol", sort=False):
        lots: deque[list] = deque()
        for row in group.sort_values("date", kind="stable").itertuples():
            if row.type == "BUY":
                lots.append([row.date, float(row.quantity), float(row.price)])
            else:
                remaining = float(row.quantity)
                while remaining > 1e-9 and lots:
                    lot = lots[0]
                    take = min(lot[1], remaining)
                    lot[1] -= take
                    remaining -= take
                    if lot[1] <= 1e-9:
                        lots.popleft()
        rows += [
            {"symbol": symbol, "purchase_date": lot_date, "quantity": qty, "price": price}
            for lot_date, qty, price in lots
            if qty > 1e-9
        ]
    lots_df = pd.DataFrame(rows, columns=["symbol", "purchase_date", "quantity", "price"])
    if lots_df.empty:
        return lots_df
    holding_days = (as_of - lots_df["purchase_date"]).dt.days
    lots_df["holding_days"] = holding_days
    lots_df["term"] = np.where(holding_days > LTCG_HOLDING_DAYS, "Long-Term", "Short-Term")
    return lots_df


def value_tax_lots(lots: pd.DataFrame, holdings: pd.DataFrame) -> pd.DataFrame:
    """Attach current price and unrealised gain to each tax lot."""
    priced = lots.copy()
    if priced.empty:
        for col in ("ltp", "invested_value", "current_value", "unrealized_gain"):
            priced[col] = pd.Series(dtype=float)
        return priced
    ltp_by_symbol = holdings.set_index("symbol")["ltp"]
    priced["ltp"] = priced["symbol"].map(ltp_by_symbol)
    priced["invested_value"] = priced["quantity"] * priced["price"]
    priced["current_value"] = priced["quantity"] * priced["ltp"]
    priced["unrealized_gain"] = priced["current_value"] - priced["invested_value"]
    return priced


@dataclass(frozen=True)
class CapitalGainsEstimate:
    """Estimated tax exposure on unrealised gains (India, listed equity)."""

    stcg_gain: float
    stcg_tax: float
    ltcg_gain: float
    ltcg_taxable: float
    ltcg_tax: float


def estimate_capital_gains(priced_lots: pd.DataFrame) -> CapitalGainsEstimate:
    """Estimate STCG (flat rate) and LTCG (above the exemption) tax on unrealised gains.

    Only the positive-gain portion of each term is taxable; a net loss in
    either bucket contributes zero tax there (it is not netted against the
    other bucket, since that requires actually realizing both, not just
    holding them).
    """
    if priced_lots.empty:
        return CapitalGainsEstimate(0.0, 0.0, 0.0, 0.0, 0.0)
    by_term = priced_lots.groupby("term")["unrealized_gain"].sum()
    stcg_gain = float(by_term.get("Short-Term", 0.0))
    ltcg_gain = float(by_term.get("Long-Term", 0.0))
    stcg_taxable = max(0.0, stcg_gain)
    ltcg_taxable = max(0.0, ltcg_gain - LTCG_EXEMPTION)
    return CapitalGainsEstimate(
        stcg_gain=stcg_gain,
        stcg_tax=stcg_taxable * STCG_TAX_RATE,
        ltcg_gain=ltcg_gain,
        ltcg_taxable=ltcg_taxable,
        ltcg_tax=ltcg_taxable * LTCG_TAX_RATE,
    )


@dataclass(frozen=True)
class TaxLossHarvestingSuggestion:
    """Loss-making lots and the tax saved by realising them to offset gains."""

    loss_lots: pd.DataFrame
    total_harvestable_loss: float
    current_tax: float
    harvested_tax: float
    potential_savings: float


def suggest_tax_loss_harvesting(
    priced_lots: pd.DataFrame, baseline: CapitalGainsEstimate
) -> TaxLossHarvestingSuggestion:
    """Loss-making lots, and the tax saved by realising them to offset gains.

    Models "sell every lot today", the same framing :func:`estimate_capital_gains`
    already uses: within each term, gains and losses there already net against
    each other. The one thing that function deliberately doesn't do - because
    it requires actually *selling* to be tax-law-valid, not just holding a
    paper loss - is let a net Short-Term loss spill over to offset a
    Long-Term gain (Indian tax law allows this one-directional carry; a
    Long-Term loss can only ever offset LTCG, never STCG, so no such branch
    exists here). This computes that "what if you actually harvested"
    scenario as an explicit comparison against the baseline, rather than
    silently changing it - the base tax estimate elsewhere in the app is
    unaffected by this function.
    """
    if priced_lots.empty:
        return TaxLossHarvestingSuggestion(priced_lots, 0.0, 0.0, 0.0, 0.0)

    loss_lots = priced_lots[priced_lots["unrealized_gain"] < 0].sort_values("unrealized_gain")
    total_harvestable_loss = float(-loss_lots["unrealized_gain"].sum())

    short_term, long_term = baseline.stcg_gain, baseline.ltcg_gain
    if short_term < 0:
        long_term += short_term  # net ST loss offsets LT gain - never the reverse
        short_term = 0.0
    short_taxable = max(0.0, short_term)
    long_taxable = max(0.0, long_term - LTCG_EXEMPTION)
    harvested_tax = short_taxable * STCG_TAX_RATE + long_taxable * LTCG_TAX_RATE
    current_tax = baseline.stcg_tax + baseline.ltcg_tax
    return TaxLossHarvestingSuggestion(
        loss_lots=loss_lots,
        total_harvestable_loss=total_harvestable_loss,
        current_tax=current_tax,
        harvested_tax=harvested_tax,
        potential_savings=current_tax - harvested_tax,
    )


def compute_realized_dividend_events(
    trades: pd.DataFrame, dividend_history: pd.DataFrame, as_of: pd.Timestamp
) -> pd.DataFrame:
    """Dividends actually received on currently-held lots since each lot's own purchase date.

    Reuses the same FIFO tax-lot pipeline as the Tax tab (:func:`compute_tax_lots`), so a
    lot bought after a dividend's ex-date correctly never counts that dividend - a blended
    average purchase date (as :func:`clean_holdings` computes for display) would get this
    wrong for a stock bought in more than one lot.

    Returns:
        One row per (held lot, matching dividend payment): ``symbol``,
        ``ex_dividend_date``, ``amount_received``. Empty (with these columns)
        without trade history, without dividend history, or if nothing matches.
    """
    columns = ["symbol", "ex_dividend_date", "amount_received"]
    if trades.empty or dividend_history.empty:
        return pd.DataFrame(columns=columns)
    lots = compute_tax_lots(trades, as_of)
    if lots.empty:
        return pd.DataFrame(columns=columns)
    ticker_by_symbol = trades.drop_duplicates("symbol").set_index("symbol")["ticker"]
    lots = lots.assign(ticker=lots["symbol"].map(ticker_by_symbol))

    events: list[pd.DataFrame] = []
    for lot in lots.itertuples():
        paid = dividend_history[
            dividend_history["ticker"].eq(lot.ticker)
            & dividend_history["ex_dividend_date"].ge(lot.purchase_date)
            & dividend_history["ex_dividend_date"].le(as_of)
        ]
        if paid.empty:
            continue
        events.append(pd.DataFrame({
            "symbol": lot.symbol, "ex_dividend_date": paid["ex_dividend_date"].to_numpy(),
            "amount_received": paid["dividend_per_share"].to_numpy() * lot.quantity,
        }))
    if not events:
        return pd.DataFrame(columns=columns)
    return pd.concat(events, ignore_index=True)


@dataclass(frozen=True)
class DividendMetrics:
    """Portfolio-level dividend income figures."""

    total_realized: float | None  # None when there is no trade history to compute it from
    estimated_annual_income: float
    portfolio_yield_pct: float | None  # current-value-weighted; None if current value is 0
    top_payer_symbol: str | None
    top_payer_annual: float


def compute_dividend_metrics(
    holdings_with_dividends: pd.DataFrame, realized_events: pd.DataFrame | None
) -> DividendMetrics:
    """Aggregate per-stock dividend figures into portfolio-level totals.

    ``realized_events`` is ``None`` (as opposed to merely empty) specifically to mean "no
    trade history was uploaded, so realised dividends cannot be computed at all" -
    distinct from "trade history exists but no dividends were realised" (a genuine ₹0),
    which is an empty-but-not-``None`` DataFrame. Callers must preserve that distinction.
    """
    total_current_value = float(holdings_with_dividends["current_value"].sum())
    portfolio_yield = None
    if total_current_value > 0:
        weighted = holdings_with_dividends["dividend_yield_pct"].fillna(0.0)
        portfolio_yield = float(
            (weighted * holdings_with_dividends["current_value"]).sum() / total_current_value
        )

    total_realized = None
    if realized_events is not None:
        total_realized = float(realized_events["amount_received"].sum())

    top_symbol, top_annual = None, 0.0
    payers = holdings_with_dividends[holdings_with_dividends["estimated_annual_dividend"] > 0]
    if not payers.empty:
        top_row = payers.loc[payers["estimated_annual_dividend"].idxmax()]
        top_symbol, top_annual = str(top_row["symbol"]), float(top_row["estimated_annual_dividend"])

    return DividendMetrics(
        total_realized=total_realized,
        estimated_annual_income=float(holdings_with_dividends["estimated_annual_dividend"].sum()),
        portfolio_yield_pct=portfolio_yield,
        top_payer_symbol=top_symbol,
        top_payer_annual=top_annual,
    )


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------


def _is_finite(value: float | None) -> bool:
    return value is not None and bool(np.isfinite(value))


def fmt_inr(value: float | None, signed: bool = False) -> str:
    """Format rupees as ``₹1,234.50`` (``+₹`` / ``-₹`` when ``signed``)."""
    if not _is_finite(value):
        return "—"
    body = INR_FORMAT.format(abs(value))
    if value < 0:
        return f"-{body}"
    return f"+{body}" if signed else body


def fmt_inr_compact(value: float | None, signed: bool = False) -> str:
    """Card-sized rupees in Indian units: ``₹85,432``, ``₹2.12 L``, ``₹1.35 Cr``."""
    if not _is_finite(value):
        return "—"
    magnitude = abs(value)
    if magnitude >= 1e7:
        body = f"₹{magnitude / 1e7:,.2f} Cr"
    elif magnitude >= 1e5:
        body = f"₹{magnitude / 1e5:.2f} L"
    else:
        body = f"₹{magnitude:,.0f}"
    sign = "-" if value < 0 else ("+" if signed else "")
    return f"{sign}{body}"


def fmt_pct(value: float | None) -> str:
    """Format a percentage with an explicit sign, e.g. ``+12.34%``."""
    return PCT_FORMAT.format(value) if _is_finite(value) else "N/A"


def fmt_pct_plain(value: float | None) -> str:
    """Format a non-directional percentage with no sign, e.g. a yield: ``3.50%``."""
    return f"{value:.2f}%" if _is_finite(value) else "N/A"


def fmt_quantity(value: float) -> str:
    """Whole shares without decimals, fractional units with four."""
    if not _is_finite(value):
        return "—"
    return f"{value:,.0f}" if float(value).is_integer() else f"{value:,.4f}"


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------


def _is_dark_theme() -> bool:
    """Best-effort detection of the active Streamlit theme."""
    try:
        return st.context.theme.type == "dark"
    except Exception:  # older Streamlit versions have no st.context.theme
        return False


def _donut_chart(
    labels: list[str], values: list[float], colors: list[str], dark: bool, center_caption: str
) -> go.Figure:
    """Shared doughnut builder for both the stock- and sector-level allocation charts."""
    total = sum(values)
    weights = [v / total * 100 if total > 0 else 0.0 for v in values]
    fig = go.Figure(
        go.Pie(
            labels=[f"{n}  {w:.1f}%" for n, w in zip(labels, weights, strict=True)],
            values=values,
            hole=0.62,
            sort=False,
            direction="clockwise",
            textinfo="none",
            marker={"colors": colors,
                    "line": {"color": SURFACE_DARK if dark else SURFACE_LIGHT, "width": 2}},
            customdata=[fmt_inr(v) for v in values],
            hovertemplate="<b>%{label}</b><br>%{customdata}<extra></extra>",
        )
    )
    top_weight = weights[0] if weights else 0.0
    fig.update_layout(
        height=360,
        margin={"l": 10, "r": 10, "t": 10, "b": 10},
        legend={"orientation": "v", "yanchor": "middle", "y": 0.5},
        annotations=[{
            "text": (f"<b>{top_weight:.1f}%</b><br>"
                     f"<span style='font-size:12px'>{center_caption}</span>"),
            "showarrow": False, "font": {"size": 20},
        }],
    )
    return fig


def _top_slices_with_others(
    ranked: pd.DataFrame, label_col: str
) -> tuple[list[str], list[float], list[str]]:
    """Split a value-ranked breakdown into up to :data:`ALLOCATION_SLICES` named
    slices plus a folded "Others" slice, with the matching validated colors."""
    dark = _is_dark_theme()
    top, rest = ranked.head(ALLOCATION_SLICES), ranked.iloc[ALLOCATION_SLICES:]
    names = top[label_col].tolist()
    values = top["current_value"].tolist()
    colors = list((CATEGORICAL_DARK if dark else CATEGORICAL_LIGHT)[: len(top)])
    if not rest.empty:
        names.append(f"Others ({len(rest)})")
        values.append(float(rest["current_value"].sum()))
        colors.append(OTHERS_COLOR)
    return names, values, colors


def allocation_figure(holdings: pd.DataFrame) -> go.Figure:
    """Doughnut of current-value weights: top holdings plus an "Others" slice."""
    ranked = holdings.sort_values("current_value", ascending=False)
    names, values, colors = _top_slices_with_others(ranked, "symbol")
    return _donut_chart(names, values, colors, _is_dark_theme(), "largest<br>position")


def sector_allocation_figure(sector_breakdown: pd.DataFrame) -> go.Figure:
    """Doughnut of current-value weights by sector, plus an "Others" slice."""
    ranked = sector_breakdown.sort_values("current_value", ascending=False)
    names, values, colors = _top_slices_with_others(ranked, "sector")
    return _donut_chart(names, values, colors, _is_dark_theme(), "largest<br>sector")


def sector_drilldown_weight_figure(detail: pd.DataFrame) -> go.Figure:
    """Horizontal bars of each stock's weight within the selected sector.

    A single fixed categorical colour (not the GAIN/LOSS pair): this is a
    weight comparison, not a P&L view - reusing GAIN_COLOR here would
    wrongly imply "good" the way it does on the actual P&L bar charts.
    """
    dark = _is_dark_theme()
    color = (CATEGORICAL_DARK if dark else CATEGORICAL_LIGHT)[0]
    data = detail.dropna(subset=["pct_of_sector"]).sort_values("pct_of_sector")
    fig = go.Figure(
        go.Bar(
            x=data["pct_of_sector"],
            y=data["symbol"],
            orientation="h",
            marker={"color": color, "cornerradius": 4},
            customdata=np.column_stack([
                [fmt_inr(v) for v in data["current_value"]],
                [fmt_pct_plain(v) for v in data["pct_of_portfolio"]],
            ]),
            hovertemplate=(
                "<b>%{y}</b><br>% of sector: %{x:.2f}%<br>Value: %{customdata[0]}"
                "<br>% of portfolio: %{customdata[1]}<extra></extra>"
            ),
        )
    )
    fig.update_layout(
        height=max(360, 26 * len(data) + 80),
        margin={"l": 10, "r": 10, "t": 10, "b": 10},
        bargap=0.35,
        xaxis={"title": "% of sector allocation", "ticksuffix": "%"},
        yaxis={"title": None, "automargin": True},
        showlegend=False,
    )
    return fig


def mutual_fund_allocation_figure(mutual_funds: pd.DataFrame) -> go.Figure:
    """Doughnut of current-value weights by fund, plus an "Others" slice."""
    ranked = mutual_funds.sort_values("current_value", ascending=False)
    names, values, colors = _top_slices_with_others(ranked, "fund_name")
    return _donut_chart(names, values, colors, _is_dark_theme(), "largest<br>fund")


_MARKET_CAP_ORDER = ["Large Cap", "Mid Cap", "Small Cap", UNCATEGORIZED]


def market_cap_figure(cap_breakdown: pd.DataFrame) -> go.Figure:
    """Vertical bar of portfolio weight by market-cap bucket, Large -> Small -> Uncategorized."""
    dark = _is_dark_theme()
    ordered = (
        cap_breakdown.set_index("market_cap_bucket")
        .reindex(_MARKET_CAP_ORDER)
        .dropna(how="all")
    )
    palette = CATEGORICAL_DARK if dark else CATEGORICAL_LIGHT
    bucket_colors = {"Large Cap": palette[0], "Mid Cap": palette[1], "Small Cap": palette[2],
                     UNCATEGORIZED: OTHERS_COLOR}

    def _pct_label(w: float) -> str:
        return f"{w:.1f}%" if np.isfinite(w) else "—"

    fig = go.Figure(
        go.Bar(
            x=ordered.index,
            y=ordered["weight_pct"],
            marker={"color": [bucket_colors[b] for b in ordered.index], "cornerradius": 4},
            customdata=np.column_stack([
                [fmt_inr(v) for v in ordered["current_value"]],
                ordered["holdings_count"],
            ]),
            hovertemplate=(
                "<b>%{x}</b><br>Weight: %{y:.1f}%<br>Value: %{customdata[0]}"
                "<br>%{customdata[1]:.0f} holding(s)<extra></extra>"
            ),
            text=[_pct_label(w) for w in ordered["weight_pct"]],
            textposition="outside",
        )
    )
    fig.update_layout(
        height=360,
        margin={"l": 10, "r": 10, "t": 30, "b": 10},
        yaxis={"title": "Portfolio weight", "ticksuffix": "%", "rangemode": "tozero"},
        xaxis={"title": None},
        showlegend=False,
    )
    return fig


def pnl_bar_figure(holdings: pd.DataFrame) -> go.Figure:
    """Horizontal bars of unrealised return % per stock, best at the top."""
    data = holdings.dropna(subset=["unrealized_pnl_pct"]).sort_values("unrealized_pnl_pct")
    colors = np.where(data["unrealized_pnl_pct"] >= 0, GAIN_COLOR, LOSS_COLOR)
    fig = go.Figure(
        go.Bar(
            x=data["unrealized_pnl_pct"],
            y=data["symbol"],
            orientation="h",
            marker={"color": colors, "cornerradius": 4},
            customdata=np.column_stack([
                [fmt_inr(v, signed=True) for v in data["unrealized_pnl"]],
                [fmt_inr(v) for v in data["current_value"]],
            ]),
            hovertemplate=(
                "<b>%{y}</b><br>Return: %{x:+.2f}%<br>P&L: %{customdata[0]}"
                "<br>Value: %{customdata[1]}<extra></extra>"
            ),
        )
    )
    fig.update_layout(
        height=max(360, 26 * len(data) + 80),
        margin={"l": 10, "r": 10, "t": 10, "b": 10},
        bargap=0.35,
        xaxis={"title": "Unrealised return", "ticksuffix": "%", "zeroline": True,
               "zerolinewidth": 1},
        yaxis={"title": None, "automargin": True},
        showlegend=False,
    )
    return fig


def dividend_income_by_month_figure(realized_events: pd.DataFrame) -> go.Figure:
    """Vertical bars of realised dividend cash flow, summed by calendar month."""
    monthly = (
        realized_events.assign(
            month=realized_events["ex_dividend_date"].dt.to_period("M").dt.to_timestamp()
        )
        .groupby("month", as_index=False)["amount_received"].sum()
        .sort_values("month")
    )
    fig = go.Figure(
        go.Bar(
            x=monthly["month"],
            y=monthly["amount_received"],
            marker={"color": GAIN_COLOR, "cornerradius": 4},
            customdata=[fmt_inr(v) for v in monthly["amount_received"]],
            hovertemplate="<b>%{x|%b %Y}</b><br>%{customdata}<extra></extra>",
        )
    )
    fig.update_layout(
        height=360,
        margin={"l": 10, "r": 10, "t": 10, "b": 10},
        bargap=0.25,
        xaxis={"title": None, "tickformat": "%b %Y"},
        yaxis={"title": "Dividends received", "tickprefix": "₹", "rangemode": "tozero"},
        showlegend=False,
    )
    return fig


def dividend_contributors_figure(holdings_with_dividends: pd.DataFrame) -> go.Figure:
    """Horizontal bars of estimated annual dividend income, top contributors first."""
    data = holdings_with_dividends[holdings_with_dividends["estimated_annual_dividend"] > 0]
    data = data.sort_values("estimated_annual_dividend")
    fig = go.Figure(
        go.Bar(
            x=data["estimated_annual_dividend"],
            y=data["symbol"],
            orientation="h",
            marker={"color": GAIN_COLOR, "cornerradius": 4},
            customdata=np.column_stack([
                [fmt_pct(v) for v in data["dividend_yield_pct"]],
                [fmt_inr(v) for v in data["ttm_dividend_per_share"]],
            ]),
            hovertemplate=(
                "<b>%{y}</b><br>Est. annual: %{x:,.0f}<br>Yield: %{customdata[0]}"
                "<br>TTM/share: %{customdata[1]}<extra></extra>"
            ),
        )
    )
    fig.update_layout(
        height=max(360, 26 * len(data) + 80),
        margin={"l": 10, "r": 10, "t": 10, "b": 10},
        bargap=0.35,
        xaxis={"title": "Estimated annual dividend", "tickprefix": "₹"},
        yaxis={"title": None, "automargin": True},
        showlegend=False,
    )
    return fig


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class SidebarControls:
    """Values collected from the sidebar widgets."""

    demo_mode: bool
    holdings_file: object | None  # streamlit UploadedFile
    trades_file: object | None
    use_live_prices: bool
    refresh_prices: bool


def render_sidebar() -> SidebarControls:
    """Draw sidebar widgets and return their current values."""
    with st.sidebar:
        st.header("Data sources")
        demo_mode = st.toggle(
            "Demo / Synthetic Data Mode", value=False,
            help="Explore the dashboard with a generated portfolio - no files needed.",
        )
        holdings_file = st.file_uploader(
            "Angel One Holdings Report", type=["xlsx", "csv"], disabled=demo_mode,
            help="Needs symbol, quantity and average price columns; LTP is optional.",
        )
        trades_file = st.file_uploader(
            "Angel One Trade History / P&L Report", type=["xlsx", "csv"], disabled=demo_mode,
            help="Trade book (date, symbol, buy/sell, qty, price) or P&L statement "
                 "(buy/sell dates and prices). Required for XIRR and CAGR.",
        )
        st.header("Market data")
        # Demo costs are synthetic, so real prices would produce meaningless P&L.
        use_live = st.toggle(
            "Use live prices (Yahoo Finance)", value=True, disabled=demo_mode,
            help="Off: value positions at the LTP in the uploaded file.",
        ) and not demo_mode
        refresh = st.button(
            "Refresh live prices", icon=":material/refresh:", disabled=not use_live,
            width="stretch",
        )
        if demo_mode:
            st.caption("Demo mode values positions at synthetic prices.")
        else:
            st.caption(f"Prices are cached for {PRICE_CACHE_TTL_SECONDS // 60} minutes.")
    return SidebarControls(
        demo_mode=demo_mode,
        holdings_file=None if demo_mode else holdings_file,
        trades_file=None if demo_mode else trades_file,
        use_live_prices=use_live,
        refresh_prices=refresh,
    )


def render_landing() -> None:
    """Instructions shown before any data is loaded."""
    st.info(
        "Upload your **Angel One Holdings Report** in the sidebar, or switch on "
        "**Demo / Synthetic Data Mode** to explore the dashboard.",
        icon=":material/upload_file:",
    )
    with st.expander("Which files and columns are supported?", expanded=True):
        st.markdown(
            """
- **Holdings report** (`.xlsx` / `.csv`, required): Symbol / Stock name,
  Quantity / Qty, Avg. Price / Average cost, and optionally LTP and Exchange.
- **Trade history / P&L report** (`.xlsx` / `.csv`, optional - enables XIRR & CAGR):
  either a *trade book* (Date, Symbol, Buy/Sell, Qty, Price or Value) or a
  *P&L statement* with Buy Date / Buy Price / Sell Date / Sell Price columns.
- **Mutual funds** - no separate upload needed: if the holdings file is Angel
  One's fuller "Portfolio" export, its Mutual Fund sheet is detected
  automatically and shown on its own tab.

Header rows are detected automatically - client details, date ranges and
disclaimers above or below the table are skipped. Symbols are mapped to Yahoo
Finance tickers (`RELIANCE` -> `RELIANCE.NS`, BSE rows -> `.BO`).
"""
        )


def load_portfolio(
    controls: SidebarControls, as_of: pd.Timestamp
) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame, list[str]] | None:
    """Return ``(holdings, trades, mutual_funds, notes)`` from demo data or uploads, or ``None``.

    ``mutual_funds`` is always a DataFrame (empty when none were found), so
    callers never need a ``None`` check the way they do for ``trades``.
    """
    if controls.demo_mode:
        holdings, trades = generate_demo_data(as_of.date())
        return holdings, trades, generate_demo_mutual_funds(), []
    if controls.holdings_file is None:
        render_landing()
        return None

    upload = controls.holdings_file
    try:
        holdings, notes = parse_holdings_file(upload.getvalue(), upload.name)
    except ParseError as exc:
        st.error(f"**Could not read the holdings report.** {exc}", icon=":material/error:")
        return None
    except Exception as exc:  # unexpected - keep the app alive and show why
        st.error(f"Unexpected error while reading '{upload.name}': {exc}")
        return None
    notes = [f"Holdings: {n}" for n in notes]

    # Angel One's fuller "Portfolio" export bundles a Mutual Fund sheet in
    # the same workbook as equity holdings - look for it in the same upload.
    # Its absence (the normal case for a plain equity holdings file) is not
    # an error, so nothing is shown unless parsing actively fails.
    mutual_funds = pd.DataFrame(columns=MUTUAL_FUND_COLUMNS)
    try:
        mutual_funds, mf_notes = parse_mutual_funds_file(upload.getvalue(), upload.name)
        notes += [f"Mutual funds: {n}" for n in mf_notes]
    except ParseError as exc:
        st.warning(f"**Mutual fund holdings ignored.** {exc}", icon=":material/warning:")
    except Exception as exc:
        st.warning(f"Mutual fund holdings ignored - unexpected error: {exc}")

    trades = None
    if controls.trades_file is not None:
        upload = controls.trades_file
        try:
            trades, trade_notes = parse_trades_file(upload.getvalue(), upload.name)
            notes += [f"Trade history: {n}" for n in trade_notes]
        except ParseError as exc:
            st.warning(f"**Trade history ignored.** {exc}", icon=":material/warning:")
        except Exception as exc:
            st.warning(f"Trade history ignored - unexpected error: {exc}")
    return holdings, trades, mutual_funds, notes


def resolve_prices(
    holdings: pd.DataFrame, use_live: bool
) -> tuple[pd.DataFrame, dict[str, float], str]:
    """Fetch live prices if enabled; returns ``(holdings, prices, status caption)``.

    Before fetching, tickers (and the displayed symbol) guessed from a
    truncated scrip name are upgraded via ISIN where possible (see
    :func:`resolve_isin_tickers`) - the returned ``holdings`` carries these
    corrected values. In the rare case that two different raw names resolve
    to the same real symbol, both rows are kept separately rather than
    merged; portfolio totals stay correct, only the per-symbol breakdown
    would show two rows for one stock.
    """
    if not use_live:
        return holdings, {}, "Live prices off - valuing positions at the LTP from the file."

    isins = tuple(sorted(set(holdings.loc[holdings["isin"] != "", "isin"])))
    if isins:
        with st.spinner("Resolving ticker symbols from ISIN..."):
            isin_map = resolve_isin_tickers(isins)
        if isin_map:
            holdings = holdings.copy()
            better_ticker = holdings["isin"].map(isin_map)
            holdings["symbol"] = better_ticker.map(
                lambda t: t.rsplit(".", 1)[0] if pd.notna(t) else None
            ).fillna(holdings["symbol"])
            holdings["ticker"] = better_ticker.fillna(holdings["ticker"])

    tickers = tuple(sorted(holdings["ticker"].unique()))
    with st.spinner("Fetching live prices from Yahoo Finance..."):
        prices, fetched_at, error = fetch_live_prices(tickers)
    if error:
        st.warning(f"{error} Falling back to prices from the file.", icon=":material/cloud_off:")
        return holdings, {}, f"Live prices unavailable (attempted {fetched_at})."
    return holdings, prices, (
        f"{sum(t in prices for t in tickers)}/{len(tickers)} prices live from Yahoo "
        f"Finance - fetched {fetched_at}."
    )


def resolve_mf_navs(mutual_funds: pd.DataFrame, use_live: bool) -> tuple[dict[str, float], str]:
    """Fetch live AMFI NAVs if enabled; returns ``(live_navs by ISIN, status caption)``.

    Mirrors :func:`resolve_prices`'s shape for equities, but there is no
    ticker/symbol to correct here, so only the NAV map and a status string
    are returned - the caller applies it via :func:`enrich_mutual_funds`.
    """
    if mutual_funds.empty:
        return {}, ""
    if not use_live:
        return {}, "Live NAVs off - using the NAV from the file."

    isins = tuple(sorted(set(mutual_funds.loc[mutual_funds["isin"] != "", "isin"])))
    if not isins:
        return {}, "No ISIN in the file for these funds - using the NAV from the file."

    with st.spinner("Fetching live NAVs from AMFI..."):
        live_navs = fetch_live_mf_navs(isins)
    found = sum(isin in live_navs for isin in isins)
    if not found:
        return {}, "No live NAVs found on AMFI - using the NAV from the file."
    return live_navs, f"{found}/{len(isins)} NAVs live from AMFI (mfapi.in)."


def render_metrics(
    metrics: PortfolioMetrics, as_of: pd.Timestamp, estimated_annual_dividend: float
) -> None:
    """Top row of executive metric cards."""
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric(
        "Total Current Value", fmt_inr_compact(metrics.total_current),
        delta=fmt_inr(metrics.net_pnl, signed=True), border=True,
        help=f"{fmt_inr(metrics.total_current)} at current prices.",
    )
    c2.metric(
        "Total Invested Capital", fmt_inr_compact(metrics.total_invested), border=True,
        help=f"{fmt_inr(metrics.total_invested)} - sum of quantity x average price.",
    )
    c3.metric(
        "Net P&L (%)", fmt_pct(metrics.net_pnl_pct),
        delta=fmt_inr(metrics.net_pnl, signed=True), border=True,
        help="Unrealised P&L on current holdings relative to their cost basis.",
    )

    if metrics.xirr_method == "unavailable":
        xirr_help = "Upload a trade history / P&L report to compute XIRR."
    elif metrics.xirr_method == "fallback":
        xirr_help = "Solver did not converge - showing the simple portfolio return instead."
    else:
        xirr_help = ("Money-weighted annual return: every trade cash flow plus today's "
                     "market value as a final inflow.")
    c4.metric(
        "Portfolio XIRR (%)",
        fmt_pct(metrics.xirr_pct) + (" *" if metrics.xirr_method == "fallback" else ""),
        border=True, help=xirr_help,
    )

    if metrics.cagr_start is None:
        cagr_help = "Upload a trade history / P&L report to date the first purchase."
    else:
        years = (as_of - metrics.cagr_start).days / DAYS_PER_YEAR
        cagr_help = (f"(Current / Invested)^(1/t) - 1 with t = {years:.2f} years since the "
                     f"first purchase on {metrics.cagr_start:%d %b %Y}.")
        if years < 1:
            cagr_help += " Periods under a year are annualised (extrapolated)."
    c5.metric("Portfolio CAGR (%)", fmt_pct(metrics.cagr_pct), border=True, help=cagr_help)
    c6.metric(
        "Est. Annual Dividends", fmt_inr_compact(estimated_annual_dividend), border=True,
        help=f"{fmt_inr(estimated_annual_dividend)}/year, projected from each holding's "
             "trailing-twelve-month dividend per share at current quantities - see the "
             "Dividends & Corporate Actions tab for the per-stock breakdown.",
    )

    if metrics.xirr_method == "fallback":
        st.caption("\\* XIRR did not converge for these cash flows; showing net P&L % instead.")


def _mover_cards(movers: pd.DataFrame, empty_message: str) -> None:
    """Render a row of metric cards for gainers or laggards."""
    if movers.empty:
        st.caption(empty_message)
        return
    for column, (_, row) in zip(st.columns(TOP_N_MOVERS), movers.iterrows(), strict=False):
        column.metric(
            row["symbol"], fmt_pct(row["unrealized_pnl_pct"]),
            delta=fmt_inr_compact(row["unrealized_pnl"], signed=True), border=True,
            help=(f"LTP {fmt_inr(row['ltp'])} vs average {fmt_inr(row['average_price'])}; "
                  f"unrealised P&L {fmt_inr(row['unrealized_pnl'], signed=True)}."),
        )


def render_movers(holdings: pd.DataFrame) -> None:
    """Top gainers / laggards section (one full-width row each so values never truncate)."""
    gainers, laggards = top_movers(holdings)
    st.subheader("Top gainers", anchor=False)
    _mover_cards(gainers, "Not enough priced positions to rank.")
    st.subheader("Top laggards", anchor=False)
    _mover_cards(laggards, "Not enough priced positions to rank.")


def render_performance_split(holdings: pd.DataFrame) -> None:
    """Aggregate profitable vs. loss-making holdings (count and total ₹)."""
    st.subheader("Performance Split", anchor=False)
    profitable, lossmaking = split_gainers_losers(holdings)
    c1, c2 = st.columns(2)
    c1.metric(
        "Profitable Holdings (Green)", f"{len(profitable)} stock(s)",
        delta=fmt_inr_compact(profitable["unrealized_pnl"].sum(), signed=True), border=True,
        help=f"Total unrealised gain: {fmt_inr(profitable['unrealized_pnl'].sum())}.",
    )
    c2.metric(
        "Loss-Making Holdings (Red)", f"{len(lossmaking)} stock(s)",
        delta=fmt_inr_compact(lossmaking["unrealized_pnl"].sum(), signed=True), border=True,
        help=f"Total unrealised loss: {fmt_inr(lossmaking['unrealized_pnl'].sum())}.",
    )


def render_benchmark_comparison(
    trades: pd.DataFrame | None, as_of: pd.Timestamp, metrics: PortfolioMetrics
) -> None:
    """NIFTY 50 XIRR and Alpha - silently omitted without a genuine solved portfolio XIRR
    (no trade history, or the solver fell back), since comparing a fallback net-P&L%
    figure against a real benchmark XIRR would be comparing two different things."""
    if trades is None or trades.empty or metrics.xirr_method != "xirr" or metrics.xirr_pct is None:
        return
    comparison = compute_benchmark_alpha(trades, as_of, metrics.xirr_pct)
    if comparison is None:
        return

    st.subheader("Benchmark Comparison", anchor=False)
    c1, c2 = st.columns(2)
    c1.metric(
        f"{comparison.benchmark_name} XIRR (%)", fmt_pct(comparison.benchmark_xirr_pct),
        border=True,
        help=(f"What the same rupee amounts, invested in {comparison.benchmark_name} on the "
              "same dates as your actual trades, would be worth today - a like-for-like "
              "comparison, not just the index's own price return."),
    )
    c2.metric(
        "Alpha vs. Benchmark (%)", fmt_pct(comparison.alpha_pct), border=True,
        help=f"Portfolio XIRR minus the {comparison.benchmark_name} XIRR above. Positive means "
             "you beat the benchmark on a same-cash-flow-timing basis.",
    )


def render_charts(holdings: pd.DataFrame) -> None:
    """Allocation doughnut and per-stock return bars."""
    left, right = st.columns(2, gap="large")
    with left:
        st.subheader("Allocation", anchor=False)
        st.plotly_chart(allocation_figure(holdings), width="stretch")
        weights = holdings["weight_pct"].sort_values(ascending=False)
        if weights.notna().any():
            st.caption(
                f"Largest position {weights.iloc[0]:.1f}% · top 5 = "
                f"{weights.head(5).sum():.1f}% of portfolio value."
            )
    with right:
        st.subheader("Unrealised return by stock", anchor=False)
        st.plotly_chart(pnl_bar_figure(holdings), width="stretch")


def _pnl_color(value: float) -> str:
    """Styler CSS for signed P&L cells."""
    if not _is_finite(value) or value == 0:
        return ""
    return f"color: {GAIN_COLOR if value > 0 else LOSS_COLOR}"


def render_breakdown_table(breakdown: pd.DataFrame, key_col: str, key_label: str) -> None:
    """Formatted table shared by the sector and market-cap breakdown tabs."""
    columns = {
        key_col: key_label, "holdings_count": "Stocks", "invested_value": "Invested",
        "current_value": "Current Value", "unrealized_pnl": "P&L",
        "unrealized_pnl_pct": "P&L %", "weight_pct": "Weight",
    }
    table = breakdown[list(columns)].rename(columns=columns)
    styler = table.style.format(
        {"Invested": INR_FORMAT, "Current Value": INR_FORMAT,
         "P&L": lambda v: fmt_inr(v, signed=True), "P&L %": PCT_FORMAT, "Weight": "{:.2f}%"},
        na_rep="—",
    ).map(_pnl_color, subset=["P&L", "P&L %"])
    st.dataframe(styler, hide_index=True, width="stretch")


def render_tab_all_holdings(holdings: pd.DataFrame) -> None:
    """Full holdings table, filterable by symbol search, sector, market cap and P&L status."""
    st.subheader("Holdings", anchor=False)
    f_search, f_sector, f_cap, f_pnl = st.columns([2, 1, 1, 1])
    with f_search:
        query = st.text_input(
            "Search by symbol", placeholder="e.g. RELIANCE", label_visibility="collapsed",
            icon=":material/search:",
        )
    with f_sector:
        sectors = sorted(holdings["sector"].unique())
        sector_choice = st.selectbox("Sector", ["All Sectors", *sectors])
    with f_cap:
        cap_values = holdings["market_cap_bucket"].to_numpy()
        present_caps = [b for b in _MARKET_CAP_ORDER if b in cap_values]
        cap_choice = st.selectbox("Market Cap", ["All Market Caps", *present_caps])
    with f_pnl:
        pnl_choice = st.selectbox("P&L Status", ["All", "Gainers", "Losers"])

    view = holdings
    if query.strip():
        view = view[view["symbol"].str.contains(query.strip(), case=False, regex=False)]
    if sector_choice != "All Sectors":
        view = view[view["sector"] == sector_choice]
    if cap_choice != "All Market Caps":
        view = view[view["market_cap_bucket"] == cap_choice]
    if pnl_choice == "Gainers":
        view = view[view["unrealized_pnl"] > 0]
    elif pnl_choice == "Losers":
        view = view[view["unrealized_pnl"] < 0]
    if view.empty:
        st.caption("No holdings match the current filters.")
        return

    columns = {
        "symbol": "Symbol", "sector": "Sector", "market_cap_bucket": "Market Cap",
        "quantity": "Qty", "average_price": "Avg Price", "ltp": "LTP",
        "invested_value": "Invested", "current_value": "Current Value",
        "unrealized_pnl": "Unrealised P&L", "unrealized_pnl_pct": "P&L %",
        "weight_pct": "Weight", "price_source": "Price Source",
    }
    table = (
        view.sort_values("current_value", ascending=False)[list(columns)]
        .rename(columns=columns)
        .reset_index(drop=True)
    )
    money = {c: INR_FORMAT for c in ("Avg Price", "LTP", "Invested", "Current Value")}
    styler = (
        table.style.format(
            {**money, "Qty": fmt_quantity, "Unrealised P&L": lambda v: fmt_inr(v, signed=True),
             "P&L %": PCT_FORMAT, "Weight": "{:.2f}%"},
            na_rep="—",
        )
        .map(_pnl_color, subset=["Unrealised P&L", "P&L %"])
    )
    st.dataframe(styler, hide_index=True, width="stretch")
    st.caption(f"{len(view)} of {len(holdings)} positions shown.")


def render_trades_table(trades: pd.DataFrame) -> None:
    """Collapsible view of the parsed trade log so users can verify parsing."""
    with st.expander(f"Parsed trade history ({len(trades)} trades)"):
        table = trades.drop(columns=["ticker"]).rename(
            columns=lambda c: c.replace("_", " ").title()
        )
        st.dataframe(
            table.style.format(
                {"Date": "{:%d %b %Y}", "Quantity": fmt_quantity, "Price": INR_FORMAT,
                 "Cash Flow": lambda v: fmt_inr(v, signed=True)}
            ).map(_pnl_color, subset=["Cash Flow"]),
            hide_index=True, width="stretch",
        )


def render_tab_overview(
    metrics: PortfolioMetrics, holdings: pd.DataFrame, trades: pd.DataFrame | None,
    dividend_metrics: DividendMetrics, as_of: pd.Timestamp,
) -> None:
    """Tab 1: top metrics, benchmark alpha, performance split, movers and allocation charts."""
    render_metrics(metrics, as_of, dividend_metrics.estimated_annual_income)
    render_benchmark_comparison(trades, as_of, metrics)
    st.divider()
    render_performance_split(holdings)
    st.divider()
    render_movers(holdings)
    st.divider()
    render_charts(holdings)


def render_tab_sector(holdings: pd.DataFrame) -> None:
    """Tab 2: sector allocation doughnut, HHI concentration risk, an interactive per-sector
    performance table, and a drill-down into the selected sector's individual holdings."""
    sector_df = group_breakdown(holdings, "sector")
    all_uncategorized = (sector_df["sector"] == UNCATEGORIZED).all()

    left, right = st.columns([3, 2], gap="large")
    with left:
        st.subheader("Sector Allocation", anchor=False)
        st.plotly_chart(sector_allocation_figure(sector_df), width="stretch")
    with right:
        st.subheader("Concentration Risk", anchor=False)
        if all_uncategorized:
            st.info(
                "Sector data unavailable for every holding (offline, or Yahoo Finance "
                "returned no sector info) - concentration risk can't be assessed.",
                icon=":material/info:",
            )
        else:
            hhi = compute_hhi(sector_df["weight_pct"])
            risk = hhi_risk_level(hhi)
            st.metric(
                "Herfindahl-Hirschman Index", f"{hhi:,.0f}", border=True,
                help="Sum of squared sector-weight percentages (0-10,000 scale): "
                     f"< {HHI_MODERATE_THRESHOLD:,.0f} Low, < {HHI_HIGH_THRESHOLD:,.0f} "
                     "Moderate, otherwise High concentration.",
            )
            banner = {"Low Concentration": st.success, "Moderate Concentration": st.warning,
                      "High Concentration": st.error}[risk]
            banner(f"**{risk}**")

    st.subheader("Sector Performance", anchor=False)
    st.caption("Click a row, or use the dropdown below, to drill into that sector's holdings.")
    table_selected = render_sector_summary_table(sector_df)
    selected_sector = render_sector_selector(sector_df, table_selected)

    st.divider()
    render_sector_drilldown(holdings, selected_sector)


def render_sector_summary_table(sector_df: pd.DataFrame) -> str | None:
    """Interactive per-sector performance table.

    Clicking a row selects that sector for the drill-down below. Returns
    the clicked sector, or ``None`` if no row is selected yet - the caller
    (:func:`render_sector_selector`) then defaults to the top sector by
    value, per the existing sort order of ``sector_df``.
    """
    columns = {
        "sector": "Sector", "holdings_count": "Stocks", "invested_value": "Invested",
        "current_value": "Current Value", "unrealized_pnl": "P&L",
        "unrealized_pnl_pct": "P&L %", "weight_pct": "Weight",
    }
    table = sector_df[list(columns)].rename(columns=columns).reset_index(drop=True)
    styler = table.style.format(
        {"Invested": INR_FORMAT, "Current Value": INR_FORMAT,
         "P&L": lambda v: fmt_inr(v, signed=True), "P&L %": PCT_FORMAT, "Weight": "{:.2f}%"},
        na_rep="—",
    ).map(_pnl_color, subset=["P&L", "P&L %"])

    event = st.dataframe(
        styler, hide_index=True, width="stretch",
        on_select="rerun", selection_mode="single-row", key="sector_table_select",
    )
    selected_rows = event.selection.rows if event and event.selection else []
    return str(table.iloc[selected_rows[0]]["Sector"]) if selected_rows else None


def render_sector_selector(sector_df: pd.DataFrame, table_selected: str | None) -> str:
    """Sector to drill into: a table row click takes priority over the dropdown below it,
    and is synced into the dropdown's own state so the two never show different sectors.
    Defaults to the top sector by value (``sector_df`` is already sorted that way)."""
    sectors = sector_df["sector"].tolist()
    if table_selected in sectors:
        st.session_state["sector_drilldown_choice"] = table_selected
    current = st.session_state.get("sector_drilldown_choice", sectors[0])
    if current not in sectors:
        current = sectors[0]
    return st.selectbox(
        "Or pick a sector to drill down:", sectors,
        index=sectors.index(current), key="sector_drilldown_choice",
    )


def render_sector_drilldown(holdings: pd.DataFrame, selected_sector: str) -> None:
    """Per-stock detail for one sector: metric cards, a table and a weight-comparison chart."""
    detail = compute_sector_drilldown(holdings, selected_sector)
    st.markdown(f"### Sector Drill-Down: {selected_sector}")
    if detail.empty:
        st.caption("No holdings in this sector.")
        return

    total_value = float(detail["current_value"].sum())
    total_invested = float(detail["invested_value"].sum())
    pnl = total_value - total_invested
    pnl_pct = pnl / total_invested * 100 if total_invested > 0 else None

    c1, c2, c3 = st.columns(3)
    c1.metric(
        "Total Sector Value", fmt_inr_compact(total_value), border=True,
        help=fmt_inr(total_value),
    )
    c2.metric(
        "Sector P&L (%)", fmt_pct(pnl_pct), delta=fmt_inr(pnl, signed=True), border=True,
    )
    c3.metric("Number of Holdings", f"{len(detail)} stock(s)", border=True)

    columns = {
        "symbol": "Symbol", "quantity": "Quantity", "average_price": "Avg. Price (₹)",
        "ltp": "Current LTP (₹)", "current_value": "Current Value (₹)",
        "unrealized_pnl": "Unrealized P&L (₹)", "unrealized_pnl_pct": "Unrealized P&L (%)",
        "pct_of_sector": "% of Sector", "pct_of_portfolio": "% of Portfolio",
    }
    table = detail[list(columns)].rename(columns=columns)
    styler = table.style.format(
        {"Quantity": fmt_quantity, "Avg. Price (₹)": INR_FORMAT, "Current LTP (₹)": INR_FORMAT,
         "Current Value (₹)": INR_FORMAT, "Unrealized P&L (₹)": lambda v: fmt_inr(v, signed=True),
         "Unrealized P&L (%)": PCT_FORMAT, "% of Sector": "{:.2f}%", "% of Portfolio": "{:.2f}%"},
        na_rep="—",
    ).map(_pnl_color, subset=["Unrealized P&L (₹)", "Unrealized P&L (%)"])
    st.dataframe(styler, hide_index=True, width="stretch")

    st.plotly_chart(sector_drilldown_weight_figure(detail), width="stretch")


def render_tab_marketcap(holdings: pd.DataFrame) -> None:
    """Tab 3: market-cap bucket distribution chart and weightage table."""
    cap_df = group_breakdown(holdings, "market_cap_bucket")
    st.subheader("Market Cap Distribution", anchor=False)
    st.plotly_chart(market_cap_figure(cap_df), width="stretch")
    st.subheader("Weightage Breakdown", anchor=False)
    render_breakdown_table(cap_df, "market_cap_bucket", "Market Cap")
    if UNCATEGORIZED in cap_df["market_cap_bucket"].to_numpy():
        st.caption("'Uncategorized' holdings have no market-cap data from Yahoo Finance "
                   "(common for ETFs) or the app is offline.")


def render_tab_tax(
    holdings: pd.DataFrame, trades: pd.DataFrame | None, as_of: pd.Timestamp
) -> None:
    """Tab 4: LTCG/STCG exposure estimate and a per-lot holding-period taxonomy table."""
    if trades is None or trades.empty:
        st.info(
            "Upload a Trade History / P&L report to see holding-period and tax classification.",
            icon=":material/info:",
        )
        return

    lots = compute_tax_lots(trades, as_of)
    priced_lots = value_tax_lots(lots, holdings)
    estimate = estimate_capital_gains(priced_lots)

    c1, c2 = st.columns(2)
    c1.metric(
        f"STCG Exposure (≤{LTCG_HOLDING_DAYS} days)", fmt_inr_compact(estimate.stcg_tax),
        border=True,
        help=(f"Unrealised short-term gain {fmt_inr(estimate.stcg_gain, signed=True)}, "
              f"taxed at a flat {STCG_TAX_RATE:.0%} (no exemption)."),
    )
    c2.metric(
        f"LTCG Exposure (>{LTCG_HOLDING_DAYS} days)", fmt_inr_compact(estimate.ltcg_tax),
        border=True,
        help=(f"Unrealised long-term gain {fmt_inr(estimate.ltcg_gain, signed=True)}; "
              f"{fmt_inr(estimate.ltcg_taxable)} taxable above the "
              f"{fmt_inr(LTCG_EXEMPTION)} exemption, at {LTCG_TAX_RATE:.1%}."),
    )
    st.caption(
        "Estimates only, as if every position were sold today; not tax advice. Ignores "
        "realised gains/losses, equity held outside this portfolio, and any exemption "
        "already used elsewhere this financial year."
    )
    if priced_lots.empty:
        st.caption("No held lots could be matched from the trade history.")
        return

    mismatch = reconcile_positions(holdings, trades)
    if not mismatch.empty:
        st.warning(
            f"Trade history doesn't fully reconcile with holdings for {len(mismatch)} "
            "stock(s) - see Data quality notes above. The tax lots below may be incomplete "
            "for those stocks.",
            icon=":material/warning:",
        )

    st.subheader("Holding Period & Tax Taxonomy", anchor=False)
    columns = {
        "symbol": "Symbol", "purchase_date": "Purchase Date", "quantity": "Qty",
        "holding_days": "Days Held", "term": "Term", "invested_value": "Invested",
        "current_value": "Current Value", "unrealized_gain": "Unrealised Gain",
    }
    table = (
        priced_lots.sort_values(["term", "symbol"])[list(columns)]
        .rename(columns=columns)
        .reset_index(drop=True)
    )
    styler = table.style.format(
        {"Purchase Date": "{:%d %b %Y}", "Qty": fmt_quantity, "Days Held": "{:,.0f}",
         "Invested": INR_FORMAT, "Current Value": INR_FORMAT,
         "Unrealised Gain": lambda v: fmt_inr(v, signed=True)},
        na_rep="—",
    ).map(_pnl_color, subset=["Unrealised Gain"])
    st.dataframe(styler, hide_index=True, width="stretch")

    render_loss_harvesting(priced_lots, estimate)
    render_trades_table(trades)


def render_loss_harvesting(priced_lots: pd.DataFrame, estimate: CapitalGainsEstimate) -> None:
    """Tax Loss Harvesting section: loss-making lots and the tax saved by realising them."""
    st.subheader("Tax Loss Harvesting Opportunities", anchor=False)
    suggestion = suggest_tax_loss_harvesting(priced_lots, estimate)
    if suggestion.loss_lots.empty:
        st.caption("No loss-making positions to harvest right now.")
        return

    c1, c2 = st.columns(2)
    c1.metric(
        "Harvestable Loss", fmt_inr_compact(suggestion.total_harvestable_loss), border=True,
        help=f"Total unrealised loss across {len(suggestion.loss_lots)} lot(s) if sold today.",
    )
    c2.metric(
        "Potential Tax Savings", fmt_inr_compact(suggestion.potential_savings), border=True,
        help=(f"Estimated tax drops from {fmt_inr(suggestion.current_tax)} to "
              f"{fmt_inr(suggestion.harvested_tax)} if these losses are realised: a "
              "Short-Term loss can offset both STCG and LTCG gains; a Long-Term loss can "
              "only offset LTCG (never a Short-Term gain)."),
    )
    st.caption(
        "Shows what selling these loss-making lots today, and using the loss to offset your "
        "gains above, would save versus the baseline estimate - not a recommendation to sell. "
        "India has no wash-sale rule, but rebuying resets that lot's holding period from "
        "scratch, so a long-term position bought back becomes short-term again."
    )
    columns = {
        "symbol": "Symbol", "purchase_date": "Purchase Date", "quantity": "Qty",
        "term": "Term", "invested_value": "Invested", "current_value": "Current Value",
        "unrealized_gain": "Unrealised Loss",
    }
    table = suggestion.loss_lots[list(columns)].rename(columns=columns).reset_index(drop=True)
    styler = table.style.format(
        {"Purchase Date": "{:%d %b %Y}", "Qty": fmt_quantity,
         "Invested": INR_FORMAT, "Current Value": INR_FORMAT,
         "Unrealised Loss": lambda v: fmt_inr(v, signed=True)},
        na_rep="—",
    ).map(_pnl_color, subset=["Unrealised Loss"])
    st.dataframe(styler, hide_index=True, width="stretch")


def render_tab_dividends(
    holdings_with_dividends: pd.DataFrame, realized_events: pd.DataFrame,
    dividend_metrics: DividendMetrics, has_trades: bool,
) -> None:
    """Tab 5: dividend income summary, realised cash flow, top contributors and a table."""
    c1, c2, c3, c4 = st.columns(4)
    if not has_trades:
        realized_display, realized_help = (
            "N/A", "Upload a Trade History / P&L report to compute realised dividends.",
        )
    else:
        realized_display = fmt_inr_compact(dividend_metrics.total_realized)
        realized_help = (f"{fmt_inr(dividend_metrics.total_realized)} received on currently "
                          "held lots since each lot's own purchase date.")
    c1.metric("Total Realized Dividends", realized_display, border=True, help=realized_help)
    c2.metric(
        "Est. Annual Passive Income", fmt_inr_compact(dividend_metrics.estimated_annual_income),
        border=True,
        help=f"{fmt_inr(dividend_metrics.estimated_annual_income)}/year at current holdings "
             "and trailing-twelve-month dividend rates.",
    )
    c3.metric(
        "Portfolio Dividend Yield", fmt_pct_plain(dividend_metrics.portfolio_yield_pct),
        border=True,
        help="Trailing-twelve-month dividend yield, weighted by each holding's current value.",
    )
    c4.metric(
        "Top Dividend Payer", dividend_metrics.top_payer_symbol or "—", border=True,
        help=(f"Estimated annual dividend {fmt_inr(dividend_metrics.top_payer_annual)}."
              if dividend_metrics.top_payer_symbol else "No dividend-paying holdings found."),
    )

    if (holdings_with_dividends["estimated_annual_dividend"] <= 0).all():
        st.info(
            "No dividend history found for any holding - either none of these stocks paid a "
            "dividend in the trailing year, or Yahoo Finance was unreachable.",
            icon=":material/info:",
        )
        return

    st.divider()
    left, right = st.columns(2, gap="large")
    with left:
        st.subheader("Monthly Dividend Income (Realised)", anchor=False)
        if not has_trades:
            st.caption("Upload a Trade History / P&L report to see realised dividend cash "
                       "flow by month.")
        elif realized_events.empty:
            st.caption("No realised dividends matched to your currently held lots yet.")
        else:
            st.plotly_chart(dividend_income_by_month_figure(realized_events), width="stretch")
    with right:
        st.subheader("Top Dividend Contributors (Est. Annual)", anchor=False)
        st.plotly_chart(dividend_contributors_figure(holdings_with_dividends), width="stretch")

    st.subheader("Dividend Details", anchor=False)
    realized_by_symbol = (
        realized_events.groupby("symbol")["amount_received"].sum()
        if has_trades and not realized_events.empty else pd.Series(dtype=float)
    )
    table_df = holdings_with_dividends.copy()
    table_df["total_realized_dividend"] = (
        table_df["symbol"].map(realized_by_symbol) if has_trades else np.nan
    )
    columns = {
        "symbol": "Stock", "quantity": "Shares Held", "ltp": "Current LTP",
        "ttm_dividend_per_share": "TTM Dividend/Share", "dividend_yield_pct": "Dividend Yield %",
        "estimated_annual_dividend": "Est. Annual Dividend (₹)",
        "total_realized_dividend": "Total Realized Dividend (₹)",
    }
    table = (
        table_df.sort_values("estimated_annual_dividend", ascending=False)[list(columns)]
        .rename(columns=columns)
        .reset_index(drop=True)
    )
    styler = table.style.format(
        {"Shares Held": fmt_quantity, "Current LTP": INR_FORMAT,
         "TTM Dividend/Share": INR_FORMAT, "Dividend Yield %": "{:.2f}%",
         "Est. Annual Dividend (₹)": INR_FORMAT, "Total Realized Dividend (₹)": INR_FORMAT},
        na_rep="—",
    )
    st.dataframe(styler, hide_index=True, width="stretch")


def render_tab_mutual_funds(mutual_funds: pd.DataFrame, use_live: bool) -> None:
    """Tab 6: mutual fund holdings, parsed from a Mutual Fund sheet in the same file."""
    if mutual_funds.empty:
        st.info(
            "No mutual fund holdings were found in the uploaded file. Angel One's fuller "
            "'Portfolio' export (not the plain Holdings CSV) includes a Mutual Fund sheet "
            "alongside equity holdings, which is picked up automatically when present.",
            icon=":material/info:",
        )
        return

    live_navs, nav_status = resolve_mf_navs(mutual_funds, use_live)
    enriched = enrich_mutual_funds(mutual_funds, live_navs)
    metrics = compute_mutual_fund_metrics(enriched)

    st.caption(nav_status)
    at_file = enriched.loc[enriched["nav_source"] == "File", "fund_name"].tolist()
    if use_live and at_file and len(at_file) < len(enriched):
        st.warning(f"No live NAV for {', '.join(at_file)} - using the NAV from the file.",
                   icon=":material/warning:")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric(
        "Total Current Value", fmt_inr_compact(metrics.total_current),
        delta=fmt_inr(metrics.net_pnl, signed=True), border=True,
        help=f"{fmt_inr(metrics.total_current)} across {len(enriched)} fund(s).",
    )
    c2.metric(
        "Total Invested Capital", fmt_inr_compact(metrics.total_invested), border=True,
        help=f"{fmt_inr(metrics.total_invested)}.",
    )
    c3.metric(
        "Net P&L (%)", fmt_pct(metrics.net_pnl_pct),
        delta=fmt_inr(metrics.net_pnl, signed=True), border=True,
    )
    c4.metric(
        "Weighted Avg. XIRR (%)", fmt_pct(metrics.weighted_xirr_pct), border=True,
        help="Each fund's own reported XIRR, averaged by invested value - not a true "
             "portfolio XIRR (that needs every fund's full cash-flow history, which a "
             "holdings snapshot doesn't carry).",
    )

    st.divider()
    st.subheader("Fund Allocation", anchor=False)
    st.plotly_chart(mutual_fund_allocation_figure(enriched), width="stretch")

    st.subheader("Funds", anchor=False)
    columns = {
        "fund_name": "Fund", "units": "Units", "average_nav": "Avg NAV",
        "current_nav": "Current NAV", "invested_value": "Invested",
        "current_value": "Current Value", "unrealized_pnl": "Unrealised P&L",
        "unrealized_pnl_pct": "P&L %", "xirr_pct": "XIRR %", "weight_pct": "Weight",
        "nav_source": "NAV Source",
    }
    table = (
        enriched.sort_values("current_value", ascending=False)[list(columns)]
        .rename(columns=columns)
        .reset_index(drop=True)
    )
    styler = table.style.format(
        {"Units": "{:,.3f}", "Avg NAV": INR_FORMAT, "Current NAV": INR_FORMAT,
         "Invested": INR_FORMAT, "Current Value": INR_FORMAT,
         "Unrealised P&L": lambda v: fmt_inr(v, signed=True), "P&L %": PCT_FORMAT,
         "XIRR %": PCT_FORMAT, "Weight": "{:.2f}%"},
        na_rep="—",
    ).map(_pnl_color, subset=["Unrealised P&L", "P&L %"])
    st.dataframe(styler, hide_index=True, width="stretch")


def render_data_warnings(
    holdings: pd.DataFrame, trades: pd.DataFrame | None, notes: list[str], use_live: bool
) -> None:
    """Surface pricing gaps, reconciliation gaps and parser notes."""
    if use_live:
        at_file = holdings.loc[holdings["price_source"] == "File", "symbol"].tolist()
        if at_file and len(at_file) < len(holdings):
            st.warning(f"No live price for {', '.join(at_file)} - using the LTP from the file.",
                       icon=":material/warning:")
    at_cost = holdings.loc[holdings["price_source"] == "Cost", "symbol"].tolist()
    if at_cost:
        st.warning(
            f"No price available for {', '.join(at_cost)} - valued at average cost (0% P&L).",
            icon=":material/warning:",
        )

    if trades is not None and not trades.empty:
        mismatch = reconcile_positions(holdings, trades)
        if not mismatch.empty:
            details = ", ".join(
                f"{r.symbol} (held {fmt_quantity(r.held)} vs "
                f"net traded {fmt_quantity(r.traded)})"
                for r in mismatch.head(8).itertuples()
            )
            more = f" and {len(mismatch) - 8} more" if len(mismatch) > 8 else ""
            notes = [
                f"Trade history does not reconcile with holdings for {len(mismatch)} stock(s): "
                f"{details}{more}. XIRR/CAGR assume the uploaded trades built today's "
                "portfolio, so missing older trades, bonus/split shares or transfers "
                "will distort them.",
                *notes,
            ]
    if notes:
        with st.expander(f"Data quality notes ({len(notes)})", icon=":material/info:"):
            for note in notes:
                st.markdown(f"- {note}")


def main() -> None:
    """Streamlit entry point."""
    st.set_page_config(
        page_title="Angel One Portfolio Manager", page_icon=":material/monitoring:",
        layout="wide",
    )
    controls = render_sidebar()
    st.title("Angel One Portfolio Manager", anchor=False)

    if controls.refresh_prices:
        fetch_live_prices.clear()
        fetch_live_mf_navs.clear()
        fetch_benchmark_series.clear()
        st.toast("Refreshing market prices...", icon=":material/refresh:")

    as_of = pd.Timestamp(datetime.now(IST).date())
    loaded = load_portfolio(controls, as_of)
    if loaded is None:
        return
    holdings, trades, mutual_funds, notes = loaded

    holdings, prices, price_status = resolve_prices(holdings, controls.use_live_prices)
    if controls.demo_mode:
        price_status = "Demo mode - synthetic holdings, trades and prices for illustration."
    enriched = enrich_holdings(holdings, prices)
    enriched = enrich_sector_and_cap(enriched)
    metrics = compute_portfolio_metrics(enriched, trades, as_of)

    dividend_history = fetch_dividend_history(tuple(sorted(enriched["ticker"].unique())))
    holdings_with_dividends = enrich_dividend_metrics(enriched, dividend_history, as_of)
    has_trades = trades is not None and not trades.empty
    realized_events = (
        compute_realized_dividend_events(trades, dividend_history, as_of) if has_trades
        else pd.DataFrame(columns=["symbol", "ex_dividend_date", "amount_received"])
    )
    dividend_metrics = compute_dividend_metrics(
        holdings_with_dividends, realized_events if has_trades else None
    )
    if not enriched.empty and dividend_history.empty:
        notes.append(
            "Dividends: no dividend data found for any holding (offline, rate-limited, "
            "or none of these stocks pay dividends)."
        )

    st.caption(price_status)
    render_data_warnings(enriched, trades, notes, controls.use_live_prices)

    tab_overview, tab_sector, tab_cap, tab_tax, tab_div, tab_mf, tab_all = st.tabs([
        "Overview & Top Metrics", "Sector Breakdown", "Market Cap Breakdown",
        "Tax & Holding Period", "Dividends & Corporate Actions", "Mutual Funds", "All Holdings",
    ])
    with tab_overview:
        render_tab_overview(metrics, enriched, trades, dividend_metrics, as_of)
    with tab_sector:
        render_tab_sector(enriched)
    with tab_cap:
        render_tab_marketcap(enriched)
    with tab_tax:
        render_tab_tax(enriched, trades, as_of)
    with tab_div:
        render_tab_dividends(holdings_with_dividends, realized_events, dividend_metrics, has_trades)
    with tab_mf:
        render_tab_mutual_funds(mutual_funds, controls.use_live_prices)
    with tab_all:
        render_tab_all_holdings(enriched)


if __name__ == "__main__":
    main()
