# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-file Streamlit dashboard (`app.py`) for Angel One traders: upload a holdings export
and (optionally) a trade history / P&L export, and it enriches positions with live NSE/BSE
prices from Yahoo Finance and reports XIRR, CAGR, per-stock P&L, sector/market-cap breakdowns,
an LTCG/STCG tax-exposure estimate, and mutual fund holdings — across six `st.tabs()` (Overview,
Sector Breakdown, Market Cap Breakdown, Tax & Holding Period, Mutual Funds, All Holdings). It
also has a "Demo / Synthetic Data Mode" that generates a self-consistent fake portfolio so the
UI can be exercised without any files.

There is no test suite, build step, or CI in the repo — verification is done by running the
app directly (see below).

## Commands

```powershell
# Setup (venv already exists at .venv if present; otherwise create it)
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt

# Run the app
streamlit run app.py
# or, for headless/background runs (e.g. while iterating):
python -m streamlit run app.py --server.headless true --server.port 8765

# Lint (project convention used throughout development; not pinned in a config file)
ruff check --line-length 100 --select E,W,F,B,UP,N,SIM,I --target-version py310 app.py
```

There are no `pytest` tests checked into the repo. When adding or verifying logic, write a
throwaway test script (e.g. under a scratch/temp directory, not committed) that imports `app`
directly — every function is a plain, testable unit (parsing, `solve_xirr`, `compute_cagr`,
`enrich_holdings`, `compute_tax_lots`, `group_breakdown`, etc. take/return plain DataFrames and
values, no Streamlit state required). For UI-level checks, `streamlit.testing.v1.AppTest.from_file(...)`
can drive the app headlessly — set widget values, `.run()`, then inspect `at.metric`,
`at.exception`, `at.tabs[i].dataframe`, etc. — without a browser. **Prefer `AppTest` widget
manipulation over browser-clicking a `st.selectbox`**: the BaseWeb dropdown used by
`st.selectbox` has been flaky to drive via click automation (options not registering); driving
it through `AppTest`'s `.set_value(...).run()` is both faster and more reliable for filter
widgets like the ones on the "All Holdings" tab.

## Architecture

Everything lives in `app.py`, organized top-to-bottom as a pipeline, marked by `# ---` section
comments in this order:

1. **Column schemas** (`TableSchema`, `HOLDINGS_SCHEMA`, `TRANSACTION_SCHEMA`,
   `PAIRED_PNL_SCHEMA`) — alias tables mapping many real-world broker column spellings
   (`"Avg. Price (₹)"`, `"Qty."`, `"Buy/Sell"`, ...) to canonical field names. **This is the
   first place to look/edit when a real export fails to parse** — add the missing header
   spelling (normalized form: lowercase, alnum tokens only, `rs`/`inr` stripped) to the
   relevant alias tuple.
2. **Low-level cleaning helpers** — `normalize_header`, `clean_text`, `to_number` (handles
   `1,234.50`, `(1,234)` accounting negatives, `₹`/`Rs.` prefixes), `_parse_one_date` (handles
   Excel serials, ISO and day-first text dates), `to_yahoo_ticker` (symbol → `SYMBOL.NS`/`.BO`).
3. **File reading & header detection** — `read_sheets` loads raw header-less grids from
   `.xlsx`/`.csv`; `extract_table` scans the first `HEADER_SCAN_ROWS` rows of every sheet
   against every candidate schema and picks the best-scoring header row. This is what lets the
   app skip client-name/date-range/disclaimer rows that brokers prepend, and what produces the
   "missing column X" `ParseError` when a file doesn't match any schema.
4. **Holdings & trade-history parsing** — `clean_holdings`, `clean_trades`,
   `_unpivot_paired_pnl` (splits a P&L statement's buy+sell-leg-per-row format into one row per
   leg), `clean_mutual_funds` (Angel One's separate Mutual Fund sheet - see below), and the
   cached entry points `parse_holdings_file`/`parse_trades_file`/`parse_mutual_funds_file`.
5. **Demo data** — `generate_demo_data` builds a synthetic portfolio from synthetic trades
   (average-cost accounting) so it's internally consistent for XIRR/CAGR.
6. **Market data** — `fetch_live_prices` (bulk `yfinance.download`, cached
   `PRICE_CACHE_TTL_SECONDS`, retries the other exchange suffix on a miss) and
   `resolve_isin_tickers` (see below), then `enrich_holdings` computes invested/current
   value, P&L, and weights.
7. **Sector & market-cap enrichment** — `fetch_sector_and_marketcap` (threaded
   `yf.Ticker(...).get_info()`, cached a day) plus the `SECTOR_FALLBACK` dict for tickers Yahoo
   doesn't classify (mostly ETFs); `classify_market_cap` buckets by ₹ crore thresholds;
   `enrich_sector_and_cap` attaches `sector`/`market_cap`/`market_cap_bucket` to holdings.
8. **Financial metrics** — `solve_xirr` (bisection-grid + Brent, Newton fallback) and
   `compute_cagr`, aggregated in `compute_portfolio_metrics`.
9. **Portfolio breakdowns (sector, market cap, tax)** — `group_breakdown` (generic
   groupby-and-aggregate shared by the sector and market-cap tabs), `compute_hhi`/
   `hhi_risk_level` (concentration risk), `split_gainers_losers`, and the FIFO tax-lot pipeline:
   `compute_tax_lots` → `value_tax_lots` → `estimate_capital_gains`.
10. **Formatting**, **Charts** (Plotly, colors validated against the project's data-viz
    accessibility standard — see below), and **UI** (Streamlit render functions - one
    `render_tab_*` per tab, wired up in `main()` at the bottom via `st.tabs()`).

### Key design decisions worth knowing before changing things

- **Symbol vs. ticker resolution is three-tiered.** `to_yahoo_ticker` first checks
  `ANGEL_ONE_SYMBOL_MAP` (condensed/loose match via `_condense` - strips everything but
  letters/digits, so `"Kotak Gold ETF"`, `"KOTAK GOLD ETF"` and `"KOTAKGOLDETF"` are the same
  key), then falls back to a naive space-stripping guess. The static map exists because some
  names never resolve any other way: Kotak Gold ETF's ISIN isn't indexed by Yahoo's ISIN search
  at all (confirmed directly against the live API - text search finds it, ISIN search doesn't),
  and Angel One's fuller "Portfolio" export spells out full company names (e.g. "Reliance
  Industr", truncated) instead of trading symbols, which the naive guess mangles. The map is
  **skipped when the row explicitly says BSE** (`"BSE" not in exchange`), so an explicit
  exchange hint always wins over the map's (NSE) ticker. When an `isin` column is present,
  `resolve_isin_tickers` (in `resolve_prices`) then upgrades both the ticker *and* the displayed
  symbol via Yahoo's own search endpoint, run in parallel via `ThreadPoolExecutor`, cached 24h -
  this still runs after the map and overrides it when it finds a better answer, so ISIN search
  > static map > naive guess. Don't add an entry to the map without verifying it against a live
  `yf.Ticker(...).get_info()` call first (a wrong static mapping silently misprices a holding,
  which is worse than the LTP-from-file fallback it would otherwise get).
- **Cache invalidation:** live prices are `@st.cache_data(ttl=300)`; the sidebar "Refresh live
  prices" button calls `fetch_live_prices.clear()`. ISIN resolution is cached for a day
  (effectively static). Parsing (`parse_holdings_file`/`parse_trades_file`) is cached on file
  bytes, so re-uploading the same file is free.
- **Data-quality notes over silent failure:** parsing functions return `(data, notes)` where
  `notes` is a list of human-readable strings (skipped rows, merged duplicates, detected
  layout). `render_data_warnings` also flags stocks priced at cost (no price found) and
  holdings that don't reconcile against the uploaded trade history (`reconcile_positions`) —
  keep this pattern when adding new parsing edge cases rather than dropping rows silently.
- **XIRR fallback:** if `solve_xirr` can't converge, `compute_portfolio_metrics` falls back to
  simple net P&L % and marks it (`xirr_method == "fallback"`); the UI shows a `*` annotation.
  Don't let a solver failure raise/crash the page.
- **Cash-flow sign convention:** BUY = negative cash flow, SELL = positive (investor's
  perspective) — used consistently in `clean_trades` and `solve_xirr` inputs.
- **Chart colors** follow the project's dataviz skill: a validated categorical palette
  (`CATEGORICAL_LIGHT`/`CATEGORICAL_DARK`, one hue per allocation slice, fixed order) plus
  fixed status colors (`GAIN_COLOR`/`LOSS_COLOR`) for P&L — don't pick new chart colors ad hoc;
  if adding a chart, validate any new palette choice the same way (see `dataviz` skill). The two
  doughnuts (stock allocation, sector allocation) share one builder, `_donut_chart`, via
  `_top_slices_with_others`, which folds anything past `ALLOCATION_SLICES` into "Others" using
  `OTHERS_COLOR` — reuse that helper for any future doughnut rather than duplicating the Plotly
  `Pie` setup.
- **Tax lots are FIFO, not average-cost.** `compute_tax_lots` walks each symbol's trades in date
  order and consumes the *earliest* remaining BUY lot(s) on every SELL (a `deque` per symbol),
  preserving each lot's own purchase date for LTCG/STCG classification (`> LTCG_HOLDING_DAYS`
  days = Long-Term). This is deliberately different from `generate_demo_data`'s average-cost
  accounting (which only needs one blended cost per symbol, not dated lots) and from
  `clean_holdings`'s quantity-weighted merge — don't conflate the three. A symbol whose sells
  exceed its recorded buys just runs out of lots silently (no negative quantities, no raise);
  pair this with `reconcile_positions` when the trade history might be incomplete.
- **Tax and concentration numbers are estimates, not advice**, and the UI says so explicitly
  (`render_tab_tax`'s caption, `render_tab_sector`'s "Uncategorized" guard). `estimate_capital_gains`
  taxes only the *positive* gain in each term bucket and never nets an STCG loss against an LTCG
  gain (that requires realizing both) — keep that asymmetry if you touch the tax math. Sector
  concentration (`hhi_risk_level`) reuses the standard DOJ/FTC HHI bands (0–10,000 scale) but
  the UI suppresses the risk *label* (not the chart) when every holding is `"Uncategorized"`,
  since a degenerate single-bucket HHI there would misrepresent missing data as concentration
  risk.
- **Mutual funds are a separate, optional table found in the same upload, not a second file.**
  Angel One's fuller "Portfolio" export is a multi-sheet workbook; `parse_mutual_funds_file`
  looks for a Mutual Fund sheet in whatever file was uploaded as the *holdings* report and
  returns an empty (never `None`) DataFrame when there isn't one - that's the normal case for a
  plain equity holdings file, so it's silent, not a warning. `load_portfolio` now returns a
  4-tuple (`holdings, trades, mutual_funds, notes`) - update all call sites if that signature
  changes again. There's no live-price step for funds (Yahoo Finance doesn't carry Indian MF
  NAVs); the file's own NAV/value is used as-is via `enrich_mutual_funds`, and the per-fund XIRR
  Angel One already computes is kept rather than recomputed (a snapshot has no cash-flow
  history to recompute it from) - `compute_mutual_fund_metrics`'s portfolio-level XIRR is only
  an invested-value-weighted average of those, not a true portfolio XIRR. When merging duplicate
  fund rows, `invested_value` is **summed from the file's own column**, not recomputed as
  `units * average_nav` - the NAV column is rounded to 2dp for display, so recomputing from it
  loses precision the file's own invested_value doesn't have (this was a real bug caught during
  development; don't reintroduce it elsewhere in the file).
- **Header date suffixes are stripped generically.** `normalize_header` (via
  `_strip_trailing_date`) drops a trailing run of numeric tokens from any header when it
  contains a 4-digit (year-like) token and isn't the whole header - e.g. `"NAV as on
  2026-09-08"` → `"nav as on"`. This exists because Angel One bakes the download date into two
  Mutual Fund column headers, which would otherwise need a new literal alias every single day.
  It's applied to every schema, not just mutual funds, so a new alias with a genuinely meaningful
  trailing 4-digit number would also get stripped - none currently exist, but keep this in mind
  before adding one.

## Codex config detected

An OpenAI Codex config exists at `~/.codex/config.toml` (user-level, outside this repo). If you
want to import anything from it (MCP servers, instructions, etc.) into Claude Code, reply
`/import` to scan it and list what's importable, then `/import --yes=<digest>` to apply.
