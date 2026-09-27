# Indian Equity Portfolio Tracker

A Streamlit dashboard for Angel One traders. Upload your holdings export (and, optionally,
your trade history / P&L export) and it enriches your positions with live NSE/BSE prices from
Yahoo Finance, then reports:

- Total invested capital, current value, net P&L, XIRR and CAGR
- Sector allocation with a Herfindahl-Hirschman concentration-risk score
- Market cap breakdown (Large / Mid / Small Cap)
- An estimated LTCG/STCG tax exposure on unrealised gains, with a per-lot holding-period table
- Mutual fund holdings, if your export includes them
- A searchable, filterable holdings table

A **Demo / Synthetic Data Mode** is built in, so you can explore the whole dashboard without
uploading anything.

## Setup

Requires Python 3.10+.

```powershell
git clone https://github.com/arun0404/indian-equity-portfolio-tracker.git
cd indian-equity-portfolio-tracker

python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

(On macOS/Linux, activate with `source .venv/bin/activate` instead.)

## Running it

```powershell
streamlit run app.py
```

This opens the dashboard in your browser at `http://localhost:8501`. From there:

- Toggle **Demo / Synthetic Data Mode** in the sidebar to explore it immediately, or
- Upload your **Angel One Holdings Report** (`.xlsx`/`.csv`) — required
- Optionally upload your **Trade History / P&L Report** (`.xlsx`/`.csv`) to unlock XIRR, CAGR
  and the tax/holding-period breakdown

Live prices are fetched from Yahoo Finance and cached for 5 minutes; use the sidebar's
**Refresh live prices** button to force an update.

## What files are supported

- **Holdings report**: needs a symbol/stock name, quantity and average price column; LTP and
  ISIN are optional but improve accuracy (see below).
- **Trade history / P&L report**: either a trade book (date, symbol, buy/sell, quantity, price)
  or a P&L statement (buy/sell dates and prices per position).
- **Mutual funds**: no separate upload needed — if your holdings file is Angel One's fuller
  "Portfolio" export (a multi-sheet workbook with a Mutual Fund sheet alongside Equity), it's
  detected automatically.

Header rows are detected automatically, so client details, date ranges and disclaimers above or
below the actual table don't need to be trimmed out first.

## Notes

- Nothing you upload leaves your machine except the ticker symbols sent to Yahoo Finance to
  fetch prices — no holdings data is transmitted anywhere else.
- Tax and concentration-risk figures are estimates for information only, not financial or tax
  advice.
- `*.xlsx` and `*.csv` files are gitignored in this repo so a real export never gets committed
  by accident.

## For contributors

If you're working on this project with Claude Code, a `CLAUDE.md` with the codebase
architecture and key design decisions may exist locally — it's gitignored, so it isn't part of
this repo and won't come through on a fresh clone.
