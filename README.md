# Stock Journal

A powerful Python-based trading journal and portfolio management system designed to process Excel tradebooks, fetch live market data from Yahoo Finance, and generate comprehensive dashboards and analytical reports.

## Features

- **Automated Data Merging**: Appends new trades from a standard template seamlessly into your master database with smart duplicate-checking.
- **Grouped Trade Analysis (Tranches/Cheats)**: Automatically groups trades of the same symbol to assign dynamic "Tranche" levels based on cumulative investment size.
- **Live Market Data Integration**: Fetches real-time Last Traded Price (LTP) and Exponential Moving Averages (EMA 9, 10, 11, 21) dynamically from Yahoo Finance.
- **Benchmark Returns Tracking**: Extracts an `INVEST_START_DATE` and calculates comparative returns against customizable indices like Nifty 50, Nifty Midcap, and Nifty Smallcap.
- **Dynamic Stop-Loss Calculation**: Computes trailing and average-price based Stop Loss values dependent on the stock's current Tranche level.
- **Smart Categorization**: Classifies stocks by Sector, Market Cap, and as "Core or Satellite" based on `.cfg` and reference Excel files automatically.
- **Corporate Action Handling**: Automatically detects stock splits and bonuses (post-purchase) from Yahoo Finance, prompting the user to auto-adjust historical trade data directly.
- **Excel Dashboard Generation**: Outputs deeply formatted, Excel-native dashboards including Allocation breakdowns (Sector, Market Cap, Core/Satellite), KPI Summaries, top gainers/losers, and PnL metrics.
- **Google Sheets Integration & Styled Dashboard**: Authorizes with Google APIs using service account credentials (`credentials.json`) to export all transactional tables and build a premium, highly formatted Google Sheets dashboard (featuring harmonious HSL-based palettes).
- **Automated Telegram Scanner**:
  - **Headless & Persistent Session**: Relies on a persistent Telethon session string (`SESSION_STR`) to bypass serverless-ready two-factor authentication challenges.
  - **Fuzzy/Smart Matcher**: Connects to the Google Sheet Core/Satellite watchlists and extracts signals using flexible, regex-based symbol and company name pattern matching.
  - **Context-Aware Thread Reconstruction**: Automatically trace-merges thread replies with parent messages to keep the discussion readable and intact.
  - **Multi-Topic & Multi-Channel Support**: Navigates forum categories (e.g., the Technofunda community) and independent announcement channels systematically to bypass pinned/hidden thread limitations.

---

## File Structure

```text
stockjournal/
├── main.py                     # Main tradebook CLI orchestrator
├── input.cfg                   # Portfolio limits, tranche sizing, and configurations
├── applied_splits.json         # History of processed splits (prevents duplicate adjustments)
├── Tradebook Template.xlsx     # Inputs for broker trades
├── Transformed_Tradebook.xlsx  # Local Excel output file
├── Equity_Master.xlsx          # Reference data for Sector and Core/Satellite categorization
├── credentials.json            # Google Service Account credentials (required for Google Sheets)
├── src/                        # Codebase source directory
│   ├── calculations.py         # Tranche sizing, P&L, and Stop-Loss engines
│   ├── config.py               # Config parsing logic
│   ├── corporate_actions.py    # Yahoo Finance split/bonus adjuster
│   ├── dashboard.py            # Local Excel dashboard layout & styling
│   ├── excel_writer.py         # Openpyxl utility to save formatted Workbooks
│   ├── gsheet_writer.py        # Exporter for raw data tables to Google Sheets
│   ├── gsheet_dashboard.py     # Custom styling and dashboard builder on Google Sheets
│   ├── market_api.py           # Yahoo Finance live fetching client
│   └── telegram_analyzer/      # Telegram monitoring sub-module
│       ├── scanner.py          # Main community topics/channels scanner
│       ├── generate_session.py # Helper tool to authorize & generate Telethon strings
│       ├── surgical_scan.py    # Targeted scan script
│       ├── check_telethon.py   # Telethon connection validation script
│       └── debug_access.py     # Diagnostic tool for checking GSpread access
└── tests/                      # Automated test suite
```

---

## Command Line Interface (CLI) Modes

To run the application, open your terminal in the project directory and execute your desired command.

### 1. Standard Processing Mode
```bash
.\.venv\Scripts\python.exe main.py
```
**What it does:** Reads configurations, merges new trades, prompts for corporate action adjustments, fetches live prices, and saves local Excel outputs.

### 2. Google Sheets Direct Export
```bash
.\.venv\Scripts\python.exe main.py --gsheet "YOUR_SPREADSHEET_NAME_OR_ID"
```
**What it does:** Performs standard tradebook processing, uploads the raw tables to your Google Sheet, and formats/styles a modern portfolio dashboard on Google Sheets automatically.

### 3. Continuous Watch Mode
```bash
.\.venv\Scripts\python.exe main.py --watch 5
```
*(Replace `5` with your desired interval in minutes)*

**What it does:** Puts the engine into a background loop, pulling the newest market prices and EMAs every `N` minutes, updating your tradebook without closing.

### 4. Telegram Stock Discussion Scanner
```bash
.\.venv\Scripts\python.exe src/telegram_analyzer/scanner.py
```
**What it does:** Scans selected Telegram forum topics and channels, identifies mentioned stocks from your watchlists, and appends the annotated message logs directly to your Google Sheet.

* **Scan all messages (Skip Watchlist Filter):**
  ```bash
  .\.venv\Scripts\python.exe src/telegram_analyzer/scanner.py --dump-all
  ```
* **Customize Lookback Window (in days):**
  ```bash
  .\.venv\Scripts\python.exe src/telegram_analyzer/scanner.py --days 3
  ```

### 5. Persistent Telegram Session Generator
```bash
.\.venv\Scripts\python.exe src/telegram_analyzer/generate_session.py
```
**What it does:** Authenticates once with your Telegram app, requesting an OTP, and generates a permanent `SESSION_STR` to be stored inside `scanner.py` or cloud environment variables.

### 6. Run Test Suite
```bash
.\.venv\Scripts\python.exe main.py --test
```
**What it does:** Executes the automated Python `unittest` suite (validating trade math, configurations, split handling, and integrations).

---

## Troubleshooting

- **`PermissionError` when saving:** Ensure `Transformed_Tradebook.xlsx` is closed before running processing scripts.
- **Google Sheets API Authorization Errors:** Ensure you have placed a valid `credentials.json` service account file in the root directory and shared your target Google Sheet with the Service Account email.
- **Telegram Connection Issues:** If your Telethon session expires, run `generate_session.py` again to get a fresh `SESSION_STR`.
- **Missing Yahoo Finance Data:** Ensure all custom indices or ticker symbols end with the appropriate exchange suffix (e.g. `.NS` for NSE, `.BO` for BSE).
