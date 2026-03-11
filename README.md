# Stock Journal

A powerful Python-based trading journal and portfolio management system designed to process Excel tradebooks, fetch live market data from Yahoo Finance, and generate comprehensive dashboards and analytical reports.

## Features

- **Automated Data Merging**: Appends new trades from a standard template seamlessly into your master database without duplicates.
- **Grouped Trade Analysis (Tranches/Cheats)**: Smartly groups trades of the same symbol to incrementally assign logical "Tranche" levels based on cumulative investment size.
- **Live Market Data Integration**: Fetches real-time Last Traded Price (LTP) and Exponential Moving Averages (EMA 9, 10, 11, 21) dynamically from Yahoo Finance.
- **Dynamic Stop-Loss Calculation**: Computes trailing and average-price based Stop Loss values dependent on the stock's current Tranche level.
- **Corporate Action Handling**: Automatically detects stock splits and bonuses (post-purchase) from Yahoo Finance and interacts with the user to auto-adjust historical trade data directly.
- **Excel Dashboard Generation**: Outputs deeply formatted, Excel-native dashboards including Allocation breakdowns (Sector, Market Cap, Core/Satellite), KPI Summaries, top gainers/losers, and PnL metrics.
- **Parameterized Live Polling**: Can run as a continuous background process, periodically updating your Excel file with the freshest prices and EMAs.

---

## File Structure

```text
stockjournal/
├── main.py                     # The main orchestrator CLI application
├── input.cfg                   # User configuration (portfolio sizing, caps, rules)
├── applied_splits.json         # Tracks already-processed stock splits to prevent double adjustments
├── Tradebook Template.xlsx     # Drop your raw broker trades in here
├── Transformed_Tradebook.xlsx  # The final output master file (created automatically)
├── Equity_Master.xlsx          # (Optional) Reference data for Sector and Core/Satellite classifications
├── src/                        # Core application modules
│   ├── calculations.py         # Math logic (PnL, Tranche grouping, SL)
│   ├── config.py               # Config parsing logic
│   ├── corporate_actions.py    # Auto-adjustment logic for splits and bonuses
│   ├── dashboard.py            # Dashboard worksheet layout and data summary logic
│   ├── data_io.py              # Excel reading, merging, and deduping logic
│   ├── excel_writer.py         # Final openpyxl workbook formatting and saving
│   └── market_api.py           # Yahoo Finance API connection
└── tests/                      # Automated test suite
```

---

## Command Line Interface (CLI) Modes

To run the application, open your terminal (Command Prompt, PowerShell, or Git Bash) in the project directory and use one of the following commands:

### 1. Standard Processing Mode
```bash
python main.py
```
**What it does:** 
Reads your configuration, merges any new trades from the `Tradebook Template.xlsx`, checks for stock splits (prompting you if found), fetches the latest market prices, calculates all PnL and Stop Loss values, and generates/updates the `Transformed_Tradebook.xlsx` file. It then gracefully exits.

### 2. Live "Watch" Updates
```bash
python main.py --watch 5
```
*(Replace `5` with your desired interval in minutes)*

**What it does:** 
Runs the Standard Processing flow immediately, but instead of exiting, it puts the program into a continuous background loop. Every `N` minutes, it wakes up, queries Yahoo Finance for the latest LTP and EMA data, recalculates your current portfolio, and transparently updates your `Transformed_Tradebook.xlsx`. 
* **Note:** If you have the Excel file currently open while it tries to save, it will print a friendly warning and pause until the next cycle. Press `Ctrl+C` in the terminal to cancel the loop.

### 3. Test Mode
```bash
python main.py --test
```
**What it does:** 
Runs the internal automated Python `unittest` suite (testing grouping logic, config parsing, math, and Excel generation) without affecting your actual data files. Very useful if you decide to modify the code in the future and want to ensure you didn't break anything.

---

## Configuration (`input.cfg`)

The `input.cfg` file allows you to define your core portfolio values and logical limits without touching the Python code:

- `TOTAL_PORTFOLIO`: Base amount used for percentage calculations.
- `TRANCH`: Denotes the size of a standard investment bracket.
- `TRANCH_TOLERANCE`: Allowed drift for a Tranche size.
- `CHEAT`: Maximum size for smaller fractional trades before they roll into a Tranch.
- `MAX_POSITION_SIZE`: A safeguard limit per symbol.
- `SMALL_CAP / MEDIUM_CAP / LARGE_CAP`: Market cap categorization definitions (in ₹ Crores).
- `EQUITY_MASTER`: The internal filename for your Sector mapping tracker.

---

## Troubleshooting

- **`PermissionError` when saving:** Ensure `Transformed_Tradebook.xlsx` is closed before running `python main.py`.
- **Missing LTP or `#N/A`**: Ensure your symbols in the Excel file have the `.NS` or `.BO` suffix required by Yahoo Finance.
- **Stock Split Prompt not appearing**: The program only prompts you for splits that happened *after* your very first buy date for that stock, and prevents you from applying the same split twice via `applied_splits.json`.
