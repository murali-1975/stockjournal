"""
Stock Journal - Main Orchestrator
==================================

Entry point for the Stock Journal trade processing engine.

Usage:
    python main.py                   Run the full trade processing pipeline once
    python main.py --test            Run all tests before processing
    python main.py --update          Only update LTP, EMA, and Dashboard (skips loading new trades)
    python main.py --watch <mins>    Run in a continuous loop, updating LTP/EMA every N minutes
    python main.py --gsheet <name>   Export results to a Google Sheet (by name or ID)

This script:
    1. Loads the configuration from `input.cfg`.
    2. Reads any existing master data from `Transformed_Tradebook.xlsx`.
    3. Loads new trades from `Tradebook Template.xlsx`.
    4. Merges new trades into the master database (with deduplication).
    5. Processes grouped trades with Tranche/Cheat classification.
    6. Calculates portfolio holdings, PnL, and dynamic Stop Loss levels.
    7. Saves all results back to `Transformed_Tradebook.xlsx` with proper
       Excel formatting, preserving any user-created custom sheets.

File Structure:
    main.py              - This orchestrator script
    src/
        __init__.py      - Package docstring
        config.py        - Configuration file parser
        data_io.py       - Excel I/O and data merging
        market_api.py    - Yahoo Finance data fetcher
        calculations.py  - Trade grouping, PnL, and Stop Loss math
        excel_writer.py  - Excel workbook formatting and export
    tests/
        test_config.py        - Config parser tests
        test_data_io.py       - Data I/O and dedup tests
        test_calculations.py  - Tranche, PnL, and SL tests
        test_excel_writer.py  - Excel formatting tests
    input.cfg            - User-defined portfolio configuration
"""

import os
import sys
import warnings

# Suppress harmless openpyxl warning about unsupported Data Validation extensions
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from src.config import load_config
from src.data_io import load_data, load_master_database, merge_and_deduplicate
from src.calculations import process_grouped_trades, calculate_portfolios
from src.excel_writer import save_workbook
from src.corporate_actions import process_splits
from src.gsheet_writer import save_to_gsheet
from src.gsheet_dashboard import build_gsheet_dashboard


def run_tests() -> bool:
    """
    Discovers and runs all test cases in the tests/ directory.

    Returns:
        True if all tests passed, False otherwise.
    """
    import unittest

    print("=" * 60)
    print("  RUNNING TEST SUITE")
    print("=" * 60)

    loader = unittest.TestLoader()
    suite = loader.discover('tests', pattern='test_*.py')
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    print()
    if result.wasSuccessful():
        print("[SUCCESS] ALL TESTS PASSED!")
    else:
        print("[FAIL] SOME TESTS FAILED!")
        print(f"   Failures: {len(result.failures)}")
        print(f"   Errors:   {len(result.errors)}")

    print("=" * 60)
    return result.wasSuccessful()


def ensure_files_accessible(filepaths: list[str], mode: str = 'a') -> None:
    """
    Checks if files are accessible in the given mode ('a' for write/append, 'r' for read).
    If they are locked (e.g., open in Excel), prompts the user interactively
    or waits/retries in non-interactive mode.
    """
    import time
    import sys
    
    for filepath in filepaths:
        if not os.path.exists(filepath):
            continue
            
        attempt = 0
        while True:
            try:
                if mode == 'a':
                    # Try opening in append mode to check write lock
                    with open(filepath, 'a'):
                        pass
                else:
                    # Try opening in read-binary mode to check read lock
                    with open(filepath, 'rb'):
                        pass
                break # Successfully verified accessibility, check next file
            except PermissionError:
                action_word = "open or modify" if mode == 'a' else "read"
                if sys.stdin and sys.stdin.isatty():
                    print(f"\n⚠️  [FILE LOCKED] '{os.path.basename(filepath)}' is currently open or locked in Excel.")
                    print(f"   Please save and close the Excel file so the program can {action_word} it, then press Enter to retry...")
                    try:
                        input()
                    except KeyboardInterrupt:
                        print("\nAborted by user.")
                        sys.exit(1)
                else:
                    attempt += 1
                    if attempt % 6 == 1: # Print every 30 seconds
                        print(f"⚠️  [FILE LOCKED] Waiting for '{os.path.basename(filepath)}' to be released (locked by another process)...")
                    time.sleep(5)


def main(update_only: bool = False, gsheet_target: str = None):
    """
    Main execution pipeline.

    Orchestrates the trade processing workflow.
    If update_only is True, skips loading/merging new trades from the template
    and only refreshes market data (LTP, EMAs) and the dashboard.
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(current_dir, 'Tradebook Template.xlsx')
    output_path = os.path.join(current_dir, 'Transformed_Tradebook.xlsx')
    config_path = os.path.join(current_dir, 'input.cfg')

    # Verify that files are accessible before performing calculations
    if not update_only:
        ensure_files_accessible([input_path], mode='r')
    ensure_files_accessible([output_path], mode='a')

    # Step 1: Load configuration
    config = load_config(config_path)

    # Step 2: Load existing master database
    master_df = load_master_database(output_path)

    if update_only:
        print("LTP/EMA Update Mode: Skipping new trade ingestion.")
        df = master_df
    else:
        # Step 3: Load new trades from the input template
        new_df = load_data(input_path)
        if new_df is None:
            print(f"\n❌ ERROR: Failed to load trade data from '{os.path.basename(input_path)}'.")
            print("Please ensure the template is not corrupted or locked, and close Excel if it is open.\n")
            sys.exit(1)

        # Step 4: Merge and deduplicate
        df = merge_and_deduplicate(master_df, new_df)
    
    if df is None or df.empty:
        print("No trade data found to process.")
        return

    # Step 4.5: Check for auto-adjustments (Stock Splits/Bonuses)
    if not df.empty:
        was_modified, df = process_splits(df, base_dir=current_dir)
        if was_modified:
            print("Historical trades adjusted in-memory. Will be saved to Excel.")

    # Step 5: Process grouped trades with Tranche/Cheat labels
    print("Processing data...")
    grouped_df = process_grouped_trades(df, config)

    # Step 6: Calculate portfolios, PnL, and Stop Loss
    from src.data_io import load_price_updates
    price_updates = load_price_updates(output_path)
    portfolio_df, overall_df = calculate_portfolios(df, grouped_df, config, price_updates)

    # Step 6.5: Fetch Benchmark Returns
    start_date_str = config.get('INVEST_START_DATE', '')
    custom_benchmarks = config.get('BENCHMARK_INDEX', None)
    from src.market_api import fetch_benchmark_returns
    benchmark_returns = fetch_benchmark_returns(start_date_str, custom_benchmarks) if start_date_str else None

    # Step 7: Save everything to Excel
    save_workbook(df, grouped_df, portfolio_df, overall_df, output_path, benchmark_returns)

    # Step 8: Save to Google Sheets (if requested)
    if gsheet_target:
        data_to_export = {
            "Raw_Tradebook": df,
            "Transaction": grouped_df,
            "Current_Portfolio": portfolio_df,
            "Overall_Portfolio": overall_df
        }
        print(f"Exporting to Google Sheets: {gsheet_target}...")
        
        # 1. Export Data Sheets
        save_to_gsheet(data_to_export, gsheet_target)
        
        # 2. Build the Styled Dashboard
        import gspread
        from google.oauth2.service_account import Credentials
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
        client = gspread.authorize(creds)
        sh = client.open(gsheet_target)
        build_gsheet_dashboard(sh, portfolio_df, overall_df, benchmark_returns)


if __name__ == "__main__":
    if '--test' in sys.argv:
        success = run_tests()
        if not success:
            print("\n⚠️  Tests failed. Fix issues before running the actual processing.")
            sys.exit(1)

        # If --test is the only flag, exit after tests
        if len(sys.argv) == 2:
            print("\nTests complete. Run without --test to process trades.")
            sys.exit(0)

    if '--recommend' in sys.argv:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.join(current_dir, 'Transformed_Tradebook.xlsx')
        from src.actions import execute_recommendation_bypass

        # Parse filter
        idx = sys.argv.index('--recommend')
        rec_filters = []
        i = idx + 1
        while i < len(sys.argv) and not sys.argv[i].startswith('-'):
            parts = [p.strip().lower() for p in sys.argv[i].split(',')]
            for p in parts:
                if p:
                    rec_filters.append(p)
            i += 1

        valid_options = {'all', 'add', 'buy', 'sell'}
        cleaned_filters = []
        for f in rec_filters:
            if f in valid_options:
                cleaned_filters.append(f)
            else:
                print(f"⚠️ Invalid recommendation filter '{f}' ignored. Valid values: all, add, buy, sell.")
        
        if not cleaned_filters:
            rec_filter = 'all'
        elif 'all' in cleaned_filters:
            rec_filter = 'all'
        else:
            rec_filter = ','.join(cleaned_filters)

        # Verify that the output workbook is writable before performing recommendations
        ensure_files_accessible([output_path], mode='a')

        try:
            execute_recommendation_bypass(output_path, rec_filter=rec_filter)
        except Exception as e:
            print(f"❌ Error running recommendations: {e}")
            sys.exit(1)
        sys.exit(0)

    if '--update' in sys.argv:
        main(update_only=True)
    elif '--watch' in sys.argv:
        try:
            idx = sys.argv.index('--watch')
            interval_minutes = float(sys.argv[idx + 1])
        except (ValueError, IndexError):
            print("Usage: python main.py --watch <minutes>")
            sys.exit(1)

        import time
        from datetime import datetime
        print(f"\n👀 Starting Watch Mode: Updating every {interval_minutes} minutes.")
        print("Press Ctrl+C to stop.\n")

        while True:
            try:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Starting update cycle...")
                main(update_only=True)
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Update complete. Waiting {interval_minutes} minutes...")
            except PermissionError as e:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] ⚠️ Permission Error: Could not save the Excel file. Is it open? Please close it to allow updates.")
            except Exception as e:
                import traceback
                print(f"[{datetime.now().strftime('%H:%M:%S')}] ❌ Error during update: {e}")
                traceback.print_exc()

            try:
                time.sleep(interval_minutes * 60)
            except KeyboardInterrupt:
                print("\nWatch mode stopped.")
                sys.exit(0)
    else:
        gsheet_target = None
        if '--gsheet' in sys.argv:
            try:
                idx = sys.argv.index('--gsheet')
                gsheet_target = sys.argv[idx + 1]
            except IndexError:
                print("Usage: python main.py --gsheet <sheet_name_or_id>")
                sys.exit(1)
        
        main(gsheet_target=gsheet_target)

