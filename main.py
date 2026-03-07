"""
Stock Journal - Main Orchestrator
==================================

Entry point for the Stock Journal trade processing engine.

Usage:
    python main.py

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
    input.cfg            - User-defined portfolio configuration
"""

import os

from src.config import load_config
from src.data_io import load_data, load_master_database, merge_and_deduplicate
from src.calculations import process_grouped_trades, calculate_portfolios
from src.excel_writer import save_workbook


def main():
    """
    Main execution pipeline.

    Orchestrates the full trade processing workflow from data ingestion
    through to formatted Excel output.
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(current_dir, 'Tradebook Template.xlsx')
    output_path = os.path.join(current_dir, 'Transformed_Tradebook.xlsx')
    config_path = os.path.join(current_dir, 'input.cfg')

    # Step 1: Load configuration
    config = load_config(config_path)

    # Step 2: Load existing master database (if any)
    master_df = load_master_database(output_path)

    # Step 3: Load new trades from the input template
    new_df = load_data(input_path)

    # Step 4: Merge and deduplicate
    df = merge_and_deduplicate(master_df, new_df)
    if df is None:
        return

    # Step 5: Process grouped trades with Tranche/Cheat labels
    print("Processing data...")
    grouped_df = process_grouped_trades(df, config)

    # Step 6: Calculate portfolios, PnL, and Stop Loss
    portfolio_df, overall_df = calculate_portfolios(df, grouped_df)

    # Step 7: Save everything to Excel
    save_workbook(df, grouped_df, portfolio_df, overall_df, output_path)


if __name__ == "__main__":
    main()
