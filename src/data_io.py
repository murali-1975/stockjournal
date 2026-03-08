"""
Data I/O Module
===============

Handles all Excel file reading and data merging operations. This module is
responsible for:
    - Loading trade data from Excel tradebook templates
    - Reading the existing Raw_Tradebook master database
    - Appending new trades and deduplicating entries
    - Safely coercing ID columns (Order ID, Trade ID) to prevent
      scientific notation corruption in large numeric fields

All pandas read_excel calls route through this module to ensure
consistent dtype handling across the application.
"""

import os
import pandas as pd


def load_data(input_file: str) -> pd.DataFrame | None:
    """
    Loads trade data from an Excel file into a pandas DataFrame.

    Automatically detects columns containing 'ID' in their name and reads
    them as strings to prevent pandas from converting large integers
    (e.g., 16-digit Order IDs) into lossy float64 values.

    Args:
        input_file: Path to the input Excel file (.xlsx).

    Returns:
        A DataFrame containing the trade data, or None if the file is
        missing, unreadable, or lacks required columns.

    Required columns:
        Trade Date, Symbol, Trade Type, Quantity, Price
    """
    print(f"Reading {input_file}...")
    try:
        # Detect ID columns and force them to string to prevent scientific notation
        df_cols = pd.read_excel(input_file, nrows=0).columns
        dtypes = {col: str for col in df_cols if 'ID' in col.upper()}
        df = pd.read_excel(input_file, dtype=dtypes)
    except FileNotFoundError:
        print(f"Error: {input_file} not found.")
        return None
    except Exception as e:
        print(f"Error reading {input_file}: {e}")
        return None

    # Validate required columns
    required_cols = ['Trade Date', 'Symbol', 'Trade Type', 'Quantity', 'Price']
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        print(f"Error: The input file is missing required columns: {missing_cols}")
        return None

    print("Data loaded successfully.")
    return df


def load_master_database(output_file: str) -> pd.DataFrame | None:
    """
    Loads the existing Raw_Tradebook sheet from the master database file.

    This is the persistent store of all historical trades. If the file
    or the 'Raw_Tradebook' sheet does not exist, returns None so the
    caller can fall back to initial bootstrapping.

    Args:
        output_file: Path to the Transformed_Tradebook.xlsx master file.

    Returns:
        A DataFrame of existing raw trades, or None if unavailable.
    """
    if not os.path.exists(output_file):
        return None

    try:
        df_cols = pd.read_excel(output_file, sheet_name='Raw_Tradebook', nrows=0).columns
        dtypes = {col: str for col in df_cols if 'ID' in col.upper()}
        master_df = pd.read_excel(output_file, sheet_name='Raw_Tradebook', dtype=dtypes)
        print(f"Loaded existing master database from {output_file} ('Raw_Tradebook' sheet).")
        return master_df
    except Exception:
        print(f"No existing 'Raw_Tradebook' sheet found in {output_file}. Starting fresh.")
        return None


def merge_and_deduplicate(master_df: pd.DataFrame | None, new_df: pd.DataFrame | None) -> pd.DataFrame | None:
    """
    Merges existing master data with newly provided trades and removes duplicates.

    Deduplication is based on the 'Trade ID' column, which is a unique
    exchange-assigned identifier for each trade. This prevents double-counting
    if the same inbox file is accidentally processed twice.

    Args:
        master_df: The existing Raw_Tradebook DataFrame (can be None).
        new_df:    The newly loaded trades DataFrame (can be None).

    Returns:
        A combined, deduplicated DataFrame, or None if both inputs are None.
    """
    if master_df is None and new_df is None:
        print("No valid data available to process. Exiting.")
        return None

    if master_df is not None and new_df is not None:
        print("Appending new trades to the master database...")
        df = pd.concat([master_df, new_df], ignore_index=True)
        before_len = len(df)
        if 'Trade ID' in df.columns:
            df = df.drop_duplicates(subset=['Trade ID'])
        else:
            df = df.drop_duplicates()
        after_len = len(df)
        if before_len > after_len:
            print(f"Dropped {before_len - after_len} duplicate trades.")
        return df
    elif new_df is not None:
        return new_df
    else:
        return master_df


def load_equity_master(config: dict) -> pd.DataFrame | None:
    """
    Loads the Equity Master reference sheet and returns a lookup DataFrame.

    Reads the 'Equity Master' sheet from the file specified in config['EQUITY_MASTER'],
    strips whitespace from column headers, and returns a DataFrame with columns:
        - Symbol (renamed from 'Stock Symbol')
        - TF_Sector (from 'TF Sector Classification')
        - TF_Classification (from 'TF Stock Classfication')

    Args:
        config: The application configuration dict. Must contain 'EQUITY_MASTER' key.

    Returns:
        A DataFrame with Symbol/TF_Sector/TF_Classification, or None if unavailable.
    """
    equity_file = config.get('EQUITY_MASTER', '')
    if not equity_file or not os.path.exists(equity_file):
        print(f"Equity Master file '{equity_file}' not found. Skipping TF columns.")
        return None

    try:
        df = pd.read_excel(equity_file, sheet_name='Equity Master')
        df.columns = df.columns.str.strip()

        required = ['Stock Symbol', 'TF Sector Classification', 'TF Stock Classfication']
        missing = [c for c in required if c not in df.columns]
        if missing:
            print(f"Equity Master missing columns: {missing}. Skipping TF columns.")
            return None

        result = df[['Stock Symbol', 'TF Sector Classification', 'TF Stock Classfication']].copy()
        result.columns = ['Symbol', 'TF_Sector', 'TF_Classification']
        result['Symbol'] = result['Symbol'].str.strip()
        return result
    except Exception as e:
        print(f"Error loading Equity Master: {e}")
        return None
