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
        # Detect sheets and prioritize 'Equity'
        with pd.ExcelFile(input_file) as xls:
            sheet_names = xls.sheet_names
            sheet_to_load = 'Equity' if 'Equity' in sheet_names else sheet_names[0]
            
            # Detect ID columns and force them to string to prevent scientific notation
            df_cols = pd.read_excel(xls, sheet_name=sheet_to_load, nrows=0).columns
            dtypes = {col: str for col in df_cols if 'ID' in col.upper()}
            df = pd.read_excel(xls, sheet_name=sheet_to_load, dtype=dtypes)
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
        with pd.ExcelFile(output_file) as xls:
            df_cols = pd.read_excel(xls, sheet_name='Raw_Tradebook', nrows=0).columns
            dtypes = {col: str for col in df_cols if 'ID' in col.upper()}
            master_df = pd.read_excel(xls, sheet_name='Raw_Tradebook', dtype=dtypes)
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

        required = ['Stock Id', 'TF Sector Classification', 'TF Stock Classfication']
        missing = [c for c in required if c not in df.columns]
        if missing:
            print(f"Equity Master missing columns: {missing}. Skipping TF columns.")
            return None

        result = df[['Stock Id', 'TF Sector Classification', 'TF Stock Classfication']].copy()
        result.columns = ['Symbol', 'TF_Sector', 'TF_Classification']
        result['Symbol'] = result['Symbol'].str.strip().str.upper()
        result = result.drop_duplicates(subset=['Symbol'])
        return result
    except PermissionError:
        print(f"Error loading Equity Master: Permission denied. Please close '{equity_file}' if it is open in Excel and try again.")
        return None
    except Exception as e:
        print(f"Error loading Equity Master: {e}")
        return None


def load_price_updates(output_file: str) -> dict:
    """
    Loads custom price updates from the 'Price_Update' sheet of the master workbook.
    Looks for the 'Symbol' (or Stock Symbol) column and the 'LTP' column, regardless of their position.
    Scans the first 10 rows to find the actual header row.
    """
    import os
    if not os.path.exists(output_file):
        return {}
    try:
        with pd.ExcelFile(output_file) as xls:
            # Load without headers first to find the header row
            df_raw = pd.read_excel(xls, sheet_name='Price_Update', header=None)
            
            header_row_idx = -1
            for idx, row in df_raw.head(10).iterrows():
                row_vals = [str(val).strip().upper() for val in row.values]
                if any(h in row_vals for h in ['SYMBOL', 'STOCK SYMBOL', 'STOCK_SYMBOL']):
                    header_row_idx = idx
                    break
                    
            if header_row_idx == -1:
                print("Warning: 'Price_Update' sheet is missing 'Symbol' column. Skipping local price updates.")
                return {}
                
            # Reload with the correct header
            df = pd.read_excel(xls, sheet_name='Price_Update', header=header_row_idx)
            df.columns = df.columns.astype(str).str.strip()
            
            # Find column matching Symbol
            symbol_col = None
            for col in df.columns:
                if col.upper() in ['SYMBOL', 'STOCK SYMBOL', 'STOCK_SYMBOL']:
                    symbol_col = col
                    break
            
            # Find column matching LTP
            ltp_col = None
            for col in df.columns:
                if col.upper() in ['LTP', 'LAST TRADED PRICE', 'LAST_TRADED_PRICE', 'PRICE']:
                    ltp_col = col
                    break
                    
            if not symbol_col or not ltp_col:
                print("Warning: 'Price_Update' sheet is missing 'LTP' column. Skipping local price updates.")
                return {}
                
            # Build dictionary
            price_dict = {}
            for _, row in df.iterrows():
                sym = str(row[symbol_col]).strip().upper()
                val = row[ltp_col]
                try:
                    price_dict[sym] = float(val)
                except (ValueError, TypeError):
                    pass # Skip invalid prices
        print(f"Loaded {len(price_dict)} price updates from 'Price_Update' sheet.")
        return price_dict
    except Exception:
        # Sheet not found or error loading, return empty dict silently
        return {}
