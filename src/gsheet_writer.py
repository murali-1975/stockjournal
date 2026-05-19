"""
Stock Journal - Google Sheets Writer
====================================

Handles exporting DataFrames to Google Sheets using service account authentication.
"""

import gspread
import pandas as pd
import numpy as np
from google.oauth2.service_account import Credentials
import os
from gspread_formatting import *

def format_value(val):
    """Formats a value for Google Sheets compatibility."""
    if pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp, np.datetime64)):
        return str(val)
    return val

def save_to_gsheet(data_frames, sheet_name_or_id, credentials_path="credentials.json"):
    """
    Exports a dictionary of DataFrames to a Google Sheet.
    
    Args:
        data_frames (dict): Dictionary mapping tab names to pandas DataFrames.
        sheet_name_or_id (str): The name or ID of the Google Sheet.
        credentials_path (str): Path to the service account JSON file.
    """
    if not os.path.exists(credentials_path):
        print(f"Error: Credentials file not found at {credentials_path}")
        return

    # 1. Setup Auth
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    try:
        creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
        client = gspread.authorize(creds)
    except Exception as e:
        print(f"Error authenticating with Google: {e}")
        return

    # 2. Open Spreadsheet
    try:
        # Try opening as ID first, then as Name
        if len(sheet_name_or_id) > 30 and "-" in sheet_name_or_id:
            sh = client.open_by_key(sheet_name_or_id)
        else:
            sh = client.open(sheet_name_or_id)
    except gspread.SpreadsheetNotFound:
        print(f"Error: Spreadsheet '{sheet_name_or_id}' not found. Make sure it is shared with the service account email.")
        return
    except Exception as e:
        print(f"Error opening spreadsheet: {e}")
        return

    # 3. Write each DataFrame to its own tab
    for tab_name, df in data_frames.items():
        print(f"  -> Updating tab: {tab_name}...")
        
        # Ensure worksheet exists
        try:
            worksheet = sh.worksheet(tab_name)
        except gspread.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=tab_name, rows="100", cols="20")

        # Prepare data (convert to list of lists, handle NaNs and dates)
        # We use values.tolist() but need to clean it up first
        clean_df = df.copy()
        
        # Format all columns
        for col in clean_df.columns:
            clean_df[col] = clean_df[col].apply(format_value)
            
        data = [clean_df.columns.values.tolist()] + clean_df.values.tolist()

        # Update worksheet
        worksheet.clear()
        worksheet.update(data)
        
        # --- Apply Professional Styling ---
        _apply_sheet_styles(worksheet, clean_df)

    print(f"Successfully exported to Google Sheets: {sh.url}")

def _apply_sheet_styles(worksheet, df):
    """Applies styling using batched requests to avoid API quota limits."""
    # Collect all formatting rules
    formats = []
    
    # 1. Header Style
    header_range = f'A1:{gspread.utils.rowcol_to_a1(1, len(df.columns))}'
    formats.append({
        "range": header_range,
        "format": cellFormat(
            backgroundColor={"red": 47/255, "green": 84/255, "blue": 150/255},
            textFormat=textFormat(bold=True, foregroundColor={"red": 1.0, "green": 1.0, "blue": 1.0}),
            horizontalAlignment='CENTER'
        )
    })

    # 2. Currency / Percent / Integer formats
    for i, col in enumerate(df.columns):
        col_letter = gspread.utils.rowcol_to_a1(1, i+1).rstrip('0123456789')
        col_range = f'{col_letter}2:{col_letter}1000'
        
        # Detection logic
        is_num = any(x in col.upper() for x in ['VALUE', 'PRICE', 'PNL', 'LTP', 'SL', 'EMA', 'QTY', 'QUANTITY', 'HOLDING', 'CLOSE'])
        is_pct = 'PCT' in col.upper() or '%' in col
        is_int = any(x in col.upper() for x in ['QTY', 'QUANTITY', 'COUNT', 'ID'])
        
        if is_num:
            if is_pct:
                formats.append({"range": col_range, "format": cellFormat(numberFormat=numberFormat(type='PERCENT', pattern='0.00%'), horizontalAlignment='RIGHT')})
            elif is_int:
                formats.append({"range": col_range, "format": cellFormat(numberFormat=numberFormat(type='NUMBER', pattern='#,##0'), horizontalAlignment='RIGHT')})
            else:
                formats.append({"range": col_range, "format": cellFormat(numberFormat=numberFormat(type='CURRENCY', pattern='₹#,##0.00'), horizontalAlignment='RIGHT')})
        else:
            formats.append({"range": col_range, "format": cellFormat(horizontalAlignment='LEFT')})

    # 3. Add Borders to all data
    data_range = f'A1:{gspread.utils.rowcol_to_a1(len(df)+1, len(df.columns))}'
    formats.append({"range": data_range, "format": cellFormat(borders=borders(top=border('SOLID'), bottom=border('SOLID'), left=border('SOLID'), right=border('SOLID')))})

    # Apply all formats in ONE call
    if formats:
        format_cell_ranges(worksheet, [(f['range'], f['format']) for f in formats])

    # Miscellaneous
    try:
        set_frozen(worksheet, rows=1)
        set_column_widths(worksheet, [
            (gspread.utils.rowcol_to_a1(1, j+1).rstrip('0123456789'), 120) for j in range(len(df.columns))
        ])
    except:
        pass
