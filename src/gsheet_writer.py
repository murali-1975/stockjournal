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

        # --- Apply custom color formatting for Satellite stocks based on Satellite_Watchlist ---
        if tab_name == 'Current_Portfolio':
            _apply_gsheet_satellite_colors(sh, worksheet, clean_df)

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

    # 4. Dynamic Conditional Formatting for Prev_Day_Close and Prev_Week_Close
    if worksheet.title == "Current_Portfolio":
        col_map = {col_name: col_idx for col_idx, col_name in enumerate(df.columns, 1)}
        if 'LTP' in col_map:
            ltp_idx = col_map['LTP']
            ltp_letter = gspread.utils.rowcol_to_a1(1, ltp_idx).rstrip('0123456789')
            
            # Fetch existing rules and clear them to avoid duplicate builds
            rules = get_conditional_format_rules(worksheet)
            rules.clear()
            
            for target_col in ['Prev_Day_Close', 'Prev_Week_Close']:
                if target_col in col_map:
                    target_idx = col_map[target_col]
                    target_letter = gspread.utils.rowcol_to_a1(1, target_idx).rstrip('0123456789')
                    
                    gr = GridRange.from_a1_range(f'{target_letter}2:{target_letter}{len(df)+1}', worksheet)
                    
                    # Custom formulas relative to row 2
                    green_formula = f"=${ltp_letter}2>${target_letter}2"
                    pink_formula = f"=${ltp_letter}2<=${target_letter}2"
                    
                    rule_green = ConditionalFormatRule(
                        ranges=[gr],
                        booleanRule=BooleanRule(
                            condition=BooleanCondition(type='CUSTOM_FORMULA', values=[green_formula]),
                            format=CellFormat(backgroundColor=Color(226/255, 239/255, 218/255))
                        )
                    )
                    
                    rule_pink = ConditionalFormatRule(
                        ranges=[gr],
                        booleanRule=BooleanRule(
                            condition=BooleanCondition(type='CUSTOM_FORMULA', values=[pink_formula]),
                            format=CellFormat(backgroundColor=Color(252/255, 228/255, 214/255))
                        )
                    )
                    
                    rules.append(rule_green)
                    rules.append(rule_pink)
            
            rules.save()

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


def _apply_gsheet_satellite_colors(sh, current_portfolio_ws, df):
    """
    Reads the Satellite_Watchlist sheet from the Google Sheet, finds the latest
    color code per stock symbol, and styles the Symbol column in the Current_Portfolio
    sheet with premium light background and dark text colors.
    """
    from gspread_formatting import cellFormat, textFormat, color, format_cell_ranges
    import pandas as pd
    
    # Predefined color mappings in float RGB for Google Sheets Color API
    GS_COLOR_MAP = {
        'BLUE':   {'bg': (221/255, 235/255, 247/255), 'font': (31/255, 78/255, 120/255)},
        'ORANGE': {'bg': (252/255, 228/255, 214/255), 'font': (198/255, 89/255, 17/255)},
        'GREEN':  {'bg': (226/255, 239/255, 218/255), 'font': (55/255, 86/255, 35/255)},
        'RED':    {'bg': (250/255, 219/255, 216/255), 'font': (169/255, 50/255, 38/255)},
        'PINK':   {'bg': (250/255, 219/255, 216/255), 'font': (169/255, 50/255, 38/255)},
        'YELLOW': {'bg': (255/255, 242/255, 204/255), 'font': (127/255, 96/255, 0/255)},
        'PURPLE': {'bg': (225/255, 213/255, 231/255), 'font': (96/255, 73/255, 122/255)}
    }

    try:
        # Load Satellite_Watchlist sheet from Google Sheet
        watchlist_ws = sh.worksheet('Satellite_Watchlist')
        records = watchlist_ws.get_all_records()
        if not records:
            return
            
        watchlist_df = pd.DataFrame(records)
    except Exception as e:
        print(f"Note: Satellite_Watchlist sheet could not be loaded from Google Sheet: {e}")
        return

    try:
        watchlist_df = watchlist_df.dropna(subset=['Stock', 'Color'])
        watchlist_df['Stock'] = watchlist_df['Stock'].astype(str).str.strip()
        watchlist_df['Color'] = watchlist_df['Color'].astype(str).str.strip().str.upper()
        
        # Sort chronologically by Date (newest first) to find the latest
        watchlist_df['Date'] = pd.to_datetime(watchlist_df['Date'], dayfirst=True, errors='coerce')
        watchlist_df = watchlist_df.sort_values(by='Date', ascending=False)
        
        # Deduplicate to keep only the latest color code per stock symbol
        latest_colors = watchlist_df.drop_duplicates(subset=['Stock']).set_index('Stock')['Color'].to_dict()
    except Exception as e:
        print(f"Error parsing Google Sheets Satellite_Watchlist data: {e}")
        return

    col_map = {col_name: col_idx for col_idx, col_name in enumerate(df.columns, 1)}
    if 'Symbol' not in col_map or 'TF_Classification' not in col_map:
        return

    symbol_idx = col_map['Symbol']
    class_idx = col_map['TF_Classification']

    formats = []
    
    # Loop over the dataframe rows
    for i, row in df.iterrows():
        row_num = i + 2 # Header is row 1
        class_val = str(row.get('TF_Classification', '')).strip()
        if class_val == 'Satellite':
            symbol = str(row.get('Symbol', '')).strip()
            if symbol in latest_colors:
                color_name = latest_colors[symbol]
                if color_name in GS_COLOR_MAP:
                    bg = GS_COLOR_MAP[color_name]['bg']
                    fg = GS_COLOR_MAP[color_name]['font']
                    
                    symbol_col_letter = gspread.utils.rowcol_to_a1(1, symbol_idx).rstrip('0123456789')
                    cell_range = f'{symbol_col_letter}{row_num}'
                    
                    formats.append({
                        "range": cell_range,
                        "format": cellFormat(
                            backgroundColor=color(bg[0], bg[1], bg[2]),
                            textFormat=textFormat(bold=True, foregroundColor=color(fg[0], fg[1], fg[2]))
                        )
                    })
                    
    if formats:
        try:
            format_cell_ranges(current_portfolio_ws, [(f['range'], f['format']) for f in formats])
        except Exception as e:
            print(f"Error applying cell formats in Google Sheets: {e}")
