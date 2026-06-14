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
import logging
from gspread_formatting import *

logger = logging.getLogger(__name__)

def print(*args, **kwargs):
    msg = " ".join(str(arg) for arg in args)
    msg_upper = msg.upper()
    if "ERROR" in msg_upper or "FAILED" in msg_upper:
        logger.error(msg)
    elif "WARNING" in msg_upper or "NOTE:" in msg_upper:
        logger.warning(msg)
    else:
        logger.info(msg)

import datetime

def format_value(val):
    """Formats a value for Google Sheets compatibility."""
    if pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp, np.datetime64, datetime.datetime, datetime.date)):
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

    # Delete legacy sheets to clean up the workbook
    legacy_sheets = ['Sheet1', 'Grouped_Trades', 'Portfolio']
    for sheet_name in legacy_sheets:
        try:
            ws_to_del = sh.worksheet(sheet_name)
            sh.del_worksheet(ws_to_del)
            print(f"Removed legacy worksheet: {sheet_name}")
        except gspread.WorksheetNotFound:
            pass
        except Exception as e:
            print(f"Note: Could not delete legacy sheet '{sheet_name}': {e}")

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
            _apply_gsheet_core_colors(sh, worksheet, clean_df)

    # --- Update Satellite_Watchlist sheet if present in Google Sheet ---
    _update_gsheet_satellite_watchlist(sh)

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
        col_upper = str(col).upper()
        is_num = any(x in col_upper for x in ['VALUE', 'PRICE', 'PNL', 'LTP', 'SL', 'EMA', 'QTY', 'QUANTITY', 'HOLDING', 'CLOSE'])
        is_pct = 'PCT' in col_upper or '%' in str(col)
        is_int = any(x in col_upper for x in ['QTY', 'QUANTITY', 'COUNT', 'ID'])
        
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
            
            for target_col in ['Prev_Day_Close', 'Prev_Week_Close', 'Prev_Month_Close']:
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
            
            # Apply conditional formatting for Trend column
            if 'Trend' in col_map:
                trend_idx = col_map['Trend']
                trend_letter = gspread.utils.rowcol_to_a1(1, trend_idx).rstrip('0123456789')
                gr_t = GridRange.from_a1_range(f'{trend_letter}2:{trend_letter}{len(df)+1}', worksheet)
                
                # Align Trend column to center
                formats.append({
                    "range": f'{trend_letter}2:{trend_letter}{len(df)+1}',
                    "format": cellFormat(horizontalAlignment='CENTER')
                })
                
                rule_up = ConditionalFormatRule(
                    ranges=[gr_t],
                    booleanRule=BooleanRule(
                        condition=BooleanCondition(type='TEXT_EQ', values=['▲']),
                        format=CellFormat(textFormat=textFormat(bold=True, foregroundColor=Color(20/255, 122/255, 30/255)))
                    )
                )
                rule_down = ConditionalFormatRule(
                    ranges=[gr_t],
                    booleanRule=BooleanRule(
                        condition=BooleanCondition(type='TEXT_EQ', values=['▼']),
                        format=CellFormat(textFormat=textFormat(bold=True, foregroundColor=Color(192/255, 0/255, 0/255)))
                    )
                )
                rules.append(rule_up)
                rules.append(rule_down)
            
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
        watchlist_df['Stock'] = watchlist_df['Stock'].astype(str).str.strip().str.upper()
        watchlist_df['Color'] = watchlist_df['Color'].astype(str).str.strip().str.upper()
        
        watchlist_df['Date'] = pd.to_datetime(watchlist_df['Date'], dayfirst=True, errors='coerce')
        watchlist_df = watchlist_df.dropna(subset=['Date'])
        
        # Enforce strict latest date filtering
        latest_colors = {}
        if not watchlist_df.empty:
            latest_date = watchlist_df['Date'].max()
            watchlist_df = watchlist_df[watchlist_df['Date'] == latest_date]
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
            symbol = str(row.get('Symbol', '')).strip().upper()
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
                else:
                    symbol_col_letter = gspread.utils.rowcol_to_a1(1, symbol_idx).rstrip('0123456789')
                    cell_range = f'{symbol_col_letter}{row_num}'
                    formats.append({
                        "range": cell_range,
                        "format": cellFormat(
                            backgroundColor=color(1.0, 1.0, 1.0),
                            textFormat=textFormat(bold=False, foregroundColor=color(0, 0, 0))
                        )
                    })
            else:
                # Stock not in the latest date's watchlist: Clear background fill and restore standard style
                symbol_col_letter = gspread.utils.rowcol_to_a1(1, symbol_idx).rstrip('0123456789')
                cell_range = f'{symbol_col_letter}{row_num}'
                formats.append({
                    "range": cell_range,
                    "format": cellFormat(
                        backgroundColor=color(1.0, 1.0, 1.0),
                        textFormat=textFormat(bold=False, foregroundColor=color(0, 0, 0))
                    )
                })
                    
    if formats:
        try:
            format_cell_ranges(current_portfolio_ws, [(f['range'], f['format']) for f in formats])
        except Exception as e:
            print(f"Error applying cell formats in Google Sheets: {e}")


def _apply_gsheet_core_colors(sh, current_portfolio_ws, df):
    """
    Reads the Core_Watchlist sheet from the Google Sheet, finds the trend status
    per stock symbol, and styles the Symbol column in the Current_Portfolio
    sheet with the corresponding user color and bold black font.
    """
    from gspread_formatting import cellFormat, textFormat, color, format_cell_ranges
    import pandas as pd
    
    # Predefined color mappings in float RGB for Google Sheets Color API
    GS_COLOR_MAP = {
        'Strong Trend- Green': (0.0, 176/255, 80/255),
        'Medium Trend':        (1.0, 192/255, 0.0),
        'Weak Trend- Red':    (220/255, 57/255, 57/255),
        'Core-Weekly':         (177/255, 160/255, 199/255)
    }

    watchlist_ws = None
    for name in ['Core_Watchlist', 'Core_Portfolio']:
        try:
            watchlist_ws = sh.worksheet(name)
            break
        except Exception:
            pass

    if not watchlist_ws:
        return

    try:
        records = watchlist_ws.get_all_records()
        if not records:
            return
            
        watchlist_df = pd.DataFrame(records)
    except Exception as e:
        print(f"Note: Core watchlist sheet could not be loaded from Google Sheet: {e}")
        return

    try:
        if 'Company' not in watchlist_df.columns or 'Trend Status' not in watchlist_df.columns:
            return
        watchlist_df = watchlist_df.dropna(subset=['Company', 'Trend Status'])
        watchlist_df['Company'] = watchlist_df['Company'].astype(str).str.strip().str.upper()
        watchlist_df['Trend Status'] = watchlist_df['Trend Status'].astype(str).str.strip()
        
        # Enforce strict latest month filtering
        if 'Month' in watchlist_df.columns:
            watchlist_df['Month'] = pd.to_datetime(watchlist_df['Month'], errors='coerce')
            latest_month = watchlist_df['Month'].max()
            watchlist_df = watchlist_df[watchlist_df['Month'] == latest_month]
            
        watchlist_df = watchlist_df.dropna(subset=['Company', 'Trend Status'])
        watchlist_df['Company'] = watchlist_df['Company'].astype(str).str.strip().str.upper()
        watchlist_df['Trend Status'] = watchlist_df['Trend Status'].astype(str).str.strip()
        
        latest_trends = watchlist_df.drop_duplicates(subset=['Company']).set_index('Company')['Trend Status'].to_dict()
    except Exception as e:
        print(f"Error parsing Google Sheets core watchlist data: {e}")
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
        if class_val == 'Core':
            symbol = str(row.get('Symbol', '')).strip().upper()
            if symbol in latest_trends:
                trend = latest_trends[symbol]
                if trend in GS_COLOR_MAP:
                    bg = GS_COLOR_MAP[trend]
                    
                    symbol_col_letter = gspread.utils.rowcol_to_a1(1, symbol_idx).rstrip('0123456789')
                    cell_range = f'{symbol_col_letter}{row_num}'
                    
                    formats.append({
                        "range": cell_range,
                        "format": cellFormat(
                            backgroundColor=color(bg[0], bg[1], bg[2]),
                            textFormat=textFormat(bold=True, foregroundColor=color(0, 0, 0))
                        )
                    })
            else:
                symbol_col_letter = gspread.utils.rowcol_to_a1(1, symbol_idx).rstrip('0123456789')
                cell_range = f'{symbol_col_letter}{row_num}'
                formats.append({
                    "range": cell_range,
                    "format": cellFormat(
                        backgroundColor=color(1.0, 1.0, 1.0),
                        textFormat=textFormat(bold=False, foregroundColor=color(0, 0, 0))
                    )
                })
                    
    if formats:
        try:
            format_cell_ranges(current_portfolio_ws, [(f['range'], f['format']) for f in formats])
        except Exception as e:
            print(f"Error applying cell formats in Google Sheets: {e}")


def _update_gsheet_satellite_watchlist(sh) -> None:
    """
    Updates the Google Sheet 'Satellite_Watchlist' tab with weekly market data (Columns G, H, I, J),
    preserving Columns D, E, F formulas.
    """
    try:
        watchlist_ws = sh.worksheet('Satellite_Watchlist')
    except Exception:
        # Sheet not found, skip gracefully
        return

    try:
        values = watchlist_ws.get_all_values()
        if not values or len(values) < 2:
            return
            
        # Get unique symbols from Column B (index 1)
        symbols = []
        for row in values[1:]:
            if len(row) > 1:
                sym = str(row[1]).strip().upper()
                if sym and sym not in symbols:
                    symbols.append(sym)
                    
        if not symbols:
            return

        from src.market_api import fetch_market_data_from_yahoo
        classifications = {sym: 'Satellite' for sym in symbols}
        market_data = fetch_market_data_from_yahoo(symbols, classifications=classifications)
        
        # Ensure header row has the headers for D, E, F, G, H, I, J
        headers = values[0]
        # Make sure the header row has at least 10 columns
        while len(headers) < 10:
            headers.append("")
        headers[3] = "Company Name"
        headers[4] = "Price"
        headers[5] = "Previous Close "
        headers[6] = "Previous week Close"
        headers[7] = "EMA 9 (weekly)"
        headers[8] = "EMA 11 (weekly)"
        headers[9] = "EMA 21 (weekly)"
        
        # Update headers in Google Sheet
        watchlist_ws.update('A1:J1', [headers[:10]])
        
        # Build 2D values list for D2:J
        row_updates = []
        for row in values[1:]:
            sym = ""
            if len(row) > 1:
                sym = str(row[1]).strip().upper()
            
            if sym in market_data:
                data = market_data[sym]
                
                # Fetch Company Name (using Yahoo Finance long/short name if available)
                comp_name = data.get('Company_Name', '')
                if not comp_name:
                    # Fallback to existing value if it is not an error string
                    existing_name = row[3] if len(row) > 3 else ""
                    if existing_name and not str(existing_name).startswith("#"):
                        comp_name = existing_name
                    else:
                        comp_name = sym
                
                row_updates.append([
                    comp_name,
                    data.get('LTP', 0.0),
                    data.get('Prev_Day_Close', 0.0),
                    data.get('Prev_Week_Close', 0.0),
                    data.get('EMA9', 0.0),
                    data.get('EMA11', 0.0),
                    data.get('EMA21', 0.0)
                ])
            else:
                # If symbol not in market_data, preserve existing non-error values or leave empty
                existing_d = row[3] if (len(row) > 3 and not str(row[3]).startswith("#")) else ""
                existing_e = row[4] if (len(row) > 4 and not str(row[4]).startswith("#")) else ""
                existing_f = row[5] if (len(row) > 5 and not str(row[5]).startswith("#")) else ""
                row_updates.append([existing_d, existing_e, existing_f, "", "", "", ""])
                
        # Bulk update D2:J
        end_row = len(values)
        range_str = f'D2:J{end_row}'
        watchlist_ws.update(range_str, row_updates)
        
        # Format the numbers (Currency format) for Columns E, F, G, H, I, J
        # In Google Sheets, format columns E through J as Currency with ₹ symbol
        # and apply standard styling like thin borders and alignment.
        from gspread_formatting import cellFormat, numberFormat, format_cell_ranges, borders, border, textFormat
        import gspread
        
        col_formats = []
        for col_idx in [5, 6, 7, 8, 9, 10]: # E, F, G, H, I, J (5 is E, 10 is J)
            col_letter = gspread.utils.rowcol_to_a1(1, col_idx).rstrip('0123456789')
            col_range = f'{col_letter}2:{col_letter}{end_row}'
            col_formats.append({
                "range": col_range,
                "format": cellFormat(
                    numberFormat=numberFormat(type='CURRENCY', pattern='₹#,##0.00'),
                    horizontalAlignment='RIGHT'
                )
            })
            
        # Apply header styling to new columns (D1:J1)
        header_range = f'D1:J1'
        col_formats.append({
            "range": header_range,
            "format": cellFormat(
                backgroundColor={"red": 47/255, "green": 84/255, "blue": 150/255},
                textFormat=textFormat(bold=True, foregroundColor={"red": 1.0, "green": 1.0, "blue": 1.0}),
                horizontalAlignment='CENTER'
            )
        })
        
        # Apply borders to the updated D1:J range
        data_range = f'D1:J{end_row}'
        col_formats.append({
            "range": data_range,
            "format": cellFormat(
                borders=borders(top=border('SOLID'), bottom=border('SOLID'), left=border('SOLID'), right=border('SOLID'))
            )
        })
        
        format_cell_ranges(watchlist_ws, [(f['range'], f['format']) for f in col_formats])
        
        # Apply dynamic conditional formatting in Google Sheet
        _apply_gsheet_watchlist_conditional_formatting(watchlist_ws, end_row)
        
    except Exception as e:
        print(f"Error updating Google Sheets Satellite_Watchlist data: {e}")


def _apply_gsheet_watchlist_conditional_formatting(ws, end_row) -> None:
    """
    Applies custom formula conditional formatting to Google Sheet 'Satellite_Watchlist' tab.
    """
    from gspread_formatting import get_conditional_format_rules, ConditionalFormatRule, BooleanRule, BooleanCondition, CellFormat, Color, textFormat, GridRange
    
    try:
        rules = get_conditional_format_rules(ws)
        rules.clear()
        
        # Format rules for Column F (Previous Close)
        gr_f = GridRange.from_a1_range(f'F2:F{end_row}', ws)
        rule_f = ConditionalFormatRule(
            ranges=[gr_f],
            booleanRule=BooleanRule(
                condition=BooleanCondition(type='CUSTOM_FORMULA', values=['=$E2>$F2']),
                format=CellFormat(
                    backgroundColor=Color(226/255, 239/255, 218/255),
                    textFormat=textFormat(bold=True, foregroundColor=Color(55/255, 86/255, 35/255))
                )
            )
        )
        
        # Format rules for Column G (Previous week Close)
        gr_g = GridRange.from_a1_range(f'G2:G{end_row}', ws)
        rule_g = ConditionalFormatRule(
            ranges=[gr_g],
            booleanRule=BooleanRule(
                condition=BooleanCondition(type='CUSTOM_FORMULA', values=['=$F2>$G2']),
                format=CellFormat(
                    backgroundColor=Color(226/255, 239/255, 218/255),
                    textFormat=textFormat(bold=True, foregroundColor=Color(55/255, 86/255, 35/255))
                )
            )
        )
        
        rules.append(rule_f)
        rules.append(rule_g)
        rules.save()
        
    except Exception as e:
        print(f"Error applying Google Sheets watchlist conditional formatting: {e}")

