"""
Verification Script: Excel vs Google Sheets
===========================================

Compares the data in the local Transformed_Tradebook.xlsx with the 
Google Sheet to ensure everything was copied correctly.
"""

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import os

def verify_sync(excel_path, gsheet_name):
    print(f"Starting Verification: {excel_path} vs Google Sheet '{gsheet_name}'")
    
    # 1. Load Excel Data
    try:
        excel_sheets = {
            "Raw_Tradebook": pd.read_excel(excel_path, sheet_name="Raw_Tradebook"),
            "Transaction": pd.read_excel(excel_path, sheet_name="Transaction"),
            "Current_Portfolio": pd.read_excel(excel_path, sheet_name="Current_Portfolio"),
            "Overall_Portfolio": pd.read_excel(excel_path, sheet_name="Overall_Portfolio"),
            "Action Tracker": pd.read_excel(excel_path, sheet_name="Action Tracker"),
            "Core_Watchlist": pd.read_excel(excel_path, sheet_name="Core_Watchlist"),
            "Satellite_Watchlist": pd.read_excel(excel_path, sheet_name="Satellite_Watchlist")
        }
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return

    # 2. Load Google Sheets Data
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    client = gspread.authorize(creds)
    sh = client.open(gsheet_name)
    
    # 3. Compare Row Counts
    mismatches = 0
    for name, excel_df in excel_sheets.items():
        try:
            ws = sh.worksheet(name)
            gsheet_df = pd.DataFrame(ws.get_all_records())
            
            e_rows = len(excel_df)
            g_rows = len(gsheet_df)
            
            if e_rows == g_rows:
                print(f"SUCCESS: {name}: Row counts match ({e_rows} rows).")
            else:
                print(f"FAILURE: {name}: ROW COUNT MISMATCH! Excel={e_rows}, GSheet={g_rows}")
                mismatches += 1
                
        except gspread.WorksheetNotFound:
            print(f"FAILURE: {name}: Worksheet missing in Google Sheets!")
            mismatches += 1
        except Exception as e:
            print(f"FAILURE: {name}: Error comparing: {e}")
            mismatches += 1

    if mismatches == 0:
        print("\nVERIFICATION SUCCESSFUL: All data sheets are synchronized.")
    else:
        print(f"\nVERIFICATION FAILED: Found {mismatches} mismatches.")

if __name__ == "__main__":
    verify_sync("Transformed_Tradebook.xlsx", "Transformed_Tradebook")
