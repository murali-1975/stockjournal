import gspread
from google.oauth2.service_account import Credentials

def debug_access():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
        client = gspread.authorize(creds)
        
        print(f"Service Account Email: {creds.service_account_email}")
        print("\n--- Listing all accessible spreadsheets ---")
        sheets = client.openall()
        if not sheets:
            print("No spreadsheets found! You must share your Watchlists with the Service Account email above.")
        for sh in sheets:
            print(f"- {sh.title} (ID: {sh.id})")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    debug_access()
