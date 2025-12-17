import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configuration
KEY_FILE = 'service_account.json'
# The name of the file as it appears in Google Drive
TARGET_SHEET_NAME = '2024-12-07~2025-12-07_v5_LinkDB' 

def test_connection():
    print("--- 🔌 Connecting to Google Sheets ---")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
        client = gspread.authorize(creds)
        print("✅ Authentication Successful!")
    except Exception as e:
        print(f"❌ Authentication Failed: {e}")
        return

    print(f"🔍 Searching for spreadsheet: '{TARGET_SHEET_NAME}'...")
    
    sheet = None
    try:
        sheet = client.open(TARGET_SHEET_NAME)
        print(f"✅ Found Spreadsheet: {sheet.title}")
    except gspread.SpreadsheetNotFound:
        print(f"❌ Could not find spreadsheet with exact name: '{TARGET_SHEET_NAME}'")
        print("   Attempting to find by filename with .xlsx extension...")
        try:
            sheet = client.open(TARGET_SHEET_NAME + ".xlsx")
            print(f"✅ Found Spreadsheet: {sheet.title}")
        except gspread.SpreadsheetNotFound:
            print("❌ File still not found.")
            print("\n📋 Listing ALL available spreadsheets for this service account:")
            try:
                for ss in client.openall():
                    print(f"   - {ss.title}")
            except Exception as e:
                print(f"   Error listing sheets: {e}")
            print("\n⚠️ If you don't see your file, please make sure you SHARED it with the service account email.")
            return

    # If found, check DB_Raw
    if sheet:
        try:
            worksheet = sheet.worksheet("DB_Raw")
            print(f"\n✅ Successfully accessed tab: 'DB_Raw'")
            print(f"   Headers: {worksheet.row_values(1)}")
            print("\n🚀 Ready for Streamlit App Development!")
        except gspread.WorksheetNotFound:
            print(f"\n❌ 'DB_Raw' tab not found in {sheet.title}. Available tabs: {[ws.title for ws in sheet.worksheets()]}")

if __name__ == "__main__":
    test_connection()
