import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os

# Configuration
GSHEET_NAME = '20251214_수식연결_가계부엔진_최종' # The discovered sheet name
SHEET_KEY_FILE = 'service_account.json'
WORKSHEET_NAME = '📋 T_RawData' # Explicitly matching the Excel sheet name just in case

# Column Mapping Configuration
COL_MAP_ENG_TO_KOR = {
    'date': '날짜',
    'time': '시간',
    'type': '구분',
    'main_category': '대분류',
    'sub_category': '소분류',
    'content': '내용',
    'amount': '금액',
    'payment_method': '결제수단',
    'memo': '메모',
    'merchant': '거래처', # Optional
    'Is_Active': 'Is_Active' # Keep same
}
COL_MAP_KOR_TO_ENG = {v: k for k, v in COL_MAP_ENG_TO_KOR.items()}

def connect_gsheet():
    """Connects to Google Sheets using credentials from Streamlit Cloud secrets or local file."""
    try:
        # Try Streamlit Cloud secrets first
        try:
            import streamlit as st
            if 'gcp_service_account' in st.secrets:
                # Use Streamlit Secrets
                creds_dict = dict(st.secrets["gcp_service_account"])
                scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
                client = gspread.authorize(creds)
                return client
        except:
            pass  # Streamlit not available or secrets not configured
        
        # Fall back to local service_account.json
        if not os.path.exists(SHEET_KEY_FILE):
            print("Error: service_account.json not found and Streamlit secrets not configured")
            return None
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(SHEET_KEY_FILE, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        print(f"Authentication Error: {e}")
        return None

def load_data():
    """
    Loads data from Google Sheet.
    """
    client = connect_gsheet()
    if not client:
        return pd.DataFrame()

    try:
        # Open Spreadsheet
        sh = client.open(GSHEET_NAME)
        
        # Open Worksheet (Try 'T_RawData', else 'Sheet1', else First)
        try:
            ws = sh.worksheet(WORKSHEET_NAME)
        except:
             try: ws = sh.worksheet('Sheet1')
             except: ws = sh.get_worksheet(0)
                 
        # Robust Read: Use get_all_values to avoid header errors
        rows = ws.get_all_values()
        
        if not rows or len(rows) < 2:
             return pd.DataFrame(columns=['날짜', '시간', '구분', '대분류', '소분류', '내용', '금액', '결제수단', '메모', 'Flow_Filter', 'Is_Active'])
             
        # ENFORCE SCHEMA (Hardcoded)
        # We assume the data starts from some row, but let's try to detect START of data or header.
        # If Row 0 has '날짜', but others are empty, it's the header row but broken.
        # Let's assume standard structure if '날짜' is in 1st col.
        
        header_row_idx = 0
        for i, r in enumerate(rows[:5]):
             if '날짜' in str(r[0]) or 'Date' in str(r[0]):
                 header_row_idx = i
                 break
                 
        # Skip header row and load data
        data_rows = rows[header_row_idx+1:]
        
        # Manually construct DataFrame with Fixed Columns
        # Standard GSheet/Excel Structure for this project:
        # 0: 날짜, 1: 시간, 2: 구분, 3: 대분류, 4: 소분류, 5: 내용, 6: 금액, 7: 결제수단, 8: 메모, 9: Flow_Filter, 10: Is_Active (Maybe)
        
        expected_cols = ['날짜', '시간', '구분', '대분류', '소분류', '내용', '금액', '결제수단', '메모', 'Flow_Filter', 'Is_Active']
        
        # Pad rows if they are short
        cleaned_data = []
        for r in data_rows:
            # Ensure row has enough items
            while len(r) < len(expected_cols):
                r.append("")
            # Truncate if too long (optional, but safer to keep)
            cleaned_data.append(r[:len(expected_cols)])
            
        df = pd.DataFrame(cleaned_data, columns=expected_cols)
        
        # Remove empty rows (where Date is empty)
        df = df[df['날짜'].astype(str).str.strip() != ""]
        
        # COLUMN MAPPING (Eng -> Kor)
        # Check if english cols exist
        if 'date' in df.columns:
            df = df.rename(columns=COL_MAP_ENG_TO_KOR)
            
        # Type Conversion
        if '날짜' in df.columns:
            df['날짜'] = pd.to_datetime(df['날짜'], errors='coerce')
            
        def parse_time(val):
            if pd.isna(val) or val == "": return None
            if isinstance(val, str):
                try: return pd.to_datetime(val, format='%H:%M:%S').time()
                except: 
                    try: return pd.to_datetime(val).time()
                    except: return None
            return val
            
        if '시간' in df.columns:
            df['시간'] = df['시간'].apply(parse_time)
            
        if '금액' in df.columns:
            # Remove commas if string
            if df['금액'].dtype == object:
                 df['금액'] = df['금액'].astype(str).str.replace(',', '')
            df['금액'] = pd.to_numeric(df['금액'], errors='coerce').fillna(0)
            
        if 'Flow_Filter' in df.columns:
            df['Is_Active'] = df['Flow_Filter'].apply(lambda x: str(x).strip()== '1')
        elif 'Is_Active' in df.columns:
            # Already boolean or string?
            # GSheet stores TRUE/FALSE as boolean usually
            df['Flow_Filter'] = df['Is_Active'].apply(lambda x: 1 if x else 0)
        else:
            df['Is_Active'] = True
            df['Flow_Filter'] = 1
            
        return df

    except Exception as e:
        print(f"GSheet Load Error: {e}")
        return pd.DataFrame()

def save_data(df):
    """
    Saves DataFrame to Google Sheet (Overwrites).
    Maps Korean columns back to English if needed.
    """
    client = connect_gsheet()
    if not client: return False
    
    try:
        sh = client.open(GSHEET_NAME)
        try:
            ws = sh.worksheet(WORKSHEET_NAME)
        except:
            ws = sh.get_worksheet(0)
            
        # Prepare Data for Upload
        save_df = df.copy()
        
        # MAPPING (Kor -> Eng)
        # We want to save consistent with the source.
        # If source was English (detected by load_data), we should save as English.
        # But here we don't know what load_data saw.
        # We can assume we want to maintain the English schema if that's what we found.
        # Let's enforce English Schema for GSheet as it seems to be the "Database" standard there.
        
        if '날짜' in save_df.columns:
            save_df['날짜'] = save_df['날짜'].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) and hasattr(x, 'strftime') else str(x) if pd.notnull(x) else "")
            
        if '시간' in save_df.columns:
            save_df['시간'] = save_df['시간'].apply(lambda x: x.strftime('%H:%M:%S') if pd.notnull(x) and hasattr(x, 'strftime') else str(x) if pd.notnull(x) else "")
            
        # Convert all to simple types
        save_df = save_df.fillna("")
        
        # Update
        # Strategy: RESTORE Headers & Save Data
        
        # 1. Clear Old Data Only (A3:K...)
        try:
             ws.batch_clear(['A3:K50000']) # Clear safe large range
        except:
             pass 
             
        # 2. Restore Title & Header (Self-Healing)
        # Row 1: Title
        ws.update('A1', [['가계부 데이터 엔진 (T_RawData)']])
        
        # Row 2: Headers (Standard Korean)
        # We need headers that match the 'save_df' columns, but in Korean.
        # save_df currently has English columns (date, time, etc.) if mapped?
        # NO, 'save_df' right now is just 'df.copy()', then renamed to ENG.
        # User wants Korean Headers visible in Sheet?
        # Image 2 shows "날짜", "시간"... (Korean).
        # But my code at line 163 renames KOR -> ENG.
        # If I save with ENG headers, the sheet will look like ENG.
        # User wants KOREAN headers in Excel.
        # So I should NOT rename to ENG if the target is to look like the Screenshot.
        # BUT, load_data assumes columns...
        # Wait, if I change save to KOR, load_data must handle KOR. 
        # My load_data handles KOR? Yes, it detects '날짜'.
        
        # Change of Plan: Save in KOREAN to match User's Original Excel.
        # This means I should SKIP the rename to ENG at line 163.
        
        # Let's revert the Rename to ENG.
        # Using Korean columns for GSheet view.
        
        # If I comment out the rename:
        # save_df columns are ['날짜', '시간', '구분'...]
        
        headers = ['날짜', '시간', '구분', '대분류', '소분류', '내용', '금액', '결제수단', '메모', 'Flow_Filter']
        # Is_Active is redundant with Flow_Filter, so we don't save it to keep sheet clean.
        
        # Ensure save_df has these columns in order
        existing_cols = [c for c in headers if c in save_df.columns]
        save_df = save_df[existing_cols]
        
        # Write Headers to A2
        ws.update('A2', [save_df.columns.values.tolist()])
             
        # Write Data (values only) starting at A3
        ws.update('A3', save_df.values.tolist())
        
        return True
        
        return True
        
    except Exception as e:
        print(f"GSheet Save Error: {e}")
        return False
        
def get_kpi_metrics(df):
    # Same logic as before
    if df.empty:
        return {'income':0, 'expense':0, 'net':0, 'other':0}
        
    # Ensure correct column names (should be Korean here as load_data maps it)
    if 'Is_Active' not in df.columns:
         df['Is_Active'] = True
    if '구분' not in df.columns and 'type' in df.columns:
         # Fallback if mapping failed?
         df['구분'] = df['type']
         df['금액'] = df['amount']
         
    active_df = df[df['Is_Active'] == True]
    
    income = active_df[active_df['구분'] == '수입']['금액'].sum()
    expense = active_df[active_df['구분'] == '지출']['금액'].sum()
    
    if 'Flow_Filter' in df.columns:
        other_df = df[(df['Is_Active'] == False) & (df['Flow_Filter'] == 1)]
    else:
        other_df = pd.DataFrame()
        
    other = other_df['금액'].sum() if not other_df.empty else 0
    
    net = income + expense
    
    return {
        'income': income,
        'expense': expense,
        'net': net,
        'other': other
    }

def add_row_optimized(new_row_dict):
    """
    Inserts a single row at the top of the data (Row 3) to maintain 'latest first' somewhat,
    without rewriting the whole sheet.
    """
    client = connect_gsheet()
    if not client: return False
    
    try:
        sh = client.open(GSHEET_NAME)
        try: ws = sh.worksheet(WORKSHEET_NAME)
        except: ws = sh.get_worksheet(0)
        
        # Prepare Row Data matching Headers
        # Headers: ['날짜', '시간', '구분', '대분류', '소분류', '내용', '금액', '결제수단', '메모', 'Flow_Filter']
        
        # Helper to format
        def fmt(val): return str(val) if val is not None else ""
        
        row_values = [
            new_row_dict.get('날짜', '').strftime('%Y-%m-%d') if hasattr(new_row_dict.get('날짜'), 'strftime') else fmt(new_row_dict.get('날짜')),
            new_row_dict.get('시간', '').strftime('%H:%M:%S') if hasattr(new_row_dict.get('시간'), 'strftime') else fmt(new_row_dict.get('시간')),
            fmt(new_row_dict.get('구분')),
            fmt(new_row_dict.get('대분류')),
            fmt(new_row_dict.get('소분류')),
            fmt(new_row_dict.get('내용')),
            new_row_dict.get('금액', 0), # Int/Float
            fmt(new_row_dict.get('결제수단')),
            fmt(new_row_dict.get('메모')),
            new_row_dict.get('Flow_Filter', 1)
        ]
        
        # Insert at Row 3 (Pushing others down)
        # This is strictly faster than overwriting 3000 rows.
        ws.insert_row(row_values, index=3)
        return True
        
    except Exception as e:
        print(f"GSheet Add Row Error: {e}")
        return False
