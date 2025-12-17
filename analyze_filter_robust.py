import pandas as pd
import os

# Use raw string for path
DATA_FILE = r'c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\20251213_수식연결_가계부엔진.xlsx'

print(f"Reading: {DATA_FILE}")

if not os.path.exists(DATA_FILE):
    print("File not found!")
    exit(1)

try:
    # Read with openpyxl engine (default for xlsx)
    df = pd.read_excel(DATA_FILE, sheet_name='📋 T_RawData', header=1, engine='openpyxl')
    
    # Check headers
    # print("Headers:", df.columns.tolist())
    
    filter_col_candidates = [c for c in df.columns if 'Filter' in str(c) or 'Flow' in str(c)]
    # print("Filter Cols:", filter_col_candidates)
    
    # Use 'Flow_Filter'
    target_col = 'Flow_Filter'
    
    if target_col in df.columns:
        # Group by Category and filter value
        summary = df.groupby(['대분류'])[target_col].max().reset_index()
        
        excluded = summary[summary[target_col] == 0]['대분류'].tolist()
        print("\n[EXCLUDED Categories (Filter=0)]")
        print(excluded)
        
        included = summary[summary[target_col] == 1]['대분류'].tolist()
        print("\n[INCLUDED Categories (Filter=1)]")
        print(included)
    else:
        print("Flow_Filter column missing")

except Exception as e:
    print(f"Error: {e}")
