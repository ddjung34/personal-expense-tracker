import pandas as pd
import shutil
import os

# 1. Copy the "Good" file (2013 version with cached values) to temp to avoid permission errors
src = r'c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\20251213_수식연결_가계부엔진.xlsx'
temp = r'c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\temp_analysis.xlsx'

try:
    shutil.copy2(src, temp)
    
    # 2. Read with Pandas
    # Header is at row 2 (index 1)
    df = pd.read_excel(temp, sheet_name='📋 T_RawData', header=1)
    
    # Cols: J is index 9, K is index 10.
    # Pandas columns might be named. Let's print columns.
    # print(df.columns)
    
    # The 'Flow_Filter' column might be explicitly named or implied.
    # Let's inspect columns J and K by index if possible, or name.
    # Flow_Filter is Col 10 (J) -> Index 9.
    
    # Filter for '지출'
    expense_df = df[df['구분'] == '지출'].copy()
    
    # Identify Filter Value
    # Assuming column 'Flow_Filter' exists
    if 'Flow_Filter' in expense_df.columns:
        # Group to see which categories have 0
        summary = expense_df.groupby('대분류')['Flow_Filter'].max().reset_index()
        
        print("=== Categories where Flow_Filter is 0 (EXCLUDED) ===")
        excluded = summary[summary['Flow_Filter'] == 0]['대분류'].tolist()
        print(excluded)
        
        print("\n=== Categories where Flow_Filter is 1 (INCLUDED) ===")
        included = summary[summary['Flow_Filter'] == 1]['대분류'].tolist()
        print(included)
        
    else:
        print("Column 'Flow_Filter' not found by name. Checking column index 9...")
        # fallback by index
        col_name = df.columns[9]
        print(f"Using column '{col_name}' as filter.")
        summary = expense_df.groupby('대분류')[col_name].max().reset_index()
        print(summary[summary[col_name] == 0])

except Exception as e:
    print(e)
finally:
    if os.path.exists(temp):
        os.remove(temp)
