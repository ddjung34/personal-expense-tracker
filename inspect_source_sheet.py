import pandas as pd

FILE_PATH = r"c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\2024-12-07~2025-12-07.xlsx"

def inspect_sheet():
    print(f"Loading {FILE_PATH}...")
    try:
        # Load specifically the sheet name mentioned by user
        # Note: User said "가계부 내역", let's verify exact name match or find close match
        xl = pd.ExcelFile(FILE_PATH, engine='openpyxl')
        print(f"All Sheet Names: {xl.sheet_names}")
        
        target_name = "가계부 내역"
        if target_name not in xl.sheet_names:
            print(f"Warning: '{target_name}' not found. checking similar names...")
            for name in xl.sheet_names:
                if "가계부" in name or "내역" in name:
                    target_name = name
                    break
        
        print(f"Inspecting Sheet: {target_name}")
        df = pd.read_excel(FILE_PATH, sheet_name=target_name, engine='openpyxl')
        
        print("\n--- Columns ---")
        print(df.columns.tolist())
        
        print("\n--- First 5 Rows ---")
        print(df.head(5).to_string())
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    inspect_sheet()
