import pandas as pd
import openpyxl

FILE_PATH = r"c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\2024-12-07~2025-12-07.xlsx"

def inspect_excel():
    print(f"Analyzing: {FILE_PATH}")
    
    # Load Excel file
    xl = pd.ExcelFile(FILE_PATH, engine='openpyxl')
    
    print(f"\n📊 Available Sheets: {xl.sheet_names}")
    
    # Inspect each sheet
    for sheet_name in xl.sheet_names:
        print(f"\n{'='*50}")
        print(f"Sheet: {sheet_name}")
        print(f"{'='*50}")
        
        df = xl.parse(sheet_name, nrows=10)  # Read first 10 rows
        
        print(f"\nColumns: {df.columns.tolist()}")
        print(f"Total Rows (estimated): {len(xl.parse(sheet_name))}")
        print(f"\nFirst 5 rows:")
        print(df.head().to_string())
        
if __name__ == "__main__":
    inspect_excel()
