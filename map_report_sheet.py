import pandas as pd

# Load the v2 file
FILE_PATH = r"c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\2024-12-07~2025-12-07_v2.xlsx"

try:
    # Read all sheets to find the report sheet
    xl = pd.ExcelFile(FILE_PATH, engine='openpyxl')
    print(f"Sheet names: {xl.sheet_names}")
    
    # Assuming the report is in the first original sheet (index 1 now, since DB_Raw is 0)
    # But let's check names. Based on previous analysis, it might be unnamed or have a specific name.
    # We will look for the sheet containing "2024-12" and "식비" (Food)
    
    target_sheet = None
    df = None
    
    for sheet in xl.sheet_names:
        if sheet == "DB_Raw": continue
        
        temp_df = xl.parse(sheet, header=None)
        # Search for key markers
        if temp_df.astype(str).apply(lambda x: x.str.contains("2024-12", na=False)).any().any():
            target_sheet = sheet
            df = temp_df
            print(f"Found report in sheet: {sheet}")
            break
            
    if df is not None:
        print("\n--- Locating Coordinates ---")
        # Find "2024-12" (Target Column)
        for col in df.columns:
            mask = df[col].astype(str).str.contains("2024-12", na=False)
            if mask.any():
                row_idx = mask.idxmax()
                print(f"Found '2024-12' at: Row {row_idx+1}, Col {col}") # +1 for 1-based index
                
        # Find "식비" (Target Row)
        # Note: The user didn't explicitly say "식비" is in the file, but implied it. 
        # Let's search for any text to map the structure roughly.
        print("\n--- Sample Header Row Implementation ---")
        # We need to find the header row containing dates
        header_row_idx = None
        for idx, row in df.iterrows():
            if row.astype(str).str.contains("2024-12").any():
                header_row_idx = idx
                print(f"Header Row found at index: {idx} (Excel Row {idx+1})")
                print(row.tolist())
                break
                
except Exception as e:
    print(f"Error: {e}")
