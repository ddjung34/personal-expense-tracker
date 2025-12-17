import pandas as pd
import os

# Use raw string for Windows path
file_path = r"c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\2024-12-07~2025-12-07.xlsx"

print(f"Analyzing file: {file_path}")

try:
    # Attempt to read the file
    # engine='openpyxl' is usually required for xlsx
    df = pd.read_excel(file_path, engine='openpyxl')
    
    print("\n--- DataFrame Info ---")
    df.info()
    
    print("\n--- Column Headers ---")
    print(df.columns.tolist())
    
    print("\n--- First 10 Rows ---")
    print(df.head(10).to_string())
    
    print("\n--- Basic Statistics ---")
    print(df.describe(include='all').to_string())

except ImportError as e:
    print(f"Missing dependency: {e}")
    print("Please install pandas and openpyxl: pip install pandas openpyxl")
except Exception as e:
    print(f"Error reading excel file: {e}")
