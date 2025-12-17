import pandas as pd

FILE_PATH = r"c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\2024-12-07~2025-12-07_v2.xlsx"

def analyze_expenses():
    print("--- 📊 App Engine Starting ---")
    print(f"Reading Database from: {FILE_PATH}")
    
    # Read ONLY the DB_Raw sheet (The "Database")
    try:
        df = pd.read_excel(FILE_PATH, sheet_name="DB_Raw", engine='openpyxl')
    except Exception as e:
        print(f"Error reading DB: {e}")
        return

    print(f"\n✅ Loaded {len(df)} transactions.")
    
    # 1. Basic Stats
    total_income = df[df['type'] == '수입']['amount'].sum()
    total_expense = df[df['type'] == '지출']['amount'].sum()
    
    print(f"\n💰 Total Income: {total_income:,} won")
    print(f"💸 Total Expense: {total_expense:,} won")
    print(f"📉 Balance: {total_income - total_expense:,} won")
    
    # 2. Category Analysis (Pivot Table)
    print("\n--- 📂 Expense by Category ---")
    expense_df = df[df['type'] == '지출']
    if not expense_df.empty:
        category_sum = expense_df.groupby('main_category')['amount'].sum().sort_values(ascending=False)
        print(category_sum.to_string())
    else:
        print("No expenses found.")
        
    print("\n--- 🚀 Ready for App Development ---")
    print("This script proves we can read the raw data programmatically.")
    print("Next Step: Build a UI (Web/App) to add these rows without opening Excel.")

if __name__ == "__main__":
    analyze_expenses()
