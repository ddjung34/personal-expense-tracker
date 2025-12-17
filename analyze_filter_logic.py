import pandas as pd
import openpyxl

# Load the file with VALID cached values
source_file = r'c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\20251213_수식연결_가계부엔진.xlsx'

print(f"Analyzing Filter Logic from: {source_file}")

wb = openpyxl.load_workbook(source_file, data_only=True)
ws = wb['📋 T_RawData']

data = []
for row in ws.iter_rows(min_row=3, values_only=True):
    if row[0] is None: continue
    
    # 2=구분, 3=대분류, 9=J(Filter), 10=K(Filter?)
    flow = row[2]
    category = row[3]
    val_j = row[9]
    val_k = row[10]
    
    # Check filter value (1 or 0/None)
    is_valid = False
    if str(val_j).strip() == '1' or str(val_j).strip() == '1.0':
        is_valid = True
    elif str(val_k).strip() == '1' or str(val_k).strip() == '1.0':
        is_valid = True
        
    data.append({
        'Flow': flow,
        'Category': category,
        'Filter_Value': 1 if is_valid else 0
    })

df = pd.DataFrame(data)

# Group by Category and see Filter Status
# We care about '지출'
expense_df = df[df['Flow'] == '지출']

print("\n[Analysis] Categories with Filter = 0 (Excluded):")
excluded = expense_df[expense_df['Filter_Value'] == 0]['Category'].unique()
print(excluded)

print("\n[Analysis] Categories with Filter = 1 (Included):")
included = expense_df[expense_df['Filter_Value'] == 1]['Category'].unique()
print(included)

# Verify if a category is ALWAYS 0 or ALWAYS 1
print("\n[Analysis] Mixed Categories (Sometimes 0, Sometimes 1):")
for cat in expense_df['Category'].unique():
    stats = expense_df[expense_df['Category'] == cat]['Filter_Value'].unique()
    if len(stats) > 1:
        print(f" - {cat}: {stats}")
