import pandas as pd
import openpyxl

# Load Excel file
wb = openpyxl.load_workbook(
    r'c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\20251213_수식연결_가계부엔진.xlsx',
    data_only=True
)

ws = wb['📋 T_RawData']

# Extract data
data = []
for row in ws.iter_rows(min_row=3, values_only=True):
    if row[0] and row[2] and row[6]:
        data.append({
            '날짜': row[0],
            '구분': row[2],
            '대분류': row[3],
            '소분류': row[4],
            '내용': row[5],
            '금액': row[6],
            'Flow_Filter': row[9]
        })

df = pd.DataFrame(data)
df_filtered = df[df['Flow_Filter'] == 1].copy()
expense = df_filtered[df_filtered['구분'] == '지출'].copy()

# Create analysis report
with open(r'c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\지출분석_보고서.txt', 'w', encoding='utf-8') as f:
    f.write("=" * 80 + "\n")
    f.write("📊 지출 분석 보고서 (Flow_Filter = 1)\n")
    f.write("=" * 80 + "\n\n")
    
    f.write(f"총 지출 금액: {expense['금액'].sum():,.0f}원\n")
    f.write(f"거래 건수: {len(expense)}건\n\n")
    
    # Category analysis
    f.write("=" * 80 + "\n")
    f.write("🔥 대분류별 지출 (높은 순)\n")
    f.write("=" * 80 + "\n")
    cat_sum = expense.groupby('대분류')['금액'].sum().sort_values(ascending=False)
    cat_count = expense.groupby('대분류').size()
    total_expense = expense['금액'].sum()
    
    for idx, (cat_name, amt) in enumerate(cat_sum.items(), 1):
        pct = (amt / total_expense * 100) if total_expense > 0 else 0
        count = cat_count[cat_name]
        avg = amt / count if count > 0 else 0
        f.write(f"{idx}. {cat_name:15s}: {amt:12,.0f}원 ({pct:5.1f}%) | {count:3d}건 | 평균 {avg:,.0f}원\n")
    
    # Sub-category analysis
    f.write("\n" + "=" * 80 + "\n")
    f.write("📝 소분류별 지출 Top 20\n")
    f.write("=" * 80 + "\n")
    sub_sum = expense.groupby('소분류')['금액'].sum().sort_values(ascending=False)
    sub_count = expense.groupby('소분류').size()
    
    for idx, (sub_name, amt) in enumerate(sub_sum.head(20).items(), 1):
        pct = (amt / total_expense * 100) if total_expense > 0 else 0
        count = sub_count[sub_name]
        avg = amt / count if count > 0 else 0
        f.write(f"{idx:2d}. {sub_name:25s}: {amt:12,.0f}원 ({pct:5.1f}%) | {count:3d}건 | 평균 {avg:,.0f}원\n")
    
    # Monthly analysis
    f.write("\n" + "=" * 80 + "\n")
    f.write("📅 월별 지출 (높은 순)\n")
    f.write("=" * 80 + "\n")
    expense['월'] = pd.to_datetime(expense['날짜']).dt.to_period('M')
    monthly_sum = expense.groupby('월')['금액'].sum().sort_values(ascending=False)
    
    for month, amt in monthly_sum.items():
        f.write(f"{month}: {amt:12,.0f}원\n")

print("보고서 저장 완료: 지출분석_보고서.txt")
print("\n=== 요약 ===")
print(f"총 지출: {expense['금액'].sum():,.0f}원")
print(f"\nTop 3 대분류:")
for idx, (cat, amt) in enumerate(cat_sum.head(3).items(), 1):
    pct = (amt / total_expense * 100) if total_expense > 0 else 0
    print(f"{idx}. {cat}: {amt:,.0f}원 ({pct:.1f}%)")

print(f"\nTop 5 소분류:")
for idx, (sub, amt) in enumerate(sub_sum.head(5).items(), 1):
    pct = (amt / total_expense * 100) if total_expense > 0 else 0
    print(f"{idx}. {sub}: {amt:,.0f}원 ({pct:.1f}%)")
