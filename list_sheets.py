import openpyxl

wb = openpyxl.load_workbook(
    r'c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\20251214_수식연결_가계부엔진_최종.xlsx',
    data_only=False
)

with open('sheets.txt', 'w', encoding='utf-8') as f:
    f.write("시트 목록:\n")
    f.write("=" * 60 + "\n")
    for idx, sheet_name in enumerate(wb.sheetnames, 1):
        f.write(f"{idx}. {sheet_name}\n")
        print(f"{idx}. {sheet_name}")
    f.write("=" * 60 + "\n")

print("\nsheets.txt에 저장 완료")
