import shutil
import os
import time

src = r'c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\20251214_Dashboard_v13_solved.xlsx'
dst = r'c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\v13_temp_work.xlsx'

print(f"Attempting to copy {src} to {dst}...")
try:
    shutil.copy2(src, dst)
    print("Copy successful via shutil.")
except Exception as e:
    print(f"Copy failed: {e}")
