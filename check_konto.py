# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import openpyxl

# Load WITHOUT data_only to see actual cell values (not just formula cache)
wb = openpyxl.load_workbook(
    r"D:\Workstation_Yuhong\Yuhong's_File\Python\Buchhaltung\Kontoauszug_Ebay.xlsx",
    data_only=False)
ws = wb["Ebay_Umsatz"]
print(f"Total rows (incl. header): {ws.max_row}")
for i, row in enumerate(ws.iter_rows(min_row=1, values_only=False)):
    if any(c.value is not None for c in row):
        vals = [c.value for c in row]
        print(f"  row{i+1}: {vals}")
