"""Дамп листов Excel-книги через COM (только чтение): python xl_dump.py <книга> <лист> [диапазон] [formulas]"""
import sys
import pythoncom
import win32com.client as wc

path, sheet = sys.argv[1], sys.argv[2]
rng = sys.argv[3] if len(sys.argv) > 3 else "A1:L40"
formulas = len(sys.argv) > 4 and sys.argv[4] == "formulas"
sys.stdout.reconfigure(encoding="utf-8")
pythoncom.CoInitialize()
xl = wc.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(path, ReadOnly=True)
    if sheet == "?":
        for ws in wb.Worksheets:
            print(ws.Name, ws.UsedRange.Address)
    else:
        ws = wb.Worksheets(sheet)
        print(sheet, ws.UsedRange.Address)
        for row in ws.Range(rng).Rows:
            cells = []
            for c in row.Cells:
                v = c.Formula if formulas else c.Value
                if v not in (None, ""):
                    cells.append(f"{c.GetAddress(False, False)}={str(v)[:90]}")
            if cells:
                print(" | ".join(cells))
    wb.Close(False)
finally:
    xl.Quit()
