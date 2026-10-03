"""Полный текст формул выбранных ячеек: python xl_cells.py <книга> <лист> <ячейка> [<ячейка> ...]"""
import sys
import pythoncom
import win32com.client as wc

sys.stdout.reconfigure(encoding="utf-8")
path, sheet, cells = sys.argv[1], sys.argv[2], sys.argv[3:]
pythoncom.CoInitialize()
xl = wc.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(path, ReadOnly=True)
    ws = wb.Worksheets(sheet)
    for c in cells:
        print(c, "::", ws.Range(c).Formula)
    wb.Close(False)
finally:
    xl.Quit()
