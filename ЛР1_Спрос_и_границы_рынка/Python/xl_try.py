import sys, pythoncom, win32com.client as wc
sys.stdout.reconfigure(encoding="utf-8")
path = sys.argv[1]
pythoncom.CoInitialize()
xl = wc.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
for kw in ({}, {"ReadOnly": False}, {"CorruptLoad": 1}, {"UpdateLinks": 0, "ReadOnly": False}):
    try:
        wb = xl.Workbooks.Open(path, **kw)
        print("OK with", kw, wb.Worksheets.Count, wb.Worksheets("Расчет").Range("A5").Value)
        wb.Close(False)
        break
    except Exception as e:
        print("FAIL", kw, str(e)[:100])
xl.Quit()
