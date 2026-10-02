import sys, shutil, pythoncom, win32com.client as wc
sys.stdout.reconfigure(encoding="utf-8")
src, dst = sys.argv[1], sys.argv[2]
shutil.copy(src, dst)
pythoncom.CoInitialize()
xl = wc.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
wb = xl.Workbooks.Open(dst)
wb.Save(); wb.Close(False)
try:
    wb = xl.Workbooks.Open(dst)
    print("reopen OK")
    wb.Close(False)
except Exception as e:
    print("reopen FAIL", str(e)[:80])
xl.Quit()
