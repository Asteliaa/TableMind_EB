import sys, shutil, pythoncom, win32com.client as wc
sys.stdout.reconfigure(encoding="utf-8")
T = sys.argv[1]
base = r"C:\Users\User\AppData\Local\Temp\xt"
pythoncom.CoInitialize()
xl = wc.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False

def test(name, fn):
    dst = f"{base}\{name}.xlsx"
    shutil.copy(T, dst)
    wb = xl.Workbooks.Open(dst)
    fn(wb)
    wb.Save(); wb.Close(False)
    try:
        w2 = xl.Workbooks.Open(dst); w2.Close(False); print(name, "OK")
    except Exception as e:
        print(name, "FAIL")

def data(wb):
    ws = wb.Worksheets("Ввод_данных")
    for r in range(5, 65):
        ws.Cells(r, 1).Value2 = 44500.0 + r
        ws.Cells(r, 2).Value = "x"
        ws.Cells(r, 3).Value = 5.5
        ws.Cells(r, 4).Value = 100
def datafmt(wb):
    ws = wb.Worksheets("Ввод_данных")
    for r in range(5, 65):
        ws.Cells(r, 1).NumberFormat = "mm.yyyy"
def journal(wb):
    jl = wb.Worksheets("Журнал_выгрузки")
    jl.Range("A4:H20").ClearContents()
    jl.Cells(4, 1).Value = "2026-10-02"
def clr(wb):
    ws = wb.Worksheets("Ввод_данных")
    ws.Cells(10, 3).ClearContents()
def props(wb):
    wb.BuiltinDocumentProperties("Author").Value = "Р. В. Земляник"
test("t_data", data); test("t_fmt", datafmt); test("t_journal", journal); test("t_clr", clr); test("t_props", props)
xl.Quit()
