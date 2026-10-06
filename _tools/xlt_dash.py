"""ЛР1: в книгах «только GT» и «только YW» пустой столбец второго источника заполняется прочерком; формулы листа «Расчет» учитывают прочерк как отсутствие данных."""
import re
import sys
from pathlib import Path

import pythoncom
import win32com.client as wc


def run(path):
    pythoncom.CoInitialize()
    xl = wc.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(path))
        xl.Calculation = -4135
        v = wb.Worksheets("Ввод_данных")
        c = wb.Worksheets("Расчет")
        h = v.Range("A1:A20").Find("Период", LookAt=1)
        r1 = h.Row + 1
        n = v.Cells(v.Rows.Count, 1).End(-4162).Row
        for r in range(r1, n + 1):
            for col in (3, 4):
                if v.Cells(r, col).Value2 in (None, ""):
                    v.Cells(r, col).Value = "-"
                    v.Cells(r, col).HorizontalAlignment = -4108
        hc = c.Range("A1:A20").Find("Период", LookAt=1)
        cr1 = hc.Row + 1
        for r in range(cr1, cr1 + (n - r1) + 1):
            for col in (4, 5):
                f = c.Cells(r, col).Formula
                nf = re.sub(r'(Ввод_данных!([CD])(\d+))=""', r'OR(\1="",\1="-")', f, count=1)
                if nf != f:
                    c.Cells(r, col).Formula = nf
        xl.Calculation = -4105
        xl.Calculate()
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    print("OK", Path(path).name)


if __name__ == "__main__":
    for a in sys.argv[1:]:
        run(a)
