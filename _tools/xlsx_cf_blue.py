"""Убирает цветные условные форматы (красный/зеленый) в книгах Excel, гистограммы оставляет синими. Запуск: python xlsx_cf_blue.py книга ..."""
import sys
from pathlib import Path

import pythoncom
import win32com.client as wc


def run(path):
    pythoncom.CoInitialize()
    xl = wc.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    n = 0
    try:
        wb = xl.Workbooks.Open(str(path))
        for ws in wb.Worksheets:
            try:
                fcs = ws.Cells.FormatConditions
                for k in range(fcs.Count, 0, -1):
                    fc = fcs.Item(k)
                    if fc.Type == 4:
                        fc.BarColor.Color = int("D5" + "9B" + "5B", 16)
                    else:
                        fc.Delete()
                        n += 1
            except Exception:
                pass
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    print("OK", Path(path).name, "условных форматов удалено", n)


if __name__ == "__main__":
    for a in sys.argv[1:]:
        run(a)
