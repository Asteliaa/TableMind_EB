"""Помощники работы с Excel через COM (копии шаблонов, запись значений, вставка строк, чтение результатов)."""
import shutil
from pathlib import Path

import pythoncom
import win32com.client as wc

ROOT = Path(__file__).resolve().parent.parent
TPL = Path(r"D:\Уник\2 ВЫШКА\ЭБ\Условия\ЛР3")
OUT = ROOT / "Excel"
AUTHOR = "Р. В. Земляник"


class XL:
    def __init__(self):
        pythoncom.CoInitialize()
        self.app = wc.DispatchEx("Excel.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.wb = None

    def open_copy(self, tpl_name, out_name):
        OUT.mkdir(exist_ok=True)
        dst = OUT / out_name
        shutil.copyfile(TPL / tpl_name, dst)
        self.wb = self.app.Workbooks.Open(str(dst))
        # валюта шаблонов - BYN/руб.; расчет ЛР2-3 ведется в EUR, демонстрационная дата заменяется датой снимка
        for ws in self.wb.Worksheets:
            for old, new in (("BYN", "EUR"), ("руб.", "EUR"), ("2026-05-02", "2026-10-03")):
                ws.Cells.Replace(What=old, Replacement=new, LookAt=2)
        return self.wb

    def ws(self, name):
        return self.wb.Worksheets(name)

    def save_close(self):
        try:
            self.wb.BuiltinDocumentProperties("Author").Value = AUTHOR
            self.wb.BuiltinDocumentProperties("Last Author").Value = AUTHOR
        except Exception:
            pass
        self.app.CalculateFull()
        self.wb.Save()
        self.wb.Close(False)
        self.wb = None

    def quit(self):
        try:
            self.app.Quit()
        except Exception:
            pass


def put(ws, addr, value):
    """Запись значения; строка с '=' - формула; None - очистка содержимого."""
    c = ws.Range(addr)
    if value is None:
        c.ClearContents()
    elif isinstance(value, str) and value.startswith("="):
        c.Formula = value
    else:
        c.Value2 = value


def put_row(ws, row, col0, values):
    """Записывает значения слева направо с колонки col0 (буква); None пропускается (оставляет шаблон)."""
    c0 = ord(col0) - 64
    for i, v in enumerate(values):
        if v is NOSET:
            continue
        put(ws, f"{chr(64 + c0 + i)}{row}", v)


NOSET = object()


def insert_rows_copy(ws, at_row, n, src_row, last_col="Z"):
    """Вставляет n строк перед строкой at_row (внутри диапазонов суммирования) и копирует в них формат и формулы строки src_row.

    Если src_row >= at_row, после вставки номер исходной строки сдвигается на n."""
    ws.Rows(f"{at_row}:{at_row + n - 1}").Insert()
    src = src_row + n if src_row >= at_row else src_row
    for r in range(at_row, at_row + n):
        ws.Range(f"A{src}:{last_col}{src}").Copy(ws.Range(f"A{r}:{last_col}{r}"))
    return src


def val(ws, addr):
    return ws.Range(addr).Value2


def errors(wb):
    """Список ячеек с ошибками формул во всей книге."""
    res = []
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        vals = ur.Value2
        if vals is None:
            continue
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        r0, c0 = ur.Row, ur.Column
        for i, row in enumerate(vals):
            for j, v in enumerate(row):
                if isinstance(v, int) and v < -2146826000:
                    res.append(f"{ws.Name}!{ws.Cells.Item(r0 + i, c0 + j).GetAddress(False, False)}")
    return res
