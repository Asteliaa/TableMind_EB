"""Общие помощники ЛР6: копия шаблона, заполнение таблицы с пропуском формул, чтение результатов."""
import os
import shutil
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT6 = HERE.parent
sys.path.insert(0, str(ROOT6.parent / "ЛР2-3_Объем_рынка" / "Python"))
sys.stdout.reconfigure(encoding="utf-8")
import pythoncom
import win32com.client as wc
from xlh import errors, val  # noqa: F401

TPLDIR = Path(r"D:\Уник\2 ВЫШКА\ЭБ\Условия\ЛР6")
SERIAL = (date(2026, 10, 3) - date(1899, 12, 30)).days
NOSET = object()


def start():
    pythoncom.CoInitialize()
    app = wc.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    return app


def open_copy(app, tpl, out):
    dst = ROOT6 / "Excel" / out
    dst.parent.mkdir(exist_ok=True)
    shutil.copyfile(TPLDIR / tpl, dst)
    wb = app.Workbooks.Open(str(dst))
    for ws in wb.Worksheets:
        ws.Cells.Replace(What="Республика Беларусь", Replacement="Англоязычные страны и ЕС", LookAt=2)
    return wb


def fill_table(ws, first_row, rows, ncols, clear_rows=0):
    """Записывает rows начиная с first_row; ячейки с формулами не трогает; лишние демо-строки (clear_rows) очищает по входным столбцам."""
    total = max(len(rows), clear_rows)
    for i in range(total):
        r = first_row + i
        row = rows[i] if i < len(rows) else None
        for j in range(ncols):
            cell = ws.Cells.Item(r, j + 1)
            if cell.HasFormula:
                continue
            if row is None:
                cell.ClearContents()
                continue
            v = row[j] if j < len(row) else None
            if v is NOSET:
                continue
            if v is None:
                cell.ClearContents()
            else:
                cell.Value2 = v


def finish(wb):
    wb.BuiltinDocumentProperties("Author").Value = "Р. В. Земляник"
    wb.BuiltinDocumentProperties("Last Author").Value = "Р. В. Земляник"
    errs = errors(wb)
    wb.Save()
    wb.Close(False)
    return errs
