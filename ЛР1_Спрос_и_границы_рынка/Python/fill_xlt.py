"""Заполнение рабочих копий Excel-шаблона «тренд, сезонность, деловые циклы» через COM (openpyxl не используется).

Запуск: python fill_xlt.py            # все книги
Результат: Excel/XLT_<код>_TableMind.xlsx и results/xlt_<код>.json (значения листов Панель, Сезонность, Оценка_тренда, Деловые_циклы).
"""
import datetime as dt
import json
import shutil
import sys
from pathlib import Path

import pythoncom
import win32com.client as wc

from data_prep import CLUSTERS, MONTHS, cluster_series, despike

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "Материалы" / "shablon_excel_ocenka_trenda_sezonnosti_delovyh_ciklov_google_trends_wordstat.xlsx"
OUT = ROOT / "Excel"
RES = ROOT / "Python" / "results"
OUT.mkdir(exist_ok=True)
RES.mkdir(exist_ok=True)
AUTHOR = "Р. В. Земляник"

# код книги: (кластер, вариант)
BOOKS = {
    "A_проверка_таблиц": ("A", "full"),
    "A_только_GT": ("A", "gt"),
    "A_только_YW": ("A", "yw"),
    "B_AI_для_Excel": ("B", "full"),
    "C_финмодели": ("C", "full"),
}

GT_FILES = {
    "A": "S_excel_audit_5y.txt; S_spreadsheet_audit_5y.txt",
    "B": "S_excel_copilot_5y.txt",
    "C": "S_financial_model_audit_5y.txt",
}
YW_LABEL = "YW_динамика_месяцы_2018-01_2026-09.csv"


def book_data(cluster, variant):
    g, y = cluster_series(cluster)
    notes = [""] * 60
    if cluster == "C":
        y, rep = despike(y)
        for m, old, new in rep:
            notes[MONTHS.index(m)] = f"разовый выброс Вордстата {old} заменен на {new:.0f} (правило: значение > 3 x второго по величине заменяется средним трех предыдущих месяцев)"
    if variant == "gt":
        y = [None] * 60
    if variant == "yw":
        g = [None] * 60
    return g, y, notes


def fill(code, cluster, variant):
    g, y, notes = book_data(cluster, variant)
    dst = OUT / f"XLT_{code}_TableMind.xlsx"
    shutil.copy(TEMPLATE, dst)
    pythoncom.CoInitialize()
    xl = wc.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(dst))
        ws = wb.Worksheets("Ввод_данных")
        c = CLUSTERS[cluster]
        label = c["name"]
        if variant == "gt":
            terr = "мир (Google Trends)"
        elif variant == "yw":
            terr = "Россия (Вордстат)"
        else:
            terr = "GT: мир; Вордстат: Россия"
        for i, m in enumerate(MONTHS):
            r = 5 + i
            y_, mo = int(m[:4]), int(m[5:])
            ws.Cells(r, 1).Value2 = float((dt.date(y_, mo, 1) - dt.date(1899, 12, 30)).days)
            ws.Cells(r, 2).Value = label
            if g[i] is None:
                ws.Cells(r, 3).ClearContents()
            else:
                ws.Cells(r, 3).Value = round(g[i], 2)
            if y[i] is None:
                ws.Cells(r, 4).ClearContents()
            else:
                ws.Cells(r, 4).Value = round(y[i], 2)
            ws.Cells(r, 5).Value = terr
            ws.Cells(r, 6).Value = notes[i]
        # журнал выгрузки
        jl = wb.Worksheets("Журнал_выгрузки")
        jl.Range("A4:H20").ClearContents()
        rows = []
        if variant in ("full", "gt"):
            rows.append(("2026-10-02", "Google Trends", c["gt_label"], "весь мир", "2021-10 - 2026-09 (5 лет, недели -> месяцы)", "текст (сжатый, base62)", GT_FILES[cluster], "веб-поиск, все категории; недели усреднены по месяцам; неполная последняя неделя исключена"))
        if variant in ("full", "yw"):
            rows.append(("2026-10-02", "Яндекс Вордстат", "; ".join(c["yw_queries"]), "Россия (код 225)", "2021-10 - 2026-09 (месяцы)", "CSV/JSON", YW_LABEL, "вкладка «Динамика», все устройства; значения запросов кластера суммированы"))
        for k, row in enumerate(rows):
            for j, v in enumerate(row):
                jl.Cells(4 + k, 1 + j).Value = v
        xl.Calculate()
        # результаты
        out = {"code": code, "cluster": cluster, "variant": variant, "name": label}
        p = wb.Worksheets("Панель")
        out["panel"] = {p.Cells(r, 2).Value: p.Cells(r, 3).Value for r in range(6, 14)}
        out["summary_text"] = p.Cells(17, 1).Value
        s = wb.Worksheets("Сезонность")
        out["season"] = [[s.Cells(r, 1).Value, s.Cells(r, 2).Value, s.Cells(r, 3).Value, s.Cells(r, 4).Value, s.Cells(r, 6).Value] for r in range(4, 16)]
        out["season_stats"] = {s.Cells(r, 1).Value: s.Cells(r, 2).Value for r in range(19, 25)}
        t = wb.Worksheets("Оценка_тренда")
        out["trend"] = {t.Cells(r, 1).Value: t.Cells(r, 2).Value for r in range(4, 12)}
        d = wb.Worksheets("Деловые_циклы")
        out["cycles"] = {d.Cells(r, 1).Value: d.Cells(r, 2).Value for r in range(5, 11)}
        rc = wb.Worksheets("Расчет")
        cols = {"G": 7, "H": 8, "I": 9, "J": 10, "K": 11, "L": 12}
        out["calc"] = [[(str(rc.Cells(5 + i, 1).Value)[:7])] + [rc.Cells(5 + i, c_).Value for c_ in cols.values()] for i in range(60)]
        # ошибки Excel
        errs = []
        for sh in wb.Worksheets:
            used = sh.UsedRange
            vals = used.Value
            if vals is None:
                continue
            for ri, row in enumerate(vals, start=1):
                for ci, v in enumerate(row, start=1):
                    if isinstance(v, int) and v < -2146826000:
                        errs.append(f"{sh.Name}!R{ri}C{ci}")
        out["errors"] = errs
        wb.BuiltinDocumentProperties("Author").Value = AUTHOR
        wb.BuiltinDocumentProperties("Last Author").Value = AUTHOR
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()

    def clean(o):
        if isinstance(o, (dt.datetime, dt.date)):
            return o.strftime("%Y-%m-%d")
        if isinstance(o, dict):
            return {str(k): clean(v) for k, v in o.items()}
        if isinstance(o, (list, tuple)):
            return [clean(x) for x in o]
        if hasattr(o, "isoformat"):
            return str(o)
        return o

    (RES / f"xlt_{code}.json").write_text(json.dumps(clean(out), ensure_ascii=False, indent=1), encoding="utf-8")
    print("OK", dst.name, "errors:", len(errs), "| trend:", out["panel"].get("Классификация тренда"), "| season:", out["panel"].get("Сила сезонности"), "| phase:", out["panel"].get("Текущая фаза"))


if __name__ == "__main__":
    only = sys.argv[1:]
    for code, (cl, var) in BOOKS.items():
        if only and code not in only:
            continue
        fill(code, cl, var)
