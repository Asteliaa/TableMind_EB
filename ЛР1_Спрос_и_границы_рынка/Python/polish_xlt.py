"""Доработка пяти книг XLT после fill_xlt.py (через Excel COM).

Что делает.
1. Заполняет пустые ячейки данными: столбец «Комментарий» листа «Ввод_данных» (по рядам), пояснения для первых 11 месяцев
   на листах «Расчет» и «Деловые_циклы» (окно скользящего среднего 12 месяцев еще не накоплено), источники трех разделов
   листа «Методика_и_источники».
2. Приводит тексты шаблона к стилю без двоеточий внутри предложений и без длинного тире.
Запуск: python polish_xlt.py
"""
import sys
from pathlib import Path

import pythoncom
import win32com.client as wc

ROOT = Path(__file__).resolve().parent.parent
BOOKS = sorted((ROOT / "Excel").glob("XLT_*_TableMind.xlsx"))
NL = chr(10)

TEXTS = {
    ("Панель", 3): "Производится анализ интереса аудитории на основе выгрузок из Google Trends и Яндекс Вордстат.",
    ("Ввод_данных", 2): "Google Trends вводится как индекс 0-100. " + NL + "Яндекс Вордстат вводится как абсолютное число запросов по выбранному семантическому кластеру. ",
    ("Расчет", 2): "Расчет построен по мультипликативной логике. " + NL + "Наблюдаемый интерес = тренд × сезонность × циклическое отклонение × случайные колебания. " + NL
                   + "Для практической оценки используются скользящее среднее, сезонные индексы и коэффициент отклонения от трендово-сезонного уровня.",
    ("Деловые_циклы", 2): "В таблице оценивается поисковый цикл спроса, то есть периоды, когда интерес находится выше или ниже трендово-сезонного уровня. "
                         ,
}
CELLS = {
    ("Панель", "D5"): "Использование",
    ("Панель", "A16"): "Вывод:",
    ("Ввод_данных", "C4"): "Google Trends, 0-100",
    ("Расчет", "F4"): "Вордстат норм. 0-100",
    ("Расчет", "H4"): "Скользящее среднее 12 мес. (тренд)",
}
SUMMARY = ('=Оценка_тренда!B11&" Для ряда характерна "&Сезонность!B22&", максимум интереса приходится на "&Сезонность!B23'
           '&", минимум на "&Сезонность!B24&". Текущая фаза спроса - "&Деловые_циклы!B7&"."')
SOURCES = {
    6: "https://otexts.com/fpp3/moving-averages.html",
    7: "https://otexts.com/fpp3/classical-decomposition.html",
    8: "https://otexts.com/fpp3/tspatterns.html",
}
WINDOW_NOTE = "окно скользящего среднего 12 мес. еще не накоплено, фаза цикла не определяется"


def rel(v, series, prev):
    """Описание значения по положению в ряду (квартили) и скачку к предыдущему месяцу."""
    vals = sorted(x for x in series if x is not None)
    mean = sum(vals) / len(vals)
    q1 = vals[len(vals) // 4]
    q2 = vals[len(vals) // 2]
    q3 = vals[(3 * len(vals)) // 4]
    if v == 0:
        s = "нулевое значение, ниже порога учета выгрузки"
    elif v == vals[-1]:
        s = "максимум ряда"
    elif v == vals[0]:
        s = "минимум ряда"
    elif v >= q3:
        s = "в верхней четверти значений ряда"
    elif v >= q2:
        s = "выше медианы ряда"
    elif v > q1:
        s = "ниже медианы ряда"
    else:
        s = "в нижней четверти значений ряда"
    if prev and prev >= 0.25 * mean and v > 0:
        ch = (v / prev - 1) * 100
        if ch >= 30:
            s += f", рост к предыдущему месяцу на {ch:.0f} %"
        elif ch <= -25:
            s += f", снижение к предыдущему месяцу на {-ch:.0f} %"
    return s


def polish(path):
    pythoncom.CoInitialize()
    xl = wc.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(path))
        if wb.Worksheets("Ввод_данных").Cells(5, 6).Value:
            print("уже обработана", path.name)
            wb.Close(False)
            return
        for (sh, r), t in TEXTS.items():
            wb.Worksheets(sh).Cells(r, 1).Value = t
        for (sh, a), t in CELLS.items():
            wb.Worksheets(sh).Range(a).Value = t
        wb.Worksheets("Панель").Range("A17").Formula = SUMMARY
        vv = wb.Worksheets("Ввод_данных")
        g = [vv.Cells(5 + i, 3).Value for i in range(60)]
        y = [vv.Cells(5 + i, 4).Value for i in range(60)]
        for i in range(60):
            r = 5 + i
            e = vv.Cells(r, 5).Value or ""
            if e.startswith("GT:"):
                vv.Cells(r, 5).Value = "Google Trends - весь мир, Вордстат - Россия"
            parts = []
            if g[i] is not None:
                parts.append("Google Trends " + rel(g[i], g, g[i - 1] if i else None))
            if y[i] is not None:
                parts.append("Вордстат " + rel(y[i], y, y[i - 1] if i else None))
            old = vv.Cells(r, 6).Value or ""
            if old:
                old = old.replace("(правило: значение > 3 x", "по правилу, значение больше 3 x")
                if old.endswith(")"):
                    old = old[:-1]
                parts.append(old)
            if i == 0:
                parts.append("начало ряда")
            if i == 59:
                parts.append("конец ряда, выгрузка от 02.10.2026")
            vv.Cells(r, 6).Value = "; ".join(parts)
        rc = wb.Worksheets("Расчет")
        dc = wb.Worksheets("Деловые_циклы")
        for r in range(5, 16):
            rc.Cells(r, 13).Formula = f'=IF(G{r}="","","{WINDOW_NOTE}")'
        for r in range(14, 25):
            h = dc.Cells(r, 8).Formula
            dc.Cells(r, 8).Formula = f'=IF(B{r}="","","окно 12 мес. не накоплено, период входит в расчет сезонности после 12-го месяца")'
        ms = wb.Worksheets("Методика_и_источники")
        for r, u in SOURCES.items():
            ms.Cells(r, 6).Value = u
        xl.Calculate()
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    print("OK", path.name)


if __name__ == "__main__":
    only = sys.argv[1:]
    for b in BOOKS:
        if only and not any(o in b.name for o in only):
            continue
        polish(b)
