"""ЛР4: заполнение шаблона «оценка конкуренции и барьеров входа Similarweb» через COM, результаты в results/excel_results.json."""
import csv
import json
import os
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT4 = HERE.parent
sys.path.insert(0, str(ROOT4.parent / "ЛР2-3_Объем_рынка" / "Python"))
sys.stdout.reconfigure(encoding="utf-8")

import pythoncom
import win32com.client as wc

from inputs import *          # noqa: E402  (данные снимка Similarweb и допущения ЛР2-3)
from xlh import put, put_row, val, errors, NOSET  # noqa: E402

TPL = Path(r"D:\Уник\2 ВЫШКА\ЭБ\Условия\ЛР4\shablon_excel_ocenka_konkurencii_i_barerov_vhoda_similarweb.xlsx")
OUT = ROOT4 / "Excel" / "ЛР4_1_конкуренция_и_барьеры_входа_TableMind.xlsx"

# --- вспомогательные данные ---
def secs(s):
    h, m, sec = s.split(":")
    return int(h) * 3600 + int(m) * 60 + int(sec)


REP = {}
with open(ROOT4 / "Материалы_собранные" / "репутация_2026-10-03.csv", encoding="utf-8-sig") as f:
    for r in csv.DictReader(f, delimiter=";"):
        REP[r["domain"]] = r

# экспертные баллы технологической зрелости и автоматизации 1-5 (допущение Д-17; проверить по сайтам)
TECH = {"perfectxl.com": (3, 3), "operis.com": (3, 2), "arixcel.com": (2, 2), "spreadsheetdetective.com": (2, 1), "cimcon.com": (4, 4),
        "shortcut.ai": (5, 5), "rows.com": (5, 4), "numerous.ai": (4, 4), "formulabot.com": (3, 3), "ajelix.com": (3, 3), "datarails.com": (5, 5), "julius.ai": (5, 4)}
TYPE_RU = {"прямой": "прямой конкурент", "частичный": "частичный конкурент", "заменитель": "заменитель"}

QUERIES = [  # запрос, намерение, файл GT (индекс, среднее за 12 мес.)
    ("excel audit", "коммерческий", "S_excel_audit_5y"), ("spreadsheet audit", "коммерческий", "S_spreadsheet_audit_5y"),
    ("excel formula checker", "инструментальный", "S_excel_formula_checker_5y"), ("financial model audit", "коммерческий", "S_financial_model_audit_5y"),
    ("excel copilot", "AI-альтернатива", "S_excel_copilot_5y"), ("perfectxl", "брендовый", "S_perfectxl_5y"),
]


# ключевые слова Similarweb Keyword research (ставка клика в EUR - середина диапазона, конкуренция PPC 0-1, сложность SEO 0-100)
KW = {"excel formulas": (3.87, 0.01, 15), "financial modeling": (4.57, 0.26, 41), "excel copilot": (4.73, 0.01, 35),
      "chatgpt excel": (5.48, 0.30, 33), "excel ai": (6.37, 0.09, 42)}
FIN = {  # финансирование (EUR по курсу НБРБ) и размер по открытым источникам (деловая пресса, каталоги)
    "datasnipper.com": ("89 млн EUR (Series B, 2024)", "около 289 чел."), "datarails.com": ("155 млн EUR (привлечено всего)", "более 400 чел."),
    "shortcut.ai": ("более 35 млн EUR (seed и Series A)", "-"), "operis.com": ("выручка 6,7 млн EUR (2022)", "51-100 чел."),
    "rows.com": ("около 39 млн EUR", "около 50 чел."), "arixcel.com": ("не раскрыто", "1-10 чел."), "perfectxl.com": ("не раскрыто", "около 4 чел."),
    "numerous.ai": ("сумма не раскрыта", "1-10 чел."), "formulabot.com": ("без внешнего финансирования", "-"),
}
QUERIES2 = [("excel formulas", "информационный", None), ("financial modeling", "информационный", None), ("excel copilot", "AI-альтернатива", "S_excel_copilot_5y"),
            ("chatgpt excel", "AI-альтернатива", None), ("excel ai", "AI-альтернатива", None),
            ("excel audit", "коммерческий", "S_excel_audit_5y"), ("spreadsheet audit", "коммерческий", "S_spreadsheet_audit_5y"),
            ("excel formula checker", "инструментальный", "S_excel_formula_checker_5y"), ("financial model audit", "коммерческий", "S_financial_model_audit_5y"),
            ("perfectxl", "брендовый", "S_perfectxl_5y")]


def gt_avg(stem):
    sys.path.insert(0, str(ROOT4.parent / "ЛР1_Спрос_и_границы_рынка" / "Python"))
    import data_prep as dp
    s = dp.gt(stem)
    return sum(s[m] for m in dp.MONTHS[-12:]) / 12


def main():
    sw = load_sw()
    n = len(sw)
    OUT.parent.mkdir(exist_ok=True)
    pythoncom.CoInitialize()
    app = wc.DispatchEx("Excel.Application"); app.Visible = False; app.DisplayAlerts = False
    # шаблон открывается только с восстановлением (CorruptLoad = xlRepairFile)
    wb = app.Workbooks.Open(os.path.normpath(str(TPL)), 0, False, None, None, None, None, None, None, None, None, None, None, None, 1)
    wb.SaveAs(str(OUT), 51)
    wb.Close(False)
    wb = app.Workbooks.Open(str(OUT))
    try:
        for ws in wb.Worksheets:
            for old, new in (("Республика Беларусь", "Англоязычные страны и ЕС"), ("2026-05-02", "2026-10-03")):
                ws.Cells.Replace(What=old, Replacement=new, LookAt=2)
        W = lambda name: wb.Worksheets(name)

        ws = W("01_Параметры")
        put(ws, "B5", "Онлайн-проверка Excel-моделей и финансовых таблиц (TableMind)")
        put(ws, "B6", "Англоязычные страны (США, Великобритания, Канада, Австралия, Ирландия) и ЕС-27")
        put(ws, "B7", "06-08.2026 (три месяца), весь мир; снимок 03.10.2026")
        # исправление шаблона: строки 8-10 содержали формулы со сдвигом ссылок
        put(ws, "B8", SNAP_SERIAL)
        put(ws, "B9", "=COUNTA('02_Конкуренты_SW'!A4:A33)")
        put(ws, "B10", "=SUM('02_Конкуренты_SW'!E4:E33)")
        put(ws, "C8", "дата"); put(ws, "C9", "сайтов"); put(ws, "C10", "визитов/мес.")
        put(ws, "E8", "Дата снимка Similarweb"); put(ws, "E9", "Число сайтов выборки"); put(ws, "E10", "Релевантный трафик: визиты x география x релевантность x качество")

        ws = W("02_Конкуренты_SW")
        put(ws, "D3", "Доля целевой географии x релевантность x качество")
        for i in range(30):
            row = 4 + i
            if i < n:
                r = sw[i]
                ch = r["ch"]
                put_row(ws, row, "A", [r["domain"], TYPE_RU[r["type"]] + ": " + r["segment"], r["visits_month"], round(r["geo_share"] * r["rel"] * r["qual"], 4), NOSET, NOSET,
                                       secs(r["duration"]), float(r["pages_visit"]), (float(r["bounce_pct"]) / 100) if r["bounce_pct"] else None,
                                       ch["direct"], ch["org_search"], ch["paid_search"], ch["org_social"] + ch["paid_social"], ch["referrals"], ch["display"], None,
                                       "Similarweb Pro, 06-08.2026, 03.10.2026; брендовый поиск не снимался"])
            else:
                put_row(ws, row, "A", [None, None, None, None, NOSET, NOSET, None, None, None, None, None, None, None, None, None, None, None])

        ws = W("04_Поиск_реклама")
        # данные CPC, конкурентности PPC, SEO difficulty, выдачи и рекламных библиотек недоступны без Semrush/Google Ads: ДОСНЯТЬ
        for r_ in range(4, 34):
            for c in "ABCDEFGHIJM":
                put(ws, f"{c}{r_}", None)
        for i, (q, intent, stem) in enumerate(QUERIES2):
            r_ = 4 + i
            kw = KW.get(q)
            gtv = round(gt_avg(stem), 1) if stem else "-"
            put_row(ws, r_, "A", [q, intent, "Google Trends (индекс 0-100, ЛР1)" + ("; Similarweb Keyword research" if kw else ""), gtv,
                                   kw[0] if kw else "-", kw[1] if kw else "-", kw[2] if kw else "-", "-", "-", "-"])
            put(ws, f"M{r_}", "ставка клика пересчитана в EUR по курсу НБРБ; конкуренция PPC и сложность SEO из Similarweb Keyword research" if kw else "запрос с малым объемом, данные ставок и сложности в Similarweb отсутствуют")
        for r_ in range(4, 34):
            put(ws, f"K{r_}", f'=IF(COUNT(E{r_}:J{r_})<2,"",MIN(1,(IF(ISNUMBER(E{r_}),MIN(E{r_}/2,1),0)+IF(ISNUMBER(F{r_}),F{r_},0)+IF(ISNUMBER(G{r_}),G{r_}/100,0)+IF(ISNUMBER(H{r_}),MIN(H{r_}/10,1),0)+IF(ISNUMBER(J{r_}),MIN(J{r_}/50,1),0))/MAX(1,COUNT(E{r_}:H{r_},J{r_}))))')
            put(ws, f"L{r_}", f'=IF(K{r_}="","",IF(K{r_}>=0.7,5,IF(K{r_}>=0.4,3,1)))')
        put(ws, "D3", "Индекс Google Trends (спрос)")
        put(ws, "B37", '=IFERROR(AVERAGE(K4:K33),"-")')
        put(ws, "C37", '=IF(ISNUMBER(B37),IF(B37>=0.7,"высокое",IF(B37>=0.4,"среднее","низкое")),"-")')
        put(ws, "B38", '=IFERROR(COUNTIF(L4:L33,5)/COUNT(L4:L33),"-")')
        put(ws, "C38", '=IF(ISNUMBER(B38),IF(B38>=0.5,"высокая доля",IF(B38>=0.25,"средняя доля","низкая доля")),"-")')
        put(ws, "D39", "Ставки и сложность выдачи получены для пяти массовых запросов; в индексе барьеров используется канальный прокси Similarweb (лист 07, строка 8), так как выборка ставок мала.")

        ws = W("05_Отзывы_рейтинги")
        for r_ in range(4, 34):
            for c in "ABCDEFGHIJKN":
                put(ws, f"{c}{r_}", None)
        for i, r in enumerate(sw):
            r_ = 4 + i
            d = REP[r["domain"]]
            num = lambda s: float(s) if s not in ("", None) else None
            g2r, g2n = num(d["g2_rating"]), num(d["g2_reviews"])
            put_row(ws, r_, "A", [r["domain"], g2r, g2n, num(d["capterra_rating"]), num(d["capterra_reviews"]), num(d["trustpilot_rating"]), num(d["trustpilot_reviews"]), None, None, None, None])
            put(ws, f"N{r_}", d["note"])
        for r_ in range(4, 34):
            put(ws, f"L{r_}", f'=IF(COUNT(B{r_}:K{r_})=0,"",MIN(1,(IF(COUNT(B{r_},D{r_},F{r_},H{r_})>0,AVERAGE(B{r_},D{r_},F{r_},H{r_})/5,0)+IF(COUNT(C{r_},E{r_},G{r_},I{r_})>0,MIN(SUM(C{r_},E{r_},G{r_},I{r_})/300,1),0)+IF(ISNUMBER(J{r_}),MIN(J{r_}/30,1),0)+IF(ISNUMBER(K{r_}),MIN(K{r_}/10,1),0))/MAX(1,(COUNT(B{r_},D{r_},F{r_},H{r_})>0)+(COUNT(C{r_},E{r_},G{r_},I{r_})>0)+ISNUMBER(J{r_})+ISNUMBER(K{r_}))))')
            put(ws, f"M{r_}", f'=IF(L{r_}="","",IF(L{r_}>=0.7,5,IF(L{r_}>=0.4,3,1)))')
        put(ws, "B38", "=SUM(C4:C33,E4:E33,G4:G33,I4:I33)/COUNTA(A4:A33)")
        put(ws, "D38", "Среднее число отзывов на конкурента (G2, Capterra, Trustpilot); отзывы Google, кейсы и возраст бренда не собирались")

        ws = W("06_Технологии_финансы")
        for r_ in range(4, 34):
            for c in "ABCDEFGHIJM":
                put(ws, f"{c}{r_}", None)
        for i, r in enumerate(sw):
            r_ = 4 + i
            put(ws, f"A{r_}", r["domain"]); put(ws, f"D{r_}", TECH[r["domain"]][0]); put(ws, f"E{r_}", TECH[r["domain"]][1])
            fin = FIN.get(r["domain"], ("-", "-"))
            put(ws, f"B{r_}", fin[0]); put(ws, f"C{r_}", fin[1])
            put(ws, f"M{r_}", "D и E - экспертные баллы (Д-17); финансирование и штат - деловая пресса (приложение Д ЛР5), вакансии и объявления не собирались")
        for r_ in range(4, 34):
            put(ws, f"K{r_}", f'=IF(COUNTA(B{r_}:J{r_})=0,"",MIN(1,(IF(ISNUMBER(B{r_}),MIN(B{r_}/500000,1),0)+IF(ISNUMBER(C{r_}),MIN(C{r_}/60,1),0)+IF(ISNUMBER(D{r_}),D{r_}/5,0)+IF(ISNUMBER(E{r_}),E{r_}/5,0)+IF(F{r_}="да",1,0)+IF(G{r_}="да",1,0)+IF(ISNUMBER(H{r_}),MIN(H{r_}/50,1),0)+IF(ISNUMBER(I{r_}),MIN(I{r_}/6,1),0)+IF(ISNUMBER(J{r_}),IF(J{r_}>0,1,0),0))/MAX(1,ISNUMBER(B{r_})+ISNUMBER(C{r_})+ISNUMBER(D{r_})+ISNUMBER(E{r_})+(F{r_}<>"")+(G{r_}<>"")+ISNUMBER(H{r_})+ISNUMBER(I{r_})+ISNUMBER(J{r_}))))')
            put(ws, f"L{r_}", f'=IF(K{r_}="","",IF(K{r_}>=0.7,5,IF(K{r_}>=0.4,3,1)))')
        put(ws, "B38", '=(COUNTIF(F4:F33,"да")+COUNTIF(G4:G33,"да"))/(2*COUNTA(A4:A33))')

        ws = W("07_Барьеры_входа")
        # исправления: сумма весов 1,04 нормируется; поисково-рекламная конкуренция при отсутствии данных - прокси по каналам Similarweb
        put(ws, "C8", "=AVERAGE(MIN('03_Каналы_трафика'!C5/'01_Параметры'!B15,1),MIN('03_Каналы_трафика'!C6/'01_Параметры'!B14,1))")
        put(ws, "B8", "Поисково-рекламное давление: органика и платный поиск Similarweb к порогам (ставки и сложность SEO приведены на листе 04)")
        put(ws, "C11", 0.6); put(ws, "B11", "Доля планируемых каналов через внешние платформы (AppSource, Product Hunt, поиск) из каналов ПЗ1")
        put(ws, "C12", "=IFERROR(1-SUMPRODUCT(('02_Конкуренты_SW'!B4:B33<>\"\")*(LEFT('02_Конкуренты_SW'!B4:B33,6)=\"прямой\")*'02_Конкуренты_SW'!E4:E33)/SUM('02_Конкуренты_SW'!E4:E33),0)")
        put(ws, "B12", "Доля релевантного трафика, приходящаяся на частичных конкурентов и заменителей")
        for r_, w_ in zip(range(4, 13), (0.14, 0.10, 0.06, 0.10, 0.12, 0.12, 0.14, 0.08, 0.14)):
            put(ws, f"E{r_}", w_)
        put(ws, "F13", "=SUM(F4:F12)/SUM(E4:E12)"); put(ws, "C13", "=F13")
        put(ws, "H13", "Интегральная оценка давления - сумма взвешенных баллов; веса заданы с суммой 1,00.")

        ws = W("08_Дашборд")
        put(ws, "A15", '="По результатам оценки цифровой конкуренции индекс барьеров входа составляет "&ROUND(B5,2)&" из 5, что соответствует уровню: "&B6&". "&"Наиболее значимые факторы давления определены по листу 07_Барьеры_входа. "&"Данные Similarweb и альтернативных источников являются оценочной базой и требуют проверки "&"через фактические коммерческие данные, цены, конверсии и платежеспособность клиентов."')

        ws = W("09_Источники")
        for r_ in range(4, 20):
            if ws.Range(f"A{r_}").Value2:
                put(ws, f"F{r_}", "2026-10-03")
        app.CalculateFull()

        w2, w3, w5, w6, w7, w8 = (W(x) for x in ("02_Конкуренты_SW", "03_Каналы_трафика", "05_Отзывы_рейтинги", "06_Технологии_финансы", "07_Барьеры_входа", "08_Дашборд"))
        res = {
            "n": val(W("01_Параметры"), "B9"), "traffic": val(W("01_Параметры"), "B10"),
            "cr3": val(w2, "B38"), "cr5": val(w2, "B39"), "hhi": val(w2, "B40"), "dur": val(w2, "B41"),
            "shares": [[val(w2, f"{c}{4 + i}") for c in "AEF"] for i in range(n)],
            "channels": [[val(w3, f"{c}{4 + i}") for c in "ACDE"] for i in range(7)],
            "reviews": [[val(w5, f"{c}{4 + i}") for c in "ACEGLM"] for i in range(n)],
            "rev_idx": val(w5, "B37"), "rev_avg": val(w5, "B38"),
            "tech": [[val(w6, f"{c}{4 + i}") for c in "AKL"] for i in range(n)], "tech_idx": val(w6, "B37"), "maturity": val(w6, "B38"),
            "barrier": [[val(w7, f"{c}{4 + i}") for c in "ACDEFG"] for i in range(9)],
            "index": val(w7, "F13"), "level": val(w7, "G13"), "conclusion": val(w8, "A15"),
            "errors": errors(wb),
        }
        wb.BuiltinDocumentProperties("Author").Value = "Р. В. Земляник"
        wb.BuiltinDocumentProperties("Last Author").Value = "Р. В. Земляник"
        wb.Save()
        (HERE / "results").mkdir(exist_ok=True)
        (HERE / "results" / "excel_results.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
        print("index", res["index"], res["level"], "errors", res["errors"])
    finally:
        wb.Close(False)
        app.Quit()


if __name__ == "__main__":
    main()
