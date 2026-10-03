"""Книга 7: оценка рынка через поисковый спрос (контрольный русскоязычный сегмент: Вордстат Россия + Беларусь, Google Trends мир)."""
import sys
from datetime import date

from inputs import *
from xlh import *

sys.path.insert(0, str(ROOT.parent / "ЛР1_Спрос_и_границы_рынка" / "Python"))
import data_prep as dp  # noqa: E402

TPL_NAME = "shablon_excel_ocenka_rynka_cherez_poiskovyy_spros.xlsx"
OUT_NAME = "ЛР2-3_7_поисковый_спрос_TableMind.xlsx"
COMM = 0.109   # доля запросов с коммерческим намерением (проверить + услуга) по структуре намерений ЛР1
CTR = dict(cons=0.04, base=0.08, opt=0.12)

MONTHS12 = dp.MONTHS[-12:]
WS_GROUPS = {
    "B": ["аудит excel", "проверка финансовой модели", "аудит финансовой модели"],
    "C": ["ошибки в формулах excel", "ошибки в excel"],
    "D": ["excel copilot", "нейросеть для excel", "chatgpt excel", "excel искусственный интеллект"],
    "E": ["проверка формул excel", "проверить excel на ошибки"],
}
GT_COLS = {"B": "S_excel_audit_5y", "C": "S_spreadsheet_audit_5y", "D": "S_excel_formula_checker_5y", "E": "S_financial_model_audit_5y"}


def serial(m):
    y, mo = map(int, m.split("-"))
    return (date(y, mo, 1) - date(1899, 12, 30)).days


def ws_series(queries):
    tot = []
    for m in dp.MONTHS:
        tot.append(sum(dp.yw(q, 225)[m] + dp.yw(q, 149)[m] for q in queries))
    ds = dp.despike(tot)
    if isinstance(ds, tuple):
        ds = ds[0]
    series = dict(zip(dp.MONTHS, ds))
    return [series[m] for m in MONTHS12]


def avg12(q):
    return sum(dp.yw(q, 225)[m] + dp.yw(q, 149)[m] for m in MONTHS12) / 12


CORE = [  # запрос, группа, намерение, вес релевантности, коммерческая близость, ключ GT (ближайший англоязычный аналог)
    ("проверка формул excel", "Инструментальный спрос", "инструментальное", 1.0, 0.70, "D"),
    ("проверить excel на ошибки", "Инструментальный спрос", "инструментальное", 1.0, 0.70, "D"),
    ("аудит excel", "Услуговый спрос", "услуговое", 1.0, 0.90, "B"),
    ("проверка финансовой модели", "Услуговый спрос", "услуговое", 0.9, 0.90, "E"),
    ("аудит финансовой модели", "Услуговый спрос", "услуговое", 0.9, 0.90, "E"),
    ("ошибки в формулах excel", "Проблемный спрос", "проблемное", 0.8, 0.35, "C"),
    ("ошибки в excel", "Проблемный спрос", "проблемное", 0.7, 0.30, "C"),
    ("excel copilot", "AI-инструменты", "AI", 0.5, 0.25, "D"),
    ("нейросеть для excel", "AI-инструменты", "AI", 0.5, 0.25, "D"),
    ("chatgpt excel", "AI-инструменты", "AI", 0.5, 0.25, "D"),
]


def fill(x: XL, sw, results):
    wb = x.open_copy(TPL_NAME, OUT_NAME)
    chk = arppu("base")[0]
    ws = x.ws("01_Параметры")
    put(ws, "B4", "Россия и Беларусь (Вордстат; контрольный русскоязычный сегмент), Google Trends - весь мир")
    put(ws, "B5", "последние 12 месяцев (10.2025-09.2026)")
    put(ws, "B6", round(chk, 2)); put(ws, "C6", "EUR"); put(ws, "D6", "Выручка первого года на платящего (Д-09)")
    put(ws, "B7", CTR["base"]); put(ws, "B8", FUNNEL["v2l"]["base"]); put(ws, "B9", FUNNEL["l2p"]["base"]); put(ws, "B10", 1.0)
    put(ws, "D10", "1: годовой чек уже учитывает платежи в течение года")
    put(ws, "B11", 0.5); put(ws, "B12", 0.5)
    put(ws, "D11", "Веса источников 0,5/0,5 (допущение Д-06 ЛР1)")

    gt = {k: dp.gt(v) for k, v in GT_COLS.items()}
    ws = x.ws("03_Google_Trends")
    put(ws, "B3", "excel audit"); put(ws, "C3", "spreadsheet audit"); put(ws, "D3", "excel formula checker"); put(ws, "E3", "financial model audit")
    for i, m in enumerate(MONTHS12):
        r = 4 + i
        put(ws, f"A{r}", serial(m))
        for c in "BCDE":
            put(ws, f"{c}{r}", round(gt[c][m], 2))
    put(ws, "G4", "Google Trends, весь мир, недельные данные приведены к месяцам (ЛР1); индекс каждого запроса нормирован отдельно")

    wsr = {c: ws_series(q) for c, q in WS_GROUPS.items()}
    ws = x.ws("04_Wordstat")
    put(ws, "B3", "Услуговые и аудиторские запросы"); put(ws, "C3", "Проблемные запросы (ошибки)"); put(ws, "D3", "AI-инструменты"); put(ws, "E3", "Проверка формул")
    for i, m in enumerate(MONTHS12):
        r = 4 + i
        put(ws, f"A{r}", serial(m))
        for c in "BCDE":
            put(ws, f"{c}{r}", wsr[c][i])
        put(ws, f"G{r}", COMM)
    put(ws, "G3", "Коммерческий коэффициент (инструментальные и услуговые запросы, ЛР1)")

    ws = x.ws("02_Семантика")
    for i, (q, grp, intent, rel, comm, gk) in enumerate(CORE):
        r = 4 + i
        gt_idx = sum(gt[gk][m] for m in MONTHS12) / 12
        put_row(ws, r, "A", [i + 1, q, grp, intent, "РФ+РБ", round(avg12(q)), round(gt_idx, 1), rel, comm])
    put(ws, "G3", "Индекс Google Trends (ближайший англоязычный аналог, среднее 12 мес.)")

    ws = x.ws("06_Воронка")
    for sc, c in zip(SC, "BCD"):
        put(ws, f"{c}5", CTR[sc]); put(ws, f"{c}7", FUNNEL["v2l"][sc]); put(ws, f"{c}9", FUNNEL["l2p"][sc])
        put(ws, f"{c}11", round(arppu(sc)[0], 2)); put(ws, f"{c}12", 1.0)

    ws = x.ws("09_Источники")
    for r in range(4, 10):
        put(ws, f"D{r}", SNAP_SERIAL)
    put(ws, "C4", "весь мир, 5 лет (данные ЛР1)"); put(ws, "C6", "Россия 225, Беларусь 149, 2018-2026 (данные ЛР1)")
    put(ws, "E8", "Рекламных кабинетов нет: ДОСНЯТЬ"); put(ws, "E9", "Сайта и CRM пока нет: ДОСНЯТЬ после MVP")

    x.app.CalculateFull()
    w5 = x.ws("05_Нормализация"); w6 = x.ws("06_Воронка"); w8 = x.ws("08_Дашборд"); w2 = x.ws("02_Семантика")
    res = {
        "months": MONTHS12,
        "idx": [[val(w5, f"{c}{4 + i}") for c in "BCDEFG"] for i in range(12)],
        "weighted_avg": val(x.ws("04_Wordstat"), "H4") and sum(val(x.ws("04_Wordstat"), f"H{4 + i}") for i in range(12)) / 12,
        "core": [[val(w2, f"{c}{4 + i}") for c in "BFGJ"] for i in range(len(CORE))],
        "funnel": {sc: [val(w6, f"{c}{r}") for r in range(4, 15)] for sc, c in zip(SC, "BCD")},
        "dash": [val(w8, f"B{r}") for r in range(4, 11)],
        "errors": errors(wb),
    }
    results["se"] = res
    x.save_close()
