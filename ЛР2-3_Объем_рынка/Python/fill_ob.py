"""Книга 6: оценка объема рынка по данным Similarweb."""
from inputs import *
from xlh import *

TPL_NAME = "metodika_ocenki_obema_rynka_po_similarweb.xlsx"
OUT_NAME = "ЛР2-3_6_объем_рынка_Similarweb_TableMind.xlsx"

COVER = dict(cons=0.70, base=0.55, opt=0.40)
SOM = dict(cons=0.005, base=0.015, opt=0.04)
INTENT = [  # доля, группа, тип намерения, вес, включать, комментарий (ЛР1, выборка топов Вордстата, 4 578 запросов)
    ("запросы о проверке и аудите (инструментальные и услуговые)", 0.109, "инструментальный и услуговый", 1.0, "включить", "Прямой спрос: проверить файл, заказать аудит"),
    ("проблемные запросы (ошибки в формулах)", 0.366, "проблемный", 0.6, "включить", "Потребность есть, нужна коммерческая проверка"),
    ("AI-запросы (Copilot, ChatGPT, нейросети для Excel)", 0.453, "AI-инструменты", 0.25, "включить частично", "Конкурентное поле заменителей"),
    ("прочие запросы выборки", 0.072, "прочее", 0.1, "исключить", "Не подтверждают платежеспособный спрос"),
]


def tgt_geo_rows(sw):
    tot = sum(r["visits_month"] for r in sw)
    acc = {}
    for r in sw:
        for c, v in r["top"]:
            acc[c] = acc.get(c, 0) + r["visits_month"] * v
    us = acc.get("US", 0) / tot
    gb = acc.get("GB", 0) / tot
    eu = sum(v for c, v in acc.items() if c in EU27) / tot
    other_c = sum(v for c, v in acc.items() if c not in TARGET) / tot
    unknown = 1 - (us + gb + eu + sum(v for c, v in acc.items() if c in {"CA", "AU"}) / tot + other_c)
    ca_au = sum(v for c, v in acc.items() if c in {"CA", "AU"}) / tot
    return [("США", us), ("Великобритания", gb), ("ЕС-27 и Канада, Австралия (из топ-5 стран сайтов)", eu + ca_au), ("Вне границ проекта или не раскрыто (Индия, Бразилия, остальные страны)", other_c + unknown)]


def fill(x: XL, sw, results):
    wb = x.open_copy(TPL_NAME, OUT_NAME)
    n = len(sw)
    chk = {sc: arppu(sc) for sc in SC}
    ws = x.ws("01_Параметры")
    put(ws, "B4", "Онлайн-проверка Excel-моделей и финансовых таблиц (TableMind)")
    put(ws, "B5", "Англоязычные страны и ЕС-27")
    put(ws, "B6", "последние 3 месяца (06-08.2026)")
    put(ws, "B7", SNAP_SERIAL)
    # исправление шаблона: суммирование начиналось со строки 5 и пропускало первую строку данных
    put(ws, "B9", "=SUM('02_Конкуренты_SW'!G4:G23)")
    put(ws, "B10", "=SUM('02_Конкуренты_SW'!H4:H23)")
    put(ws, "B11", COVER["base"])

    ws = x.ws("02_Конкуренты_SW")
    put(ws, "F3", "Доля целевой географии x релевантность x качество")
    for i in range(20):
        row = 4 + i
        if i < n:
            r = sw[i]
            ch = r["ch"]
            put_row(ws, row, "A", [r["domain"], r["name"], r["segment"], r["type"], r["visits_month"], round(r["geo_share"] * r["rel"] * r["qual"], 4), NOSET, NOSET, NOSET,
                                   (float(r["bounce_pct"]) / 100) if r["bounce_pct"] else None, float(r["pages_visit"]), r["duration"],
                                   ch["direct"], ch["org_search"], ch["referrals"], ch["org_social"] + ch["paid_social"], ch["paid_search"], ch["display"],
                                   "Similarweb Pro, 06-08.2026, 03.10.2026"])
        else:
            put_row(ws, row, "A", [None, None, None, None, None, None, NOSET, NOSET, NOSET, None, None, None, None, None, None, None, None, None, None])

    ws = x.ws("04_География")
    for i, (nm, sh) in enumerate(tgt_geo_rows(sw)):
        put(ws, f"A{4 + i}", nm); put(ws, f"B{4 + i}", round(sh, 4))
    put(ws, "E4", "высокая"); put(ws, "E5", "высокая"); put(ws, "E6", "высокая"); put(ws, "E7", "н/д")
    put(ws, "F4", "высокая"); put(ws, "F5", "средняя"); put(ws, "F6", "средняя"); put(ws, "F7", "низкая")
    put(ws, "H4", "Основной рынок: англоязычные страны"); put(ws, "H5", "Англоязычный рынок вне США"); put(ws, "H6", "Расширение на ЕС после MVP"); put(ws, "H7", "Индия и другие страны вне границ проекта")
    put(ws, "B3", "Доля визитов 12 сайтов (по топ-5 стран каждого)")

    ws = x.ws("05_Запросы")
    for i, t in enumerate(INTENT):
        r = 4 + i
        put_row(ws, r, "A", [t[0], "ЛР1: структура намерений выборки топов Вордстата (4 578 запросов)", t[1], NOSET, t[2], t[3], NOSET, t[2], t[4], t[5]])
    for r in range(4 + len(INTENT), 11):
        put_row(ws, r, "A", [None, None, 0, NOSET, None, 0, NOSET, None, None, None])

    ws = x.ws("06_Воронка_сценарии")
    for sc, r in zip(SC, (5, 6, 7)):
        put(ws, f"B{r}", COVER[sc]); put(ws, f"D{r}", FUNNEL["v2l"][sc]); put(ws, f"E{r}", FUNNEL["l2p"][sc])
        put(ws, f"F{r}", round(chk[sc][2], 2)); put(ws, f"G{r}", round(chk[sc][1], 3))
    for sc, r in zip(SC, (12, 13, 14)):
        put(ws, f"C{r}", SOM[sc])
    put(ws, "H19", "EUR/год")

    ws = x.ws("08_Источники")
    for r in range(4, 19):
        v = ws.Range(f"E{r}").Value2
    x.app.CalculateFull()
    w1 = x.ws("01_Параметры"); w3 = x.ws("03_Каналы"); w6 = x.ws("06_Воронка_сценарии"); w2 = x.ws("02_Конкуренты_SW"); w5 = x.ws("05_Запросы"); w7 = x.ws("07_Дашборд")
    res = {
        "params": {k: val(w1, a) for k, a in (("n", "B8"), ("obs_m", "B9"), ("obs_y", "B10"), ("full_m", "B12"), ("full_y", "B13"))},
        "shares": [[val(w2, f"{c}{4 + i}") for c in "AEGI"] for i in range(n)],
        "w_channels": {k: val(w2, f"{c}25") for k, c in zip(("direct", "org", "ref", "soc", "paid", "disp"), "MNOPQR")},
        "channels": [[val(w3, f"{c}{4 + i}") for c in "ABCDH"] for i in range(6)],
        "geo": [[val(x.ws("04_География"), f"{c}{4 + i}") for c in "ABCG"] for i in range(4)],
        "queries": [[val(w5, f"{c}{4 + i}") for c in "ACDFG"] for i in range(len(INTENT))],
        "scen": {sc: [val(w6, f"{c}{5 + i}") for c in "BCDEFGHIJ"] for i, sc in enumerate(SC)},
        "som": {sc: [val(w6, f"{c}{12 + i}") for c in "BCD"] for i, sc in enumerate(SC)},
        "conclusion": val(w6, "B22"),
        "dash": [val(w7, a) for a in ("J5", "B6", "D6", "F6")],
        "errors": errors(wb),
    }
    results["ob"] = res
    x.save_close()
