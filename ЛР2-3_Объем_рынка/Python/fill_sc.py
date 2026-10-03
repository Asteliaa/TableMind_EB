"""Книга 5: сценарная оценка рынка с данными Similarweb."""
from inputs import *
from xlh import *

TPL_NAME = "shablon_excel_scenarnaya_ocenka_rynka_s_similarweb.xlsx"
OUT_NAME = "ЛР2-3_5_сценарная_оценка_Similarweb_TableMind.xlsx"

UNCOUNTED = dict(cons=1.3, base=round(1 / FUNNEL["cover"], 3), opt=2.5)        # во сколько раз рынок шире найденных сайтов
RELIAB = dict(cons=0.75, base=0.90, opt=1.0)
UNAVAIL = dict(cons=0.8, base=1.0, opt=1.15)
SOM_SHARE = dict(cons=0.005, base=0.015, opt=0.04)
APP = dict(reg=dict(cons=0.01, base=0.03, opt=0.06), deals=dict(cons=1, base=2, opt=4), fee=0.03)


def fill(x: XL, sw, results, ext=None):
    wb = x.open_copy(TPL_NAME, OUT_NAME)
    n = len(sw)
    chk = {sc: arppu(sc) for sc in SC}
    base_chk = chk["base"][2]
    ws = x.ws("01_Параметры")
    put(ws, "B4", "TableMind: онлайн-проверка Excel-моделей")
    put(ws, "B5", "Англоязычные страны и ЕС-27")
    put(ws, "B6", "EUR")
    put(ws, "B8", round(base_chk, 2)); put(ws, "C8", "EUR/платеж")
    put(ws, "E8", "Структура платящих (Д-09) и тарифы ПЗ1")
    put(ws, "B9", round(chk["base"][1], 3)); put(ws, "C9", "платежей/год")
    tot_v = sum(r["visits_month"] for r in sw)
    put(ws, "B10", round(sum(r["visits_month"] * r["geo_share"] for r in sw) / tot_v, 4))
    put(ws, "B11", round(sum(r["visits_month"] * r["rel"] for r in sw) / tot_v, 4))
    put(ws, "B12", round(sum(r["visits_month"] * r["qual"] for r in sw) / tot_v, 4))
    put(ws, "B13", SNAP_SERIAL)
    for r, d in ((16, UNCOUNTED), (17, RELIAB), (18, UNAVAIL)):
        put_row(ws, r, "B", [d[sc] for sc in SC])
    put_row(ws, 19, "B", [FUNNEL["v2l"][sc] for sc in SC])
    put_row(ws, 20, "B", [FUNNEL["l2p"][sc] for sc in SC])
    put_row(ws, 21, "B", [round(chk[sc][2] / base_chk, 4) for sc in SC])
    put_row(ws, 22, "B", [round(chk[sc][1], 3) for sc in SC])
    put_row(ws, 23, "B", [SOM_SHARE[sc] for sc in SC])
    put(ws, "E16", "Во сколько раз рынок шире 12 найденных сайтов: 1/покрытие 0,55 в базовом сценарии")

    ws = x.ws("02_Similarweb")
    for i in range(15):
        row = 5 + i
        if i < n:
            r = sw[i]
            put_row(ws, row, "A", [i + 1, r["name"], r["domain"], r["type"] + ": " + r["segment"], r["visits_month"], r["geo_share"], r["rel"], r["qual"], NOSET, NOSET, NOSET,
                                   "https://www.similarweb.com/website/" + r["domain"] + "/", "Similarweb Pro, 06-08.2026, мир, снято 03.10.2026; география по топ-5 стран"])
        else:
            put_row(ws, row, "A", [i + 1, None, None, None, 0, 0, 0, 0])

    ws = x.ws("05_Платформа")
    for sc, row in zip(SC, (5, 6, 7)):
        put(ws, f"C{row}", APP["reg"][sc]); put(ws, f"E{row}", APP["deals"][sc]); put(ws, f"F{row}", round(chk[sc][2], 2))
        put(ws, f"H{row}", APP["fee"]); put(ws, f"J{row}", SOM_SHARE[sc])
    put(ws, "A1", "Канал AppSource: оценка оборота магазина надстроек Excel и комиссии (интерпретация платформенной модели)")

    ws = x.ws("06_Сверка")
    if ext:
        for sc, c in zip(SC, "BCD"):
            put(ws, f"{c}7", round(ext["search"][sc]))
            put(ws, f"{c}8", round(ext["pc"][sc]))
            put(ws, f"{c}9", round(ext["tb"][sc]))
    # исправление шаблона: платформенная строка (оборот и комиссия магазина надстроек) не является оценкой выручки нового бизнеса и исключена из итогов
    for c in "BCD":
        put(ws, f"{c}12", f"=MIN({c}5,{c}7:{c}9)"); put(ws, f"{c}13", f"=AVERAGE({c}5,{c}7:{c}9)"); put(ws, f"{c}14", f"=MAX({c}5,{c}7:{c}9)")
    put(ws, "A6", "Платформенная модель (комиссия AppSource, не входит в итог)")
    put(ws, "A7", "Поисковый спрос (контрольный русскоязычный сегмент, ЛР2-3 книга 7), EUR/год")
    put(ws, "A8", "Количество потенциальных клиентов (книга 1): SOM, EUR/год")
    put(ws, "A9", "Снизу вверх (книга 2): SAM x доля SOM, EUR/год")

    x.app.CalculateFull()
    w3 = x.ws("03_Сценарии"); w4 = x.ws("04_Воронка"); w5 = x.ws("05_Платформа"); w6 = x.ws("06_Сверка")
    res = {
        "sw_month": val(x.ws("02_Similarweb"), "I21"), "sw_year": val(x.ws("02_Similarweb"), "J21"),
        "scen": {sc: [val(w3, f"{c}{5 + i}") for c in "BCDEFGHI"] for i, sc in enumerate(SC)},
        "funnel": {sc: [val(w4, f"{c}{5 + i}") for c in "BCDEFGHIJ"] for i, sc in enumerate(SC)},
        "platform": {sc: [val(w5, f"{c}{5 + i}") for c in "BCDEFGHIJK"] for i, sc in enumerate(SC)},
        "sverka": {k: [val(w6, f"{c}{r}") for c in "BCD"] for k, r in (("sw", 5), ("plat", 6), ("search", 7), ("pc", 8), ("tb", 9), ("min", 12), ("avg", 13), ("max", 14))},
        "errors": errors(wb),
    }
    results["sc"] = res
    x.save_close()
