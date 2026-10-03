"""Книга 4: оценка рынка через цифровую воронку."""
from inputs import *
from xlh import *

TPL_NAME = "shablon_excel_metodika_ocenki_rynka_cherez_cifrovuyu_voronku.xlsx"
OUT_NAME = "ЛР2-3_4_цифровая_воронка_TableMind.xlsx"

CHANNELS = [
    ("Прямой трафик", "Similarweb", ["direct"]),
    ("Органический поиск", "Similarweb; Google Trends и Вордстат (ЛР1)", ["org_search"]),
    ("Платный поиск", "Similarweb", ["paid_search"]),
    ("Реферальные площадки", "Similarweb: каталоги, партнеры, AppSource", ["referrals"]),
    ("Социальные сети", "Similarweb: органические и платные", ["org_social", "paid_social"]),
    ("Медийная реклама и email", "Similarweb", ["display", "email"]),
    ("AI-поиск и партнерские сети", "Similarweb: Gen AI и affiliates", ["gen_ai", "affiliates"]),
]
GMV = {  # платежи через магазин надстроек AppSource: сделок в месяц, комиссия площадки (допущение Д-14: ДОСНЯТЬ условия Microsoft)
    "cons": 2, "base": 8, "opt": 20,
}
FEE = 0.03


def fill(x: XL, sw, results):
    wb = x.open_copy(TPL_NAME, OUT_NAME)
    n = len(sw)
    ws = x.ws("01_Параметры")
    put(ws, "B4", "TableMind: онлайн-проверка Excel-моделей")
    put(ws, "B5", "Англоязычные страны и ЕС-27")
    put(ws, "B6", "месяц / год"); put(ws, "B7", "EUR")
    put(ws, "B8", "финансовые аналитики, бухгалтеры, консультанты, МСП")
    check = {sc: arppu(sc)[2] for sc in SC}
    put(ws, "B9", round(check["base"], 2)); put(ws, "B10", round(check["cons"], 2)); put(ws, "B11", round(check["opt"], 2))
    put(ws, "C9", "EUR/платеж"); put(ws, "C10", "EUR/платеж"); put(ws, "C11", "EUR/платеж")
    put(ws, "D9", "Структура платящих (Д-09) и тарифы ПЗ1")
    # доля релевантного трафика и покрытие
    tot_v = sum(r["visits_month"] for r in sw)
    rel_avg = sum(r["visits_month"] * r["rel"] for r in sw) / tot_v
    put(ws, "B12", round(rel_avg, 4)); put(ws, "D12", "Средневзвешенная по визитам доля релевантного трафика 12 сайтов (inputs.SITES)")
    put(ws, "B13", FUNNEL["cover"]); put(ws, "D13", "Допущение Д-13: 12 сайтов покрывают около 55 % трафика сегмента")
    put(ws, "B14", SNAP_SERIAL)
    for r, sc in zip((17, 18, 19), SC):
        put_row(ws, r, "B", [FUNNEL["reach"][sc], FUNNEL["v2l"][sc], FUNNEL["l2p"][sc], NOSET, arppu(sc)[1]])
    put(ws, "B16", "Достижимая доля трафика сегмента")
    put(ws, "F16", "Платежей в год на платящего")

    ws = x.ws("02_Трафик_охват")
    for i, (nm, src, keys) in enumerate(CHANNELS):
        r = 4 + i
        c_visits = sum(s["visits_month"] * sum(s["ch"][k] for k in keys) for s in sw)
        c_geo = sum(s["visits_month"] * sum(s["ch"][k] for k in keys) * s["geo_share"] for s in sw)
        c_rel = sum(s["visits_month"] * sum(s["ch"][k] for k in keys) * s["geo_share"] * s["rel"] * s["qual"] for s in sw)
        # C = визиты в канале; D = доля целевой географии; F = релевантность и качество в целевой географии
        put_row(ws, r, "A", [nm, src, round(c_visits), round(c_geo / c_visits, 4) if c_visits else 0, NOSET, round(c_rel / c_geo, 4) if c_geo else 0, NOSET,
                             "среднее", "Сумма по 12 сайтам; D - доля целевой географии; F - релевантность x качество"])
    put(ws, "C3", "Визиты по каналу в месяц (12 сайтов)"); put(ws, "D3", "Доля целевой географии"); put(ws, "F3", "Доля релевантной аудитории (релевантность x качество)")

    ws = x.ws("03_Конкуренты_SW")
    insert_rows_copy(ws, 6, n - 5, 4)
    for i, r in enumerate(sw):
        row = 4 + i
        put_row(ws, row, "A", [r["domain"], r["type"], r["visits_month"], r["geo_share"], round(r["rel"] * r["qual"], 4), NOSET,
                               FUNNEL["v2l"]["base"] * FUNNEL["l2p"]["base"], round(arppu("base")[0], 2), NOSET, "Similarweb Pro, 06-08.2026; E = релевантность x качество"])

    # исправление шаблона: формулы C12 и C14 делили на B12 (доля релевантного трафика) вместо B13 (коэффициент покрытия)
    put(ws, f"C{n + 7}", f"=C{n + 6}/'01_Параметры'!B13")
    put(ws, f"C{n + 9}", f"=C{n + 8}/'01_Параметры'!B13*12")
    ws = x.ws("04_Воронка")
    put(ws, "A4", "Релевантные посещения сегмента в месяц (оценка полного трафика)")
    for c in "BCD":
        put(ws, f"{c}4", f"='03_Конкуренты_SW'!C{n + 7}")   # оценка полного трафика сегмента (с учетом покрытия)
    put(ws, "E4", "Берется с листа 03_Конкуренты_SW: релевантные посещения 12 сайтов / покрытие.")

    ws = x.ws("07_Платформа_GMV")
    for c, sc in zip("BCD", SC):
        put(ws, f"{c}4", GMV[sc]); put(ws, f"{c}5", round(arppu(sc)[2], 2)); put(ws, f"{c}7", FEE)
    put(ws, "A7", "Комиссия магазина надстроек (AppSource)"); put(ws, "A8", "Удержание магазином надстроек в месяц")
    put(ws, "A9", "Удержание магазином надстроек в год")
    put(ws, "F4", "Допущение Д-14: оплаты через AppSource; ДОСНЯТЬ условия Microsoft")

    ws = x.ws("09_Источники")
    for r in range(4, 11):
        put(ws, f"D{r}", SNAP_SERIAL)
    put(ws, "F4", "Своего сайта пока нет: ДОСНЯТЬ после запуска MVP (31.03.2027)")
    put(ws, "F5", "Рекламных кабинетов нет: ДОСНЯТЬ тестовые кампании")

    x.app.CalculateFull()
    w3 = x.ws("03_Конкуренты_SW"); w4 = x.ws("04_Воронка"); w8 = x.ws("08_Дашборд")
    res = {
        "channels": [[val(x.ws("02_Трафик_охват"), f"{c}{4 + i}") for c in "ACDEFG"] for i in range(len(CHANNELS))],
        "traffic": {k: val(x.ws("02_Трафик_охват"), a) for k, a in (("total_m", "C13"), ("rel_m", "C14"), ("rel_y", "C15"))},
        "sw": {k: val(w3, f"C{n + 6 + i}") for i, k in enumerate(("rel_visits", "full_traffic", "rev_month", "rev_year"))},
        "funnel": {sc: [val(w4, f"{c}{r}") for r in range(4, 14)] for sc, c in zip(SC, "BCD")},
        "ltv": {sc: [val(x.ws("06_Повторы_LTV"), f"{c}{4 + i}") for c in "BCDEF"] for i, sc in enumerate(SC)},
        "gmv": {sc: [val(x.ws("07_Платформа_GMV"), f"{c}{r}") for r in (6, 8, 9)] for sc, c in zip(SC, "BCD")},
        "dash": [val(w8, f"B{r}") for r in range(4, 13)],
        "errors": errors(wb),
    }
    results["fun"] = res
    x.save_close()
