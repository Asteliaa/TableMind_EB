"""Книга 3: проверка рынка через платежеспособность."""
from inputs import *
from xlh import *

TPL_NAME = "shablon_excel_metodika_proverki_rynka_cherez_platezhesposobnost.xlsx"
OUT_NAME = "ЛР2-3_3_платежеспособность_TableMind.xlsx"

USD = lambda v: round(v / EUR_USD, 2)

# ценовой коридор (EUR/мес.): мин, медиана, макс (по открытым ценам 03.10.2026), себестоимость (допущение Д-11), цена TableMind
PRICE_ROWS = [
    ("Разовая проверка файла (аналог: месяц дешевого AI-помощника)", 7.10, 8.88, 19.17, 1.5, PRICE["audit"],
     "Прямых разовых аналогов с публичной ценой нет; ориентир - Rows Plus 8 USD, Numerous 10 USD, медиана Pro-класса", "rows.com/pricing, numerous.ai/pricing"),
    ("Персональная подписка уровня Pro (в месяц)", 7.10, 19.17, 88.76, 4.0, PRICE["pro"],
     "Rows Plus 8 USD; Numerous 10 USD; Copilot Business 18-25,2 USD; PerfectXL от 69 EUR; Shortcut Pro 100 USD", "страницы цен, 03.10.2026"),
    ("Командное место Team (в месяц за место)", 7.10, 22.37, 88.76, 4.5, PRICE["team"],
     "Rows Plus на пользователя 8 USD; Copilot Business 25,2 USD; Shortcut Teams 100 USD за место", "страницы цен, 03.10.2026"),
    ("Специализированный аудит таблиц (класс PerfectXL)", 69.0, 69.0, 69.0, 4.0, PRICE["pro"],
     "PerfectXL отдельный инструмент от 69 EUR в месяц; Arixcel и Operis цены не публикуют (ДОСНЯТЬ)", "perfectxl.com/pricing"),
]
SEG_KEYS = ["fin_analysts", "accountants", "consultants", "sme"]
HOURS = {"fin_analysts": 60, "accountants": 50, "consultants": 50, "sme": 40}  # часов в месяц на работу с таблицами
CHECK_SHARE = 0.20   # доля времени на проверку и поиск ошибок
SAVE_SHARE = 0.45    # доля экономии за счет сервиса


def fill(x: XL, sw, results):
    wb = x.open_copy(TPL_NAME, OUT_NAME)
    ws = x.ws("01_Параметры")
    put(ws, "B5", "TableMind: онлайн-проверка Excel-моделей")
    put(ws, "B6", "Англоязычные страны и ЕС-27")
    put(ws, "B7", "EUR")
    put(ws, "B9", FUNNEL["l2p"]["base"])
    put(ws, "B11", 0.80); put(ws, "C11", "Целевая валовая маржа SaaS (допущение Д-11): цена минус вычисления и LLM")
    put(ws, "B12", SNAP_SERIAL)

    ws = x.ws("02_Сегменты")
    for i, s in enumerate(SEGMENTS):
        r = 5 + i
        put_row(ws, r, "A", [s["name"], round(s["total"] * s["border"]), s["need"], s["digital"], s["payable"], s["freq"], s["check"], NOSET, NOSET, NOSET,
                             s["note"], s["src"]])
    put(ws, "A9", None); put(ws, "B9", None)
    for c in "CDEFG":
        put(ws, f"{c}9", None)
    put(ws, "K9", None); put(ws, "L9", None)

    ws = x.ws("03_Ценовой_коридор")
    for i, p in enumerate(PRICE_ROWS):
        r = 5 + i
        put_row(ws, r, "A", [p[0], p[1], p[2], p[3], p[4], NOSET, p[5], NOSET, p[6], p[7]])
    for c in "ABCDEGIJ":
        put(ws, f"{c}9", None)
    put(ws, "H3", "Индекс цены TableMind к медиане")

    ws = x.ws("04_Экономика_покупки")
    put(ws, "B3", "Стоимость месячной работы с таблицами, EUR")
    put(ws, "C3", "Доля времени на проверку и поиск ошибок")
    put(ws, "D3", "Доля экономии времени за счет сервиса")
    put(ws, "E3", "Экономия, EUR/мес.")
    for i, k in enumerate(SEG_KEYS):
        r = 5 + i
        put(ws, f"A{r}", SEGMENTS[i]["name"])
        put(ws, f"B{r}", HOURS[k] * HOURLY[k if k != "sme" else "sme"] if k in HOURLY else HOURS[k] * HOURLY["consultants"])
        put(ws, f"C{r}", CHECK_SHARE); put(ws, f"D{r}", SAVE_SHARE)
        put(ws, f"F{r}", PAY[i]["price"])
        put(ws, f"J{r}", "Эффект - экономия времени проверки; не учитывает стоимость пропущенной ошибки (она выше)")
    put(ws, "A9", None)
    for c in "BCDF":
        put(ws, f"{c}9", None)
    put(ws, "J9", None)

    ws = x.ws("05_Платежеспособность")
    for i, p in enumerate(PAY):
        r = 5 + i
        put_row(ws, r, "A", [p["seg"], p["budget"], p["price"], NOSET, NOSET, p["share_budget"], NOSET, p["annual"], NOSET, p["note"]])
    for c in "ABCFH":
        put(ws, f"{c}9", None)
    put(ws, "J9", None)

    ws = x.ws("06_Сценарии")
    # исправление шаблона: пустая строка сегмента давала ошибку #ЗНАЧ! (ссылка на пустую ячейку возвращает 0, а не пустую строку)
    for r in range(13, 21):
        put(ws, f"A{r}", f"=IF('02_Сегменты'!A{r - 8}=\"\",\"\",'02_Сегменты'!A{r - 8})")
        put(ws, f"B{r}", f"=IF(A{r}=\"\",\"\",'02_Сегменты'!J{r - 8})")
    sc = PC_SCEN
    for c, k in zip("BCD", SC):
        put(ws, f"{c}5", sc[k]["need"]); put(ws, f"{c}6", sc[k]["pay"]); put(ws, f"{c}7", sc[k]["freq"]); put(ws, f"{c}8", sc[k]["check"]); put(ws, f"{c}9", sc[k]["share"])

    ws = x.ws("07_Similarweb")
    n = len(sw)
    insert_rows_copy(ws, 8, n - 8, 5)
    for i, r in enumerate(sw):
        row = 5 + i
        eff = r["geo_share"] * r["rel"] * r["qual"]
        put_row(ws, row, "A", [r["domain"], r["visits_month"], round(eff, 4), NOSET, r["ch"]["org_search"], r["ch"]["direct"],
                               r["ch"]["org_social"] + r["ch"]["paid_social"], FUNNEL["v2l"]["base"] * FUNNEL["l2p"]["base"], NOSET,
                               round(arppu("base")[0], 2), NOSET, "Similarweb Pro, 03.10.2026; C = география x релевантность x качество; H = визит - платящий"])
    put(ws, "C4", "Эффективная доля (география x релевантность x качество)"); put(ws, "H4", "Конверсия визит - платящий"); put(ws, "I4", "Оценка новых платящих/мес.")
    put(ws, "J4", "Выручка первого года на платящего, EUR"); put(ws, "K4", "Оценка выручки/мес. (первый год), EUR")

    ws = x.ws("09_Источники")
    for r in range(5, 11):
        put(ws, f"E{r}", "2026-10-03")
    put(ws, "A5", "BLS, SBA, ЕК, UK BPE (вместо Белстата)"); put(ws, "D5", "bls.gov/ooh; advocacy.sba.gov; single-market-economy.ec.europa.eu; gov.uk")
    put(ws, "A8", "Яндекс Вордстат и Google Trends (ЛР1)"); put(ws, "D8", "ЛР1, Материалы_собранные")
    put(ws, "A10", "Опросы / интервью (расчетная модель)"); put(ws, "F10", "Расчетные данные по решению автора (Р-8); реального опроса не проводилось")

    x.app.CalculateFull()
    w5 = x.ws("05_Платежеспособность"); w8 = x.ws("08_Итоговая_панель")
    res = {
        "volume": val(x.ws("02_Сегменты"), "J16"), "payers": val(x.ws("02_Сегменты"), "H16"),
        "seg": [[val(x.ws("02_Сегменты"), f"{c}{5 + i}") for c in "HIJ"] for i in range(len(SEGMENTS))],
        "price": [[val(x.ws("03_Ценовой_коридор"), f"{c}{5 + i}") for c in "BCDFGH"] for i in range(len(PRICE_ROWS))],
        "econ": [[val(x.ws("04_Экономика_покупки"), f"{c}{5 + i}") for c in "BEFGHI"] for i in range(len(SEGMENTS))],
        "idx": [[val(w5, f"{c}{5 + i}") for c in "DEGI"] for i in range(len(SEGMENTS))], "idx_avg": val(w5, "D14"),
        "panel": {k: val(w8, a) for k, a in (("avg_idx", "B7"), ("roi", "B8"), ("som_c", "B9"), ("som_b", "B10"), ("som_o", "B11"), ("sw_rev_m", "B12"))},
        "scen": {k: [val(x.ws("06_Сценарии"), f"{c}22") for c in "CDEFGH"] for k in ["all"]},
        "errors": errors(wb),
    }
    results["pay"] = res
    x.save_close()
