"""Книга 2: методики «сверху вниз» и «снизу вверх» с данными Similarweb."""
from inputs import *
from xlh import *

TPL_NAME = "shablon_excel_metodika_sverhu_vniz_snizu_vverh_similarweb.xlsx"
OUT_NAME = "ЛР2-3_2_сверху_вниз_снизу_вверх_TableMind.xlsx"

# потенциальные пользователи: специалисты (три профессиональных сегмента) и по одному лицу, принимающему решения, в МСП
SPEC = sum(s["total"] for s in SEGMENTS[:3])
SME = SEGMENTS[3]["total"]
TD_USERS = {"cons": 0.85, "base": 1.0, "opt": 1.15}   # коэффициент неопределенности статистической базы
TD_PROD = dict(cons=0.07, base=0.10, opt=0.13)
TD_NEED = dict(cons=0.20, base=0.30, opt=0.40)
TD_DIG = dict(cons=0.55, base=0.65, opt=0.75)
SC_COEF = dict(cons=0.6, base=1.0, opt=1.4)


def fill(x: XL, sw, results):
    wb = x.open_copy(TPL_NAME, OUT_NAME)
    base_check = arppu("base")[0]
    ws = x.ws("01_Параметры")
    put(ws, "B5", "Онлайн-проверка Excel-моделей TableMind (аудит таблиц)")
    put(ws, "B6", "Англоязычные страны и ЕС-27")
    put(ws, "B7", "EUR")
    put(ws, "B8", round(base_check, 2)); put(ws, "C8", "EUR/платящий/год")
    put(ws, "E8", "Тарифы ПЗ1 9/19/29 EUR и структура платящих (Д-09)")
    put(ws, "B9", 1)
    put(ws, "B10", FUNNEL["v2l"]["base"]); put(ws, "E10", "Допущение Д-12: визит - регистрация Free")
    put(ws, "B11", FUNNEL["l2p"]["base"]); put(ws, "E11", "Допущение Д-12: регистрация - платящий")
    put(ws, "B12", 1.0); put(ws, "E12", "Выручка первого года уже включает платежи в течение года")
    put(ws, "B13", round(1 / FUNNEL["cover"], 4)); put(ws, "E13", "Список из 12 сайтов покрывает около 55 % трафика сегмента (допущение Д-13)")
    put(ws, "B14", SC_COEF["cons"]); put(ws, "B15", SC_COEF["base"]); put(ws, "B16", SC_COEF["opt"])

    ws = x.ws("02_Сверху_вниз")
    put(ws, "A5", "Потенциальные пользователи: специалисты и лица, принимающие решения в МСП с наемными работниками")
    for c, sc in zip("BCD", SC):
        put(ws, f"{c}5", round((SPEC + SME) * TD_USERS[sc]))
        put(ws, f"{c}6", TD_PROD[sc]); put(ws, f"{c}7", 1.0)
        put(ws, f"{c}8", TD_DIG[sc]); put(ws, f"{c}9", TD_NEED[sc])
    put(ws, "F5", "Занятые (BLS 2025 и масштаб) + МСП с наемными (SBA, UK, ЕК); коэффициент неопределенности 0,85/1/1,15")
    put(ws, "F6", "Доля работающих с моделями, где ошибка критична (допущение Д-10)")
    put(ws, "F7", "База уже ограничена территорией проекта (Р-1)")

    ws = x.ws("03_Снизу_вверх")
    for i, s in enumerate(SEGMENTS):
        r = 5 + i
        put_row(ws, r, "A", [s["name"], round(s["total"] * s["border"] * s["payable"]), s["need"], s["digital"], 1, s["check"], NOSET, NOSET,
                             "Клиенты в границах рынка x доля платежеспособных; " + s["src"], "Заполнено"])
    for r in range(5 + len(SEGMENTS), 11):
        put_row(ws, r, "A", ["Резерв", 0, 0, 0, 1, 0, NOSET, NOSET, "Свободная строка шаблона", "Не используется"])

    ws = x.ws("04_Similarweb")
    n = len(sw)
    insert_rows_copy(ws, 10, n - 10 if n > 10 else 0, 5) if n > 10 else None
    for i, r in enumerate(sw):
        row = 5 + i
        ch_rel = r["ch"]["direct"] + r["ch"]["org_search"] + r["ch"]["paid_search"] + r["ch"]["referrals"]
        put_row(ws, row, "A", [r["domain"], r["type"], r["visits_month"], r["geo_share"], round(ch_rel, 4), r["rel"], r["qual"], NOSET, NOSET, NOSET,
                               "https://www.similarweb.com/website/" + r["domain"] + "/", "2026-10-03", "Similarweb Pro, 06-08.2026, мир; география по топ-5 стран (нижняя оценка)"])
    for r in range(5 + n, 15 + max(0, n - 10)):
        put_row(ws, r, "A", [None, None, 0, 0, 0, 0, 0])

    ws = x.ws("08_Источники")
    for r in range(5, 11):
        put(ws, f"E{r}", "2026-10-03")
    put(ws, "H5", "Снято 03.10.2026 (Pro, пробный доступ)")
    put(ws, "H8", "BLS 2025, SBA 2025, ЕК 2025/2026, UK BPE 2025; см. внешние_данные_2026-10-03.md")
    put(ws, "H9", "PerfectXL от 69 EUR, Numerous 10 USD, Rows 8 USD, Shortcut 100 USD, Copilot 18 USD (03.10.2026)")
    put(ws, "H10", "Нет доступа к рекламным кабинетам: ДОСНЯТЬ тестовые кампании")

    x.app.CalculateFull()
    w6 = x.ws("06_Сверка"); w7 = x.ws("07_Дашборд")
    res = {
        "td": {sc: [val(x.ws("02_Сверху_вниз"), f"{c}{r}") for r in (11, 12)] for sc, c in zip(SC, "BCD")},
        "bu": {"total": val(x.ws("03_Снизу_вверх"), "G11"), "clients": val(x.ws("03_Снизу_вверх"), "B11"),
               "by_seg": [[val(x.ws("03_Снизу_вверх"), f"{c}{5 + i}") for c in "BG"] for i in range(len(SEGMENTS))]},
        "sw": {"visits_year": val(x.ws("04_Similarweb"), f"I{15 + max(0, n - 10)}"), "funnel": {sc: [val(x.ws("05_Воронка_SW"), f"{c}{r}") for r in (5, 7, 9, 11, 14)] for sc, c in zip(SC, "BCD")}},
        "sverka": {k: [val(w6, f"{c}{r}") for c in "BCD"] for k, r in (("td", 5), ("bu", 6), ("sw", 7), ("min", 8), ("avg", 9), ("max", 10))},
        "errors": errors(wb),
    }
    results["tb"] = res
    x.save_close()
