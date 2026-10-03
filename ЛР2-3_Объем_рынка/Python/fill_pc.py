"""Книга 1: оценка рынка через количество потенциальных клиентов."""
from inputs import *
from xlh import *

TPL_NAME = "shablon_excel_metodika_ocenki_rynka_cherez_kolichestvo_potencialnyh_klientov.xlsx"
OUT_NAME = "ЛР2-3_1_потенциальные_клиенты_TableMind.xlsx"

SRC_NOTE = "Similarweb Pro, 06-08.2026, мир; снимок 03.10.2026"


def fill(x: XL, sw, results):
    wb = x.open_copy(TPL_NAME, OUT_NAME)
    ws = x.ws("01_Параметры")
    put(ws, "B4", "TableMind: онлайн-сервис независимой проверки (аудита) Excel-моделей")
    put(ws, "B5", "Проверка таблиц и финансовых моделей")
    put(ws, "B6", "Веб-аудит файла Excel и надстройка: поиск ошибок формул, структуры и логики; тарифы Free, 9 EUR, Pro 19 EUR, Team 29 EUR")
    put(ws, "B7", "Англоязычные страны (США, Великобритания, Канада, Австралия, Ирландия) и ЕС-27; контрольная территория Беларусь и русскоязычный сегмент")
    put(ws, "B8", "B2B и B2C (профессионалы и малый бизнес), подписка freemium")
    put(ws, "B9", "EUR")
    put(ws, "C9", "Курс НБРБ 03.10.2026: 1 EUR = 3,3860 BYN")
    put(ws, "B11", SNAP_SERIAL)
    put(ws, "B12", "Р. В. Земляник, гр. 60131")

    ws = x.ws("02_Сегменты")
    for i, s in enumerate(SEGMENTS):
        r = 4 + i
        put_row(ws, r, "A", [s["name"], "B2B/B2C", s["src"], s["total"], s["border"], NOSET, s["need"], s["digital"], s["payable"],
                             NOSET, s["check"], s["freq"], NOSET, s["share"], NOSET, s["rel"], s["note"]])
    # лишние демо-строки шаблона обнуляются
    for r in range(4 + len(SEGMENTS), 11):
        put_row(ws, r, "A", ["Резерв", "B2B", "не используется", 0, 0, NOSET, 0, 0, 0, NOSET, 0, 0, NOSET, 0, NOSET, "Нет данных", "Свободная строка шаблона."])

    ws = x.ws("03_Сценарии")
    for i, sc in enumerate(SC):
        r = 4 + i
        m = PC_SCEN[sc]
        put_row(ws, r, "B", [m["need"], m["online"], m["pay"], m["check"], m["freq"], m["share"]])
    put(ws, "K4", "Низкие коэффициенты потребности, готовности и платежеспособности; осторожный выход на рынок.")
    put(ws, "K5", "Основной расчет на базе листа сегментов.")
    put(ws, "K6", "Высокая конверсия в платящих, больше мест в Team, заметная доля рынка.")

    ws = x.ws("05_Similarweb")
    n = len(sw)
    insert_rows_copy(ws, 6, n - 5, 4)  # было 5 строк (4-8), вставка внутри диапазона сумм
    for i, r in enumerate(sw):
        row = 4 + i
        ch_rel = r["ch"]["direct"] + r["ch"]["org_search"] + r["ch"]["paid_search"] + r["ch"]["referrals"]
        put_row(ws, row, "A", [r["name"], "https://www.similarweb.com/website/" + r["domain"] + "/", r["type"] + ": " + r["segment"],
                               r["visits_month"], r["geo_share"], NOSET, r["rel"], r["qual"], NOSET,
                               FUNNEL["v2l"]["base"], FUNNEL["l2p"]["base"], arppu("base")[0]])

    ws = x.ws("07_Источники")
    rows = [
        ("Количество клиентов", "Занятые по профессиям и число предприятий", "BLS OOH 2025 (США); SBA Advocacy 2025; European Commission SME Report 2025/2026; UK Business Population Estimates 2025",
         "см. Материалы_собранные/внешние_данные_2026-10-03.md", "Факт и допущение (масштаб по занятости 0,45 и 1,30)", "Средняя"),
        ("Коэффициент потребности", "Доля с потребностью проверять сложные таблицы", "Допущение Д-10 на основе ПЗ1 (ошибки в 86-94 % таблиц) и ЛР1", "журнал_допущений.md", "Экспертное допущение", "Низкая"),
        ("Цифровая готовность", "Доля покупающих сервис онлайн", "Допущение Д-10; Similarweb: каналы поиск и прямые заходы", "similarweb_снимок_2026-10-03.csv", "Оценка", "Средняя"),
        ("Платежеспособность", "Доля клиентов с бюджетом", "Цены конкурентов (PerfectXL от 69 EUR, Numerous 10 USD, Copilot 18 USD) и тарифы ПЗ1", "внешние_данные_2026-10-03.md", "Оценка", "Средняя"),
        ("Средний чек", "Выручка первого года на платящего", "Структура платящих (допущение Д-09) и тарифы 9/19/29 EUR", "inputs.py", "Расчет по допущению", "Средняя"),
        ("Similarweb", "Трафик конкурентов", "Similarweb Pro (пробный доступ), 06-08.2026", "https://www.similarweb.com/website/", "Оценка", "Средняя"),
    ]
    for i, t in enumerate(rows):
        r = 4 + i
        put(ws, f"A{r}", t[0]); put(ws, f"B{r}", t[1]); put(ws, f"C{r}", t[2]); put(ws, f"D{r}", t[3])
        put(ws, f"E{r}", SNAP_SERIAL); put(ws, f"F{r}", t[4]); put(ws, f"G{r}", t[5])
        put(ws, f"H{r}", "Заполнено при выполнении ЛР2-3; детали в приложении отчета.")

    x.app.CalculateFull()
    res = {
        "tam": val(x.ws("04_Расчет"), "C4"), "in_border": val(x.ws("04_Расчет"), "C5"), "addr": val(x.ws("04_Расчет"), "C6"),
        "sam": val(x.ws("04_Расчет"), "C7"), "som": val(x.ws("04_Расчет"), "C8"), "som_to_sam": val(x.ws("04_Расчет"), "C9"),
        "addr_share": val(x.ws("04_Расчет"), "C14"), "sam_per_client": val(x.ws("04_Расчет"), "C15"),
        "som_to_sw": val(x.ws("04_Расчет"), "C16"), "sw_year": val(x.ws("05_Similarweb"), f"N{n + 7}"),
        "scen": {sc: [val(x.ws("03_Сценарии"), f"{c}{4 + i}") for c in "HIJ"] for i, sc in enumerate(SC)},
        "seg": [[val(x.ws("02_Сегменты"), f"{c}{4 + i}") for c in "DFJMO"] for i in range(len(SEGMENTS))],
        "errors": errors(wb),
    }
    # дополнительно: сводка сайта (итоговая строка листа Similarweb)
    wsw = x.ws("05_Similarweb")
    tot = n + 5
    res["sw_row"] = {c: val(wsw, f"{c}{tot}") for c in "DFIMNO"}
    results["pc"] = res
    x.save_close()
