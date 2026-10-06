"""ЛР6.1 (ценностное предложение и цифровой товар) и ЛР6.2 (Канвас)."""
import json
import sys

from xl6 import *
from data6a import *

RES = {}


def do_vp(app):
    wb = open_copy(app, "Шаблон_формулирования_ценностного_предложения_и_цифрового_товара.xlsx", "ЛР6_1_ценностное_предложение_и_цифровой_товар_TableMind.xlsx")
    g = wb.Worksheets
    fill_table(g("Исходные данные"), 5, [[s[0], SERIAL, s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8]] for s in SRC], 10, 5)
    fill_table(g("Проблемы клиентов"), 5, [[p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8], p[9], None, None, p[10]] for p in PROB], 13, 5)
    fill_table(g("Конкурентные решения"), 5, [[c[0], "https://" + c[1], c[2], c[3], c[4], c[5], c[6], c[7], c[8], None, None, c[9]] for c in COMPSOL], 12, 5)
    fill_table(g("Ценностные предложения"), 5, [[v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7], v[8], v[9], None, None, v[10]] for v in VP], 13, 5)
    fill_table(g("Цифровой товар"), 5, [[p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8], None, None, None, p[9]] for p in PROD], 13, 5)
    fill_table(g("Единица монетизации"), 5, [[m[0], m[1], m[2], m[3], m[4], m[5], m[6], None, m[7], m[8], m[9], m[10], None, None] for m in MONET], 14, 5)
    fill_table(g("Бизнес-модель"), 5, [[b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7], b[8], b[9], b[10], None, b[11]] for b in BMODEL], 13, 5)
    fill_table(g("Проверка связности"), 5, [[c[0], c[1], c[2], c[3], c[4], c[5], None, None, c[6], c[7]] for c in CHECKS], 10, 5)
    app.CalculateFull()
    RES["vp"] = {
        "problems": [[val(g("Проблемы клиентов"), f"{c}{5 + i}") for c in "ADKL"] for i in range(len(PROB))],
        "solutions": [[val(g("Конкурентные решения"), f"{c}{5 + i}") for c in "ACJK"] for i in range(len(COMPSOL))],
        "vp": [[val(g("Ценностные предложения"), f"{c}{5 + i}") for c in "ABKL"] for i in range(len(VP))],
        "product": [[val(g("Цифровой товар"), f"{c}{5 + i}") for c in "ACJKL"] for i in range(len(PROD))],
        "monet": [[val(g("Единица монетизации"), f"{c}{5 + i}") for c in "ACFGHMN"] for i in range(len(MONET))],
        "bm": [[val(g("Бизнес-модель"), f"{c}{5 + i}") for c in "AL"] for i in range(len(BMODEL))],
        "checks": [[val(g("Проверка связности"), f"{c}{5 + i}") for c in "AGH"] for i in range(len(CHECKS))],
        "dash": [[val(g("Дашборд"), f"{c}{r}") for c in "AB"] for r in range(5, 15)],
    }
    RES["vp"]["errors"] = finish(wb)


def do_canvas(app):
    wb = open_copy(app, "Шаблон_сборки_бизнес_модели_по_Канвас.xlsx", "ЛР6_2_канвас_бизнес_модели_TableMind.xlsx")
    g = wb.Worksheets
    blocks = ["Клиентские сегменты", "Ценностное предложение", "Каналы", "Отношения с клиентами", "Потоки доходов", "Ключевые ресурсы", "Ключевые виды деятельности", "Ключевые партнеры", "Структура затрат"]
    blk_of = {"Отчет ЛР1": "Клиентские сегменты", "Отчет ЛР2-3": "Потоки доходов", "Отчет ЛР4": "Каналы", "Отзывы": "Отношения с клиентами", "Страница цен": "Потоки доходов", "Главная страница": "Ценностное предложение", "Каталог": "Каналы", "Документ проекта": "Ключевые ресурсы", "Отчет ЛР5": "Ценностное предложение"}
    fill_table(g("01_Данные_анализа"), 3, [[NOSET, s[1], s[3], s[2], s[4], s[5], s[6], blk_of.get(s[1], "Ценностное предложение"), s[7], 5, 4, None, "использовано", s[0]] for s in SRC], 14, 0)
    ws = g("02_Канвас")
    for i, (blk, form, ev, rel, comp, risk, chk) in enumerate(CANVAS):
        r = 3 + blocks.index(blk)
        ws.Range(f"D{r}").Value2 = form
        ws.Range(f"E{r}").Value2 = ev
        ws.Range(f"F{r}").Value2 = rel
        ws.Range(f"G{r}").Value2 = comp
        ws.Range(f"H{r}").Value2 = risk
        ws.Range(f"I{r}").Value2 = chk
    fill_table(g("03_Сегменты"), 3, [[NOSET, s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8], None, None] for s in SEGS], 11, 0)
    fill_table(g("04_Ценность_товар"), 3, [
        [NOSET, "P-001", PROB[0][3], "ручная проверка, PerfectXL, универсальный AI", "модель без скрытых ошибок перед передачей", VP[0][7], PROD[0][2], "веб-сервис", "поиск ошибок формул", "веб-аудит с отчетом", "объяснения, исправления, история", "Обязательное", 4, 4, 5, None, None],
        [NOSET, "P-002", PROB[1][3], "Copilot, ChatGPT, ручная проверка результата", "уверенность в числах, полученных с помощью AI", VP[1][7], PROD[1][2], "надстройка", "проверка книги в Excel", "подсветка находок в Excel", "проверка после правки", "Привлекательное", 5, 3, 4, None, None],
        [NOSET, "P-005", PROB[4][3], "отказ от внешних сервисов", "проверка без утечки данных", VP[2][7], PROD[0][2], "веб-сервис", "обработка без хранения", "политика данных GDPR", "план SOC 2", "Обязательное", 3, 4, 4, None, None],
    ], 17, 0)
    fill_table(g("05_Каналы_отношения"), 3, [[NOSET, c[1], c[2], c[3], c[4], c[5], c[6], c[7], c[8], None, None] for c in CHAN], 11, 0)
    fill_table(g("06_Монетизация"), 3, [
        [NOSET, "Файл (разовый аудит)", "оплата проверки файла", "число проверенных файлов", "9 EUR за файл", "разовые платежи", "AI-помощники от 8-10 USD в месяц", "результат проверки", 4, 5, 2, 4, None, None],
        [NOSET, "Пользователь Pro", "ежемесячная подписка", "число активных пользователей", "19 EUR в месяц", "подписка", "Rows Plus 8 USD, Copilot 18 USD", "постоянная проверка", 4, 5, 5, 4, None, None],
        [NOSET, "Место Team", "ежемесячная подписка за место", "число мест", "29 EUR за место", "подписка по местам", "Shortcut Teams 100 USD за место", "командная работа", 4, 5, 5, 3, None, None],
    ], 14, 0)
    fill_table(g("07_Операционная_основа"), 3, [[NOSET, o[1], o[2], o[3], o[4], o[5], o[6], o[7], o[8], o[9], None] for o in OPS], 11, 0)
    fill_table(g("08_Затраты"), 3, [[NOSET, c[1], c[2], c[3], c[4], c[5], c[6], c[7], c[8], c[9], None] for c in COSTS], 11, 0)
    ws = g("09_Проверка_связности")
    for i, (sc, evid, todo) in enumerate(CANV_CHECK):
        r = 3 + i
        ws.Range(f"D{r}").Value2 = evid
        ws.Range(f"E{r}").Value2 = sc
        ws.Range(f"G{r}").Value2 = todo
    ws = g("10_Паспорт")
    pas = {
        4: "Скрытые ошибки в формулах Excel-моделей и невозможность подтвердить результат AI-помощников (P-001, P-002)",
        5: VP[0][7],
        6: "TableMind: веб-аудит файла Excel и надстройка Excel с поиском ошибок, объяснениями и отчетом",
        7: "файл (9 EUR), пользователь Pro (19 EUR в месяц), место Team (29 EUR в месяц)",
        8: "поиск по проблемным запросам, AppSource, партнерские фирмы, Product Hunt",
        9: "бесплатный старт, самообслуживание, шаблоны отчетов, документация и обучение",
        10: "разовые платежи за аудит и подписки Pro и Team",
        11: "движок проверки, база правил и eval-набор, соответствие GDPR, команда разработки",
        12: "развитие правил и движка, оценка точности, контент и партнерства",
        13: "Microsoft AppSource, бухгалтерские и консалтинговые фирмы, провайдеры LLM и облака",
        14: "вычисления и LLM, разработка, продвижение, поддержка, резерв 10 %",
        15: "независимое доказательство каждой находки движком при прозрачной цене и бесплатном входе",
        16: "Финансисты платят 19 EUR в месяц за проверку с доказательствами, если получают первый результат бесплатно в течение одного сеанса",
    }
    for r, t in pas.items():
        ws.Range(f"B{r}").Value2 = t
    app.CalculateFull()
    RES["canvas"] = {
        "canvas": [[val(g("02_Канвас"), f"{c}{3 + i}") for c in "AFGHJ"] for i in range(9)],
        "segments": [[val(g("03_Сегменты"), f"{c}{3 + i}") for c in "BJK"] for i in range(len(SEGS))],
        "value": [[val(g("04_Ценность_товар"), f"{c}{3 + i}") for c in "GPQ"] for i in range(3)],
        "channels": [[val(g("05_Каналы_отношения"), f"{c}{3 + i}") for c in "DJK"] for i in range(len(CHAN))],
        "monet": [[val(g("06_Монетизация"), f"{c}{3 + i}") for c in "BMN"] for i in range(3)],
        "ops": [[val(g("07_Операционная_основа"), f"{c}{3 + i}") for c in "CK"] for i in range(len(OPS))],
        "costs": [[val(g("08_Затраты"), f"{c}{3 + i}") for c in "BK"] for i in range(len(COSTS))],
        "checks": [[val(g("09_Проверка_связности"), f"{c}{3 + i}") for c in "EF"] for i in range(12)],
        "dash": [[val(g("11_Дашборд"), f"{c}{r}") for c in "AB"] for r in range(4, 10)],
        "n_src": len(SRC),
    }
    RES["canvas"]["errors"] = finish(wb)


def main(keys):
    app = start()
    try:
        for k in keys:
            {"vp": do_vp, "canvas": do_canvas}[k](app)
            print(k, "errors:", RES[k]["errors"][:6], len(RES[k]["errors"]))
    finally:
        app.Quit()
        out = HERE / "results"
        out.mkdir(exist_ok=True)
        f = out / "excel_results.json"
        old = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
        old.update(RES)
        f.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main(sys.argv[1:] or ["vp", "canvas"])
