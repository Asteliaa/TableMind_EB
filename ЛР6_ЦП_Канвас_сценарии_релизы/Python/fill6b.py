"""ЛР6.3 (бизнес-логика и сценарии) и ЛР6.5 (релизы)."""
import json
import sys
from datetime import date

from xl6 import *
from data6b import *

RES = {}
RN = {r[0]: r[1] for r in ROLES}


def ser(s):
    y, m, d = map(int, s.split("-"))
    return (date(y, m, d) - date(1899, 12, 30)).days


def do_logic(app):
    wb = open_copy(app, "Шаблон_бизнес_логики_и_пользовательских_сценариев.xlsx", "ЛР6_3_бизнес_логика_и_сценарии_TableMind.xlsx")
    g = wb.Worksheets
    fill_table(g("Банк анализа"), 4, [[b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7], b[8], None, b[9], b[10], None, b[11]] for b in BANK], 14, 4)
    fill_table(g("Бизнес-логика"), 4, [[l[0], l[1], l[2], l[3], l[4], l[5], l[6], l[7], l[8], l[9], l[10], None, None, "Требует проверки", l[11]] for l in LOGIC], 15, 4)
    fill_table(g("Роли и сегменты"), 4, [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], None, r[8], r[9], r[10], r[11], r[12]] for r in ROLES], 14, 4)
    fill_table(g("Пользовательские сценарии"), 4, [[s[0], s[1], RN[s[2]], s[3], s[4], s[5], s[6], s[7], s[8], s[9], s[10], s[11], s[12], s[13], s[14], s[15], s[16], None, None, "Требует проверки", "A-001"] for s in SCEN], 21, 4)
    fill_table(g("Карта сценариев"), 4, [[s[0], RN[s[2]], s[1], "поиск, партнеры", "страница продукта", "по желанию после результата", "минимальная", s[10], "тарифы" if s[4] == "Монетизационный" else "после ценности", "повторная проверка", "справка", "письмо с отчетом", s[9], s[3], "ЛР6"] for s in SCEN], 15, 4)
    fill_table(g("Функциональные требования"), 4, [[r[0], r[1], r[2], RN[r[3]], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], None, None, None, "Требует проверки", r[12]] for r in REQ], 17, 4)
    fill_table(g("Приоритизация"), 4, [[p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8], p[9], None, None, p[10], p[11]] for p in PRIO], 14, 4)
    app.CalculateFull()
    RES["logic"] = {
        "bank": [[val(g("Банк анализа"), f"{c}{4 + i}") for c in "ADJM"] for i in range(len(BANK))],
        "logic": [[val(g("Бизнес-логика"), f"{c}{4 + i}") for c in "ABLM"] for i in range(len(LOGIC))],
        "roles": [[val(g("Роли и сегменты"), f"{c}{4 + i}") for c in "ABI"] for i in range(len(ROLES))],
        "scen": [[val(g("Пользовательские сценарии"), f"{c}{4 + i}") for c in "ABERS"] for i in range(len(SCEN))],
        "req": [[val(g("Функциональные требования"), f"{c}{4 + i}") for c in "ABGMNO"] for i in range(len(REQ))],
        "prio": [[val(g("Приоритизация"), f"{c}{4 + i}") for c in "ABKL"] for i in range(len(PRIO))],
        "conn": [[val(g("Связность"), f"{c}{r}") for c in "AFGH"] for r in range(4, 14)],
        "dash": [[val(g("Дашборд"), f"{c}{r}") for c in "AB"] for r in range(5, 15)],
        "dist": [[val(g("Дашборд"), f"{c}{r}") for c in "FGHI"] for r in range(5, 10)],
    }
    RES["logic"]["errors"] = finish(wb)


def do_release(app):
    wb = open_copy(app, "Шаблон_деления_разработки_сайта_на_релизы.xlsx", "ЛР6_5_деление_на_релизы_TableMind.xlsx")
    g = wb.Worksheets
    bank = [[b[0], b[1], b[2], b[3], b[4], "финансовые специалисты", b[5], b[5], b[6] + 1 if b[6] < 5 else 5, b[7], "учесть при составе релизов", ""] for b in BANK]
    fill_table(g("Банк_анализа"), 2, bank, 12, 3)
    reg = []
    for u in UNITS:
        reg.append([u[0], u[1], u[2], RN[u[3]], u[4], u[5]] + u[6] + [None, None, None, None, u[7]])
    fill_table(g("Реестр_единиц"), 2, reg, 21, 5)
    fill_table(g("Первый_релиз"), 4, [list(f) for f in FIRST], 8, 4)
    fill_table(g("Следующие_релизы"), 4, [list(n) for n in NEXT], 9, 4)
    fill_table(g("Дорожная_карта"), 4, [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], ser(r[9]), ser(r[10])] for r in ROAD], 11, 4)
    fill_table(g("Проверка_ценности"), 4, [list(v) for v in VALCHK], 9, 4)
    fill_table(g("Риски_зависимости"), 4, [[r[0], r[1], r[2], r[3], r[4], r[5], None, r[6], r[7], r[8]] for r in RISKS], 10, 4)
    app.CalculateFull()
    RES["release"] = {
        "units": [[val(g("Реестр_единиц"), f"{c}{2 + i}") for c in "ACQRST"] for i in range(len(UNITS))],
        "eval": [[val(g("Оценка_релизов"), f"{c}{r}") for c in "ABCDEFGHIJ"] for r in range(4, 7)],
        "dash": [[val(g("Дашборд"), f"{c}{r}") for c in "AB"] for r in range(4, 10)],
        "crit": [[val(g("Дашборд"), f"{c}{r}") for c in "AB"] for r in range(13, 20)],
        "risks": [[val(g("Риски_зависимости"), f"{c}{4 + i}") for c in "ABCG"] for i in range(len(RISKS))],
    }
    RES["release"]["errors"] = finish(wb)


def main(keys):
    app = start()
    try:
        for k in keys:
            {"logic": do_logic, "release": do_release}[k](app)
            print(k, "errors:", RES[k]["errors"][:6], len(RES[k]["errors"]))
    finally:
        app.Quit()
        f = HERE / "results" / "excel_results.json"
        old = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
        old.update(RES)
        f.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main(sys.argv[1:] or ["logic", "release"])
