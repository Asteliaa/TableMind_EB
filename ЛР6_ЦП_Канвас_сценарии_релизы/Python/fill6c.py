"""ЛР6.4: заполнение шаблона перехода к страницам и экранам."""
import json
import sys

from xl6 import *
from data6c import *

RES = {}
NAMES = {u[0]: u[2] for u in UNITS_UI}


def main():
    app = start()
    try:
        wb = open_copy(app, "Шаблон_перехода_к_страницам_и_экранам.xlsx", "ЛР6_4_страницы_и_экраны_TableMind.xlsx")
        g = wb.Worksheets
        CL = 20
        fill_table(g("Банк_выводов"), 5, [[b[0], b[1], b[2], b[3], b[4], 4 if b[6] >= 3 else 3, b[7], UI_BANK[b[0]][0], rl(ROLES[0][0]) if i % 2 == 0 else rl(ROLES[1][0]), "SC-001", UI_BANK[b[0]][1]] for i, b in enumerate(BANK)], 11, CL)
        fill_table(g("Роли_и_сегменты"), 5, [[rl(r[0]), r[1], r[2], r[3], r[4], "Средняя", "периодически", r[7], r[8], r[10]] for r in ROLES], 10, CL)
        fill_table(g("Пользовательские_сценарии"), 5, [[sc(s[0]), rl(s[2]), s[3], s[6], s[7], s[8], s[9], "файл Excel, находки, отчет", s[11], s[14], s[13], 3, None, None, None] for s in SCEN], 15, CL)
        rows = []
        for u in UNITS_UI:
            rows.append([u[0], u[1], u[2], u[3], sc(u[4]), rl(u[5]), u[6], u[7], u[8], u[9], u[10], u[11], "нормальное; пустое; загрузка; ошибка; успех", "см. лист «Состояния_и_ошибки»", u[12], u[13], None, u[15], u[14], None, None])
        fill_table(g("Страницы_и_экраны"), 5, rows, 21, CL)
        fill_table(g("Матрица_сценарий_интерфейс"), 5, [[f"M-{i + 1:03d}", sc(l[0]), l[1], l[2], l[3], l[4], l[5], l[6], l[7]] for i, l in enumerate(LINKS)], 9, CL)
        fill_table(g("Навигация_и_карта"), 5, [[f"N-{i + 1:03d}", n[0], n[1], n[2], n[3], rl(n[4]), n[5], n[6], n[7], n[8]] for i, n in enumerate(NAV)], 10, CL)
        fill_table(g("Состояния_и_ошибки"), 5, [[f"S-{i + 1:03d}", u[0]] + list(STATES[u[0]]) + [None] for i, u in enumerate(UNITS_UI)], 11, CL)
        fill_table(g("Контент_и_данные"), 5, [list(c) for c in CONTENT], 9, CL)
        nxt = {}
        for l in LINKS:
            nxt.setdefault(l[1], l[6])
        fill_table(g("Паспорт_экрана"), 5, [[u[0], u[2], u[1], sc(u[4]), rl(u[5]), u[3], u[7], u[8], u[11], "нормальное; пустое; загрузка; ошибка; успех", ("переход к " + NAMES[nxt[u[0]]]) if nxt.get(u[0]) in NAMES else "возврат к предыдущему шагу", u[18], u[19], None, None] for u in UNITS_UI], 15, CL)
        fill_table(g("Приоритизация"), 5, [[u[0], u[2], sc(u[4]), u[16][0], u[16][1], u[16][2], u[16][3], u[16][4], None, None, u[17], "Включить в первую версию" if u[0] in ("U-001", "U-002", "U-003", "U-004", "U-005", "U-006", "U-007", "U-008", "U-009", "U-010", "U-017") else "Отложить"] for u in UNITS_UI], 12, CL)
        app.CalculateFull()
        n = len(UNITS_UI)
        RES["screens"] = {
            "scen": [[val(g("Пользовательские_сценарии"), f"{c}{5 + i}") for c in "AMNO"] for i in range(len(SCEN))],
            "units": [[val(g("Страницы_и_экраны"), f"{c}{5 + i}") for c in "ABCQRTU"] for i in range(n)],
            "states": [[val(g("Состояния_и_ошибки"), f"{c}{5 + i}") for c in "BK"] for i in range(n)],
            "passport": [[val(g("Паспорт_экрана"), f"{c}{5 + i}") for c in "ABNO"] for i in range(n)],
            "prio": [[val(g("Приоритизация"), f"{c}{5 + i}") for c in "ABIJL"] for i in range(n)],
            "conn": [[val(g("Проверка_связности"), f"{c}{r}") for c in "ADE"] for r in range(5, 16)],
            "dash": [[val(g("Дашборд"), f"{c}{r}") for c in "AB"] for r in range(5, 20)],
            "n_links": len(LINKS), "n_nav": len(NAV),
        }
        RES["screens"]["errors"] = finish(wb)
        print("errors:", RES["screens"]["errors"][:8], len(RES["screens"]["errors"]))
    finally:
        app.Quit()
        f = HERE / "results" / "excel_results.json"
        old = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
        old.update(RES)
        f.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
