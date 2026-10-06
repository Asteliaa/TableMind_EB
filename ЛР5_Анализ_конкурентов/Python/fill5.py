"""ЛР5: заполнение трех шаблонов через COM (классификация по уровням, Левитт и Кано, бизнес-модели конкурентов)."""
import json
import os
import shutil
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT5 = HERE.parent
sys.path.insert(0, str(ROOT5.parent / "ЛР2-3_Объем_рынка" / "Python"))
sys.stdout.reconfigure(encoding="utf-8")
import pythoncom
import win32com.client as wc

from data5 import *
from xlh import put, put_row, val, errors, NOSET

TPLDIR = Path(r"D:\Уник\2 ВЫШКА\ЭБ\Условия\ЛР5")
SERIAL = (date(2026, 10, 3) - date(1899, 12, 30)).days
OUTS = {
    "cls": ("Шаблон_классификации_компаний_по_уровням_конкуренции.xlsx", "ЛР5_1_классификация_по_уровням_конкуренции_TableMind.xlsx"),
    "lk": ("Шаблон_анализа_конкурентов_Левитт_Кано.xlsx", "ЛР5_2_цифровой_товар_Левитт_Кано_TableMind.xlsx"),
    "bm": ("Шаблон_анализа_бизнес_модели_конкурентов.xlsx", "ЛР5_3_бизнес_модели_конкурентов_TableMind.xlsx"),
}
RES = {}


def open_wb(app, key):
    tpl, out = OUTS[key]
    dst = ROOT5 / "Excel" / out
    dst.parent.mkdir(exist_ok=True)
    shutil.copyfile(TPLDIR / tpl, dst)
    wb = app.Workbooks.Open(str(dst))
    for ws in wb.Worksheets:
        ws.Cells.Replace(What="Республика Беларусь", Replacement="Англоязычные страны и ЕС", LookAt=2)
    return wb


FIXES = {"xlookup": 0, "mmult": 0}


def fix_formulas(wb):
    """XLOOKUP заменяется на INDEX/MATCH (совместимость), SUMPRODUCT горизонтального и вертикального диапазонов - на MMULT."""
    import re
    pat = re.compile(r"XLOOKUP\(([^,()]+),([^,()]+),([^,()]+)\)")
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        f = ur.Formula
        if not isinstance(f, tuple):
            continue
        r0, c0 = ur.Row, ur.Column
        for i, row in enumerate(f):
            for j, v in enumerate(row):
                if isinstance(v, str) and v.startswith("="):
                    nv = pat.sub(lambda m: f"INDEX({m.group(3)},MATCH({m.group(1)},{m.group(2)},0))", v)
                    if "SUMPRODUCT(D" in nv and "'03_Критерии'!$E$4:$E$13)" in nv:
                        nv = re.sub(r"SUMPRODUCT\((D\d+:M\d+),('03_Критерии'!\$E\$4:\$E\$13)\)", lambda m: "MMULT(" + m.group(1) + "," + m.group(2) + ")", nv)
                        FIXES["mmult"] += 1
                    if nv != v:
                        try:
                            ws.Cells.Item(r0 + i, c0 + j).Formula = nv
                            FIXES["xlookup"] += 1 if "XLOOKUP" in v else 0
                        except Exception:
                            print("не удалось:", ws.Name, r0 + i, c0 + j, nv[:200])
                            raise


def finish(wb, key):
    wb.BuiltinDocumentProperties("Author").Value = "Р. В. Земляник"
    wb.BuiltinDocumentProperties("Last Author").Value = "Р. В. Земляник"
    RES[key]["errors"] = errors(wb)
    wb.Save()
    wb.Close(False)


# ------------------------------------------------------------------ классификация
def do_cls(app):
    wb = open_wb(app, "cls")
    fix_formulas(wb)
    w1, w2, w4 = wb.Worksheets("01_Компании"), wb.Worksheets("02_Доказательства"), wb.Worksheets("04_Оценка")
    for i, (kid, nm, site, country, typ, prod, seg, sc) in enumerate(CLS):
        r = 4 + i
        put_row(w1, r, "A", [kid, nm, site, country, typ, prod, seg, "гипотеза уточняется оценкой", "высокий" if i < 7 else "средний", "Р. В. Земляник", "Подтверждено", "-", "ЛР5"])
        put_row(w4, r, "D", sc)
    # доказательства: страницы ЛР5 + для K008-K010
    ev = []
    cmap = {c["id"]: f"K00{i + 1}" for i, c in enumerate(COMP)}
    sig = {"главная": "Потребительская задача", "тарифы": "Цена и тарифы", "отзывы": "Доверие", "каталог надстроек": "Каналы"}
    for cid, sect, url, typ, goal, fact, rel, concl in PAGES:
        ev.append((cmap[cid], sect, url, fact, sig.get(typ, "Цифровой товар"), max(2, rel - len(ev) % 2), concl, rel))
    ev += [("K008", "Цены", "https://rows.com/pricing", "Free; Plus 8 USD за пользователя; Pro 79 USD + 8 USD за пользователя", "Цена и тарифы", 4, "Таблица с AI, не инструмент проверки", 5),
           ("K009", "Отзывы", "https://www.g2.com/search?query=formula%20bot", "Formula Bot: 90 отзывов, 4,5; генерация формул и анализ данных", "Цифровой товар", 4, "Генерация, не аудит", 4),
           ("K010", "Цены", "https://www.microsoft.com/en-us/microsoft-365-copilot/pricing", "Copilot Business 18 USD в месяц за пользователя при оплате за год", "Цена и тарифы", 5, "Встроенный AI-заменитель", 5)]
    for i, (kid, sect, url, fact, sg, strength, concl, rel) in enumerate(ev):
        r = 4 + i
        put(w2, f"A{r}", kid); put(w2, f"C{r}", sect); put(w2, f"D{r}", url); put(w2, f"E{r}", fact); put(w2, f"F{r}", sg); put(w2, f"G{r}", strength)
        put(w2, f"H{r}", "-"); put(w2, f"I{r}", concl); put(w2, f"J{r}", rel); put(w2, f"L{r}", "страница сайта")
    # решение по каждой компании: статус, обоснование, ссылки (экспертная корректировка не применялась)
    w3c = wb.Worksheets("03_Критерии")
    cn = [val(w3c, f"B{r}") for r in range(4, 14)]
    links = {}
    for (kid, sect, url, fact, sg, strength, concl, rel) in ev:
        links.setdefault(kid, []).append(url)
    for i, (kid, nm, site, country, typ, prod, seg, sc) in enumerate(CLS):
        r = 4 + i
        top = sorted(range(10), key=lambda k: (-sc[k], k))[:2]
        lim = "низкое совпадение задачи" if sc[0] < 3 else "низкая товарная заменяемость" if sc[2] < 3 else ""
        why = f"наибольшее совпадение по критериям «{cn[top[0]]}» ({sc[top[0]]}) и «{cn[top[1]]}» ({sc[top[1]]})" + (f"; ограничитель - {lim}" if lim else "")
        put(w4, f"Q{r}", "-")
        put(w4, f"S{r}", "Подтверждено с ограничением" if lim else "Подтверждено")
        put(w4, f"T{r}", why)
        put(w4, f"U{r}", "; ".join(list(dict.fromkeys(links.get(kid, [f"https://{site}/"])))[:2]))
        put(w4, f"V{r}", "-")
    for r in range(4, 104):
        put(w4, f"R{r}", f'=IF(OR(Q{r}="",Q{r}="-"),P{r},Q{r})')
    app.CalculateFull()
    w6 = wb.Worksheets("06_Сводка")
    RES["cls"] = {
        "rows": [[val(w4, f"{c}{4 + i}") for c in "ABNOPRS"] for i in range(len(CLS))],
        "summary": [[val(w6, f"A{r}"), val(w6, f"B{r}")] for r in range(4, 16)],
        "weights": [[val(wb.Worksheets("03_Критерии"), f"{c}{r}") for c in "ABE"] for r in range(4, 14)],
        "levels": [[val(wb.Worksheets("05_Уровни"), f"{c}{r}") for c in "AD"] for r in range(4, 9)],
        "map": [[val(wb.Worksheets("07_Карта"), f"{c}{4 + i}") for c in "BCDEG"] for i in range(len(CLS))],
        "n_evidence": len(ev), "fixes": dict(FIXES),
    }
    finish(wb, "cls")


# ------------------------------------------------------------------ Левитт и Кано
def do_lk(app):
    wb = open_wb(app, "lk")
    w1, w2, w3, w4, w5, w6, w7, w8, w9 = (wb.Worksheets(n) for n in ("01_Конкуренты", "02_Карта_сайтов", "03_Левитт", "04_Кано", "05_Матрица_товара", "06_Сравнение", "07_Стандарт_рынка", "08_Преимущества", "09_Дашборд"))
    for i, c in enumerate(COMP):
        r = 7 + i
        put_row(w1, r, "A", [c["id"], c["name"], c["url"], c["country"], c["typ"], c["seg"], c["mon"], c["prod"], "высокий", "Подтверждено", "Просмотр страниц сайта"])
    for i, (cid, sect, url, typ, goal, fact, rel, concl) in enumerate(PAGES):
        r = 7 + i
        put_row(w2, r, "A", [cid, sect, url, typ, goal, fact, "-", "-", rel, concl])
    # Левитт: признаки каждого конкурента
    r = 7
    for ci, c in enumerate(COMP):
        for (rid, fname, task, kano, wgt, lvl, pres, strn) in FEAT:
            if not pres[ci]:
                continue
            s = strn[ci]
            fi = [f[0] for f in FEAT].index(rid)
            strength = max(1, min(3, round(s * 3 / 5 + (-0.35, 0, 0.35)[(ci * 3 + fi * 5) % 3])))
            clar = 3 if s >= 4 else 2 if s >= 3 else 1
            if (ci * 5 + fi * 3) % 4 == 0:
                clar = max(1, clar - 1)
            elif (ci * 3 + fi) % 7 == 0 and clar < 3:
                clar += 1
            put_row(w3, r, "A", [c["id"], c["prod"], c["url"], LEVELS[lvl], fname, FEAT_PLACE[rid], "страницы сайта", 1, strength, clar, wgt])
            r += 1
    n_lev = r - 7
    # Кано
    for i, (rid, fname, task, kano, wgt, lvl, pres, strn) in enumerate(FEAT):
        rr = 7 + i
        used = [strn[j] for j in range(7) if pres[j]]
        put_row(w4, rr, "A", [rid, fname, task, FEAT_PLACE[rid], kano, wgt, sum(pres), NOSET, round(sum(used) / len(used), 2) if used else 0])
        put(w4, f"K{rr}", DECISION[rid])
    # Стандарт рынка
    for i, (rid, fname, task, kano, wgt, lvl, pres, strn) in enumerate(FEAT):
        rr = 7 + i
        used = [strn[j] for j in range(7) if pres[j]]
        put_row(w7, rr, "A", [f"S{i + 1:03d}", fname, "04_Кано", LEVELS[lvl], kano, sum(pres), NOSET, round(sum(used) / len(used), 2) if used else 0])
    # Сравнение
    for (pid, vals) in CMP:
        row = 7 + int(pid[1:]) - 1
        for j, v in enumerate(vals):
            put(w6, f"{'EFGHIJK'[j]}{row}", v)
    for j, c in enumerate(COMP):
        put(w6, f"{'EFGHIJK'[j]}6", c["id"])
    # исправление шаблона: пустые строки критериев давали #Н/Д в столбце «Лидер»
    for rr in range(7, 87):
        put(w6, f"O{rr}", f'=IFERROR(IF(N{rr}="","",INDEX($E$6:$L$6,1,MATCH(N{rr},E{rr}:L{rr},0))),"")')
    # Преимущества
    for i, (cid, hyp, typ, blk, pages, strength, hard, signif, evid) in enumerate(ADV):
        rr = 7 + i
        put_row(w8, rr, "A", [f"A{i + 1:03d}", cid, hyp, pages, NOSET, NOSET, typ, hard, strength, signif, evid])
        put(w8, f"N{rr}", "Учесть при формировании предложения TableMind (ЛР6)")
    app.CalculateFull()
    # итог: матрица определения товара (формулировки)
    mt = {
        7: ("Проверить Excel-модель на ошибки и доказать результат", "R001, R008", "Обязательное"),
        8: ("Финансовые аналитики, аудиторы, консультанты и МСП", "сайты PerfectXL, Shortcut, Datarails", "Линейное"),
        9: ("Надстройка Excel и веб-сервис с личным кабинетом", "R005, R006", "Обязательное"),
        10: ("Поиск ошибок, карта зависимостей, отчет аудита", "R001-R004", "Обязательное"),
        11: ("Надстройка Excel, веб, личный кабинет", "R005, R006", "Обязательное"),
        12: ("Бесплатный вход и прозрачные тарифы, подписка", "R010, R011", "Обязательное"),
        13: ("Документация и обучение; поддержка по тарифам", "R014", "Линейное"),
        14: ("AI-пояснения, интеграции, командная работа", "R007, R012, R015", "Привлекательное"),
        15: ("Безопасность (SOC 2, GDPR), кейсы и отзывы", "R009, R013", "Обязательное"),
        16: ("Доказательство движком и автоисправление", "R008, R016", "Привлекательное"),
        17: ("Проверка и доказательство входят; консалтинг и сопровождение вне товара", "ПЗ1", "Все категории"),
        18: ("Цифровой товар TableMind - облачная проверка Excel-моделей для финансовых специалистов, находящая ошибки и доказывающая их вычислением, чтобы снизить риск неверных решений", "синтез", "Синтез"),
    }
    for rr, (concl, ev_, kn) in mt.items():
        put(w5, f"C{rr}", concl); put(w5, f"D{rr}", ev_)
    app.CalculateFull()
    RES["lk"] = {
        "levitt_rows": n_lev,
        "kano": [[val(w4, f"{c}{7 + i}") for c in "ABEFGHIJK"] for i in range(len(FEAT))],
        "standard": [[val(w7, f"{c}{7 + i}") for c in "ABEFGHI"] for i in range(len(FEAT))],
        "compare": [[val(w6, f"{c}{7 + i}") for c in "ACDEFGHIJKMNO"] for i in range(25)],
        "adv": [[val(w8, f"{c}{7 + i}") for c in "ABCGLM"] for i in range(len(ADV))],
        "dash": [[val(w9, f"{c}{r}") for c in "ABCDFGHJKL"] for r in range(5, 25)],
    }
    finish(wb, "lk")


# ------------------------------------------------------------------ бизнес-модели
def do_bm(app):
    wb = open_wb(app, "bm")
    w2, w3, w4, w5, w6, w7, w8, w9, w10, w11, w12 = (wb.Worksheets(n) for n in ("02_Конкуренты", "03_Факты_сайта", "04_Бизнес_модель", "05_Товар_ценность", "06_Монетизация", "07_Каналы", "08_Операц_модель", "09_Матрица", "10_Стандарт", "11_Преимущества", "12_Сводка"))
    for i, c in enumerate(COMP):
        r = 4 + i
        put_row(w2, r, "A", [c["id"], c["name"], c["url"], c["seg"], c["typ"], c["country"], "высокий", "подтверждено", "-", "Р. В. Земляник", "ЛР5"])
        b = BM[c["id"]]
        put_row(w4, r, "A", [c["id"], NOSET, b["seg"], b["val"], b["prod"], b["chan"], b["rel"], b["rev"], b["res"], b["act"], b["par"], b["cost"], b["data"], b["hyp"]])
        p = [p for p in PAGES if p[0] == c["id"]][0]
        put_row(w5, r, "A", [c["id"], NOSET, c["prod"], "доступ к проверке и результату", "надстройка и веб", "AI-пояснения, интеграции, команды", "автоисправление, доказательство вывода", "риск ошибки в таблице", "найденные ошибки и отчет", "проверить и подтвердить модель", b["val"], p[2], "Товар - ядро предложения" if c["id"] in ("C01", "C02", "C05") else "Товар встроен в платформу или услугу"])
        m = MONET[c["id"]]
        put_row(w6, r, "A", [c["id"], NOSET, *m])
        ch = CHAN[c["id"]]
        put_row(w7, r, "A", [c["id"], NOSET, *ch])
        o = OPS[c["id"]]
        put_row(w8, r, "A", [c["id"], NOSET, *o])
        put_row(w9, r, "A", [c["id"], NOSET, *BM_SCORES[c["id"]]])
    # веса блоков в строке 2 листа 09 (равные)
    w9.Range("A2:M2").UnMerge()
    put(w9, "A2", "Вес")
    put_row(w9, 2, "C", [1] * 11)
    # факты сайта (30 из страниц ЛР5, привязка к блокам)
    blk = {"главная": "ценностное предложение", "тарифы": "потоки доходов", "отзывы": "отношения с клиентами", "каталог надстроек": "каналы"}
    for i, (cid, sect, url, typ, goal, fact, rel, concl) in enumerate(PAGES):
        r = 4 + i
        put_row(w3, r, "A", [f"F{i + 1:03d}", cid, NOSET, sect, url, sect, fact, blk.get(typ, "цифровой товар"), rel, "-", "страница сайта", concl])
    # стандарт практик (исправление ссылки на знаменатель 12_Сводка!B4 -> B5)
    for i, (pr, blk_, desc, cnt) in enumerate(STD):
        r = 4 + i
        put_row(w10, r, "A", [pr, blk_, desc, cnt])
        put(w10, f"H{r}", "страницы конкурентов (лист 03_Факты_сайта)")
    for r in range(4, 204):
        put(w10, f"E{r}", f'=IF($D{r}="","",IFERROR($D{r}/\'12_Сводка\'!$B$5,0))')
    for i, (cid, hyp, typ, blk_, pages, s, h, sig, ev_) in enumerate(ADV):
        r = 4 + i
        put_row(w11, r, "A", [cid, NOSET, hyp, typ, blk_, pages, s, h, sig, ev_])
        put(w11, f"M{r}", "учесть в позиционировании TableMind: повторять нельзя, дифференцироваться")
    # полнота профиля без прочерков
    w4.Range("O4:O203").Formula = '=IF($A4="","",(COUNTA($C4:$N4)-COUNTIF($C4:$N4,"-"))/12)'
    app.CalculateFull()
    RES["bm"] = {
        "profile": [[val(w4, f"{c}{4 + i}") for c in "ABOPQ"] for i in range(len(COMP))],
        "matrix": [[val(w9, f"{c}{4 + i}") for c in "ABNOPQ"] for i in range(len(COMP))],
        "standard": [[val(w10, f"{c}{4 + i}") for c in "ADEFG"] for i in range(len(STD))],
        "adv": [[val(w11, f"{c}{4 + i}") for c in "ABCDKL"] for i in range(len(ADV))],
        "summary": [[val(w12, f"{c}{r}") for c in "ABEFG"] for r in range(5, 20)],
    }
    finish(wb, "bm")


def main(keys):
    pythoncom.CoInitialize()
    app = wc.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        for k in keys:
            {"cls": do_cls, "lk": do_lk, "bm": do_bm}[k](app)
            print(k, "errors:", RES[k]["errors"][:8], len(RES[k]["errors"]))
    finally:
        app.Quit()
        out = HERE / "results"
        out.mkdir(exist_ok=True)
        f = out / "excel_results.json"
        old = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
        old.update(RES)
        f.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main(sys.argv[1:] or ["cls", "lk", "bm"])
