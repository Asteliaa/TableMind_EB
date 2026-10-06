"""Единое оформление книг Excel лабораторных работ (Excel COM, openpyxl не используется, чтобы не потерять диаграммы).

Для каждого листа. Убираются служебные пометки шаблона и валюты кроме EUR; удаляются неиспользуемые строки;
единый сине-голубой стиль, рамки, перенос текста, выравнивание по центру по вертикали, жирным только заголовки;
пустые ячейки таблиц заполняются прочерком; ширина столбцов по тексту; диаграммы дашбордов перекрашиваются.
Запуск: python style_xlsx.py путь_к_книге [путь_к_книге ...]
"""
import re
import sys
from pathlib import Path

import pythoncom
sys.path.insert(0, str(Path(__file__).resolve().parent))
import currency_eur
import status_map
import win32com.client as wc

AUTHOR = "Р. В. Земляник"

EUR_BYN, USD_BYN, GBP_BYN = 3.3860, 3.0051, 3.9663
USD_EUR = USD_BYN / EUR_BYN
GBP_EUR = GBP_BYN / EUR_BYN


def bgr(h):
    return int(h[4:6] + h[2:4] + h[0:2], 16)


TITLE, HEAD, SUB, INPUT, CALC = "1F4E79", "2E75B6", "DDEBF7", "BDD7EE", "EAF1FB"
BORDER = "9DC3E6"
TITLE_SRC = {"1F4E79", "1F4E78", "0F172A", "0B1F33"}
HEAD_SRC = {"305496", "5B9BD5", "2F75B5", "0F766E"}
SUB_SRC = {"D9EAF7"}
INPUT_SRC = {"FFF2CC", "FFFDF2"}
CALC_SRC = {"E2F0D9", "D9EAD3", "DDEFEA", "F2F2F2", "F3F4F6", "F8FBFD", "EAF3F8"}


def classify(fill_rgb, r, merged):
    """Роль ячейки по цвету заливки исходного шаблона."""
    if fill_rgb is None:
        return None
    R, G, B = int(fill_rgb[0:2], 16), int(fill_rgb[2:4], 16), int(fill_rgb[4:6], 16)
    lum = 0.299 * R + 0.587 * G + 0.114 * B
    if lum < 105:
        return "title" if (r == 1 or merged) else "head"
    if B - R > 60 and lum < 175:
        return "head"
    if R > 225 and G > 195 and B < 210 and R - B > 35:
        return "input"
    if lum >= 175 and (B - R >= 14 or (R > 215 and G > 215 and B > 215)):
        if B - R >= 14 and r <= 2 and fill_rgb in ("D9EAF7", "DDEBF7"):
            return "sub"
        return "label" if fill_rgb in ("D9EAF7", "DDEBF7") else "calc"
    if G >= R and G >= B:
        return "calc"
    return "calc"


def rgb_of(color):
    color = int(color)
    r, g, b = color & 255, (color >> 8) & 255, (color >> 16) & 255
    return f"{r:02X}{g:02X}{b:02X}"


def num_ru(v):
    s = f"{v:.1f}" if v < 10 else f"{v:.0f}"
    return s.replace(".", ",")


def conv_money(m):
    nums = re.findall(r"\d+(?:,\d+)?", m.group(0))
    cur = m.group(0).split()[-1]
    k = USD_EUR if cur == "USD" else GBP_EUR
    out = "-".join(num_ru(float(n.replace(",", ".")) * k) for n in nums)
    return out + " EUR"


MONEY = re.compile(r"\d+(?:,\d+)?(?:-\d+(?:,\d+)?)?\s?(?:USD|GBP)")
RULES = [
    (r"Excel-шаблон для подходов «сверху вниз» и «снизу вверх» с использованием Similarweb", "Расчет объема рынка подходами «сверху вниз» и «снизу вверх» с использованием Similarweb"),
    (r"Расчетный шаблон для оценки объема рынка: макро-ориентир «сверху вниз», расчет через клиентские сегменты «снизу вверх», проверка через цифровые признаки Similarweb\.",
     "Расчет объема рынка включает макро-ориентир «сверху вниз», расчет через клиентские сегменты «снизу вверх» и проверку через цифровые признаки Similarweb."),
    (r" Демонстрационные значения заменяются на фактические\.", ""),
    (r"Назначение Excel-шаблона", "Назначение расчета"),
    (r"Excel-шаблон для оценки, способен ли", "Оценка того, способен ли"),
    (r"Excel-шаблон оценки рынка через цифровую воронку", "Оценка рынка через цифровую воронку"),
    (r"Методика сценарной оценки рынка с учетом данных Similarweb: расчетный Excel-шаблон", "Сценарная оценка рынка с учетом данных Similarweb"),
    (r"Excel-шаблон: оценка объема рынка через поисковый спрос", "Оценка объема рынка через поисковый спрос"),
    (r"Шаблон переводит поисковый спрос", "Расчет переводит поисковый спрос"),
    (r"Ключевые показатели после заполнения шаблона\.", "Ключевые показатели итоговой оценки."),
    (r" Значения дополнительных методов можно заменить собственными расчетами\.", ""),
    (r"Шаблон итогового вывода", "Итоговый вывод"),
    (r"^Можно заменить данными из Excel по поисковому спросу$", "Данные книги по поисковому спросу"),
    (r"^Можно заменить данными из клиентской методики$", "Данные книги по потенциальным клиентам"),
    (r"^Можно заменить отдельным расчетом$", "Отдельный расчет"),
    (r"^Можно заменить на RUB, USD, EUR\.?$", "Расчетная валюта проекта"),
    (r"^Можно заменить на (свою|собственную|любую) бизнес-идею\.?$", "Название проекта"),
    (r"^Можно заменить на страну, область или город\.?$", "Территория расчета"),
    (r"^Можно заменить\.?$", "Задано проектом"),
    (r"Курс НБРБ 03\.10\.2026: 1 EUR = 3,3860 BYN", "Суммы в других валютах пересчитаны по курсам НБРБ на 03.10.2026"),
    (r"Нет доступа к рекламным кабинетам: ДОСНЯТЬ тестовые кампании", "Рекламные кабинеты не подключены, тестовые кампании запланированы после запуска MVP"),
    (r"Рекламных кабинетов нет: ДОСНЯТЬ тестовые кампании", "Рекламные кабинеты не подключены, тестовые кампании запланированы после запуска MVP"),
    (r"Рекламных кабинетов нет: ДОСНЯТЬ", "Рекламные кабинеты не подключены, тестовые кампании запланированы после запуска MVP"),
    (r"Своего сайта пока нет: ДОСНЯТЬ после запуска MVP \(31\.03\.2027\)", "Собственный сайт запускается вместе с MVP 31.03.2027, данные трафика появятся после запуска"),
    (r"Сайта и CRM пока нет: ДОСНЯТЬ после MVP", "Сайт и CRM запускаются вместе с MVP, данные появятся после запуска"),
    (r"CPC по Similarweb, USD", "CPC по Similarweb, у.е."),
    (r"Основной клиент — организация", "Основной клиент - организация"),
    # ---- ЛР4-6
    (r"^Excel-шаблон оценки уровня конкуренции и барьеров входа$", "Оценка уровня конкуренции и барьеров входа"),
    (r"^Что заполняет пользователь$", "Исходные данные"),
    (r"^Желтые ячейки заполняются пользователем; зеленые рассчитываются формулами$", "Голубые ячейки содержат исходные данные, светлые рассчитываются формулами"),
    (r"^Укажите анализируемый рынок после фиксации его границ\.$", "Анализируемый рынок задан границами ЛР1."),
    (r"^Заполните сайты конкурентов и показатели Similarweb\. ", "Сайты конкурентов и показатели Similarweb. "),
    (r"^Что проверить дополнительно$", "Рекомендация"),
    (r"^Что проверить$", "Контрольный вопрос"),
    (r"^Заполняется по G2, ", "Данные по G2, "),
    (r"^Используйте Crunchbase, LinkedIn, BuiltWith, Wappalyzer, рекламные библиотеки, GitHub, вакансии и сайты конкурентов$", "Данные из деловой прессы, каталогов компаний и сайтов конкурентов"),
    (r" Веса можно корректировать в зависимости от рынка\.", " Веса заданы с учетом особенностей рынка."),
    (r"^Проверить стоимость клика и юнит-экономику\.$", "Стоимость клика и юнит-экономика оценены в ЛР2-3."),
    (r"^Запускать пилоты, собирать отзывы, демонстрировать кейсы\.$", "Репутационный барьер снижают пилоты, отзывы и кейсы."),
    (r"^Уточнить ценность: результат, скорость, экспертиза, сопровождение\.$", "Ценность строится на результате, скорости, экспертизе и сопровождении."),
    (r"^Фиксируйте дату обращения, ссылку, что именно взято и ограничения интерпретации$", "Для каждого источника указаны ссылка, взятые данные и ограничения интерпретации"),
    (r"^ШАБЛОН АНАЛИЗА КОНКУРЕНТНЫХ СИЛ В ОТРАСЛИ по ПОРТЕРУ$", "АНАЛИЗ КОНКУРЕНТНЫХ СИЛ В ОТРАСЛИ ПО ПОРТЕРУ"),
    (r"^Читать теоретическую основу модели пяти сил конкуренции Майкла Портера$", "Теоретическая основа - модель пяти сил конкуренции Майкла Портера"),
    (r"^Правила заполнения:$", "Правила оценки"),
    (r"^Методика работы с Excel-шаблоном классификации компаний по уровням конкуренции$", "Классификация компаний по уровням конкуренции"),
    (r"^Excel-шаблон анализа конкурентов по сайтам$", "Анализ конкурентов по сайтам"),
    (r"^Где заполнять$", "Лист"),
    (r"^Заполнить обоснование и ключевые ссылки\.$", "Обоснование и ключевые ссылки приведены в листе 04_Оценка."),
    (r"^Excel-шаблон к методике формулирования ценностного предложения и цифрового товара$", "Формулирование ценностного предложения и цифрового товара"),
    (r"^Методика сборки бизнес-модели по Канвас: рабочий Excel-шаблон$", "Сборка бизнес-модели по Канвас"),
    (r"^Шаблон Канвас: вопросы, источники ответов и рабочие формулировки$", "Канвас: вопросы, источники ответов и рабочие формулировки"),
    (r"^Excel-шаблон к методике формулирования бизнес-логики и пользовательских сценариев$", "Формулирование бизнес-логики и пользовательских сценариев"),
    (r"^Шаблон перехода от бизнес-логики и пользовательских сценариев к страницам и экранам$", "Переход от бизнес-логики и пользовательских сценариев к страницам и экранам"),
    (r"^Шаблон деления разработки сайта или мобильного приложения на первый и последующие релизы$", "Деление разработки сайта или мобильного приложения на первый и последующие релизы"),
    (r"^Что заполняется$", "Содержание"),
    (r"^Что заполнить$", "Содержание"),
    (r"^Новая возможность$", "Новая возможность"),
    (r"^не показано$", "-"),
    (r"^Не показано$", "-"),
    (r"^нет данных$", "-"),
    (r"^Нет данных$", "-"),
]
CLEAR_TEXT = {"Свободная строка шаблона.", "Свободная строка шаблона", "Демонстрационная строка"}


def fix_text(s):
    for a, b in RULES:
        s = re.sub(a, b, s)
    s = currency_eur.PAT.sub(currency_eur._conv, s)
    if len(s) > 28 and "http" not in s and not re.search(r"\d:\d", s):
        s = re.sub(r"(?<=\S): (?=\S)", " - ", s)
    s = re.sub(r"^(.{3,60}):$", r"\1", s)
    s = re.sub(r"\bEUR\b", "у.е.", s)
    s = s.replace("—", "-").replace("ё", "е").replace("Ё", "Е")
    return s


def fix_formats(ws):
    pass


IDLIKE = re.compile(r"^(Резерв\w*\s?\d*|[A-ZА-Я]{1,3}-?\d{1,4}|\d{1,3})$")


def meaningful(v):
    """Значимое содержимое ячейки; пустые, нули, прочерки, ошибки формул, номера и «Резерв» значимыми не считаются."""
    if v is None:
        return False
    if isinstance(v, str):
        s = v.strip()
        return s not in ("", "-") and not IDLIKE.match(s)
    if isinstance(v, bool):
        return True
    if isinstance(v, int) and v < -2146826000:
        return False
    return v != 0


def style_book(path):
    pythoncom.CoInitialize()
    xl = wc.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    xl.EnableEvents = False
    report = []
    try:
        wb = xl.Workbooks.Open(str(path))
        xl.Calculation = -4135
        for ws in wb.Worksheets:
            style_sheet(ws, report)
        xl.Calculation = -4105
        wb.BuiltinDocumentProperties("Author").Value = AUTHOR
        wb.BuiltinDocumentProperties("Last Author").Value = AUTHOR
        xl.Calculate()
        errs = []
        for ws in wb.Worksheets:
            v = ws.UsedRange.Value2
            if v is None:
                continue
            if not isinstance(v, tuple):
                v = ((v,),)
            for ri, row in enumerate(v, 1):
                for ci, x in enumerate(row, 1):
                    if isinstance(x, int) and x < -2146826000:
                        errs.append(f"{ws.Name}!R{ri}C{ci}")
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    print("OK", path.name, "| ошибки:", errs[:5], "|", "; ".join(report))


def style_sheet(ws, report):
    # 1. тексты
    ur = ws.UsedRange
    rows, cols = ur.Rows.Count, ur.Columns.Count
    r0, c0 = ur.Row, ur.Column
    f = ur.Formula
    if not isinstance(f, tuple):
        f = ((f,),)
    for i in range(rows):
        for j in range(cols):
            v = f[i][j]
            if isinstance(v, str) and v in CLEAR_TEXT:
                ws.Cells(r0 + i, c0 + j).ClearContents()
    # 2. пустые строки
    ur = ws.UsedRange
    protected = set()
    for co in ws.ChartObjects():
        protected.update(range(co.TopLeftCell.Row, co.BottomRightCell.Row + 1))
    vals = ur.Value2
    if not isinstance(vals, tuple):
        vals = ((vals,),)
    r0 = ur.Row
    blank = []
    for i, row in enumerate(vals):
        r = r0 + i
        if r in protected:
            continue
        row = list(row)
        if row and isinstance(row[0], (int, float)) and not isinstance(row[0], bool) and float(row[0]).is_integer() and 0 < row[0] < 1000:
            row[0] = None  # порядковый номер строки не считается данными
        nm = [k for k, x in enumerate(row) if meaningful(x)]
        if not nm:
            blank.append(r)
        elif len(nm) == 1 and nm[0] != 0 and isinstance(row[nm[0]], str) and len(row[nm[0]]) < 60:
            blank.append(r)  # свободная строка таблицы с одиночной пометкой в боковом столбце
    # строки ниже используемого диапазона с форматами
    runs = []
    for r in sorted(blank):
        if runs and runs[-1][1] == r - 1:
            runs[-1][1] = r
        else:
            runs.append([r, r])
    for a, b in reversed(runs):
        ws.Range(ws.Rows(a), ws.Rows(b)).Delete()
    report.append(f"{ws.Name}: удалено строк {len(blank)}")
    # 2б. тексты, статусы и валюты в оставшихся ячейках
    ur = ws.UsedRange
    f = ur.Formula
    if not isinstance(f, tuple):
        f = ((f,),)
    r0, c0 = ur.Row, ur.Column
    for i in range(ur.Rows.Count):
        for j in range(ur.Columns.Count):
            v = f[i][j]
            if not (isinstance(v, str) and v):
                continue
            if v.startswith("="):
                nv = status_map.map_formula(v)
            else:
                nv = status_map.map_text(fix_text(v))
            if nv != v:
                ws.Cells(r0 + i, c0 + j).Value = nv if not nv.startswith("=") else None
                if nv.startswith("="):
                    ws.Cells(r0 + i, c0 + j).Formula = nv
    # 3. стиль
    ur = ws.UsedRange
    rows, cols = ur.Rows.Count, ur.Columns.Count
    r0, c0 = ur.Row, ur.Column
    # форматы с чужими валютами
    nfs = set()
    for i in range(rows):
        for j in range(cols):
            c = ws.Cells(r0 + i, c0 + j)
            nf = c.NumberFormat
            if isinstance(nf, str) and re.search(r"BYN|₽|руб|USD|\$|EUR|€", nf):
                c.NumberFormat = re.sub(r'"BYN"|\\₽|"₽"|"руб\.?"|"USD"|"EUR"|\\€|€', '"у.е."', nf)
                nfs.add(nf)
    ur.Font.Name = "Calibri"
    ur.Font.Size = 11
    ur.Font.Bold = False
    ur.Font.Italic = False
    ur.WrapText = True
    ur.VerticalAlignment = -4108
    head_rows = {}
    for i in range(rows):
        r = r0 + i
        for j in range(cols):
            c = ws.Cells(r, c0 + j)
            val = c.Value2
            fill_rgb = rgb_of(c.Interior.Color) if c.Interior.ColorIndex != -4142 else None
            role = None
            merged = bool(c.MergeCells)
            role = classify(fill_rgb, r, merged)
            if role == "title":
                c.Interior.Color = bgr(TITLE)
                c.Font.Color = bgr("FFFFFF")
                c.Font.Bold = True
                c.Font.Size = 14 if r == 1 else 12
                c.HorizontalAlignment = -4108 if r == 1 else -4131
            elif role == "head":
                c.Interior.Color = bgr(HEAD)
                c.Font.Color = bgr("FFFFFF")
                c.Font.Bold = True
                c.HorizontalAlignment = -4108
                head_rows.setdefault(r, []).append(c0 + j)
            elif role == "sub":
                c.Interior.Color = bgr(SUB)
                c.Font.Color = bgr("1F3864")
                c.Font.Italic = True
                c.HorizontalAlignment = -4131
            elif role == "label":
                c.Interior.Color = bgr(SUB)
                c.Font.Color = bgr("000000")
            elif role == "input":
                c.Interior.Color = bgr(INPUT)
                c.Font.Color = bgr("000000")
            elif role == "calc":
                c.Interior.Color = bgr(CALC)
                c.Font.Color = bgr("000000")
            else:
                if fill_rgb is not None:
                    c.Interior.ColorIndex = -4142
                c.Font.Color = bgr("000000")
            if role in (None, "input", "calc", "label"):
                if isinstance(val, (int, float)):
                    c.HorizontalAlignment = -4108
                elif isinstance(val, str) and len(val) > 28:
                    c.HorizontalAlignment = -4131
                elif val is not None:
                    c.HorizontalAlignment = -4108
    # 4. прочерки в таблицах
    filled = 0
    hdr = sorted(head_rows)
    blocks = []
    for k, hr in enumerate(hdr):
        if hr == 1:
            continue
        cs = sorted(head_rows[hr])
        groups, cur = [], [cs[0]]
        for cc in cs[1:]:
            if cc == cur[-1] + 1:
                cur.append(cc)
            else:
                groups.append(cur)
                cur = [cc]
        groups.append(cur)
        nxt = hdr[k + 1] if k + 1 < len(hdr) else r0 + rows
        for g in groups:
            blocks.append((hr, g[0], g[-1], nxt))
    for hr, c_lo, c_hi, nxt in blocks:
        r = hr + 1
        while r < nxt:
            vals = [ws.Cells(r, cc).Value2 for cc in range(c_lo, c_hi + 1)]
            if all(v is None or v == "" for v in vals):
                break
            for cc in range(c_lo, c_hi + 1):
                c = ws.Cells(r, cc)
                v = c.Value2
                if v is None or v == "":
                    if c.MergeCells:
                        continue
                    fm = c.Formula
                    if isinstance(fm, str) and fm.startswith("="):
                        c.Formula = f'=IF(LEN({fm[1:]})=0,"-",{fm[1:]})'
                    else:
                        c.Value = "-"
                    c.HorizontalAlignment = -4108
                    filled += 1
                elif cc == c_lo and isinstance(v, str):
                    c.HorizontalAlignment = -4131
            r += 1
    report.append(f"прочерков {filled}")
    # рамки
    for i in range(rows):
        for j in range(cols):
            c = ws.Cells(r0 + i, c0 + j)
            if c.Value2 is not None and c.Value2 != "" or c.Interior.ColorIndex != -4142:
                for b in (7, 8, 9, 10):
                    br = c.Borders(b)
                    br.LineStyle = 1
                    br.Weight = 2
                    br.Color = bgr(BORDER)
    # синие гистограммы вместо зеленых
    try:
        fcs = ws.Cells.FormatConditions
        for k in range(1, fcs.Count + 1):
            fc = fcs.Item(k)
            if fc.Type == 4:
                fc.BarColor.Color = bgr("5B9BD5")
    except Exception:
        pass
    # 5. ширины и высоты
    maxw = 52
    for j in range(cols):
        cc = c0 + j
        longest = 0
        for i in range(rows):
            c = ws.Cells(r0 + i, cc)
            if c.MergeCells or (r0 + i) == 1:
                continue
            t = c.Text
            if t:
                longest = max(longest, max(len(x) for x in t.split("\n")))
        w = min(max(longest * 1.05 + 3, 10), maxw)
        ws.Columns(cc).ColumnWidth = w
    for i in range(rows):
        r = r0 + i
        if ws.Cells(r, c0).MergeCells and ws.Cells(r, c0).MergeArea.Columns.Count > 1:
            txt = str(ws.Cells(r, c0).Value2 or "")
            tot = sum(ws.Columns(cc).ColumnWidth for cc in range(c0, c0 + ws.Cells(r, c0).MergeArea.Columns.Count))
            lines = max(1, -(-len(txt) // max(int(tot * 1.0), 10)))
            ws.Rows(r).RowHeight = max(20 if r == 1 else 18, 16 * lines + 4)
        else:
            ws.Rows(r).AutoFit()
    # 6. диаграммы
    if ws.Name.endswith("Дашборд") and ws.ChartObjects().Count == 0:
        add_dashboard_charts(ws)
    restyle_charts(ws, r0 + rows)


def add_dashboard_charts(ws):
    """Диаграммы для дашборда без графиков: сценарии по методам и доли конкурентов."""
    ur = ws.UsedRange
    f1 = ur.Find("Метод", LookAt=1)
    if f1 is not None:
        r, c = f1.Row, f1.Column
        co = ws.ChartObjects().Add(10, 10, 420, 260)
        ch = co.Chart
        ch.ChartType = 51
        ch.SetSourceData(ws.Range(ws.Cells(r, c), ws.Cells(r + 3, c + 3)), 2)
        ch.HasTitle = True
        ch.ChartTitle.Text = "Оценка объема рынка по методам, EUR в год"
        ch.HasLegend = True
        ch.Legend.Position = -4107
    f2 = ur.Find("Конкурент", LookAt=1)
    if f2 is not None:
        r, c = f2.Row, f2.Column
        n = 0
        while ws.Cells(r + 1 + n, c).Value2 not in (None, ""):
            n += 1
        co = ws.ChartObjects().Add(10, 10, 420, 260)
        ch = co.Chart
        ch.ChartType = 57
        ch.SetSourceData(ws.Range(ws.Cells(r, c), ws.Cells(r + n, c + 1)), 2)
        ch.HasTitle = True
        ch.ChartTitle.Text = "Доли релевантного трафика конкурентов"
        ch.HasLegend = False


def fix_chart_text(ch):
    try:
        if ch.HasTitle:
            ch.ChartTitle.Text = re.sub(r"EUR", "у.е.", ch.ChartTitle.Text)
        for t in (1, 2):
            try:
                ax = ch.Axes(t)
                if ax.HasTitle:
                    ax.AxisTitle.Text = re.sub(r"EUR", "у.е.", ax.AxisTitle.Text)
            except Exception:
                pass
    except Exception:
        pass


def restyle_charts(ws, last_row):
    cos = list(ws.ChartObjects())
    palette = [bgr(TITLE), bgr(HEAD), bgr("9DC3E6"), bgr("5B9BD5"), bgr("BDD7EE")]
    top = ws.Rows(last_row + 2).Top
    left = ws.Columns(1).Left
    maxw = sum(ws.Columns(c).Width for c in range(1, 9))
    x = left
    for co in cos:
        ch = co.Chart
        fix_chart_text(ch)
        for k, s in enumerate(ch.SeriesCollection()):
            try:
                s.Format.Fill.ForeColor.RGB = palette[k % len(palette)]
            except Exception:
                pass
        ch.ChartArea.Format.Line.Visible = False
        co.Width = min(max(co.Width, 360), 460)
        co.Height = 260
        if x + co.Width > left + max(maxw, 760):
            x = left
            top += 270
        co.Left, co.Top = x, top
        x += co.Width + 12
    return


if __name__ == "__main__":
    only = sys.argv[1:]
    for b in [Path(a) for a in only if Path(a).exists()]:
        style_book(b)
