# -*- coding: utf-8 -*-
"""Проверка всех лабораторных работ: наличие отчёта и docx, нумерация таблиц и рисунков, ссылки на рисунки, ИИ-метки, ошибки в книгах Excel, автор docx."""
import glob
import os
import re
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LABS = ["ЛР1_Спрос_и_границы_рынка", "ЛР2-3_Объем_рынка", "ЛР4_Пять_сил_Портера", "ЛР5_Анализ_конкурентов", "ЛР6_ЦП_Канвас_сценарии_релизы", "ЛР7_Итоговый_отчет"]
MARK = re.compile("|".join(["cla"+"ude", "anthro"+"pic", "generated "+"with", "co-"+"authored", "chat"+"gpt", "open"+"ai"]), re.I)
rep = []
for lab in LABS:
    d = os.path.join(ROOT, lab)
    row = [lab]
    md = os.path.join(d, "ОТЧЕТ.md")
    docx = glob.glob(os.path.join(d, "ОТЧЕТ_*.docx"))
    if not os.path.exists(md):
        rep.append(row + ["НЕТ ОТЧЕТ.md"])
        continue
    t = open(md, encoding="utf-8").read()
    tabs = [int(x) for x in re.findall(r"^Таблица (\d+) – ", t, re.M)]
    figs = [int(x) for x in re.findall(r"^Рисунок (\d+) – ", t, re.M)]
    imgs = re.findall(r"!\[\]\(([^)]+)\)", t)
    missing = [i for i in imgs if not os.path.exists(os.path.join(d, i))]
    row.append(f"таблиц {len(tabs)}" + ("" if tabs == list(range(1, len(tabs) + 1)) else " (НУМЕРАЦИЯ!)"))
    row.append(f"рисунков {len(figs)}" + ("" if figs == list(range(1, len(figs) + 1)) else " (НУМЕРАЦИЯ!)") + (f", нет файлов: {missing}" if missing else ""))
    row.append("🔲 %d" % t.count("🔲"))
    if docx:
        # в CI даты файлов после checkout одинаковы, сравнение по времени не применяется
        newer = bool(os.environ.get("CI")) or os.path.getmtime(docx[0]) >= os.path.getmtime(md)
        row.append("docx актуален" if newer else "docx УСТАРЕЛ")
    else:
        row.append("docx ОТСУТСТВУЕТ")
    marks = []
    for f in glob.glob(os.path.join(d, "**", "*"), recursive=True):
        if os.path.isfile(f) and f.lower().endswith((".md", ".py", ".ps1", ".bas", ".json", ".csv")) and "Материалы_собранные" not in f:
            try:
                if MARK.search(open(f, encoding="utf-8", errors="ignore").read()):
                    marks.append(os.path.relpath(f, d))
            except OSError:
                pass
    row.append("ИИ-метки: " + (", ".join(marks[:3]) if marks else "нет"))
    errs = []
    for x in glob.glob(os.path.join(d, "Excel", "*.xlsx")):
        try:
            z = zipfile.ZipFile(x)
            n = 0
            for name in z.namelist():
                if name.startswith("xl/worksheets/sheet"):
                    n += len(re.findall(rb't="e"', z.read(name)))
            if n:
                errs.append(f"{os.path.basename(x)}: {n}")
            core = z.read("docProps/core.xml").decode("utf-8", "ignore")
            if "Земляник" not in core:
                errs.append(os.path.basename(x) + ": автор не Земляник")
        except Exception as e:
            errs.append(os.path.basename(x) + ": " + str(e)[:40])
    row.append("ошибки Excel: " + ("; ".join(errs) if errs else "нет"))
    if docx:
        try:
            from docx import Document
            a = Document(docx[0]).core_properties.author
            row.append("автор docx: " + ("OK" if "Земляник" in (a or "") else f"НЕ ТОТ ({a})"))
        except Exception as e:
            row.append("docx: " + str(e)[:40])
    rep.append(row)
for r in rep:
    print(" | ".join(r))
