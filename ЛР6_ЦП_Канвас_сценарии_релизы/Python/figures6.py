"""Рисунки ЛР6: Канвас, приоритеты требований, дорожная карта релизов, схемы экранов."""
import json
import sys
from datetime import date
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle

HERE = Path(__file__).resolve().parent
OUT = HERE.parent / "Рисунки"
OUT.mkdir(exist_ok=True)
from data6a import *
from data6b import *

X = json.loads((HERE / "results" / "excel_results.json").read_text(encoding="utf-8"))
plt.rcParams.update({"font.size": 9, "axes.spines.top": False, "axes.spines.right": False})


def wrap(t, n):
    words, lines, cur = t.split(), [], ""
    for w in words:
        if len(cur) + len(w) + 1 > n:
            lines.append(cur)
            cur = w
        else:
            cur = (cur + " " + w).strip()
    lines.append(cur)
    return "\n".join(lines)


def f1():
    fig, ax = plt.subplots(figsize=(8.2, 5.2))
    ax.set_xlim(0, 10); ax.set_ylim(0, 6.4); ax.axis("off")
    d = {c[0]: c for c in CANVAS}
    boxes = [("Ключевые партнеры", 0, 2.4, 2, 3.8), ("Ключевые виды деятельности", 2, 4.3, 2, 1.9), ("Ключевые ресурсы", 2, 2.4, 2, 1.9), ("Ценностное предложение", 4, 2.4, 2, 3.8),
             ("Отношения с клиентами", 6, 4.3, 2, 1.9), ("Каналы", 6, 2.4, 2, 1.9), ("Клиентские сегменты", 8, 2.4, 2, 3.8), ("Структура затрат", 0, 0, 5, 2.3), ("Потоки доходов", 5, 0, 5, 2.3)]
    for name, x, y, w, h in boxes:
        ax.add_patch(Rectangle((x + 0.03, y + 0.03), w - 0.06, h - 0.06, fill=True, fc="#eef3fa", ec="#1f4e79", lw=1))
        tl = wrap(name, int(w * 9))
        ax.text(x + 0.1, y + h - 0.12, tl, fontsize=7, fontweight="bold", va="top", color="#1f4e79")
        ax.text(x + 0.1, y + h - 0.12 - 0.42 * (len(tl.splitlines())) - 0.1, wrap(d[name][1], int(w * 9.5)), fontsize=5.8, va="top")
    fig.tight_layout(); fig.savefig(OUT / "01_канвас.png", dpi=200); plt.close(fig)


def f2():
    rows = X["logic"]["req"]
    rows = sorted(rows, key=lambda r: r[3])
    fig, ax = plt.subplots(figsize=(8.2, 6.0))
    cols = ["#d62728" if r[5] == "Да" else "#7f7f7f" for r in rows]
    ax.barh([wrap(f"{r[0]} {r[1]}", 40) for r in rows], [r[3] for r in rows], color=cols)
    ax.axvline(10, color="k", ls="--", lw=1)
    ax.text(10.3, 0.2, "порог MVP (высокий приоритет) = 10", fontsize=7.5)
    ax.set_xlabel("индекс приоритета требования; красные - требования MVP (8 из 18)")
    plt.setp(ax.get_yticklabels(), fontsize=7.5)
    from matplotlib.ticker import FuncFormatter
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}".replace(".", ",")))
    fig.tight_layout(); fig.savefig(OUT / "02_приоритет_требований.png", dpi=200); plt.close(fig)


def f3():
    fig, ax = plt.subplots(figsize=(8.2, 3.9))
    cols = ["#1f77b4", "#2ca02c", "#ff7f0e", "#d62728"]
    for i, r in enumerate(ROAD):
        y, m, d = map(int, r[9].split("-")); a = date(y, m, d)
        y, m, d = map(int, r[10].split("-")); b = date(y, m, d)
        ax.barh(len(ROAD) - i, (b - a).days, left=mdates.date2num(a), color=cols[i], height=0.5)
    for lab, dt in (("MVP 31.03.2027", date(2027, 3, 31)), ("бета с оплатой 01.06.2027", date(2027, 6, 1)), ("бета v1 июль 2027", date(2027, 7, 31)), ("релиз 30.09.2027", date(2027, 9, 30))):
        ax.axvline(mdates.date2num(dt), color="k", lw=0.7, ls=":")
        ax.text(mdates.date2num(dt), 6.2, lab, rotation=90, fontsize=7, va="top", ha="right")
    ax.set_yticks([len(ROAD) - i for i in range(len(ROAD))])
    ax.set_yticklabels([wrap(f"{r[0]}: {r[1]}", 30) for r in ROAD], fontsize=7.5)
    ax.xaxis_date(); ax.xaxis.set_major_formatter(mdates.DateFormatter("%m.%Y"))
    ax.set_ylim(0.3, 6.3)
    fig.tight_layout(); fig.savefig(OUT / "03_дорожная_карта.png", dpi=200); plt.close(fig)


def screen(ax, title, blocks, accent="#1f4e79"):
    ax.set_xlim(0, 10); ax.set_ylim(0, 16); ax.axis("off")
    ax.add_patch(Rectangle((0.1, 0.1), 9.8, 15.8, fill=True, fc="white", ec="#444", lw=1.2))
    ax.add_patch(Rectangle((0.1, 14.9), 9.8, 1.0, fill=True, fc=accent, ec="#444"))
    ax.text(0.4, 15.4, title, color="white", fontsize=8, va="center", fontweight="bold")
    y = 14.5
    for h, t, kind in blocks:
        y -= h
        fc = {"btn": "#2ca02c", "box": "#f2f2f2", "warn": "#fff2cc", "tbl": "#eaf1fb"}[kind]
        ax.add_patch(FancyBboxPatch((0.5, y), 9.0, h - 0.2, boxstyle="round,pad=0.02,rounding_size=0.15", fc=fc, ec="#888", lw=0.8))
        ax.text(5, y + (h - 0.2) / 2, wrap(t, 38), ha="center", va="center", fontsize=7.2, color="white" if kind == "btn" else "black")
        y -= 0.1


def f4():
    fig, axes = plt.subplots(2, 3, figsize=(8.2, 9.4))
    axes = axes.flatten()
    specs = [
        ("U-001 Страница продукта", [(2.2, "Первый экран: ошибки в Excel-моделях находятся до передачи файла", "box"), (1.2, "Проверить файл бесплатно", "btn"), (2.6, "Примеры найденных ошибок и доказательства", "tbl"), (2.0, "Чем отличается от AI-помощников", "box"), (1.8, "Тарифы Free / 9 / 19 / 29 EUR", "box"), (1.2, "Безопасность и GDPR", "warn")]),
        ("U-002 Загрузка файла", [(1.6, "Перетащите файл xlsx", "box"), (1.4, "Лимиты Free: размер, число проверок", "warn"), (1.2, "Запустить проверку", "btn"), (2.2, "Индикатор: 32 %", "tbl"), (1.6, "Ошибка: неверный формат", "warn")]),
        ("U-003 Результат", [(1.6, "Найдено: 7 критичных, 12 средних", "box"), (4.6, "Таблица находок: ячейка, тип, критичность", "tbl"), (1.4, "Открыть карточку", "btn"), (1.6, "Полный отчет - 9 EUR", "warn")]),
        ("U-004 Карточка находки", [(1.4, "Ячейка D12: сумма не совпадает", "box"), (2.4, "Доказательство: 1 250 против 1 205", "tbl"), (1.8, "Исправление", "box"), (1.2, "Принять", "btn"), (1.2, "Отклонить", "box")]),
        ("U-007 Тарифы", [(2.2, "Free: проверка с лимитом", "box"), (2.2, "Разовый аудит 9 EUR", "box"), (2.2, "Pro 19 EUR в месяц", "tbl"), (2.2, "Team 29 EUR за место", "box"), (1.2, "Выбрать тариф", "btn")]),
        ("U-005 Отчет аудита", [(1.4, "Модель: budget_2027.xlsx", "box"), (2.4, "Сводка по критичности", "tbl"), (4.2, "Перечень находок со ссылками на ячейки", "tbl"), (1.2, "Скачать PDF", "btn")]),
    ]
    for ax, (t, b) in zip(axes, specs):
        screen(ax, t, b)
    fig.tight_layout(pad=0.4); fig.savefig(OUT / "04_схемы_экранов.png", dpi=200); plt.close(fig)


if __name__ == "__main__":
    for f in (f1, f2, f3, f4):
        f()
    print(sorted(p.name for p in OUT.glob("*.png")))
