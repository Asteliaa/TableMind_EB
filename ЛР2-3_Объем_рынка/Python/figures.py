"""Рисунки ЛР2-3 (PNG, 200 dpi) по data results/analysis.json, excel_results.json и снимку Similarweb."""
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

from inputs import *

HERE = Path(__file__).resolve().parent
OUT = HERE.parent / "Рисунки"
OUT.mkdir(exist_ok=True)
A = json.loads((HERE / "results" / "analysis.json").read_text(encoding="utf-8"))
X = json.loads((HERE / "results" / "excel_results.json").read_text(encoding="utf-8"))
sw = load_sw()
plt.rcParams.update({"font.size": 9, "axes.grid": True, "grid.alpha": 0.3, "axes.spines.top": False, "axes.spines.right": False})
COL = {"прямой": "#d62728", "частичный": "#ff7f0e", "заменитель": "#1f77b4"}
C1, C2, C3, C4 = "#1f77b4", "#d62728", "#2ca02c", "#7f7f7f"


def fmt(v):
    return f"{v:,.0f}".replace(",", " ")


def fig1():
    rows = sorted(sw, key=lambda r: r["visits_month"])
    fig, ax = plt.subplots(figsize=(8.2, 4.8))
    ax.barh([r["name"] for r in rows], [r["visits_month"] for r in rows], color=[COL[r["type"]] for r in rows])
    ax.set_xscale("log")
    ax.set_xlabel("среднее число визитов в месяц, июнь - август 2026, весь мир (логарифмическая шкала)")
    for i, r in enumerate(rows):
        ax.text(r["visits_month"] * 1.08, i, fmt(r["visits_month"]), va="center", fontsize=8)
    ax.set_xlim(1000, 3e6)
    for t, c in COL.items():
        ax.add_patch(plt.Rectangle((0, 0), 0, 0, color=c, label=t))
    ax.legend(title="тип игрока", loc="lower right", fontsize=8)
    fig.tight_layout(); fig.savefig(OUT / "01_визиты_сайтов.png", dpi=200); plt.close(fig)


def fig2():
    rows = sorted(sw, key=lambda r: r["visits_month"])
    keys = [("direct", "Прямые заходы", "#4c78a8"), ("org_search", "Органический поиск", "#f58518"), ("referrals", "Реферальные", "#54a24b"),
            ("social", "Социальные сети", "#e45756"), ("paid_search", "Платный поиск", "#72b7b2"), ("other", "Прочие каналы", "#bab0ac")]
    fig, ax = plt.subplots(figsize=(8.2, 5.0))
    left = [0] * len(rows)
    for k, lab, col in keys:
        vals = []
        for r in rows:
            ch = r["ch"]
            if k == "social":
                v = ch["org_social"] + ch["paid_social"]
            elif k == "other":
                v = ch["display"] + ch["gen_ai"] + ch["email"] + ch["affiliates"]
            else:
                v = ch[k]
            vals.append(v * 100)
        ax.barh([r["name"] for r in rows], vals, left=left, color=col, label=lab)
        left = [a + b for a, b in zip(left, vals)]
    ax.set_xlabel("доля визитов, %"); ax.set_xlim(0, 100)
    ax.legend(ncol=3, fontsize=7.5, loc="upper center", bbox_to_anchor=(0.5, -0.12))
    fig.tight_layout(); fig.savefig(OUT / "02_каналы_трафика.png", dpi=200); plt.close(fig)


def fig3():
    tam, sam, som = A["tam_model"], A["final_potential"]["base"]["sam"], A["final_potential"]["base"]["som"]
    labs = ["TAM: теоретический потолок\n(все клиенты x чек)", "SAM: адресуемый рынок\n(медиана четырех методов)", "SOM: достижимая часть\n(год, базовый сценарий)"]
    vals = [tam, sam, som]
    fig, ax = plt.subplots(figsize=(8.2, 3.8))
    ax.barh(labs[::-1], vals[::-1], color=[C3, C1, C4])
    ax.set_xscale("log"); ax.set_xlim(1e4, 3e10)
    for i, v in enumerate(vals[::-1]):
        ax.text(v * 1.2, i, fmt(v) + " EUR", va="center", fontsize=9)
    ax.set_xlabel("EUR в год (логарифмическая шкала)")
    fig.tight_layout(); fig.savefig(OUT / "03_tam_sam_som.png", dpi=200); plt.close(fig)


def fig4():
    items = []
    for k, m in A["methods"].items():
        items.append((k + ": SAM", m["sam"]))
    items.append(("Сценарная оценка Similarweb (книга 5): рынок", A["traffic_methods"]["Сценарная оценка Similarweb (книга 5): рынок/год"]))
    items.append(("Объем по Similarweb (книга 6): рынок", A["traffic_methods"]["Объем по Similarweb (книга 6): рынок/год"]))
    items.append(("Цифровая воронка (книга 4): оборот TableMind", A["traffic_methods"]["Цифровая воронка (книга 4): оборот новых платящих за год"]))
    items = items[::-1]
    fig, ax = plt.subplots(figsize=(8.2, 4.6))
    for i, (n, v) in enumerate(items):
        ax.plot([v[0], v[2]], [i, i], color=C4, lw=3, solid_capstyle="round")
        ax.scatter([v[0], v[2]], [i, i], color=C1, s=25, zorder=3)
        ax.scatter([v[1]], [i], color=C2, s=45, zorder=4)
    ax.set_yticks(range(len(items))); ax.set_yticklabels([n for n, _ in items], fontsize=8)
    ax.set_xscale("log"); ax.set_xlabel("EUR в год (логарифмическая шкала); красная точка - базовый сценарий")
    fig.tight_layout(); fig.savefig(OUT / "04_диапазоны_методов.png", dpi=200); plt.close(fig)


def fig5():
    t = A["tornado"]
    rows = sorted(t["rows"], key=lambda r: abs(r[2] - r[1]))
    fig, ax = plt.subplots(figsize=(8.2, 3.8))
    base = t["base"]
    for i, (n, lo, hi) in enumerate(rows):
        ax.barh(i, lo - base, left=base, color=C2); ax.barh(i, hi - base, left=base, color=C3)
        ax.text(lo - 0.8, i, f"{lo:.0f}", ha="right", va="center", fontsize=8); ax.text(hi + 0.8, i, f"{hi:.0f}", ha="left", va="center", fontsize=8)
    ax.axvline(base, color="k", lw=1); ax.axvline(GOAL_PAY, color=C1, lw=1.5, ls="--")
    ax.text(GOAL_PAY + 0.5, len(rows) - 0.4, "цель: 50", color=C1, fontsize=8)
    ax.set_yticks(range(len(rows))); ax.set_yticklabels([r[0] for r in rows], fontsize=8)
    ax.set_xlabel(f"платящих за 4 месяца платной беты (базовый сценарий: {base:.0f})"); ax.set_xlim(0, 62)
    fig.tight_layout(); fig.savefig(OUT / "05_чувствительность.png", dpi=200); plt.close(fig)


def fig6():
    g = A["goal"]
    fig, axes = plt.subplots(1, 2, figsize=(8.2, 3.6))
    ax = axes[0]
    m = list(range(0, 7))
    for sc, col in zip(SC, (C4, C1, C3)):
        ax.plot(m, [g[sc]["reg_month"] * i for i in m], color=col, label=SC_RU[sc])
    ax.axhline(GOAL_REG, color=C2, ls="--"); ax.text(0.1, GOAL_REG * 1.05, "цель: 300 регистраций", color=C2, fontsize=8)
    ax.set_ylim(0, 3 * GOAL_REG); ax.set_xlabel("месяцев после MVP (31.03.2027)"); ax.set_ylabel("регистраций нарастающим итогом")
    ax.legend(fontsize=7.5, loc="upper left", bbox_to_anchor=(0.0, 0.88))
    ax = axes[1]
    m = list(range(0, 5))
    for sc, col in zip(SC, (C4, C1, C3)):
        ax.plot(m, [g[sc]["payers_month"] * i for i in m], color=col, label=SC_RU[sc])
    ax.axhline(GOAL_PAY, color=C2, ls="--"); ax.text(0.1, GOAL_PAY * 1.1, "цель: 50 платящих", color=C2, fontsize=8)
    ax.set_ylim(0, 3 * GOAL_PAY); ax.set_xlabel("месяцев платной беты (с 01.06.2027)"); ax.set_ylabel("платящих нарастающим итогом")
    fig.tight_layout(); fig.savefig(OUT / "06_цели_проекта.png", dpi=200); plt.close(fig)


if __name__ == "__main__":
    for f in (fig1, fig2, fig3, fig4, fig5, fig6):
        f()
    print("рисунки:", sorted(p.name for p in OUT.glob("*.png")))
