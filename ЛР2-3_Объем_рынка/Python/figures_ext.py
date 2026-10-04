"""Дополнительные рисунки ЛР2-3: качество трафика и география сайтов (снимок Similarweb 03.10.2026)."""
import csv
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

from inputs import *

HERE = Path(__file__).resolve().parent
OUT = HERE.parent / "Рисунки"
sw = load_sw()
raw = {r["domain"]: r for r in csv.DictReader(open(HERE.parent / "Материалы_собранные" / "similarweb_снимок_2026-10-03.csv", encoding="utf-8-sig"), delimiter=";")}
plt.rcParams.update({"font.size": 9, "axes.grid": True, "grid.alpha": 0.3, "axes.spines.top": False, "axes.spines.right": False})
COL = {"прямой": "#d62728", "частичный": "#ff7f0e", "заменитель": "#1f77b4"}


def secs(t):
    h, m, s = [int(x) for x in t.split(":")]
    return h * 3600 + m * 60 + s


def fig_quality():
    fig, ax = plt.subplots(figsize=(8.2, 4.8))
    for r in sw:
        d = raw[r["domain"]]
        ax.scatter(r["visits_month"], secs(d["duration"]), s=float(d["pages_visit"]) * 70, color=COL[r["type"]], alpha=0.75, edgecolor="white")
        ax.annotate(r["name"], (r["visits_month"], secs(d["duration"])), textcoords="offset points", xytext=(6, 5), fontsize=7.5)
    ax.set_xscale("log")
    ax.set_xlabel("среднее число визитов в месяц (логарифмическая шкала)")
    ax.set_ylabel("длительность визита, секунд")
    for t, c in COL.items():
        ax.scatter([], [], color=c, label=t)
    ax.legend(title="тип игрока (размер круга - страниц за визит)", fontsize=8, loc="upper left")
    fig.tight_layout(); fig.savefig(OUT / "07_качество_трафика.png", dpi=200); plt.close(fig)


def fig_geo():
    rows = sorted(sw, key=lambda r: r["geo_share"])
    fig, ax = plt.subplots(figsize=(8.2, 4.8))
    names = [r["name"] for r in rows]
    tgt = [r["geo_share"] * 100 for r in rows]
    top5 = [sum(v for _, v in r["top"]) * 100 for r in rows]
    oth = [t5 - tg for t5, tg in zip(top5, tgt)]
    rest = [100 - t5 for t5 in top5]
    ax.barh(names, tgt, color="#2ca02c", label="целевые страны (из топ-5)")
    ax.barh(names, oth, left=tgt, color="#7f7f7f", label="прочие страны из топ-5")
    ax.barh(names, rest, left=top5, color="#d9d9d9", label="остальные страны")
    ax.set_xlim(0, 100)
    ax.set_xlabel("доля визитов, %")
    ax.legend(fontsize=8, loc="lower right")
    fig.tight_layout(); fig.savefig(OUT / "08_география_сайтов.png", dpi=200); plt.close(fig)


if __name__ == "__main__":
    fig_quality(); fig_geo(); print("ok")
