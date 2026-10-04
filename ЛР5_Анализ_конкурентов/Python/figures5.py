"""Рисунки ЛР5."""
import json
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

HERE = Path(__file__).resolve().parent
OUT = HERE.parent / "Рисунки"
OUT.mkdir(exist_ok=True)
from data5 import *

A = json.loads((HERE / "results" / "analysis.json").read_text(encoding="utf-8"))
X = json.loads((HERE / "results" / "excel_results.json").read_text(encoding="utf-8"))
plt.rcParams.update({"font.size": 9, "axes.grid": True, "grid.alpha": 0.3, "axes.spines.top": False, "axes.spines.right": False})
LCOL = {"Прямой конкурент": "#d62728", "Близкий альтернативный конкурент": "#ff7f0e", "Косвенный конкурент / заменитель": "#1f77b4", "Смежная или потенциальная конкуренция": "#7f7f7f"}


MAN = {"Datarails": (-6, 6), "DataSnipper": (6, 6), "Arixcel": (6, -13), "Shortcut": (-6, 6)}


def f1():
    fig, ax = plt.subplots(figsize=(8.2, 4.8))
    seen = {}
    placed = []
    for (kid, nm, site, country, typ, prod, seg, sc), c in zip(CLS, A["cls"]):
        x = (sc[0] + sc[2]) / 2; y = (sc[1] + sc[6]) / 2
        ax.scatter(x, y, s=70, color=LCOL[c["level"]], zorder=3)
        k = seen.get((x, y), 0); seen[(x, y)] = k + 1
        near = [xx for (xx, yy) in placed if yy == y and abs(xx - x) < 1.1 and xx != x]
        left = bool(near) and min(near) < x
        placed.append((x, y))
        ax.annotate(f"{nm} ({c['score']:.2f})".replace(".ai", "#ai").replace(".", ",").replace("#ai", ".ai"), (x, y), textcoords="offset points", xytext=MAN.get(nm, (-5 if left else 5, 4 + 11 * k)), fontsize=8, ha=("right" if MAN[nm][0] < 0 else "left") if nm in MAN else ("right" if left else "left"))
    ax.axvline(3, color="k", lw=0.8, ls="--"); ax.axhline(3, color="k", lw=0.8, ls="--")
    ax.set_xlim(0.8, 5.3); ax.set_ylim(1.2, 5.3)
    ax.set_xlabel("задача и функциональная заменяемость, 0-5"); ax.set_ylabel("аудитория и география, 0-5")
    for t, c in LCOL.items():
        ax.scatter([], [], color=c, label=t)
    ax.legend(fontsize=7.5, loc="lower right")
    fig.tight_layout(); fig.savefig(OUT / "01_карта_конкуренции.png", dpi=200); plt.close(fig)


def f2():
    K = A["kano"]
    cols = {"Обязательное": "#d62728", "Линейное": "#1f77b4", "Привлекательное": "#2ca02c", "Безразличное": "#7f7f7f"}
    rows = sorted(K, key=lambda k: k["share"])
    fig, ax = plt.subplots(figsize=(8.2, 7.0))
    import textwrap
    ax.barh([textwrap.fill(r["name"], 34) for r in rows], [r["share"] * 100 for r in rows], color=[cols[r["cat"]] for r in rows])
    ax.axvline(60, color="k", ls="--", lw=1); ax.axvline(35, color="k", ls=":", lw=1)
    ax.text(61, len(rows) - 0.6, "стандарт 60 %", fontsize=7.5); ax.text(36, len(rows) - 0.6, "формируется 35 %", fontsize=7.5)
    for t, c in cols.items():
        ax.add_patch(plt.Rectangle((0, 0), 0, 0, color=c, label=t))
    ax.legend(fontsize=7.5, loc="center right"); ax.set_xlabel("доля из 7 конкурентов, у которых признак есть, %")
    plt.setp(ax.get_yticklabels(), fontsize=7.5)
    fig.tight_layout(); fig.savefig(OUT / "02_кано_доли.png", dpi=200); plt.close(fig)


GROUPS = [("Сущность", range(0, 3)), ("Функциональность", range(3, 6)), ("Ценность", range(6, 9)), ("Доверие", range(9, 12)), ("Интерфейс", range(12, 15)), ("Тарифы", range(15, 17)), ("Сопровождение", range(17, 19)), ("Расширение", range(19, 21)), ("Дифференциация", range(21, 23)), ("Стандарт", range(23, 25))]


def f3():
    M = np.array([v for _, v in CMP], dtype=float)  # 25 x 7
    G = np.array([[M[list(rg), j].mean() for j in range(7)] for _, rg in GROUPS])
    fig, ax = plt.subplots(figsize=(8.2, 4.4))
    im = ax.imshow(G, cmap="YlGnBu", vmin=0, vmax=5, aspect="auto")
    ax.set_xticks(range(7)); ax.set_xticklabels([c["name"] for c in COMP], rotation=30, ha="right", fontsize=8)
    ax.set_yticks(range(len(GROUPS))); ax.set_yticklabels([g for g, _ in GROUPS], fontsize=8)
    ax.grid(False)
    for i in range(G.shape[0]):
        for j in range(G.shape[1]):
            ax.text(j, i, f"{G[i, j]:.1f}".replace(".", ","), ha="center", va="center", fontsize=7.5, color="white" if G[i, j] > 3.2 else "black")
    fig.colorbar(im, ax=ax, fraction=0.03)
    fig.tight_layout(); fig.savefig(OUT / "03_сравнение_товаров.png", dpi=200); plt.close(fig)


def f4():
    names = [c["name"] for c in COMP]
    bm = [A["bm_index"][c["id"]] for c in COMP]
    cw = [A["compare_weighted"][c["id"]] for c in COMP]
    x = np.arange(len(names))
    fig, ax = plt.subplots(figsize=(8.2, 3.8))
    ax.bar(x - 0.2, bm, 0.4, label="индекс силы бизнес-модели (1-5)", color="#1f77b4")
    ax.bar(x + 0.2, cw, 0.4, label="взвешенная оценка цифрового товара (0-5)", color="#ff7f0e")
    ax.set_xticks(x); ax.set_xticklabels(names, rotation=20, ha="right", fontsize=8); ax.set_ylim(0, 5.5)
    for xi, v in zip(x - 0.2, bm):
        ax.text(xi, v + 0.05, f"{v:.1f}".replace(".", ","), ha="center", fontsize=7.5)
    for xi, v in zip(x + 0.2, cw):
        ax.text(xi, v + 0.05, f"{v:.1f}".replace(".", ","), ha="center", fontsize=7.5)
    ax.legend(fontsize=8, loc="upper right")
    fig.tight_layout(); fig.savefig(OUT / "04_бизнес_модель_и_товар.png", dpi=200); plt.close(fig)


if __name__ == "__main__":
    for f in (f1, f2, f3, f4):
        f()
    print(sorted(p.name for p in OUT.glob("*.png")))
