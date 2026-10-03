"""Рисунки ЛР4."""
import csv
import json
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent
ROOT4 = HERE.parent
sys.path.insert(0, str(ROOT4.parent / "ЛР2-3_Объем_рынка" / "Python"))
from inputs import *

OUT = ROOT4 / "Рисунки"
OUT.mkdir(exist_ok=True)
A = json.loads((HERE / "results" / "analysis.json").read_text(encoding="utf-8"))
sw = load_sw()
NM = {r["domain"]: r["name"] for r in sw}
TY = {r["domain"]: r["type"] for r in sw}
COL = {"прямой": "#d62728", "частичный": "#ff7f0e", "заменитель": "#1f77b4"}
plt.rcParams.update({"font.size": 9, "axes.grid": True, "grid.alpha": 0.3, "axes.spines.top": False, "axes.spines.right": False})


def f1():
    sh = A["conc"]["shares"]
    fig, ax = plt.subplots(figsize=(8.2, 4.4))
    names = [NM[d] for d, _ in sh][::-1]
    vals = [v * 100 for _, v in sh][::-1]
    cols = [COL[TY[d]] for d, _ in sh][::-1]
    ax.barh(names, vals, color=cols)
    for i, v in enumerate(vals):
        ax.text(v + 0.3, i, f"{v:.1f} %", va="center", fontsize=8)
    ax.set_xlabel(f"доля релевантного трафика, % (CR3 = {A['conc']['cr3'] * 100:.0f} %, CR5 = {A['conc']['cr5'] * 100:.0f} %, HHI = {A['conc']['hhi']:.0f})")
    for t, c in COL.items():
        ax.barh([], [], color=c, label=t)
    ax.legend(loc="lower right", fontsize=8)
    fig.tight_layout(); fig.savefig(OUT / "01_доли_трафика.png", dpi=200); plt.close(fig)


def f2():
    b = A["barrier"]
    fig, ax = plt.subplots(figsize=(8.2, 4.2))
    idx = list(range(len(b["factors"])))[::-1]
    ax.barh(idx, b["scores"], color=["#d62728" if s >= 5 else "#ff7f0e" if s >= 3 else "#2ca02c" for s in b["scores"]])
    ax.set_yticks(idx); ax.set_yticklabels([f"{f} (вес {w:.2f})" for f, w in zip(b["factors"], b["weights"])], fontsize=8)
    ax.set_xlim(0, 5.5); ax.set_xlabel(f"балл давления (1, 3 или 5); итоговый индекс барьеров = {b['index']:.2f}")
    ax.axvline(b["index"], color="k", ls="--", lw=1)
    fig.tight_layout(); fig.savefig(OUT / "02_барьеры_входа.png", dpi=200); plt.close(fig)


def f3():
    p = A["porter"]
    names = list(p["matrix"].keys()); vals = list(p["matrix"].values())
    fig, axes = plt.subplots(1, 2, figsize=(8.2, 3.8), gridspec_kw={"width_ratios": [1.1, 1]})
    ax = axes[0]
    ax.barh(names[::-1], vals[::-1], color="#1f77b4")
    ax.axvline(p["avg5"], color="k", ls="--", lw=1)
    ax.set_xlim(0, 5.5); ax.set_xlabel(f"давление, 1-5 (среднее {p['avg5']:.2f})"); ax.set_title("Пять сил", fontsize=9, loc="left")
    ax = axes[1]
    an = list(p["amps"].keys()); av = list(p["amps"].values())
    ax.barh(an[::-1], av[::-1], color="#7f7f7f")
    ax.set_xlim(0, 5.5); ax.set_xlabel(f"выраженность, 1-5 (среднее {p['amps_avg']:.1f})"); ax.set_title("Цифровые усилители", fontsize=9, loc="left")
    fig.tight_layout(); fig.savefig(OUT / "03_пять_сил.png", dpi=200); plt.close(fig)


def f4():
    rows = []
    with open(ROOT4 / "Материалы_собранные" / "репутация_2026-10-03.csv", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f, delimiter=";"):
            n = float(r["g2_reviews"]) if r["g2_reviews"] != "" else None
            rows.append((NM[r["domain"]], n, r["g2_rating"]))
    rows.sort(key=lambda t: (t[1] is None, -(t[1] or 0)))
    fig, ax = plt.subplots(figsize=(8.2, 3.9))
    xs = [t[0] for t in rows]
    ys = [t[1] if t[1] is not None else 0 for t in rows]
    ax.bar(xs, [max(y, 0.3) for y in ys], color=[COL[TY[d]] for d in [k for k, v in NM.items() if v in xs for _ in [0]] ] if False else "#1f77b4")
    ax.set_yscale("log")
    for i, t in enumerate(rows):
        lab = "нет карточки" if t[1] is None else (f"{int(t[1])}" + (f" ({t[2]})" if t[2] else ""))
        ax.text(i, max(ys[i], 0.3) * 1.15, lab, ha="center", fontsize=7)
    ax.set_ylabel("отзывов на G2 (логарифмическая шкала)")
    plt.setp(ax.get_xticklabels(), rotation=35, ha="right", fontsize=8)
    ax.set_ylim(0.2, 1500)
    fig.tight_layout(); fig.savefig(OUT / "04_отзывы_G2.png", dpi=200); plt.close(fig)


if __name__ == "__main__":
    for f in (f1, f2, f3, f4):
        f()
    print(sorted(p.name for p in OUT.glob("*.png")))
