"""Рисунки ЛР1 (PNG, 200 dpi) по данным results/analysis.json и results/xlt_*.json."""
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent
OUT = HERE.parent / "Рисунки"
OUT.mkdir(exist_ok=True)
A = json.loads((HERE / "results" / "analysis.json").read_text(encoding="utf-8"))
MONTHS = A["months"]
NAMES = {"A": "Кластер А: проверка и аудит таблиц Excel", "B": "Кластер Б: AI-инструменты для Excel", "C": "Кластер В: аудит финансовых моделей"}
MRU = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
C1, C2, C3 = "#1f77b4", "#d62728", "#2ca02c"
plt.rcParams.update({"font.size": 9, "axes.grid": True, "grid.alpha": 0.3, "axes.spines.top": False, "axes.spines.right": False})
xs = list(range(60))
ticks = [i for i, m in enumerate(MONTHS) if m.endswith("-01") or i == 0]
tl = [MONTHS[i][:4] if MONTHS[i].endswith("-01") else MONTHS[i] for i in ticks]


def fig1():
    fig, axes = plt.subplots(3, 1, figsize=(8.2, 8.4), sharex=True)
    for ax, cl in zip(axes, "ABC"):
        w = A["word"][cl]
        gt = w["GT"]["series"]
        yv = w["YW_RU"]["series"]
        ymax = max(yv)
        ax.plot(xs, gt, color=C1, lw=1.2, label="Google Trends, мир (индекс 0-100)")
        ax.plot(xs, [v / ymax * 100 for v in yv], color=C2, lw=1.2, label="Вордстат, Россия (индекс, максимум = 100)")
        ax.plot(xs, [None if v is None else v / ymax * 100 for v in w["YW_RU"]["ma12"]], color=C2, lw=2.2, ls="--", label="Вордстат: скользящая средняя 12 мес.")
        ax.plot(xs, w["GT"]["ma12"], color=C1, lw=2.2, ls="--", label="Google Trends: скользящая средняя 12 мес.")
        ax.set_title(NAMES[cl], fontsize=10, loc="left")
        ax.set_ylabel("индекс")
        if cl == "A":
            ax.legend(fontsize=7.5, loc="upper left")
    axes[-1].set_xticks(ticks)
    axes[-1].set_xticklabels(tl, rotation=45, ha="right", fontsize=8)
    fig.tight_layout()
    fig.savefig(OUT / "01_динамика_кластеров.png", dpi=200)
    plt.close(fig)


def fig2():
    fig, axes = plt.subplots(1, 3, figsize=(9.4, 3.3), sharey=True)
    for ax, cl in zip(axes, "ABC"):
        si = A["word"][cl]["YW_RU"]["season_idx"]
        cols = [C2 if v >= 110 else ("#7f7f7f" if v >= 90 else C1) for v in si]
        ax.bar(range(12), si, color=cols)
        ax.axhline(100, color="black", lw=0.8)
        ax.set_xticks(range(12))
        ax.set_xticklabels(MRU, rotation=90, fontsize=8)
        ax.set_title({"A": "Кластер А", "B": "Кластер Б", "C": "Кластер В"}[cl], fontsize=10)
    axes[0].set_ylabel("сезонный индекс, 100 = среднее")
    fig.tight_layout()
    fig.savefig(OUT / "02_сезонные_индексы.png", dpi=200)
    plt.close(fig)


def fig3():
    import data_prep as dp
    fig, axes = plt.subplots(3, 1, figsize=(8.2, 7.6), sharex=True)
    qs = ["ошибки в формулах excel", "нейросеть для excel", "проверка формул excel"]
    m = dp.YW_MONTHS
    xx = list(range(len(m)))
    tk = [i for i, s in enumerate(m) if s.endswith("-01")]
    for ax, q in zip(axes, qs):
        r = dp.yw(q, 225)
        b = dp.yw(q, 149)
        ax.plot(xx, [r[x] for x in m], color=C2, lw=1.2, label="Россия (левая шкала)")
        ax.set_ylabel("запросов в месяц, Россия")
        ax2 = ax.twinx()
        ax2.plot(xx, [b[x] for x in m], color=C1, lw=1.2, label="Беларусь (правая шкала)")
        ax2.set_ylabel("запросов в месяц, Беларусь")
        ax2.spines["right"].set_visible(True)
        ax2.grid(False)
        ax.set_title("«" + q + "»", fontsize=10, loc="left")
        if q == qs[0]:
            h1, l1 = ax.get_legend_handles_labels()
            h2, l2 = ax2.get_legend_handles_labels()
            ax.legend(h1 + h2, l1 + l2, fontsize=8, loc="upper left")
    axes[-1].set_xticks(tk)
    axes[-1].set_xticklabels([m[i][:4] for i in tk], fontsize=8)
    fig.tight_layout()
    fig.savefig(OUT / "03_вордстат_россия_беларусь.png", dpi=200)
    plt.close(fig)


def fig4():
    fig, ax = plt.subplots(figsize=(8.2, 3.6))
    for cl, col, nm in (("A", C1, "кластер А"), ("B", C2, "кластер Б"), ("C", C3, "кластер В")):
        code = {"A": "A_проверка_таблиц", "B": "B_AI_для_Excel", "C": "C_финмодели"}[cl]
        k = A["xlt"][code]["K"]
        ax.plot(xs, [None if v is None else v for v in k], color=col, lw=1.3, label=nm)
    ax.axhline(0.9, color="gray", ls=":", lw=1)
    ax.axhline(1.1, color="gray", ls=":", lw=1)
    ax.axhline(1.0, color="black", lw=0.8)
    ax.set_ylim(0, 3)
    ax.set_ylabel("циклический коэффициент")
    ax.set_xticks(ticks)
    ax.set_xticklabels(tl, rotation=45, ha="right", fontsize=8)
    ax.legend(fontsize=8, loc="upper left")
    fig.tight_layout()
    fig.savefig(OUT / "04_циклический_коэффициент.png", dpi=200)
    plt.close(fig)


def fig5():
    p5 = [q for q in A["gt_queries"] if q["pack"] == "P5"]
    p5.sort(key=lambda q: q["avg"], reverse=True)
    fig, ax = plt.subplots(figsize=(8.0, 3.2))
    ax.barh([q["query"] for q in p5][::-1], [q["avg"] for q in p5][::-1], color=C1)
    ax.set_xscale("log")
    ax.set_xlabel("средний недельный индекс Google Trends в общем пакете (логарифмическая шкала)")
    for i, q in enumerate(p5[::-1]):
        ax.text(max(q["avg"], 0.02) * 1.15, i, f"{q['avg']:.2f}", va="center", fontsize=8)
    ax.set_xlim(0.01, 300)
    fig.tight_layout()
    fig.savefig(OUT / "05_масштаб_запросов_gt.png", dpi=200)
    plt.close(fig)


def fig6():
    dev = {"десктопы": 10757, "смартфоны": 9140, "планшеты": 93}
    fig, ax = plt.subplots(figsize=(5.4, 3.0))
    tot = sum(dev.values())
    ax.bar(list(dev), list(dev.values()), color=[C1, C2, "#7f7f7f"])
    for i, v in enumerate(dev.values()):
        ax.text(i, v + 200, f"{v}\n({v / tot * 100:.1f} %)", ha="center", fontsize=8)
    ax.set_ylim(0, 13500)
    ax.set_ylabel("запросов за 12 мес.")
    fig.tight_layout()
    fig.savefig(OUT / "06_устройства.png", dpi=200)
    plt.close(fig)


if __name__ == "__main__":
    for f in (fig1, fig2, fig3, fig4, fig5, fig6):
        f()
    print("ok", sorted(p.name for p in OUT.glob("*.png")))
