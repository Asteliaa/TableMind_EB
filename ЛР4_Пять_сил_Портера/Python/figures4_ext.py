"""Дополнительный рисунок ЛР4: кривая концентрации релевантного трафика."""
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent
A = json.loads((HERE / "results" / "analysis.json").read_text(encoding="utf-8"))
plt.rcParams.update({"font.size": 9, "axes.grid": True, "grid.alpha": 0.3, "axes.spines.top": False, "axes.spines.right": False})
sh = A["conc"]["shares"]
cum, c = [], 0
for _, v in sh:
    c += v * 100
    cum.append(c)
n = list(range(1, len(sh) + 1))
fig, ax = plt.subplots(figsize=(7.6, 4.4))
ax.bar(n, [v * 100 for _, v in sh], color="#9ecae1", label="доля игрока")
ax.plot(n, cum, color="#d62728", marker="o", label="накопленная доля")
ax.axhline(60, color="#7f7f7f", linestyle="--", linewidth=1)
ax.text(len(sh) - 0.1, 61.5, "порог высокой концентрации CR3 = 60 %", ha="right", fontsize=8, color="#555")
for i in (3, 5):
    ax.annotate(f"CR{i} = {cum[i-1]:.0f} %", (i, cum[i - 1]), textcoords="offset points", xytext=(8, -14), fontsize=8)
ax.set_xticks(n)
ax.set_xticklabels([d.split(".")[0] for d, _ in sh], rotation=45, ha="right", fontsize=8)
ax.set_ylabel("доля релевантного трафика, %")
ax.set_ylim(0, 105)
ax.legend(loc="center right", fontsize=8)
fig.tight_layout(); fig.savefig(HERE.parent / "Рисунки" / "05_кривая_концентрации.png", dpi=200); plt.close(fig)
print("ok")
