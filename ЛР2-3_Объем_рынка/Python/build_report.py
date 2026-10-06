"""Сборка ОТЧЕТ.md ЛР2-3 из модулей rep_a, rep_b (числа берутся из results/analysis.json и results/excel_results.json)."""
import re
from pathlib import Path

import rep_common as C
import rep_a, rep_b

ROOT = Path(__file__).resolve().parent.parent
for m in (rep_a, rep_b):
    m.run()
text = C.resolve("\n".join(C.OUT)) + "\n"
text = re.sub(r"\n{3,}", "\n\n", text)
import sys as _sys
_sys.path.insert(0, str(ROOT.parent / "_tools"))
from renumber_sources import renumber
text = renumber(text)
import currency_eur
from nocolon import apply as _nocolon, capfirst_cells
text = currency_eur.apply(text)
text = _nocolon(text)
text = capfirst_cells(text)
(ROOT / "ОТЧЕТ.md").write_text(text, encoding="utf-8")
bad_yo = text.count("ё") + text.count("Ё")
bad_dash = text.count("—")
print(f"ОТЧЕТ.md: {len(text)} символов, таблиц {len(C.TABLES)}, рисунков {len(C.FIGS)}, 'ё': {bad_yo}, '—': {bad_dash}")
print("неразрешенных ссылок:", re.findall(r"\[\[[tf]:\w+\]\]", text))
print("auto intros:", C.AUTO)
