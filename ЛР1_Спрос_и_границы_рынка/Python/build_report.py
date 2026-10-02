"""Сборка ОТЧЕТ.md ЛР1 из модулей rep_s1..rep_s4 (числа берутся из results/analysis.json и results/xlt_*.json)."""
import re
from pathlib import Path

import rep_common as C
import rep_intro
import rep_s1, rep_s2, rep_s3, rep_s4

ROOT = Path(__file__).resolve().parent.parent
for m in (rep_s1, rep_s2, rep_s3, rep_s4):
    m.run()
text = C.resolve("\n".join(C.OUT)) + "\n"
text = re.sub(r"\n{3,}", "\n\n", text)
(ROOT / "ОТЧЕТ.md").write_text(text, encoding="utf-8")
# проверки
bad_yo = text.count("ё") + text.count("Ё")
bad_dash = text.count("—")
n_t, n_f = len(C.TABLES), len(C.FIGS)
print(f"ОТЧЕТ.md: {len(text)} символов, таблиц {n_t}, рисунков {n_f}, 'ё': {bad_yo}, '—': {bad_dash}")
unres = re.findall(r"\[\[[tf]:\w+\]\]", text)
print("неразрешенных ссылок:", unres)
print("auto intros:", C.AUTO)
