"""Заполнение всех книг ЛР2-3 через COM и сохранение результатов в results/*.json."""
import json
import sys
from pathlib import Path

from inputs import load_sw
from xlh import XL

HERE = Path(__file__).resolve().parent
(HERE / "results").mkdir(exist_ok=True)
sys.stdout.reconfigure(encoding="utf-8")

only = sys.argv[1:] or ["pc", "tb", "pay", "fun", "sc", "ob", "se"]
sw = load_sw()
x = XL()
results = {}
try:
    for key in only:
        mod = __import__("fill_" + key)
        print("fill", key)
        mod.fill(x, sw, results)
        print("  errors:", results[key].get("errors"))
finally:
    x.quit()
    out = HERE / "results" / "excel_results.json"
    old = json.loads(out.read_text(encoding="utf-8")) if out.exists() else {}
    old.update(results)
    out.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8")
