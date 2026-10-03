"""Заполнение всех книг ЛР2-3 через COM и сохранение результатов в results/*.json."""
import json
import sys
from pathlib import Path

from inputs import load_sw
from xlh import XL

HERE = Path(__file__).resolve().parent
(HERE / "results").mkdir(exist_ok=True)
sys.stdout.reconfigure(encoding="utf-8")

only = sys.argv[1:] or ["pc", "tb", "pay", "fun", "ob", "se", "sc"]
sw = load_sw()
x = XL()
results = {}
try:
    for key in only:
        mod = __import__("fill_" + key)
        print("fill", key)
        if key == "sc":
            from inputs import FUNNEL, PC_SCEN, SC, arppu
            prev = json.loads((HERE / "results" / "excel_results.json").read_text(encoding="utf-8")) if (HERE / "results" / "excel_results.json").exists() else {}
            prev.update(results)
            org_rel = prev["fun"]["channels"][1][5]
            ext = {
                "search": {sc: org_rel / FUNNEL["cover"] * 12 * FUNNEL["reach"][sc] * FUNNEL["v2l"][sc] * FUNNEL["l2p"][sc] * arppu(sc)[0] for sc in SC},
                "pc": {sc: prev["pc"]["scen"][sc][2] for sc in SC},
                "tb": {sc: prev["tb"]["sverka"]["bu"][i] * PC_SCEN[sc]["share"] for i, sc in enumerate(SC)},
            }
            mod.fill(x, sw, results, ext)
        else:
            mod.fill(x, sw, results)
        print("  errors:", results[key].get("errors"))
finally:
    x.quit()
    out = HERE / "results" / "excel_results.json"
    old = json.loads(out.read_text(encoding="utf-8")) if out.exists() else {}
    old.update(results)
    out.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8")
