"""ЛР5: независимый пересчет показателей на Python и сверка с Excel."""
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.stdout.reconfigure(encoding="utf-8")
from data5 import *

X = json.loads((HERE / "results" / "excel_results.json").read_text(encoding="utf-8"))
A = {"checks": []}


def check(name, py, xl, tol=1e-6):
    ok = abs(py - xl) <= tol * max(1.0, abs(py))
    A["checks"].append({"name": name, "python": py, "excel": xl, "ok": bool(ok)})


# --- классификация ---
W = [20, 15, 18, 10, None, None, None, None, None, None]
w = [r[2] for r in X["cls"]["weights"]]
cls = []
for (kid, nm, site, country, typ, prod, seg, sc), row in zip(CLS, X["cls"]["rows"]):
    score = sum(a * b for a, b in zip(sc, w)) / sum(w)
    check(f"Взвешенный балл {nm}", round(score, 2), row[2], 5e-3)
    lvl = "Прямой конкурент" if score >= 4.2 and sc[0] >= 3 and sc[2] >= 3 else "Близкий альтернативный конкурент" if score >= 3.4 else "Косвенный конкурент / заменитель" if score >= 2.5 else "Смежная или потенциальная конкуренция"
    cls.append(dict(id=kid, name=nm, score=score, level=row[4], my_level=lvl, limiter=row[3]))
A["cls"] = cls
A["weights"] = w
# --- Кано ---
n = len(COMP)
kano = []
for (rid, fname, task, cat, wgt, lvl, pres, strn), row in zip(FEAT, X["lk"]["kano"]):
    share = sum(pres) / n
    check(f"Доля конкурентов {rid}", share, row[5])
    used = [s for p, s in zip(pres, strn) if p]
    kano.append(dict(id=rid, name=fname, cat=cat, weight=wgt, count=sum(pres), share=share, avg=(sum(used) / len(used) if used else 0), status=row[7], decision=row[8]))
A["kano"] = kano
# --- сравнение ---
cmpw = {}
import re
for pid, vals in CMP:
    pass
comp_avg = {c["id"]: 0 for c in COMP}
# веса критериев из Excel (столбец D)
wt = [r[2] for r in X["lk"]["compare"]]
tot_w = sum(wt)
for j, c in enumerate(COMP):
    comp_avg[c["id"]] = sum(v[j] * x for (pid, v), x in zip(CMP, wt)) / tot_w
A["compare_weighted"] = comp_avg
for (pid, vals), row in zip(CMP, X["lk"]["compare"]):
    check(f"Среднее {pid}", sum(vals) / len(vals), row[10])
# --- бизнес-модель ---
bms = {}
for (cid, sc), row in zip(BM_SCORES.items(), X["bm"]["matrix"]):
    idx = sum(sc) / len(sc)
    check(f"Индекс силы БМ {cid}", idx, row[2])
    bms[cid] = idx
A["bm_index"] = bms
std = []
for (pr, blk, desc, cnt), row in zip(STD, X["bm"]["standard"]):
    sh = cnt / n
    check(f"Доля практики {pr[:20]}", sh, row[2])
    std.append(dict(practice=pr, block=blk, count=cnt, share=sh, status=row[3]))
A["std"] = std
A["all_ok"] = all(c["ok"] for c in A["checks"])
(HERE / "results" / "analysis.json").write_text(json.dumps(A, ensure_ascii=False, indent=1), encoding="utf-8")
bad = [c for c in A["checks"] if not c["ok"]]
print("проверок", len(A["checks"]), "расхождений", len(bad), bad[:3])
print({c["name"]: round(c["score"], 2) for c in cls})
