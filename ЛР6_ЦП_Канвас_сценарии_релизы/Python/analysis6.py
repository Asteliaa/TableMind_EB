"""ЛР6: независимый пересчет показателей на Python и сверка с Excel; рисунки."""
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.stdout.reconfigure(encoding="utf-8")
from data6a import *
from data6b import *
from data6c import UNITS_UI

X = json.loads((HERE / "results" / "excel_results.json").read_text(encoding="utf-8"))
A = {"checks": []}


def check(name, py, xl, tol=1e-6):
    ok = abs(py - xl) <= tol * max(1.0, abs(py))
    A["checks"].append({"name": name, "python": py, "excel": xl, "ok": bool(ok)})


# 6.1: индекс проблемы
for p, row in zip(PROB, X["vp"]["problems"]):
    idx = round((p[4] * 0.18 + p[5] * 0.22 + p[6] * 0.18 + p[7] * 0.2 + p[8] * 0.14 + p[9] * 0.08) * 20)
    check(f"Индекс проблемы {p[0]}", idx, row[2])
# 6.1: единица монетизации (M = (J*0.4 + K*0.3 + L*0.3)*20)
for m, row in zip(MONET, X["vp"]["monet"]):
    check(f"Потенциал монетизации {m[0]}", round((m[8] * 0.4 + m[9] * 0.3 + m[10] * 0.3) * 20), row[5])
    check(f"Выручка {m[0]}", m[5] * m[6], row[4])
# 6.2: индекс приоритета сегмента
for s, row in zip(SEGS, X["canvas"]["segments"]):
    check(f"Индекс сегмента {s[1]}", round((s[5] + s[6] + s[7] + s[8]) / 4, 2), row[1])
# 6.3: индекс требования и MVP
mvp = 0
for r, row in zip(REQ, X["logic"]["req"]):
    m = (r[8] + r[9]) * r[11] / r[10]
    check(f"Индекс требования {r[0]}", m, row[3])
    pr = "Критический" if m >= 16 else "Высокий" if m >= 10 else "Средний" if m >= 6 else "Низкий" if m >= 3 else "Отложить"
    if pr in ("Критический", "Высокий"):
        mvp += 1
    check(f"MVP {r[0]}", 1.0 if row[5] == "Да" else 0.0, 1.0 if pr in ("Критический", "Высокий") else 0.0)
A["mvp"] = mvp
for s, row in zip(SCEN, X["logic"]["scen"]):
    check(f"Индекс сценария {s[0]}", s[13] * s[14] * s[15] * s[16] / 20, row[3])
# 6.5: баллы единиц
first = nxt = 0
for u, row in zip(UNITS, X["release"]["units"]):
    G, H, I, J, K, L, M, N, O, P = u[6]
    q = round(G * 0.18 + H * 0.15 + I * 0.12 + J * 0.08 + K * 0.13 + L * 0.1 + M * 0.12 + N * 0.12 - O * 0.07 + P * 0.1, 2)
    r_ = round(I * 0.16 + J * 0.2 + K * 0.12 + L * 0.1 + (6 - M) * 0.08 + (6 - N) * 0.08 + O * 0.06 + P * 0.14 + G * 0.06, 2)
    rel = "Первый релиз" if (G >= 4 and H >= 4 and K >= 3 and P >= 4) or q >= 3.7 else ("Следующий релиз" if r_ >= 3.6 else "Не включать / проверить")
    check(f"Балл первого релиза {u[0]}", q, row[2], 1e-9)
    check(f"Балл следующих релизов {u[0]}", r_, row[3], 1e-9)
    check(f"Релиз {u[0]}", 1.0 if row[4] == rel else 0.0, 1.0)
    first += rel == "Первый релиз"
    nxt += rel == "Следующий релиз"
A["release_counts"] = {"first": first, "next": nxt, "other": len(UNITS) - first - nxt}
# риски
for r, row in zip(RISKS, X["release"]["risks"]):
    check(f"Индекс риска {r[0]}", r[4] * r[5], row[3])
# 6.4: приоритет экрана
for u, row in zip(UNITS_UI, X["screens"]["prio"]):
    v, b, f, r_, c = u[16]
    idx = max(0, min(100, round(((v * 0.3) + (b * 0.3) + (f * 0.15) + (r_ * 0.15) - (c * 0.1)) * 20)))
    check(f"Приоритет экрана {u[0]}", idx, row[2])
A["all_ok"] = all(c["ok"] for c in A["checks"])
(HERE / "results" / "analysis.json").write_text(json.dumps(A, ensure_ascii=False, indent=1), encoding="utf-8")
bad = [c for c in A["checks"] if not c["ok"]]
print("проверок", len(A["checks"]), "расхождений", len(bad), bad[:4], A["mvp"], A["release_counts"])
