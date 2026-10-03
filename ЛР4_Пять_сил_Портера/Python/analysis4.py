"""ЛР4: независимый пересчет показателей конкуренции и барьеров на Python, сверка с Excel, пять сил Портера, чувствительность."""
import csv
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT4 = HERE.parent
sys.path.insert(0, str(ROOT4.parent / "ЛР2-3_Объем_рынка" / "Python"))
sys.stdout.reconfigure(encoding="utf-8")
from inputs import *

X = json.loads((HERE / "results" / "excel_results.json").read_text(encoding="utf-8"))
P = json.loads((HERE / "results" / "porter_results.json").read_text(encoding="utf-8"))
sw = load_sw()
A = {"checks": []}


def check(name, py, xl, tol=1e-6):
    ok = abs(py - xl) <= tol * max(1.0, abs(py))
    A["checks"].append({"name": name, "python": py, "excel": xl, "ok": bool(ok)})


# --- трафик и концентрация ---
eff = {r["domain"]: r["visits_month"] * round(r["geo_share"] * r["rel"] * r["qual"], 4) for r in sw}
tot = sum(eff.values())
sh = {d: v / tot for d, v in eff.items()}
srt = sorted(sh.values(), reverse=True)
cr3, cr5 = sum(srt[:3]), sum(srt[:5])
hhi = sum(v * v for v in sh.values()) * 10000
check("Релевантный трафик", tot, X["traffic"], 1e-4)
check("CR3", cr3, X["cr3"], 1e-4); check("CR5", cr5, X["cr5"], 1e-4); check("HHI", hhi, X["hhi"], 1e-4)
A["conc"] = dict(total=tot, cr3=cr3, cr5=cr5, hhi=hhi, shares=sorted(sh.items(), key=lambda kv: -kv[1]))

# --- каналы (взвешенно по релевантному трафику) ---
def wch(f):
    return sum(eff[r["domain"]] * f(r) for r in sw) / tot


ch = {"direct": wch(lambda r: r["ch"]["direct"]), "org": wch(lambda r: r["ch"]["org_search"]), "paid": wch(lambda r: r["ch"]["paid_search"]),
      "soc": wch(lambda r: r["ch"]["org_social"] + r["ch"]["paid_social"]), "ref": wch(lambda r: r["ch"]["referrals"]), "disp": wch(lambda r: r["ch"]["display"])}
for k, i in (("direct", 0), ("org", 1), ("paid", 2), ("soc", 3), ("ref", 4), ("disp", 5)):
    check(f"Канал {k}", ch[k], X["channels"][i][1], 1e-4)
A["channels"] = ch

# --- индекс барьеров ---
def sc(v, thr):
    return 5 if v >= thr else (3 if v >= thr * 0.65 else 1)


thr = dict(cr3=0.6, cr5=0.75, hhi=2500, paid=0.2, org=0.45, brand=0.5, direct=0.3, soc=0.2, ref=0.15, disp=0.1)
d_conc = 5 if (cr3 >= thr["cr3"] or hhi >= thr["hhi"]) else (3 if (cr3 >= thr["cr3"] * 0.65 or hhi >= thr["hhi"] * 0.65) else 1)
d_seo = sc(ch["org"], thr["org"])
paid_c = ch["paid"]
paid_idx = paid_c
d_paid = 5 if paid_c >= 0.6 else (3 if paid_c >= 0.35 else 1)
d_brand = sc(ch["direct"], thr["brand"])
search_proxy = (min(ch["org"] / thr["org"], 1) + min(ch["paid"] / thr["paid"], 1)) / 2
d_search = 5 if search_proxy >= 0.7 else (3 if search_proxy >= 0.4 else 1)
A["proxy_search"] = search_proxy
# репутация
rep = {}
with open(ROOT4 / "Материалы_собранные" / "репутация_2026-10-03.csv", encoding="utf-8-sig") as f:
    for r in csv.DictReader(f, delimiter=";"):
        rep[r["domain"]] = r
idx = []
for r in sw:
    d = rep[r["domain"]]
    rt = [float(d[k]) for k in ("g2_rating", "capterra_rating", "trustpilot_rating") if d[k] != ""]
    cnt = [float(d[k]) for k in ("g2_reviews", "capterra_reviews", "trustpilot_reviews") if d[k] != ""]
    comps = []
    if rt:
        comps.append(sum(rt) / len(rt) / 5)
    if cnt:
        comps.append(min(sum(cnt) / 300, 1))
    if comps:
        idx.append(min(1, sum(comps) / len(comps)))
rev_idx = sum(idx) / len(idx)
check("Индекс доверия", rev_idx, X["rev_idx"], 1e-6)
d_rep = 5 if rev_idx >= 0.7 else (3 if rev_idx >= 0.4 else 1)
tech_pts = []
TECH = {"perfectxl.com": (3, 3), "operis.com": (3, 2), "arixcel.com": (2, 2), "spreadsheetdetective.com": (2, 1), "cimcon.com": (4, 4),
        "shortcut.ai": (5, 5), "rows.com": (5, 4), "numerous.ai": (4, 4), "formulabot.com": (3, 3), "ajelix.com": (3, 3), "datarails.com": (5, 5), "julius.ai": (5, 4)}
for r in sw:
    a, b = TECH[r["domain"]]
    tech_pts.append((a / 5 + b / 5) / 2)
tech_idx = sum(tech_pts) / len(tech_pts)
check("Индекс ресурсной силы", tech_idx, X["tech_idx"], 1e-6)
d_tech = 5 if tech_idx >= 0.7 - 1e-12 else (3 if tech_idx >= 0.4 else 1)
plat = 0.6
d_plat = 5 if plat >= 0.7 else (3 if plat >= 0.4 else 1)
direct_share = sum(sh[r["domain"]] for r in sw if r["type"] == "прямой")
subst = 1 - direct_share
d_sub = 5 if subst >= 0.7 else (3 if subst >= 0.4 else 1)
check("Доля заменителей", subst, X["barrier"][8][1], 1e-6)
W = [0.16, 0.12, 0.12, 0.12, 0.12, 0.12, 0.12, 0.08, 0.08]
D = [d_conc, d_seo, d_paid, d_brand, d_search, d_rep, d_tech, d_plat, d_sub]
index = sum(w * d for w, d in zip(W, D)) / sum(W)
check("Индекс барьеров входа", index, X["index"], 1e-9)
A["barrier"] = dict(scores=D, weights=W, index=index, raw_weighted=sum(w * d for w, d in zip(W, D)), sumw=sum(W),
                    factors=["Концентрация трафика", "Органическое SEO-давление", "Платное рекламное давление", "Брендовая сила", "Поисково-рекламная конкуренция",
                             "Репутационный барьер", "Технологический и ресурсный барьер", "Платформенная зависимость", "Угроза заменителей"])
A["subst"] = subst; A["direct_share"] = direct_share

# --- пять сил Портера ---
lim = {"substitutes": (1, 3), "rivalry": (4, 12), "entrants": (8, 24), "buyers": (4, 12), "suppliers": (4, 8)}
pt = {"substitutes": P["substitutes"], "rivalry": P["rivalry"], "entrants": P["entrants"], "buyers": P["buyers"], "suppliers": P["suppliers"]}
for g, (lo, hi) in lim.items():
    rows = P["rows"][g]
    check(f"Сумма баллов {g}", sum(v[0] for v in rows.values()), pt[g], 1e-9)


def to5(g):
    lo, hi = lim[g]
    return 1 + 4 * (pt[g] - lo) / (hi - lo)


# шкала 1-5 для итоговой матрицы задана экспертно по результатам, нормировка проверяется
norm = {"Конкуренция игроков": to5("rivalry"), "Угроза новых участников": to5("entrants"), "Сила покупателей": to5("buyers"), "Сила поставщиков": to5("suppliers"), "Угроза заменителей": to5("substitutes")}
matrix = {k: round(v, 2) for k, v in norm.items()}
amps = {"Сетевые эффекты": 2, "Данные": 3, "Алгоритмическая видимость": 4, "Платформенная зависимость": 3, "Издержки переключения": 2, "Экосистемная связанность": 4}
avg5 = sum(matrix.values()) / 5
avg_all = sum(matrix.values()) / 5 * 0.8 + sum(amps.values()) / len(amps) * 0.2
A["porter"] = dict(sums=pt, limits=lim, matrix=matrix, normalized=norm, amps=amps, avg5=avg5, amps_avg=sum(amps.values()) / len(amps), avg_with_amps=avg_all)

# --- чувствительность индекса барьеров ---
sens = []
for name, f in (("Веса: равные", [1 / 9] * 9),):
    sens.append((name, sum(w * d for w, d in zip(f, D)) / sum(f)))
for i, fac in enumerate(A["barrier"]["factors"]):
    lo = D.copy(); lo[i] = max(1, D[i] - 2)
    hi = D.copy(); hi[i] = min(5, D[i] + 2)
    sens.append((fac, sum(w * d for w, d in zip(W, lo)) / sum(W), sum(w * d for w, d in zip(W, hi)) / sum(W)))
A["sens"] = sens
A["all_ok"] = all(c["ok"] for c in A["checks"])
(HERE / "results" / "analysis.json").write_text(json.dumps(A, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
bad = [c for c in A["checks"] if not c["ok"]]
print("проверок", len(A["checks"]), "расхождений", len(bad), bad)
print("CR3 %.3f CR5 %.3f HHI %.0f  индекс %.3f  subst %.3f" % (cr3, cr5, hhi, index, subst))
print(A["porter"])
