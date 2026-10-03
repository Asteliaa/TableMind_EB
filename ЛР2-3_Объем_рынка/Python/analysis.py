"""Независимый пересчет методик ЛР2-3 на Python (без Excel), сверка с книгами, чувствительность, сведение методов и сопоставление с целями проекта."""
import json
import sys
from pathlib import Path
from statistics import median

from inputs import *

sys.stdout.reconfigure(encoding="utf-8")
HERE = Path(__file__).resolve().parent
XLR = json.loads((HERE / "results" / "excel_results.json").read_text(encoding="utf-8"))
sw = load_sw()
TOL = 1e-6  # допуск сверки Excel и Python: относительная погрешность 1e-6 (зафиксирован до расчета)

A = {"checks": [], "tol": TOL}


def check(name, py, xl, rel=TOL):
    ok = abs(py - xl) <= rel * max(1.0, abs(py))
    A["checks"].append({"name": name, "python": py, "excel": xl, "ok": bool(ok)})
    return ok


# ---------- книга 1: потенциальные клиенты ----------
def seg_calc(s, mult=None):
    m = mult or dict(need=1, online=1, pay=1, check=1, freq=1)
    inb = s["total"] * s["border"]
    addr = inb * s["need"] * s["digital"] * s["payable"]
    sam = addr * s["check"] * s["freq"]
    return inb, addr, sam


tam = sum(s["total"] * s["check"] * s["freq"] for s in SEGMENTS)
sam = sum(seg_calc(s)[2] for s in SEGMENTS)
addr = sum(seg_calc(s)[1] for s in SEGMENTS)
som = sum(seg_calc(s)[2] * s["share"] for s in SEGMENTS)
check("PC TAM", tam, XLR["pc"]["tam"]); check("PC SAM", sam, XLR["pc"]["sam"]); check("PC адресуемые клиенты", addr, XLR["pc"]["addr"]); check("PC SOM по сегментам", som, XLR["pc"]["som"])
pc_scen = {}
for sc in SC:
    m = PC_SCEN[sc]
    t = tam * m["check"] * m["freq"]
    s_ = sam * m["need"] * m["online"] * m["pay"] * m["check"] * m["freq"]
    pc_scen[sc] = (t, s_, s_ * m["share"])
    for i, nm in enumerate(("TAM", "SAM", "SOM")):
        check(f"PC {SC_RU[sc]} {nm}", pc_scen[sc][i], XLR["pc"]["scen"][sc][i])

# ---------- Similarweb: релевантный трафик ----------
rel_m = sum(r["visits_month"] * r["geo_share"] * r["rel"] * r["qual"] for r in sw)
check("Релевантный трафик 12 сайтов, мес. (книга 6)", rel_m, XLR["ob"]["params"]["obs_m"], 1e-3)
full_m = rel_m / FUNNEL["cover"]
fun_full = XLR["fun"]["sw"]["full_traffic"]
check("Полный трафик сегмента, мес. (книга 4)", full_m, fun_full, 1e-3)

# ---------- книга 2: сверху вниз / снизу вверх ----------
SPEC = sum(s["total"] for s in SEGMENTS[:3]); SME = SEGMENTS[3]["total"]
TD_USERS = {"cons": 0.85, "base": 1.0, "opt": 1.15}
TD_PROD = dict(cons=0.07, base=0.10, opt=0.13); TD_NEED = dict(cons=0.20, base=0.30, opt=0.40); TD_DIG = dict(cons=0.55, base=0.65, opt=0.75)
base_check = round(arppu("base")[0], 2)
td = {}
for sc in SC:
    cl = round((SPEC + SME) * TD_USERS[sc]) * TD_PROD[sc] * 1.0 * TD_DIG[sc] * TD_NEED[sc]
    td[sc] = (cl, cl * base_check)
    check(f"TB сверху вниз {SC_RU[sc]} объем", td[sc][1], XLR["tb"]["td"][sc][1])
bu = sum(round(s["total"] * s["border"] * s["payable"]) * s["need"] * s["digital"] * s["check"] for s in SEGMENTS)
check("TB снизу вверх итог", bu, XLR["tb"]["bu"]["total"], 1e-6)
# воронка Similarweb книги 2
eng = {}
for r in sw:
    ch_rel = r["ch"]["direct"] + r["ch"]["org_search"] + r["ch"]["paid_search"] + r["ch"]["referrals"]
    eng[r["domain"]] = r["visits_month"] * r["geo_share"] * round(ch_rel, 4) * r["rel"] * r["qual"]
rel_year2 = sum(eng.values()) * 12
SCC = dict(cons=0.6, base=1.0, opt=1.4)
for sc in SC:
    vis = rel_year2 * SCC[sc] * round(1 / FUNNEL["cover"], 4)
    k = {"cons": (0.7, 0.8, 0.8, 0.9), "base": (1, 1, 1, 1), "opt": (1.3, 1.2, 1.2, 1.1)}[sc]
    vol = vis * FUNNEL["v2l"]["base"] * k[0] * FUNNEL["l2p"]["base"] * k[1] * base_check * k[2] * 1.0 * k[3]
    check(f"TB Similarweb-воронка {SC_RU[sc]}", vol, XLR["tb"]["sverka"]["sw"][SC.index(sc)], 1e-6)

# ---------- книга 3: платежеспособность ----------
pay_vol = sum(round(s["total"] * s["border"]) * s["need"] * s["digital"] * s["payable"] * s["freq"] * s["check"] for s in SEGMENTS)
check("PAY объем", pay_vol, XLR["pay"]["volume"], 1e-5)
for i, p in enumerate(PAY):
    check(f"PAY индекс {p['seg']}", p["budget"] / p["price"], XLR["pay"]["idx"][i][0])

# ---------- книга 4: цифровая воронка ----------
for sc in SC:
    c = arppu(sc)[2]
    vis = full_m * FUNNEL["reach"][sc]
    reg = vis * FUNNEL["v2l"][sc]
    paid = reg * FUNNEL["l2p"][sc]
    k = dict(zip(SC, ("cons", "base", "opt"))).get(sc)
    xr = XLR["fun"]["funnel"][sc]
    check(f"FUN визиты {SC_RU[sc]}", vis, xr[2], 1e-3); check(f"FUN регистрации {SC_RU[sc]}", reg, xr[4], 1e-3); check(f"FUN платящие/мес. {SC_RU[sc]}", paid, xr[6], 1e-3)
    check(f"FUN выручка в год {SC_RU[sc]}", paid * round(c, 2) * 12, xr[9], 1e-3)

# ---------- книга 5: сценарная оценка ----------
UNC = dict(cons=1.3, base=round(1 / FUNNEL["cover"], 3), opt=2.5); REL = dict(cons=0.75, base=0.90, opt=1.0); UNAV = dict(cons=0.8, base=1.0, opt=1.15)
SOMS = dict(cons=0.005, base=0.015, opt=0.04)
sc_res = {}
for sc in SC:
    mkt_y = rel_m * UNC[sc] * REL[sc] * UNAV[sc] * 12
    sales = mkt_y * FUNNEL["v2l"][sc] * FUNNEL["l2p"][sc]
    rev = sales * arppu(sc)[0]
    sc_res[sc] = dict(market_year=mkt_y, sales=sales, market_rev=rev, som_rev=rev * SOMS[sc], som_payers=sales * SOMS[sc])
    check(f"SC рыночные визиты/год {SC_RU[sc]}", mkt_y, XLR["sc"]["scen"][sc][5], 1e-3)
    check(f"SC SOM {SC_RU[sc]}", rev * SOMS[sc], XLR["sc"]["funnel"][sc][8], 5e-3)

# ---------- книга 6: объем по Similarweb ----------
COV = dict(cons=0.70, base=0.55, opt=0.40); SOMO = dict(cons=0.005, base=0.015, opt=0.04)
ob = {}
for sc in SC:
    orders = rel_m / COV[sc] * FUNNEL["v2l"][sc] * FUNNEL["l2p"][sc]
    rev_y = orders * arppu(sc)[2] * 12 * arppu(sc)[1]
    ob[sc] = (rev_y, rev_y * SOMO[sc])
    check(f"OB рынок/год {SC_RU[sc]}", rev_y, XLR["ob"]["som"][sc][0], 5e-3)

# ---------- книга 7: поисковый спрос ----------
check("SE среднее взвешенной частотности", XLR["se"]["weighted_avg"], XLR["se"]["dash"][2], 1e-6)
for sc in SC:
    wf = XLR["se"]["weighted_avg"]
    rev_m = wf * {"cons": 0.04, "base": 0.08, "opt": 0.12}[sc] * FUNNEL["v2l"][sc] * FUNNEL["l2p"][sc] * round(arppu(sc)[0], 2)
    check(f"SE выручка/мес. {SC_RU[sc]}", rev_m, XLR["se"]["funnel"][sc][9], 1e-6)

A["all_ok"] = all(c["ok"] for c in A["checks"])

# ---------- сведение методов ----------
methods = {
    "Потенциальные клиенты (книга 1)": dict(sam=[pc_scen[s][1] for s in SC], som=[pc_scen[s][2] for s in SC]),
    "Сверху вниз (книга 2)": dict(sam=[td[s][1] for s in SC], som=[td[s][1] * PC_SCEN[s]["share"] for s in SC]),
    "Снизу вверх (книга 2)": dict(sam=[bu * SCC[s] for s in SC], som=[bu * SCC[s] * PC_SCEN[s]["share"] for s in SC]),
    "Платежеспособность (книга 3)": dict(sam=[pay_vol * PC_SCEN[s]["need"] * PC_SCEN[s]["pay"] * PC_SCEN[s]["freq"] * PC_SCEN[s]["check"] for s in SC],
                                         som=[pay_vol * PC_SCEN[s]["need"] * PC_SCEN[s]["pay"] * PC_SCEN[s]["freq"] * PC_SCEN[s]["check"] * PC_SCEN[s]["share"] for s in SC]),
}
A["methods"] = methods
A["traffic_methods"] = {
    "Цифровая воронка (книга 4): оборот новых платящих за год": [XLR["fun"]["funnel"][s][9] for s in SC],
    "Сценарная оценка Similarweb (книга 5): рынок/год": [sc_res[s]["market_rev"] for s in SC],
    "Сценарная оценка Similarweb (книга 5): SOM/год": [sc_res[s]["som_rev"] for s in SC],
    "Объем по Similarweb (книга 6): рынок/год": [ob[s][0] for s in SC],
    "Объем по Similarweb (книга 6): SOM/год": [ob[s][1] for s in SC],
    "Поиск РФ+РБ (книга 7): выручка/год": [XLR["se"]["funnel"][s][10] for s in SC],
}
A["sc_res"] = sc_res
A["ob"] = {s: ob[s] for s in SC}
fin = {}
for k, i in (("low", 0), ("base", 1), ("high", 2)):
    fin[k] = dict(sam=median(m["sam"][i] for m in methods.values()), som=median(m["som"][i] for m in methods.values()))
A["final_potential"] = fin
A["tam_pz1_eur"] = 11e9 / EUR_USD * (2 / 3)
A["tam_model"] = tam

# ---------- цели проекта ----------
cover = FUNNEL["cover"]
paid_start_months = 4   # открытая бета с оплатой 01.06.2027 - 30.09.2027
reg_months = 6          # MVP 31.03.2027 - 30.09.2027
goal = {}
for sc in SC:
    payers_m = full_m * FUNNEL["reach"][sc] * FUNNEL["v2l"][sc] * FUNNEL["l2p"][sc]
    reg_m = full_m * FUNNEL["reach"][sc] * FUNNEL["v2l"][sc]
    goal[sc] = dict(payers_month=payers_m, reg_month=reg_m, payers_4m=payers_m * paid_start_months, reg_6m=reg_m * reg_months,
                    payers_12m=payers_m * 12, arppu=arppu(sc)[0], som_payers_pc=pc_scen[sc][2] / arppu(sc)[0],
                    som_payers_sc=sc_res[sc]["som_payers"])
need_payers_m = GOAL_PAY / paid_start_months
goal["need_reach"] = need_payers_m / (full_m * FUNNEL["v2l"]["base"] * FUNNEL["l2p"]["base"])
goal["need_conv"] = need_payers_m / (full_m * FUNNEL["reach"]["base"] * FUNNEL["v2l"]["base"])
goal["revenue_goal"] = GOAL_PAY * arppu("base")[0]
goal["budget_eur"] = BUDGET_BYN / EUR_BYN
A["goal"] = goal

# ---------- чувствительность (базовый сценарий цифровой воронки, платящие за 4 месяца платной беты) ----------
def payers4(reach=FUNNEL["reach"]["base"], v2l=FUNNEL["v2l"]["base"], l2p=FUNNEL["l2p"]["base"], cov=cover, rel_scale=1.0):
    return rel_m * rel_scale / cov * reach * v2l * l2p * paid_start_months


base_p = payers4()
tornado = []
for name, kw_lo, kw_hi in (
    ("Достижимая доля трафика (0,5 % - 4 %)", dict(reach=FUNNEL["reach"]["cons"]), dict(reach=FUNNEL["reach"]["opt"])),
    ("Конверсия визит - регистрация (2 % - 7 %)", dict(v2l=FUNNEL["v2l"]["cons"]), dict(v2l=FUNNEL["v2l"]["opt"])),
    ("Конверсия регистрация - платящий (3 % - 8 %)", dict(l2p=FUNNEL["l2p"]["cons"]), dict(l2p=FUNNEL["l2p"]["opt"])),
    ("Покрытие рынка найденными сайтами (70 % - 40 %)", dict(cov=0.70), dict(cov=0.40)),
    ("Релевантность трафика (-30 % / +30 %)", dict(rel_scale=0.7), dict(rel_scale=1.3)),
):
    tornado.append((name, payers4(**kw_lo), payers4(**kw_hi)))
A["tornado"] = {"base": base_p, "rows": tornado}
# чувствительность SOM (книга 1) к цене/структуре: выручка на платящего
A["arppu"] = {sc: arppu(sc) for sc in SC}
(HERE / "results" / "analysis.json").write_text(json.dumps(A, ensure_ascii=False, indent=1, default=float), encoding="utf-8")

bad = [c for c in A["checks"] if not c["ok"]]
print("проверок:", len(A["checks"]), "расхождений:", len(bad))
for c in bad:
    print("  ", c)
print("final potential", {k: {a: round(b) for a, b in v.items()} for k, v in fin.items()})
print("goal", {k: ({a: round(b, 2) for a, b in v.items()} if isinstance(v, dict) else round(v, 4)) for k, v in goal.items()})
print("tornado base", round(base_p, 2), [(n, round(a, 1), round(b, 1)) for n, a, b in tornado])
