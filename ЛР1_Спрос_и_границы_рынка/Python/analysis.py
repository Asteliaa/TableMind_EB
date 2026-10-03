"""Расчеты ЛР1: независимый пересчет шаблона XLT на Python, сверка с Excel, показатели GT и Вордстата, фазы цикла.

Результат: Python/results/analysis.json (читается построителем отчета и рисунков).
"""
import json
import math
import statistics as st
from pathlib import Path

from data_prep import CLUSTERS, MONTHS, YW, YW_MONTHS, cluster_series, despike, yw
from gt_monthly import GT_DIR, monthly, weekly

HERE = Path(__file__).resolve().parent
RES = HERE / "results"
MONTH_RU = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
SET = {"wg": 0.5, "wy": 0.5, "strong": 0.3, "weak": 0.15, "down": -0.15, "up": 0.15, "low": 0.9, "high": 1.1}


def xlt(gt, ywv):
    """Повтор расчета шаблона XLT (листы Расчет, Сезонность, Оценка_тренда, Деловые_циклы)."""
    ymax = max(v for v in ywv if v is not None) if any(v is not None for v in ywv) else None
    f = [None if v is None else (v / ymax * 100 if ymax else None) for v in ywv]
    g = []
    for a, b in zip(gt, f):
        if a is None and b is None:
            g.append(None)
        elif a is not None and b is not None:
            g.append(a * SET["wg"] + b * SET["wy"])
        else:
            g.append(a if a is not None else b)
    n = len(g)
    h = [None] * n
    for i in range(n):
        if i >= 11 and g[i] is not None:
            h[i] = sum(g[i - 11:i + 1]) / 12
    yt = [None if (h[i] in (None, 0) or g[i] is None) else g[i] / h[i] for i in range(n)]
    mon = [int(m[5:]) for m in MONTHS]
    c = []
    for mm in range(1, 13):
        vals = [yt[i] for i in range(n) if mon[i] == mm and yt[i] is not None and yt[i] > 0]
        c.append(sum(vals) / len(vals) if vals else None)
    avg = sum(x for x in c if x is not None) / len([x for x in c if x is not None])
    s = [None if x is None else x / avg for x in c]
    k = []
    for i in range(n):
        si = s[mon[i] - 1]
        if g[i] is None or h[i] in (None, 0) or not si:
            k.append(None)
        else:
            k.append(g[i] / (h[i] * si))
    first = sum(g[:12]) / 12
    last = sum(g[-12:]) / 12
    ch = (last - first) / first
    kv = [x for x in k if x is not None]
    amp = max(s) - min(s)
    slope = _slope(g)
    return {
        "G": g, "H": h, "K": k, "S": s, "first12": first, "last12": last, "change": ch, "slope": slope,
        "trend": "растущий тренд" if ch >= SET["up"] else ("снижающийся тренд" if ch <= SET["down"] else "стабильный тренд"),
        "season_amp": amp,
        "season": "выраженная сезонность" if amp >= SET["strong"] else ("умеренная сезонность" if amp >= SET["weak"] else "слабая сезонность"),
        "peak": MONTH_RU[s.index(max(s))], "trough": MONTH_RU[s.index(min(s))],
        "k_last": k[-1], "k_min": min(kv), "k_max": max(kv),
        "phase": "ниже тренда" if k[-1] < SET["low"] else ("выше тренда" if k[-1] > SET["high"] else "норма"),
    }


def _slope(y):
    pts = [(i + 1, v) for i, v in enumerate(y) if v is not None]
    xm = sum(p[0] for p in pts) / len(pts)
    ym = sum(p[1] for p in pts) / len(pts)
    return sum((x - xm) * (v - ym) for x, v in pts) / sum((x - xm) ** 2 for x, _ in pts)


def mean(x):
    return sum(x) / len(x)


def ma(series, w):
    return [None if i < w - 1 else mean(series[i - w + 1:i + 1]) for i in range(len(series))]


def season_word(series):
    """Сезонный индекс по Word-методике: среднее месяца / среднее всех месяцев x 100 (по месяцам календаря)."""
    mon = [int(m[5:]) for m in MONTHS]
    tot = mean(series)
    return [mean([series[i] for i in range(len(series)) if mon[i] == m]) / tot * 100 for m in range(1, 13)]


def phase_rule(des):
    """Фаза цикла по 12-месячной средней сезонно очищенного ряда (рабочее правило отчета)."""
    m = [x for x in ma(des, 12) if x is not None]
    now = m[-1]
    g6 = now / m[-7] - 1
    g12 = now / m[-13] - 1
    share = now / max(m)
    if share < 0.10 and g6 > 0:
        name = "зарождение интереса"
    elif share >= 0.95 and g6 <= 0.05:
        name = "пик"
    elif g6 > 0.05 and g12 > 0.05:
        name = "рост"
    elif g12 > 0.05 and -0.05 <= g6 <= 0.05:
        name = "замедление"
    elif g6 < -0.05:
        name = "спад"
    elif g6 > 0.05 and g12 <= 0.05:
        name = "восстановление"
    else:
        name = "стабильность"
    return {"phase": name, "g6": g6, "g12": g12, "share": share, "ma12_now": now}


def corr(a, b):
    ma_, mb = mean(a), mean(b)
    num = sum((x - ma_) * (y - mb) for x, y in zip(a, b))
    den = math.sqrt(sum((x - ma_) ** 2 for x in a) * sum((y - mb) ** 2 for y in b))
    return num / den if den else float("nan")


def main():
    out = {"months": MONTHS}
    # --- 1. XLT: Python vs Excel
    books = {"A_проверка_таблиц": ("A", "full"), "A_только_GT": ("A", "gt"), "A_только_YW": ("A", "yw"), "B_AI_для_Excel": ("B", "full"), "C_финмодели": ("C", "full")}
    out["xlt"] = {}
    cmp = []
    for code, (cl, var) in books.items():
        g, y = cluster_series(cl)
        if cl == "C":
            y, _ = despike(y)
        if var == "gt":
            y = [None] * 60
        if var == "yw":
            g = [None] * 60
        r = xlt(g, y)
        ex = json.loads((RES / f"xlt_{code}.json").read_text(encoding="utf-8"))
        # Excel: calc = [period, G, H, I, J, K, L]
        exK = [row[5] if isinstance(row[5], (int, float)) else None for row in ex["calc"]]
        exG = [row[1] if isinstance(row[1], (int, float)) else None for row in ex["calc"]]
        dG = max(abs(a - b) for a, b in zip(r["G"], exG) if a is not None and b is not None)
        dK = max(abs(a - b) for a, b in zip(r["K"], exK) if a is not None and b is not None)
        exS = [x[3] for x in ex["season"]]
        dS = max(abs(a - b) for a, b in zip(r["S"], exS))
        same = (r["trend"] == ex["panel"]["Классификация тренда"] and r["season"] == ex["panel"]["Сила сезонности"]
                and r["peak"] == ex["panel"]["Месяц максимума"] and r["trough"] == ex["panel"]["Месяц минимума"]
                and r["phase"] == ex["panel"]["Текущая фаза"])
        cmp.append({"book": code, "dG": dG, "dK": dK, "dS": dS, "dChange": abs(r["change"] - ex["panel"]["Изменение последних 12 мес. к первым 12 мес."]), "labels_equal": same})
        out["xlt"][code] = {k: v for k, v in r.items() if k not in ("G", "H", "K")}
        out["xlt"][code].update({"G": r["G"], "K": r["K"], "S_list": r["S"]})
    out["xlt_vs_excel"] = cmp

    # --- 2. Word-метод по кластерам и источникам
    out["word"] = {}
    for cl in "ABC":
        g, y = cluster_series(cl)
        yb = cluster_series(cl, 149)[1]
        if cl == "C":
            y, _ = despike(y)
        res = {}
        for name, s in (("GT", g), ("YW_RU", y), ("YW_BY", yb)):
            if sum(s) == 0:
                continue
            n = len(s)
            first, prev, last = mean(s[:12]), mean(s[-24:-12]), mean(s[-12:])
            si = season_word(s)
            m12 = ma(s, 12)
            m3 = ma(s, 3)
            mon = [int(m[5:]) for m in MONTHS]
            sidx = [si[mm - 1] / 100 for mm in mon]
            des = [v / x if x else 0 for v, x in zip(s, sidx)]
            # повторяемость пика: в скольких полных годах (10.2022-09.2026 по календарным годам 2022-2025) месяц пика входит в тройку лучших месяцев года
            pk = si.index(max(si)) + 1
            rep = 0
            for yr in (2022, 2023, 2024, 2025):
                vals = [(s[i], mon[i]) for i, m in enumerate(MONTHS) if m.startswith(str(yr))]
                top3 = [mm for _, mm in sorted(vals, reverse=True)[:3]]
                rep += pk in top3
            ph = phase_rule(des)
            share_up = sum(1 for i in range(1, n) if s[i] > s[i - 1]) / (n - 1)
            res[name] = {
                "mean": mean(s), "max": max(s), "first12": first, "prev12": prev, "last12": last,
                "chg_prev": (last / prev - 1) if prev else None, "chg_first": (last / first - 1) if first else None,
                "ma3_last": m3[-1], "ma12_last": m12[-1], "ma12_max": max(x for x in m12 if x is not None),
                "season_idx": si, "season_peak": MONTH_RU[pk - 1], "season_trough": MONTH_RU[si.index(min(si))],
                "season_amp": max(si) - min(si), "peak_repeat": rep, "share_up": share_up,
                "cv": st.pstdev(s) / mean(s) if mean(s) else None, "phase": ph,
                "series": s, "ma12": m12, "ma3": m3, "des": des,
            }
        out["word"][cl] = res
    # --- 3. сопоставление GT и Вордстата по кластерам
    out["compare"] = {}
    for cl in "ABC":
        g, y = cluster_series(cl)
        if cl == "C":
            y, _ = despike(y)
        w = out["word"][cl]
        out["compare"][cl] = {
            "corr_level": corr(g, y),
            "corr_ma12": corr([x for x in ma(g, 12) if x is not None], [x for x in ma(y, 12) if x is not None]),
            "gt_chg": w["GT"]["chg_prev"], "yw_chg": w["YW_RU"]["chg_prev"],
            "gt_first": w["GT"]["chg_first"], "yw_first": w["YW_RU"]["chg_first"],
        }
    # --- 4. показатели GT по запросам (недельные ряды 5 лет, без последней неполной недели)
    out["gt_queries"] = []
    pk = {"W_P1_5y": "P1", "W_P2_5y": "P2", "W_P3_5y": "P3", "W_P4_5y": "P4", "W_P5_5y": "P5"}
    for stem, tag in pk.items():
        meta, ser = weekly(GT_DIR / f"{stem}.txt")
        for q, pts in ser.items():
            v = [x for _, x in pts][:-1]
            last12 = v[-52:]
            prev12 = v[-104:-52]
            l3, p3 = v[-13:], v[-26:-13]
            out["gt_queries"].append({
                "pack": tag, "query": q, "avg": mean(v), "max": max(v), "nonzero": sum(1 for x in v if x > 0) / len(v),
                "last52": mean(last12), "prev52": mean(prev12), "chg52": (mean(last12) / mean(prev12) - 1) if mean(prev12) else None,
                "chg13": (mean(l3) / mean(p3) - 1) if mean(p3) else None,
                "cv": st.pstdev(v) / mean(v) if mean(v) else None,
            })
    for stem in ("S_excel_formula_checker_5y", "S_spreadsheet_audit_5y", "S_excel_audit_5y", "S_financial_model_audit_5y", "S_excel_copilot_5y", "S_perfectxl_5y"):
        meta, ser = weekly(GT_DIR / f"{stem}.txt")
        (q, pts), = ser.items()
        v = [x for _, x in pts][:-1]
        last12, prev12 = v[-52:], v[-104:-52]
        out["gt_queries"].append({
            "pack": "single", "query": q, "avg": mean(v), "max": max(v), "nonzero": sum(1 for x in v if x > 0) / len(v),
            "last52": mean(last12), "prev52": mean(prev12), "chg52": (mean(last12) / mean(prev12) - 1) if mean(prev12) else None,
            "chg13": None, "cv": st.pstdev(v) / mean(v) if mean(v) else None,
        })
    # --- 5. показатели Вордстата по запросам
    out["yw_queries"] = []
    queries = sorted({s["query"] for s in YW["series"]})
    for q in queries:
        row = {"query": q}
        for reg in (225, 149):
            try:
                s = yw(q, reg)
            except KeyError:
                continue
            v = [s[m] for m in YW_MONTHS]
            last12, prev12 = v[-12:], v[-24:-12]
            row[str(reg)] = {
                "sum_all": sum(v), "last12": sum(last12), "prev12": sum(prev12),
                "chg": (sum(last12) / sum(prev12) - 1) if sum(prev12) else None,
                "avg_month": mean(last12), "max": max(v), "max_month": YW_MONTHS[v.index(max(v))],
                "zero_months": sum(1 for x in v if x == 0), "sep26": v[-1],
            }
        out["yw_queries"].append(row)
    # --- 6. структура: доля Беларуси в России (последние 12 месяцев)
    out["by_share"] = []
    for q in queries:
        try:
            r, b = yw(q, 225), yw(q, 149)
        except KeyError:
            continue
        sr, sb = sum(r[m] for m in YW_MONTHS[-12:]), sum(b[m] for m in YW_MONTHS[-12:])
        out["by_share"].append({"query": q, "ru": sr, "by": sb, "share": sb / sr if sr else None})
    (RES / "analysis.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print("saved analysis.json")
    for c in out["xlt_vs_excel"]:
        print(c)
    for cl in "ABC":
        for name, d in out["word"][cl].items():
            print(cl, name, "chg_prev", None if d["chg_prev"] is None else round(d["chg_prev"], 3), "peak", d["season_peak"], "trough", d["season_trough"], "amp %.0f" % d["season_amp"], "rep", d["peak_repeat"], "phase", d["phase"]["phase"], round(d["phase"]["g6"], 3), round(d["phase"]["g12"], 3), round(d["phase"]["share"], 2))
        print(cl, "compare", {k: (None if v is None else round(v, 3)) for k, v in out["compare"][cl].items()})


if __name__ == "__main__":
    main()
