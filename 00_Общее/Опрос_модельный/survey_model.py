"""МОДЕЛЬНЫЕ ДАННЫЕ опроса (респондентов нет): генерация по заданным параметрам, seed фиксирован.
Результат - не опрос реальных людей. Выход: respondents.csv, kano_responses.csv, summary.json."""
import csv
import json
import sys
from pathlib import Path

import numpy as np

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
sys.path.insert(0, str(ROOT / "ЛР5_Анализ_конкурентов" / "Python"))
from data5 import FEAT  # noqa: E402

SEED = 20261003
rng = np.random.default_rng(SEED)

# сегменты: название, число респондентов, часов в месяц с таблицами (среднее), медиана допустимой цены Pro (EUR/мес.), доля имеющих критичную ошибку за год
SEG = [("Финансовые аналитики и FP&A", 8, 60, 22.0, 0.65),
       ("Бухгалтеры и аудиторы", 7, 50, 14.0, 0.55),
       ("Консультанты по управлению", 8, 50, 26.0, 0.70),
       ("МСП с финансовой функцией", 7, 40, 12.0, 0.60)]
TOOLS = ["только вручную", "Copilot или чат-бот", "специализированный инструмент"]
TOOLP = [0.62, 0.30, 0.08]

rows = []
rid = 0
for name, n, hours, med_pro, perr in SEG:
    for _ in range(n):
        rid += 1
        h = int(np.clip(rng.normal(hours, 12), 10, 140))
        err = int(rng.random() < perr)
        tool = rng.choice(TOOLS, p=TOOLP)
        # четыре вопроса Ван Вестендорпа для подписки Pro (EUR/мес.), порядок цен гарантируется
        base = med_pro * float(np.exp(rng.normal(0, 0.35)))
        too_cheap = round(base * rng.uniform(0.15, 0.35), 1)
        cheap = round(base * rng.uniform(0.5, 0.75), 1)
        expensive = round(base * rng.uniform(1.2, 1.6), 1)
        too_exp = round(base * rng.uniform(1.8, 2.6), 1)
        one_off = round(base * rng.uniform(0.35, 0.7), 1)   # максимум за разовый аудит
        intent = int(np.clip(round(rng.normal(3.2 + (0.5 if err else -0.2), 0.9)), 1, 5))
        pick = "не купил бы"
        if intent >= 3:
            pick = "разовый аудит 9 EUR" if expensive < 19 else ("Pro 19 EUR" if expensive < 29 * 0.95 or rng.random() < 0.5 else "Team 29 EUR")
        rows.append(dict(id=f"M{rid:02d}", segment=name, hours_month=h, critical_error_12m=err, tool=tool,
                         vw_too_cheap=too_cheap, vw_cheap=cheap, vw_expensive=expensive, vw_too_expensive=too_exp,
                         max_one_off=one_off, intent_1_5=intent, choice=pick))

# Ван Вестендорп: пересечения кумулятивных кривых (шаг 0,1 EUR)
grid = np.arange(0, 100.05, 0.1)
tc = np.array([r["vw_too_cheap"] for r in rows]); ch = np.array([r["vw_cheap"] for r in rows])
ex = np.array([r["vw_expensive"] for r in rows]); te = np.array([r["vw_too_expensive"] for r in rows])
n = len(rows)
c_tc = np.array([(tc >= g).mean() for g in grid]); c_ch = np.array([(ch >= g).mean() for g in grid])
c_ex = np.array([(ex <= g).mean() for g in grid]); c_te = np.array([(te <= g).mean() for g in grid])
def cross(a, b):
    i = np.argmin(np.abs(a - b)); return float(grid[i])
PMC = cross(c_tc, c_ex)   # точка предельной дешевизны
PME = cross(c_ch, c_te)   # точка предельной дороговизны
OPP = cross(c_tc, c_te)   # оптимальная цена
IPP = cross(c_ch, c_ex)   # безразличная цена

# Кано: функциональный и дисфункциональный ответы 1..5 (нравится, обязательно, безразлично, мирюсь, не нравится)
KT = {  # строка - функциональный, столбец - дисфункциональный; A привлекательное, O линейное, M обязательное, I безразличное, R обратное, Q сомнительное
    1: "QAAAO", 2: "RIIIM", 3: "RIIIM", 4: "RIIIM", 5: "RRRRQ"}
KNAME = {"A": "Привлекательное", "O": "Линейное", "M": "Обязательное", "I": "Безразличное", "R": "Обратное", "Q": "Сомнительное"}
def kano(f, d):
    return KT[f][d - 1]
# типичные пары ответов для категорий (f, d)
PAIRS = {"Привлекательное": [(1, 3), (1, 4), (1, 2)], "Линейное": [(1, 5), (1, 5), (2, 5)], "Обязательное": [(3, 5), (4, 5), (2, 5)],
         "Безразличное": [(3, 3), (3, 4), (4, 3)]}
AGREE = 0.58   # доля респондентов, чей ответ укладывается в аналитическую категорию (параметр модели)
OTHER = ["Привлекательное", "Линейное", "Обязательное", "Безразличное"]
kano_rows = []
summary_kano = []
for code, name, _task, cat, *_ in FEAT:
    counts = {v: 0 for v in KNAME.values()}
    for r in rows:
        c = cat if rng.random() < AGREE else str(rng.choice([x for x in OTHER if x != cat]))
        f, d = PAIRS[c][int(rng.integers(len(PAIRS[c])))]
        res = KNAME[kano(f, d)]
        counts[res] += 1
        kano_rows.append(dict(id=r["id"], feature=code, functional=f, dysfunctional=d, category=res))
    valid = {k: v for k, v in counts.items() if k in ("Привлекательное", "Линейное", "Обязательное", "Безразличное")}
    mx = max(valid.values()); modal = [k for k, v in valid.items() if v == mx]
    A, O, M, I = (valid[k] for k in ("Привлекательное", "Линейное", "Обязательное", "Безразличное"))
    tot = A + O + M + I
    better = (A + O) / tot if tot else 0; worse = -(O + M) / tot if tot else 0
    summary_kano.append(dict(code=code, name=name, analyst=cat, modal=modal[0] if len(modal) == 1 else " / ".join(modal),
                             counts=valid, better=round(better, 2), worse=round(worse, 2), agree=(cat in modal)))

by_choice = {}
for r in rows:
    by_choice[r["choice"]] = by_choice.get(r["choice"], 0) + 1
summ = dict(label="МОДЕЛЬНЫЕ ДАННЫЕ, не результат опроса респондентов", seed=SEED, n=n, agree_param=AGREE,
            segments={s[0]: s[1] for s in SEG},
            hours_mean=round(float(np.mean([r["hours_month"] for r in rows])), 1),
            critical_error_share=round(float(np.mean([r["critical_error_12m"] for r in rows])), 2),
            tools={t: sum(1 for r in rows if r["tool"] == t) for t in TOOLS},
            vw=dict(PMC=PMC, OPP=OPP, IPP=IPP, PME=PME),
            one_off_median=round(float(np.median([r["max_one_off"] for r in rows])), 1),
            one_off_share_ge_9=round(float(np.mean([r["max_one_off"] >= 9 for r in rows])), 2),
            pro_share_expensive_gt_19=round(float(np.mean([r["vw_expensive"] > 19 for r in rows])), 2),
            intent_mean=round(float(np.mean([r["intent_1_5"] for r in rows])), 2),
            intent_ge4_share=round(float(np.mean([r["intent_1_5"] >= 4 for r in rows])), 2),
            choice=by_choice, kano=summary_kano,
            kano_agree=sum(1 for k in summary_kano if k["agree"]), kano_n=len(summary_kano))
(HERE / "summary.json").write_text(json.dumps(summ, ensure_ascii=False, indent=1), encoding="utf-8")
for fn, data in (("respondents.csv", rows), ("kano_responses.csv", kano_rows)):
    with open(HERE / fn, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(data[0].keys()), delimiter=";"); w.writeheader(); w.writerows(data)
print(json.dumps({k: summ[k] for k in ("n", "vw", "one_off_median", "one_off_share_ge_9", "intent_mean", "choice", "kano_agree")}, ensure_ascii=False))
for k in summary_kano: print(k["code"], k["analyst"], "->", k["modal"], k["counts"], k["better"], k["worse"])
