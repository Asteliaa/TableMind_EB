"""МОДЕЛЬНЫЕ ЗАМЕРЫ MVP (MVP еще не написан): имитация недельных метрик воронки с 01.04.2027 по 30.09.2027.
Параметры взяты из базового сценария ЛР2-3 (визит -> регистрация 4 %, регистрация -> платящий 5 %), шум задан распределениями, seed фиксирован.
Это не измерения: файл показывает, как выглядит отчет о конверсиях и насколько цель (300 регистраций, 50 платящих) чувствительна к разбросу."""
import csv, json
from datetime import date, timedelta
from pathlib import Path
import numpy as np

HERE = Path(__file__).resolve().parent
rng = np.random.default_rng(20261004)
START, END = date(2027, 4, 1), date(2027, 9, 30)
PAY_START = date(2027, 6, 1)          # открытая бета с оплатой
V2L, L2A, A2P = 0.04, 0.60, 0.083     # визит->регистрация, регистрация->активация (первый аудит), активация->платящий (L2P = 0,6 x 0,083 = 5 %)
# каналы: доля визитов, множитель конверсии
CH = {"органический поиск": (0.38, 1.00), "прямые и закладки": (0.20, 1.25), "рефералы и сообщества": (0.17, 0.85), "платные объявления": (0.15, 0.60), "почта и партнеры": (0.10, 1.40)}
rows = []
d = START
week = 0
while d <= END:
    week += 1
    # еженедельный трафик: линейный рост от 250 до 1100 визитов, лог-нормальный шум
    base_visits = 250 + (1100 - 250) * (week - 1) / 25
    for ch, (share, mult) in CH.items():
        v = int(max(0, rng.lognormal(np.log(base_visits * share), 0.18)))
        p_reg = float(np.clip(rng.beta(40 * V2L * mult * 25 / 25 + 1, 40 * (1 - V2L * mult) + 1), 0.002, 0.2))
        reg = int(rng.binomial(v, p_reg))
        act = int(rng.binomial(reg, float(np.clip(rng.normal(L2A, 0.05), 0.2, 0.9))))
        pay = int(rng.binomial(act, float(np.clip(rng.normal(A2P * mult, 0.02), 0.0, 0.3)))) if d >= PAY_START else 0
        rows.append(dict(week=week, week_start=d.isoformat(), channel=ch, visits=v, registrations=reg, activations=act, paying=pay))
    d += timedelta(days=7)
with open(HERE / "mvp_замеры_недели.csv", "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys()), delimiter=";"); w.writeheader(); w.writerows(rows)
months = {}
for r in rows:
    m = r["week_start"][:7]
    a = months.setdefault(m, dict(visits=0, registrations=0, activations=0, paying=0))
    for k in a: a[k] += r[k]
tot = {k: sum(v[k] for v in months.values()) for k in ("visits", "registrations", "activations", "paying")}
chs = {}
for r in rows:
    a = chs.setdefault(r["channel"], dict(visits=0, registrations=0, activations=0, paying=0))
    for k in a: a[k] += r[k]
paid_visits = sum(r["visits"] for r in rows if r["week_start"] >= PAY_START.isoformat())
summ = dict(label="МОДЕЛЬНЫЕ ЗАМЕРЫ, не измерения MVP", seed=20261004, weeks=week, months=months, total=tot, channels=chs,
            conv=dict(v2l=tot["registrations"] / tot["visits"], l2a=tot["activations"] / tot["registrations"],
                      a2p=tot["paying"] / max(1, sum(r["activations"] for r in rows if r["week_start"] >= PAY_START.isoformat())),
                      l2p=tot["paying"] / max(1, sum(r["registrations"] for r in rows if r["week_start"] >= PAY_START.isoformat()))),
            goal=dict(registrations=300, paying=50, reg_ok=tot["registrations"] >= 300, pay_ok=tot["paying"] >= 50))
# разброс итога: 500 повторов тех же параметров
sim = []
for s in range(500):
    r2 = np.random.default_rng(1000 + s)
    reg = pay = 0
    for wk in range(week):
        bv = 250 + 850 * wk / 25
        wd = START + timedelta(days=7 * wk)
        for ch, (share, mult) in CH.items():
            v = int(max(0, r2.lognormal(np.log(bv * share), 0.18)))
            # неопределенность самой конверсии (общая на всю симуляцию): от 2 % до 7 % по Beta
            if wk == 0 and ch == list(CH)[0]:
                k_v2l = float(np.clip(r2.beta(8, 190), 0.01, 0.12)); k_a2p = float(np.clip(r2.beta(4, 46), 0.01, 0.25))
            g = int(r2.binomial(v, k_v2l * mult)); reg += g
            if wd >= PAY_START:
                pay += int(r2.binomial(int(g * 0.6), min(0.9, k_a2p * mult)))
    sim.append((reg, pay))
sim = np.array(sim)
summ["mc"] = dict(n=500, reg_p10=int(np.percentile(sim[:, 0], 10)), reg_p50=int(np.percentile(sim[:, 0], 50)), reg_p90=int(np.percentile(sim[:, 0], 90)),
                  pay_p10=int(np.percentile(sim[:, 1], 10)), pay_p50=int(np.percentile(sim[:, 1], 50)), pay_p90=int(np.percentile(sim[:, 1], 90)),
                  p_reg_goal=float((sim[:, 0] >= 300).mean()), p_pay_goal=float((sim[:, 1] >= 50).mean()))
(HERE / "mvp_замеры_сводка.json").write_text(json.dumps(summ, ensure_ascii=False, indent=1), encoding="utf-8")
print(json.dumps({k: summ[k] for k in ("total", "conv", "goal", "mc")}, ensure_ascii=False))
