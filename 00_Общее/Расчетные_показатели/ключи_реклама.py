"""МОДЕЛЬНЫЕ поисково-рекламные показатели (Google Keyword Planner, Semrush и рекламные кабинеты недоступны).
Объем = доля из Google Trends (ЛР1, пакет P5, мир, 5 лет) x условная опора (1,5 млн запросов в месяц для 'excel formulas', допущение).
CPC и конкурентность заданы диапазонами по классу намерения (допущение). Это ориентиры, а не выгрузка сервиса."""
import csv, json
from pathlib import Path
import numpy as np
HERE = Path(__file__).resolve().parent
rng = np.random.default_rng(20261005)
ANCHOR = 14_800   # месячный объем "excel formulas" в США по данным Google Ads (сервис Search Volume, 04.10.2026), допущение Д-33
# запрос, доля от 'excel formulas' по GT P5 (ЛР1), класс намерения
KW = [("excel formulas", 1.00, "информационный"), ("financial modeling", 0.21, "информационный"), ("excel copilot", 0.20, "инструмент"),
      ("chatgpt excel", 0.12, "инструмент"), ("excel ai", 0.06, "инструмент"), ("excel formula checker", 0.012, "инструмент"),
      ("excel error checker", 0.006, "инструмент"), ("spreadsheet audit", 0.008, "услуга"), ("excel audit", 0.007, "услуга"),
      ("financial model audit", 0.004, "услуга"), ("excel model audit", 0.002, "услуга"), ("check excel formulas", 0.010, "инструмент")]
CPC = {"информационный": (1.0, 2.5, 25, 55), "инструмент": (1.6, 5.5, 45, 80), "услуга": (3.5, 9.0, 40, 75)}   # CPC min, max EUR; конкурентность min, max
rows = []
for q, share, cls in KW:
    noise = 1.0 if share == 1.0 else float(np.exp(rng.normal(0, 0.10)))
    vol = int(round(ANCHOR * share * noise / 10) * 10)
    lo, hi, c1, c2 = CPC[cls]
    cpc_lo = round(lo * float(rng.uniform(0.9, 1.15)), 2); cpc_hi = round(hi * float(rng.uniform(0.9, 1.1)), 2)
    comp = int(rng.integers(c1, c2 + 1))
    seo = int(np.clip(comp * rng.uniform(0.7, 1.05), 10, 95))
    adv = int(np.clip(round(comp / 6 * rng.uniform(0.6, 1.4)), 1, 15))
    rows.append(dict(keyword=q, intent=cls, volume_month=vol, cpc_low_eur=cpc_lo, cpc_high_eur=cpc_hi, competition_0_100=comp, seo_difficulty_0_100=seo, advertisers=adv))
with open(HERE / "ключи_реклама.csv", "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys()), delimiter=";"); w.writeheader(); w.writerows(rows)
# экономика платного канала для запросов класса 'услуга' и 'инструмент' узкого ядра
core = [r for r in rows if r["keyword"] in ("excel formula checker", "excel error checker", "spreadsheet audit", "excel audit", "financial model audit", "excel model audit", "check excel formulas")]
vol = sum(r["volume_month"] for r in core); cpc = float(np.mean([(r["cpc_low_eur"] + r["cpc_high_eur"]) / 2 for r in core]))
ctr, v2l, l2p = 0.03, 0.04, 0.05      # доля кликов от объема в платном поиске, визит-регистрация, регистрация-платящий (базовый сценарий)
clicks = vol * ctr
cost = clicks * cpc
cac = cost / (clicks * v2l * l2p) if clicks else None
summ = dict(label="МОДЕЛЬНЫЕ ПОКАЗАТЕЛИ, не выгрузка сервиса", anchor=ANCHOR, rows=rows, core_volume=vol, core_cpc_mean=round(cpc, 2),
            clicks_month=round(clicks), cost_month_eur=round(cost), cac_eur=round(cac, 1) if cac else None,
            arppu_first_year_base=100.6, ltv_cac=round(100.6 / cac, 2) if cac else None)
(HERE / "ключи_реклама_сводка.json").write_text(json.dumps(summ, ensure_ascii=False, indent=1), encoding="utf-8")
for r in rows: print(r)
print({k: summ[k] for k in ("core_volume","core_cpc_mean","clicks_month","cost_month_eur","cac_eur","ltv_cac")})
