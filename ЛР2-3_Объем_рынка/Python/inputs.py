"""Исходные данные и допущения ЛР2-3 (единый источник для Excel, независимого пересчета и отчета).

Валюта расчета - EUR, год. Курсы НБРБ на 03.10.2026: 1 EUR = 3,3860 BYN, 1 USD = 3,0051 BYN.
Любое число здесь имеет метку источника: 'факт' (открытый источник), 'проект' (ПЗ1), 'допущение' (журнал допущений Д-09...).
"""
import csv
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SNAP_DATE = date(2026, 10, 3)
SNAP_SERIAL = (SNAP_DATE - date(1899, 12, 30)).days

EUR_BYN = 3.3860
USD_BYN = 3.0051
EUR_USD = EUR_BYN / USD_BYN

# ---- Тарифы TableMind (ПЗ1) ----
PRICE = {"audit": 9.0, "pro": 19.0, "team": 29.0}

# ---- Структура платящих и частота платежей по сценариям (допущение Д-09) ----
# доли платящих: разовый аудит / Pro / Team; платежей в год на клиента (разовый - покупок; Pro и Team - оплаченных месяцев)
MIX = {
    "cons": dict(share=(0.60, 0.37, 0.03), events=(1.2, 6.0, 6.0)),
    "base": dict(share=(0.45, 0.45, 0.10), events=(1.5, 8.0, 9.0)),
    "opt":  dict(share=(0.35, 0.50, 0.15), events=(2.0, 10.0, 11.0)),
}


def arppu(sc):
    """Выручка первого года на одного платящего и число платежей в год, средний платеж."""
    m = MIX[sc]
    prices = (PRICE["audit"], PRICE["pro"], PRICE["team"])
    rev = sum(s * e * p for s, e, p in zip(m["share"], m["events"], prices))
    ev = sum(s * e for s, e in zip(m["share"], m["events"]))
    return rev, ev, rev / ev


SC = ["cons", "base", "opt"]
SC_RU = {"cons": "Осторожный", "base": "Базовый", "opt": "Оптимистичный"}

# ---- Similarweb: снимок 03.10.2026 (06-08.2026, мир) ----
EU27 = {"AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LV", "LT", "LU", "MT",
        "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE"}
TARGET = EU27 | {"US", "GB", "CA", "AU"}  # англоязычные страны и ЕС (Р-1); IE входит в EU27

# домен -> (название, тип игрока, сегмент, доля релевантного трафика, коэффициент качества, основание)
SITES = {
    "perfectxl.com":           ("PerfectXL", "прямой", "аудит таблиц", 0.80, 0.85, "специализированный аудит Excel"),
    "operis.com":              ("Operis", "прямой", "аудит финансовых моделей", 0.50, 0.80, "консалтинг и аудит моделей, часть трафика прочие услуги"),
    "arixcel.com":             ("Arixcel", "прямой", "аудит таблиц", 0.80, 0.80, "надстройка для проверки моделей"),
    "spreadsheetdetective.com": ("Spreadsheet Detective", "прямой", "аудит таблиц", 0.80, 0.60, "малый сайт, данные Similarweb ненадежны"),
    "cimcon.com":              ("CIMCON Software", "частичный", "риски пользовательских таблиц", 0.30, 0.60, "малый сайт, широкий продуктовый ряд"),
    "shortcut.ai":             ("Shortcut", "частичный", "AI-агент для Excel", 0.35, 0.85, "AI-агент создает и правит таблицы"),
    "rows.com":                ("Rows", "заменитель", "AI-таблицы", 0.15, 0.85, "альтернативная электронная таблица"),
    "numerous.ai":             ("Numerous.ai", "заменитель", "AI-помощник в таблицах", 0.30, 0.85, "генерация формул, не проверка"),
    "formulabot.com":          ("Formula Bot", "заменитель", "AI-помощник в таблицах", 0.30, 0.80, "генерация формул, трафик из развивающихся стран"),
    "ajelix.com":              ("Ajelix", "заменитель", "AI-помощник в таблицах", 0.25, 0.85, "генерация формул и аналитика"),
    "datarails.com":           ("Datarails", "заменитель", "FP&A-платформа для Excel", 0.15, 0.90, "консолидация и планирование над Excel"),
    "julius.ai":               ("Julius AI", "заменитель", "AI-анализ данных", 0.10, 0.85, "общий анализ данных, пересечение малое"),
}


def load_sw():
    rows = []
    with open(ROOT / "Материалы_собранные" / "similarweb_снимок_2026-10-03.csv", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f, delimiter=";"):
            r = dict(r)
            r["visits_month"] = float(r["visits_month_avg"])
            ch = {}
            for k in ("direct", "org_search", "paid_search", "referrals", "display", "org_social", "paid_social", "gen_ai", "email", "affiliates"):
                ch[k] = float(r[k]) / 100 if r[k] not in ("", None) else 0.0
            r["ch"] = ch
            top = []
            for tok in r["top5_countries_share_pct"].split("|"):
                c, v = tok.split()
                top.append((c, float(v) / 100))
            r["top"] = top
            r["geo_share"] = sum(v for c, v in top if c in TARGET)   # нижняя оценка: учитываются только топ-5 стран
            name, typ, seg, rel, qual, why = SITES[r["domain"]]
            r.update(name=name, type=typ, segment=seg, rel=rel, qual=qual, why=why)
            rows.append(r)
    return rows


# ---- Потенциальные клиенты (допущения Д-10...) ----
# масштаб по занятости относительно США: Великобритания + Канада + Австралия + Ирландия (71,5 млн занятых) ~0,45; ЕС-27 ~1,3
SCALE_ANGLO = 0.45
SCALE_EU = 1.30
SCALE_ALL = 1 + SCALE_ANGLO + SCALE_EU

US_BLS = {"fin_analysts": 443_100, "accountants": 1_595_200, "mgmt_analysts": 1_077_100}
US_EMPLOYER_FIRMS = 6_395_635
UK_EMPLOYER_FIRMS = 1_420_000
EU_SME = 34_000_000
EU_EMPLOYER_SHARE = 0.25        # допущение: доля МСП ЕС с наемными работниками
CA_AU_IE_EMPLOYER_FIRMS = 1_100_000  # допущение: оценка по занятости

SEGMENTS = [
    # name, type, source, total, border, need, digital, payable, arppu(EUR/yr), purchases, share, reliability, comment
    dict(name="Финансовые аналитики и FP&A", total=round(US_BLS["fin_analysts"] * SCALE_ALL), border=0.70, need=0.30, digital=0.60,
         payable=0.35, check=120.0, freq=1.0, share=0.0015, rel="Средняя",
         src="BLS 2025 (США) и масштаб по занятости", note="Модели строят и проверяют сами; платит чаще работодатель (Pro или Team)."),
    dict(name="Бухгалтеры и аудиторы", total=round(US_BLS["accountants"] * SCALE_ALL), border=0.35, need=0.20, digital=0.60,
         payable=0.30, check=70.0, freq=1.0, share=0.0010, rel="Средняя",
         src="BLS 2025 (США) и масштаб по занятости", note="В границы входят только работающие со сложными таблицами; часто разовый аудит."),
    dict(name="Консультанты по управлению", total=round(US_BLS["mgmt_analysts"] * SCALE_ALL), border=0.40, need=0.30, digital=0.70,
         payable=0.40, check=180.0, freq=1.0, share=0.0015, rel="Низкая",
         src="BLS 2025 (США) и масштаб по занятости", note="Проверка чужих моделей клиентов; высокая готовность платить."),
    dict(name="МСП с финансовой функцией", total=US_EMPLOYER_FIRMS + UK_EMPLOYER_FIRMS + CA_AU_IE_EMPLOYER_FIRMS + round(EU_SME * EU_EMPLOYER_SHARE),
         border=0.03, need=0.30, digital=0.60, payable=0.30, check=240.0, freq=1.0, share=0.0005, rel="Низкая",
         src="SBA 2025, UK BPE 2025, ЕК 2025/2026 и допущения", note="Покупка Team на 2-3 места; основатели и операционные менеджеры."),
]

# сценарные множители листа «Сценарии» методики потенциальных клиентов (шаблон): потребность, онлайн, платежеспособность, чек, частота, доля
PC_SCEN = {
    "cons": dict(need=0.75, online=0.80, pay=0.75, check=0.80, freq=0.80, share=0.0005),
    "base": dict(need=1.0, online=1.0, pay=1.0, check=1.0, freq=1.0, share=0.0013),
    "opt":  dict(need=1.2, online=1.15, pay=1.2, check=1.15, freq=1.1, share=0.0030),
}

# ---- Сверху вниз ----
TD = {
    "orgs": dict(cons=15_000_000, base=17_400_000, opt=20_000_000),
    "prod": dict(cons=0.02, base=0.03, opt=0.05),
    "geo": dict(cons=1.0, base=1.0, opt=1.0),
    "digital": dict(cons=0.55, base=0.65, opt=0.75),
    "need": dict(cons=0.20, base=0.30, opt=0.40),
    "check": 240.0,
}

# ---- Воронка Similarweb -> продажи ----
FUNNEL = {
    "reach": dict(cons=0.005, base=0.015, opt=0.04),      # достижимая доля релевантного трафика сегмента
    "v2l": dict(cons=0.02, base=0.04, opt=0.07),          # визит -> регистрация Free
    "l2p": dict(cons=0.03, base=0.05, opt=0.08),          # регистрация -> платящий
    "cover": 0.55,                                         # доля сегмента, видимая через найденные сайты
}

# ---- Платежеспособность: сегмент, бюджет/мес., цена/мес. (EUR) ----
PAY = [
    dict(seg="Финансовые аналитики и FP&A", budget=45.0, price=19.0, share_budget=0.35, annual=120.0,
         econ=dict(rev=0, margin=0, growth=0), hours=("часов на проверку модели в месяц", 6), note="Платит работодатель; индекс бюджета к Pro."),
    dict(seg="Бухгалтеры и аудиторы", budget=25.0, price=19.0, share_budget=0.30, annual=70.0,
         econ=None, hours=("часов", 4), note="Часто собственный бюджет или бюджет фирмы; разовый аудит дешевле."),
    dict(seg="Консультанты по управлению", budget=80.0, price=19.0, share_budget=0.40, annual=180.0,
         econ=None, hours=("часов", 8), note="Стоимость включается в проект клиента."),
    dict(seg="МСП с финансовой функцией", budget=90.0, price=58.0, share_budget=0.30, annual=240.0,
         econ=None, hours=("часов", 10), note="Team на 2 места по 29 EUR."),
]

# цена часа специалиста (допущение) и экономия времени: проверка вручную 3 ч vs аудит TableMind 0,5 ч
HOURLY = {"fin_analysts": 55.0, "accountants": 40.0, "consultants": 85.0, "sme": 45.0}  # EUR/час
SAVED_HOURS_PER_MONTH = {"fin_analysts": 4.0, "accountants": 2.5, "consultants": 5.0, "sme": 3.0}

# ---- Цели проекта ----
GOAL_REG = 300
GOAL_PAY = 50
BUDGET_BYN = 517_000
