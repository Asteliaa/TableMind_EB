"""Приведение денежных сумм отчета к одной валюте (EUR).

Суммы в USD и GBP пересчитываются по официальным курсам НБРБ на 03.10.2026 (EUR 3,3860 BYN, USD 3,0051 BYN, GBP 3,9663 BYN),
упоминания BYN убираются (бюджет проекта приводится в EUR). apply(text) -> text.
"""
import re

EUR_BYN, USD_BYN, GBP_BYN = 3.3860, 3.0051, 3.9663
K = {"USD": USD_BYN / EUR_BYN, "GBP": GBP_BYN / EUR_BYN, "BYN": 1 / EUR_BYN}
NUM = r"\d{1,3}(?:[   ]\d{3})+(?:,\d+)?|\d+(?:,\d+)?"
PAT = re.compile(rf"(?P<a>{NUM})(?:(?P<sep>-| до )(?P<b>{NUM}))?(?P<t> (?:тыс\.|млн|млрд))? (?P<cur>USD|GBP|BYN)")


def _val(s):
    return float(re.sub(r"[   ]", "", s).replace(",", "."))


def fmt(v, thousands=False):
    if thousands:
        s = f"{v:.1f}" if v < 10 else f"{v:.0f}"
    elif v < 10:
        s = f"{v:.1f}"
    else:
        s = f"{v:,.0f}".replace(",", " ")
    return s.replace(".", ",")


def _conv(m):
    k = K[m.group("cur")]
    th = bool(m.group("t")) and "тыс" in m.group("t") or bool(m.group("t"))
    a = fmt(_val(m.group("a")) * k, th)
    out = a
    if m.group("b"):
        out += m.group("sep") + fmt(_val(m.group("b")) * k, th)
    return out + (m.group("t") or "") + " у.е."


MANUAL = [
    ("Все денежные значения приведены в EUR за год, если не сказано иное.", "Все денежные значения приведены в условных единицах (у.е.) за год, если не сказано иное."),
    ("Курсы НБРБ на 03.10.2026: 1 EUR = 3,3860 BYN, 1 USD = 3,0051 BYN [3].", "Суммы в других валютах пересчитаны в у.е. по официальным курсам НБРБ на 03.10.2026 [3]."),
    ("с его бюджетом (517 000 BYN)", "с его бюджетом (152 688 у.е.)"),
    ("при бюджете 517 000 BYN, то есть 152 688 EUR по курсу НБРБ [3], [4]", "при бюджете 152 688 у.е. (пересчет по курсу НБРБ [3], [4])"),
    ("валюта BYN и «руб.», демонстрационные данные по Беларуси и маркетингу", "валюты шаблонов и демонстрационные данные по Беларуси и маркетингу"),
    ("порядка нескольких долларов", "порядка нескольких у.е."),
]


def apply(text: str) -> str:
    for a, b in MANUAL:
        text = text.replace(a, b)
    text = PAT.sub(_conv, text)
    text = re.sub(r"\bEUR\b", "у.е.", text)
    return re.sub(r"\bевро\b", "у.е.", text)
