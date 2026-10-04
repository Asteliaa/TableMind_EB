"""Общие функции построителя отчета ЛР1: форматирование чисел, таблицы с автонумерацией, рисунки, ссылки."""
import json
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
A = json.loads((HERE / "results" / "analysis.json").read_text(encoding="utf-8"))
X = json.loads((HERE / "results" / "excel_results.json").read_text(encoding="utf-8"))

AUTO = []
INTRO = {}
OUT = []          # строки итогового md
TABLES = {}       # ключ -> номер
FIGS = {}         # ключ -> номер
_t = [0]
_f = [0]


def num(x, d=0):
    if x is None:
        return "н/д"
    s = f"{x:,.{d}f}".replace(",", " ").replace(".", ",")
    return s


def pct(x, d=0, sign=False):
    if x is None:
        return "н/д"
    v = x * 100
    s = f"{v:+.{d}f}" if sign else f"{v:.{d}f}"
    return s.replace(".", ",") + " %"


def P(text=""):
    """Абзац."""
    OUT.append(text)
    OUT.append("")


def H(level, text):
    OUT.append("#" * level + " " + text)
    OUT.append("")


def _lower_first(s):
    return s[0].lower() + s[1:] if s and "А" <= s[0] <= "Я" else s


def _dec(c):
    "Десятичная точка в числовой ячейке заменяется запятой."
    return re.sub(r"^(-?\d+)\.(\d+)$", r"\1,\2", c)


def table(key, caption, header, rows, after):
    """Таблица с подписью сверху и обязательным пояснением после; если ссылки на таблицу еще не было, добавляется вводная фраза."""
    if f"[[t:{key}]]" not in "\n".join(OUT):
        AUTO.append(key)
        P(INTRO.get(key) or f"В таблице [[t:{key}]] представлено: {_lower_first(caption)}.")
    _t[0] += 1
    TABLES[key] = _t[0]
    OUT.append(f"Таблица {_t[0]} – {caption}")
    OUT.append("")
    OUT.append("| " + " | ".join(header) + " |")
    OUT.append("|" + "|".join("---" for _ in header) + "|")
    for r in rows:
        cells = [_dec(str(c).replace("|", "/").replace("\n", " ")) for c in r]
        OUT.append("| " + " | ".join(cells) + " |")
    OUT.append("")
    P(after)


def figure(key, path, caption, after):
    """Рисунок: изображение, подпись снизу, пояснение после; ссылка на рисунок добавляется, если ее не было."""
    if f"[[f:{key}]]" not in "\n".join(OUT):
        AUTO.append("fig:" + key)
        P(INTRO.get("fig:" + key) or f"На рисунке [[f:{key}]] показано: {_lower_first(caption)}.")
    _f[0] += 1
    FIGS[key] = _f[0]
    OUT.append(f"![]({path})")
    OUT.append("")
    OUT.append(f"Рисунок {_f[0]} – {caption}")
    OUT.append("")
    P(after)


def resolve(text):
    """Подстановка [[t:key]] / [[f:key]] в номера таблиц и рисунков."""
    def rt(m):
        return str(TABLES[m.group(1)])

    def rf(m):
        return str(FIGS[m.group(1)])
    text = re.sub(r"\[\[t:([\w]+)\]\]", rt, text)
    text = re.sub(r"\[\[f:([\w]+)\]\]", rf, text)
    return text


