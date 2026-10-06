"""Замена двоеточий внутри предложений отчета (требование заказчика: предложения через двоеточие не пишутся).

apply(text, manual=()) -> text. Не трогаются: заголовки, блоки кода, список использованных источников,
время и отношения вида 10:30, адреса. В подписях таблиц и рисунков двоеточие заменяется запятой, в остальных
местах - тире с пробелами (дефис, как принято в документе). Пары (было, стало) из manual применяются первыми
к строке целиком, чтобы вручную переписать сложные предложения.
"""
import re

CAP = re.compile(r"^(Таблица|Рисунок) [\dА-Я.]+ [–-] ")
COLON = re.compile(r"\s?(?<![\d/])(?<!http)(?<!https):(?!//)(?!\d)\s")


def apply(text: str, manual=()) -> str:
    out = []
    in_fence = False
    in_src = False
    for line in text.split("\n"):
        if line.startswith("```"):
            in_fence = not in_fence
        if line.startswith("## "):
            in_src = line.startswith("## Список использованных источников")
        if in_fence or in_src or line.startswith("#") or line.startswith("|---"):
            out.append(line)
            continue
        for old, new in manual:
            if old in line:
                line = line.replace(old, new)
        if CAP.match(line):
            line = line.replace(": ", ", ", 1)
        line = COLON.sub(" - ", line)
        out.append(line)
    return "\n".join(out)
