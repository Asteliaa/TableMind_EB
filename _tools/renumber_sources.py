"""Перенумерация ссылок на источники по порядку первого упоминания в тексте (требование СТП БГУИР 01-2024).

renumber(text) -> text: ссылки вида [n] в основной части и пронумерованный список «Список использованных источников»
приводятся к порядку первого упоминания; источники, не упомянутые в тексте, остаются в конце."""
import re


def renumber(text: str) -> str:
    head = "## Список использованных источников"
    i = text.find(head)
    if i < 0:
        return text
    j = text.find("\n## ", i + len(head))
    if j < 0:
        j = len(text)
    body_a, srcs, body_b = text[:i], text[i:j], text[j:]
    cite = re.compile(r"(?<!\[)\[(\d+)\](?!\])")
    order = []
    for part in (body_a, body_b):
        for m in cite.finditer(part):
            n = int(m.group(1))
            if n not in order:
                order.append(n)
    entries = {}
    lines = srcs.split("\n")
    cur = None
    for ln in lines[1:]:
        m = re.match(r"^(\d+)\. (.*)$", ln)
        if m:
            cur = int(m.group(1))
            entries[cur] = m.group(2)
    for n in sorted(entries):
        if n not in order:
            order.append(n)
    mapping = {old: new for new, old in enumerate(order, 1)}
    if all(k == v for k, v in mapping.items()):
        return text

    def sub(part):
        return cite.sub(lambda m: "[" + str(mapping.get(int(m.group(1)), int(m.group(1)))) + "]", part)

    new_srcs = [lines[0], ""]
    for old in order:
        if old in entries:
            new_srcs.append(f"{mapping[old]}. {entries[old]}")
            new_srcs.append("")
    return sub(body_a) + "\n".join(new_srcs).rstrip("\n") + "\n" + sub(body_b)
