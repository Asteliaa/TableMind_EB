"""Недельные ряды Google Trends -> месячные (простое среднее недель, неделя относится к месяцу ее начала).

Крайние неполные месяцы исключаются: первый (ряд начинается 26.09.2021) и последняя неполная неделя ряда.
Результат: словарь {запрос: {'YYYY-MM': значение}} для каждого файла выгрузки.
"""
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

from gt_decode import load

GT_DIR = Path(__file__).resolve().parent.parent / "Материалы_собранные" / "GT"


def weekly(path):
    meta, kws, rows = load(path)
    start = datetime.fromtimestamp(int(meta["start_epoch"]), tz=timezone.utc)
    step = int(meta["step_sec"])
    out = {k: [] for k in kws}
    for i, row in enumerate(rows):
        d = start + timedelta(seconds=step * i)
        for k, v in zip(kws, row):
            out[k].append((d, v))
    return meta, out


def monthly(path, drop_last_week=True):
    meta, ser = weekly(path)
    res = {}
    for k, pts in ser.items():
        if drop_last_week:
            pts = pts[:-1]
        by = defaultdict(list)
        for d, v in pts:
            by[d.strftime("%Y-%m")].append(v)
        months = sorted(by)
        # первый месяц неполный (ряд начинается 26.09.2021)
        if months and meta["time"].startswith("today 5-y"):
            months = months[1:]
        res[k] = {m: sum(by[m]) / len(by[m]) for m in months}
    return meta, res


if __name__ == "__main__":
    for p in sorted(GT_DIR.glob("S_*_5y.txt")) + [GT_DIR / "W_P1_5y.txt"]:
        meta, res = monthly(p)
        for k, s in res.items():
            vals = list(s.values())
            print(p.name, k, len(vals), "avg %.1f" % (sum(vals) / len(vals)), "max %.1f" % max(vals), list(s)[0], list(s)[-1])
