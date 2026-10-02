"""Подготовка месячных рядов кластеров: Google Trends (мир) и Вордстат (Россия, Беларусь), 10.2021 - 09.2026 (60 месяцев)."""
import json
from pathlib import Path

from gt_monthly import monthly, GT_DIR

ROOT = Path(__file__).resolve().parent.parent
YW = json.loads((ROOT / "Материалы_собранные" / "Wordstat" / "YW_динамика_месяцы_2018-01_2026-09.json").read_text(encoding="utf-8"))
YW_MONTHS = [f"{2018 + i // 12}-{i % 12 + 1:02d}" for i in range(105)]
MONTHS = YW_MONTHS[45:]  # 2021-10 .. 2026-09
assert MONTHS[0] == "2021-10" and len(MONTHS) == 60


def yw(query, region):
    for s in YW["series"]:
        if s["query"] == query and s["region"] == region:
            return dict(zip(YW_MONTHS, s["values"]))
    raise KeyError((query, region))


def gt(file_stem):
    meta, res = monthly(GT_DIR / f"{file_stem}.txt")
    (k, ser), = res.items()
    return ser


CLUSTERS = {
    "A": {
        "name": "проверка и аудит таблиц Excel",
        "gt_queries": ["S_excel_audit_5y", "S_spreadsheet_audit_5y"],
        "gt_label": "среднее индексов запросов excel audit и spreadsheet audit",
        "yw_queries": ["проверка формул excel", "ошибки в формулах excel", "аудит excel", "проверить excel на ошибки"],
    },
    "B": {
        "name": "AI-инструменты для Excel",
        "gt_queries": ["S_excel_copilot_5y"],
        "gt_label": "индекс запроса excel copilot",
        "yw_queries": ["нейросеть для excel", "excel искусственный интеллект", "excel copilot", "chatgpt excel"],
    },
    "C": {
        "name": "аудит и проверка финансовых моделей",
        "gt_queries": ["S_financial_model_audit_5y"],
        "gt_label": "индекс запроса financial model audit",
        "yw_queries": ["проверка финансовой модели", "аудит финансовой модели"],
    },
}


def cluster_series(key, region=225):
    c = CLUSTERS[key]
    g = [gt(q) for q in c["gt_queries"]]
    gt_series = [sum(s[m] for s in g) / len(g) for m in MONTHS]
    ys = [yw(q, region) for q in c["yw_queries"]]
    yw_series = [sum(s[m] for s in ys) for m in MONTHS]
    return gt_series, yw_series


def despike(series, factor=3.0):
    """Разовый выброс: значение > factor x второго по величине значения ряда заменяется средним трех предыдущих месяцев."""
    srt = sorted(series, reverse=True)
    out, repl = list(series), []
    for i, v in enumerate(series):
        if len(srt) > 1 and srt[1] > 0 and v > factor * srt[1] and v == srt[0]:
            prev = series[max(0, i - 3):i] or [srt[1]]
            out[i] = sum(prev) / len(prev)
            repl.append((MONTHS[i], v, out[i]))
    return out, repl


if __name__ == "__main__":
    for k in CLUSTERS:
        g, y = cluster_series(k)
        yb = cluster_series(k, 149)[1]
        y2, rep = despike(y)
        print(k, CLUSTERS[k]["name"])
        print("  GT avg %.1f max %.1f" % (sum(g) / 60, max(g)), "| YW RU sum", sum(y), "max", max(y), "| BY sum", sum(yb), "max", max(yb), "| spikes", rep)
