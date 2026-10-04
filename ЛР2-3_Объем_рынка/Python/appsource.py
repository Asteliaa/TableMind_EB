"""Оценка сегмента через маркетплейс (методика маркетплейсов и платформ, модель через отзывы), данные Microsoft Marketplace 03.10.2026."""
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
COUNTS = [("excel", 2184), ("excel AI", 457), ("excel formula", 160), ("excel audit", 138),
          ("financial model", 87), ("spreadsheet audit", 37), ("spreadsheet error", 25), ("formula checker", 1)]
RATINGS = [9, 43, 8, 2, 3, 7, 15, 4, 1, 17, 3, 7, 8]   # карточки с рейтингом из первых 15 (запрос spreadsheet error)
LOADED, TOTAL, PAID, UNRATED = 15, 25, 9, 2
YEARS = 3                       # допущение Д-34: средний возраст предложений на рынке, лет
REVIEW = {"cons": 0.10, "base": 0.05, "opt": 0.02}      # коэффициент отзывов (методика, таблица сценариев)
CONV = {"cons": 0.05, "base": 0.10, "opt": 0.15}        # допущение Д-34: установка -> платящий
ARP = {"cons": 53.9, "base": 100.58, "opt": 149.2}      # выручка первого года на платящего, ЛР2-3 (воронка)
CHANNEL = {"cons": 2, "base": 8, "opt": 20}             # оплат в месяц канала AppSource (Д-14)


def run():
    mean = sum(RATINGS) / len(RATINGS)
    per_listing = sum(RATINGS) / LOADED          # 2 карточки без оценок учтены нулем
    total_ratings = per_listing * TOTAL
    paid_share = PAID / LOADED
    out = {"mean_rated": mean, "median_rated": sorted(RATINGS)[len(RATINGS) // 2], "sum_loaded": sum(RATINGS),
           "per_listing": per_listing, "total_ratings": total_ratings, "paid_share": paid_share, "sc": {}}
    for sc in ("cons", "base", "opt"):
        inst = total_ratings / REVIEW[sc]
        year = inst / YEARS
        payers = year * paid_share * CONV[sc]
        rev = payers * ARP[sc]
        out["sc"][sc] = {"installs": inst, "per_year": year, "payers": payers, "revenue": rev,
                         "payers_month": payers / 12, "channel_year": CHANNEL[sc] * 12,
                         "channel_vs_segment": CHANNEL[sc] * 12 / payers}
    (HERE / "results" / "appsource.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    return out


if __name__ == "__main__":
    o = run()
    print(json.dumps(o, ensure_ascii=False, indent=1))
