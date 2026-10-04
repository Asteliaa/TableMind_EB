"""ЛР4: простой шаблон анализа пяти сил Портера (баллы 1-3 по параметрам) для TableMind."""
import json
import os
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent.parent / "ЛР2-3_Объем_рынка" / "Python"))
sys.stdout.reconfigure(encoding="utf-8")
import pythoncom
import win32com.client as wc
from xlh import put, val, errors

TPL = Path(r"C:\Users\User\Downloads\шаблон анализа 5сил Портера.xlsx")
OUT = HERE.parent / "Excel" / "ЛР4_2_пять_сил_Портера_шаблон_TableMind.xlsx"

# строка оценки -> (балл, обоснование); уровни 3/2/1 лежат в столбцах C/D/E; у поставщиков 2/1 в C/D
SCORES = {
    "substitutes": {14: (3, "Частичные конкуренты и заменители (AI-помощники, FP&A-платформы, встроенный Copilot) дают около 73 % релевантного трафика сегмента; ручная проверка и Spreadsheet Inquire тоже доступны")},
    "rivalry": {24: (2, "Специализированных игроков 5 (3-10), вместе с заменителями больше 10"),
                26: (2, "Рынок табличного ПО растет на 6-8 % в год, интерес по Google Trends растет, по Вордстату снижается"),
                28: (2, "Ключевые свойства стандартизированы, но отличаются проверка структуры, доказательность и отчеты"),
                30: (2, "Цены от 8 до 100 USD; рост цен возможен в рамках ценности для сегмента")},
    "entrants": {40: (2, "Экономия на масштабе есть у Microsoft и OpenAI, в нише аудита - нет"),
                 42: (3, "Крупных брендов именно в нише аудита таблиц нет: у PerfectXL 35 тыс. визитов в месяц, отзывов почти нет"),
                 44: (2, "Есть микро-ниши: аудит моделей, EUC-риск, FP&A"),
                 46: (2, "Окупаемость разработки и продвижения 6-12 месяцев и более (бюджет 152 тыс. EUR)"),
                 48: (2, "Нужны AppSource, поиск и партнеры; платный трафик у конкурентов почти не используется"),
                 50: (2, "Регулирование низкое, но GDPR и правила AppSource обязательны"),
                 52: (1, "Конкуренты имеют бесплатные тарифы (Rows, Shortcut) и могут снизить цены"),
                 54: (3, "Рост интереса к AI-инструментам для Excel привлекает новых игроков")},
    "buyers": {65: (1, "Покупатели - отдельные специалисты и небольшие компании, крупных закупщиков мало"),
               67: (3, "Переключение на AI-помощников и бесплатные тарифы дешево"),
               69: (3, "Цена сравнивается с бесплатными и недорогими AI-подписками"),
               71: (2, "AI-ответы требуют проверки (ПЗ1), неудовлетворенность создает спрос на независимую проверку")},
    "suppliers": {81: (2, "Ключевые поставщики немногочисленны: платформа Microsoft, облако, платежи"),
                  83: (1, "Вычислительные ресурсы и LLM доступны у нескольких провайдеров"),
                  85: (1, "Принцип «LLM предлагает, движок доказывает» позволяет менять провайдера модели без смены ядра"),
                  87: (2, "Ниша аудита таблиц неприоритетна для Microsoft и провайдеров LLM")},
}
SUMS = {"substitutes": "C15", "rivalry": "C31", "entrants": "C55", "buyers": "C72", "suppliers": "C88"}


def main():
    OUT.parent.mkdir(exist_ok=True)
    shutil.copyfile(TPL, OUT)
    pythoncom.CoInitialize()
    app = wc.DispatchEx("Excel.Application"); app.Visible = False; app.DisplayAlerts = False
    wb = app.Workbooks.Open(str(OUT))
    try:
        ws = wb.Worksheets("5 сил")
        put(ws, "B6", "TableMind (облачный AI-сервис проверки Excel-моделей)")
        put(ws, "B7", "Онлайн-сервисы независимой проверки Excel-моделей; англоязычные страны и ЕС")
        res = {}
        for grp, rows in SCORES.items():
            for r, (sc, why) in rows.items():
                for c in "CDE":
                    put(ws, f"{c}{r}", None)
                col = {3: "C", 2: "D", 1: "E"}[sc] if grp != "suppliers" else {2: "C", 1: "D"}[sc]
                put(ws, f"{col}{r}", sc)
                put(ws, f"F{r}", why)
        app.CalculateFull()
        for grp, cell in SUMS.items():
            res[grp] = val(ws, cell)
        res["rows"] = {g: {str(r): list(v) for r, v in rows.items()} for g, rows in SCORES.items()}
        res["errors"] = errors(wb)
        wb.BuiltinDocumentProperties("Author").Value = "Р. В. Земляник"
        wb.BuiltinDocumentProperties("Last Author").Value = "Р. В. Земляник"
        wb.Save()
        (HERE / "results").mkdir(exist_ok=True)
        (HERE / "results" / "porter_results.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
        print({k: res[k] for k in SUMS}, res["errors"])
    finally:
        wb.Close(False)
        app.Quit()


if __name__ == "__main__":
    main()
