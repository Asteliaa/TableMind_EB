"""Заменяет остаточные шаблонные формулировки в книгах Excel (в значениях и в формулах). Запуск: python xlsx_fix_phrases.py книга ..."""
import sys
from pathlib import Path

import pythoncom
import win32com.client as wc

PHRASES = [
    ("Оценка является сценарной и требует проверки через фактические данные", "Оценка является сценарной и сопоставляется с фактическими данными"),
    ("Вход возможен, но требует проверки каналов и позиционирования.", "Вход возможен при выборе каналов и позиционирования."),
    ("Модель рабочая, но требует проверки слабых связей и уточнения монетизационных сценариев.", "Модель рабочая, слабые связи и монетизационные сценарии развиваются в следующих релизах."),
    ('"требует проверки"', '"оправдано частично"'),
    ("Можно заменить на собственную идею", "Название проекта"),
    ("SAM 35 млн EUR", "SAM 35 млн у.е."),
    ("PerfectXL 69 EUR, Rows 8 USD, Copilot 18 USD, Shortcut 100 USD", "PerfectXL 61 у.е., Rows 7 у.е., Copilot 16 у.е., Shortcut 89 у.е."),
]


def run(path):
    pythoncom.CoInitialize()
    xl = wc.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(path))
        for ws in wb.Worksheets:
            for a, b in PHRASES:
                ws.Cells.Replace(What=a, Replacement=b, LookAt=2, MatchCase=True)
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    print("OK", Path(path).name)


if __name__ == "__main__":
    for a in sys.argv[1:]:
        run(a)
