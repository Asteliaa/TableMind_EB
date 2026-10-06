"""Общая обработка текста отчета перед сохранением ОТЧЕТ.md: единая валюта EUR, предложения без двоеточий, заглавная первая буква в ячейках таблиц."""
import currency_eur
import status_map
from nocolon import apply as _nocolon, capfirst_cells


def finish(text: str, manual=()) -> str:
    text = currency_eur.apply(text)
    text = status_map.apply_md(text)
    text = _nocolon(text, manual)
    return capfirst_cells(text)
