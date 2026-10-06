"""Добавляет на лист «Панель» книг XLT (ЛР1) две диаграммы: сводный индекс с трендом и сезонные индексы. Запуск: python xlt_charts.py книга ..."""
import sys
from pathlib import Path

import pythoncom
import win32com.client as wc

PAL = [0x1F | 0x4E << 8 | 0x79 << 16, 0x2E | 0x75 << 8 | 0xB6 << 16]  # 1F4E79 и 2E75B6 (RGB в формате COM)


def find_last(ws, col):
    r = ws.Cells(ws.Rows.Count, col).End(-4162).Row
    return r


def run(path):
    pythoncom.CoInitialize()
    xl = wc.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(path))
        p = wb.Worksheets("Панель")
        calc = wb.Worksheets("Расчет")
        sez = wb.Worksheets("Сезонность")
        for co in list(p.ChartObjects()):
            co.Delete()
        top = p.Cells(p.UsedRange.Row + p.UsedRange.Rows.Count + 1, 1).Top
        hdr = calc.Range("A1:A20").Find("Период", LookAt=1)
        r1 = hdr.Row + 1
        n = calc.Cells(calc.Rows.Count, 1).End(-4162).Row
        # первая строка с трендом
        r12 = r1 + 11
        co = p.ChartObjects().Add(p.Cells(1, 1).Left, top, 470, 270)
        ch = co.Chart
        ch.ChartType = 75
        while ch.SeriesCollection().Count:
            ch.SeriesCollection(1).Delete()
        s1 = ch.SeriesCollection().NewSeries()
        s1.Name = "Сводный индекс"
        s1.XValues = calc.Range(calc.Cells(r1, 1), calc.Cells(n, 1))
        s1.Values = calc.Range(calc.Cells(r1, 7), calc.Cells(n, 7))
        s2 = ch.SeriesCollection().NewSeries()
        s2.Name = "Тренд (скользящее среднее 12 мес.)"
        s2.XValues = calc.Range(calc.Cells(r12, 1), calc.Cells(n, 1))
        s2.Values = calc.Range(calc.Cells(r12, 8), calc.Cells(n, 8))
        s1.Format.Line.ForeColor.RGB = PAL[0]
        s2.Format.Line.ForeColor.RGB = PAL[1]
        s1.Format.Line.Weight = 1.75
        s2.Format.Line.Weight = 2.25
        ch.Axes(1).TickLabels.NumberFormat = "mm.yyyy"
        ch.HasTitle = True
        ch.ChartTitle.Text = "Динамика интереса и тренд"
        ch.HasLegend = True
        ch.Legend.Position = -4107
        ch.ChartArea.Format.Line.Visible = False
        # 2. сезонность
        jan = sez.Range("B1:B20").Find("январь", LookAt=1)
        sr = jan.Row
        co2 = p.ChartObjects().Add(p.Cells(1, 1).Left + 490, top, 420, 270)
        c2 = co2.Chart
        c2.ChartType = 51
        while c2.SeriesCollection().Count:
            c2.SeriesCollection(1).Delete()
        s = c2.SeriesCollection().NewSeries()
        s.Name = "Сезонный индекс"
        s.Values = sez.Range(sez.Cells(sr, 4), sez.Cells(sr + 11, 4))
        s.XValues = sez.Range(sez.Cells(sr, 2), sez.Cells(sr + 11, 2))
        s.Format.Fill.ForeColor.RGB = PAL[1]
        c2.HasTitle = True
        c2.ChartTitle.Text = "Сезонные индексы по месяцам"
        c2.HasLegend = False
        c2.ChartArea.Format.Line.Visible = False
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    print("OK", Path(path).name)


if __name__ == "__main__":
    for a in sys.argv[1:]:
        run(a)
