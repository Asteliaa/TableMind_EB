"""docx -> PDF через Word COM: python _tools/docx_to_pdf.py <файл.docx> [<файл.pdf>]"""
import os
import sys

import pythoncom
import win32com.client as wc

src = os.path.abspath(sys.argv[1])
dst = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.splitext(src)[0] + ".pdf"
pythoncom.CoInitialize()
w = wc.DispatchEx("Word.Application")
w.Visible = False
w.DisplayAlerts = 0
try:
    d = w.Documents.Open(src, ReadOnly=True)
    d.ExportAsFixedFormat(dst, 17)
    print("pages:", d.ComputeStatistics(2))
    d.Close(False)
finally:
    w.Quit()
print(dst)
