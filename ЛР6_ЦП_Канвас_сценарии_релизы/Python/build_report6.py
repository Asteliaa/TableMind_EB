import re
from pathlib import Path

import rep_common as C
import report6

ROOT = Path(__file__).resolve().parent.parent
report6.run()
text = C.resolve("\n".join(C.OUT)) + "\n"
text = re.sub(r"\n{3,}", "\n\n", text)
text = text.replace(chr(0x2014), "-")
import sys as _s2
_s2.path.insert(0, str(ROOT.parent / "_tools"))
from postprocess import finish as _finish
text = _finish(text)
(ROOT / "ОТЧЕТ.md").write_text(text, encoding="utf-8")
print(f"ОТЧЕТ.md: {len(text)} символов, таблиц {len(C.TABLES)}, рисунков {len(C.FIGS)}, ё: {text.count('ё') + text.count('Ё')}, тире: {text.count('—')}")
print("неразрешенных ссылок:", re.findall(r"\[\[[tf]:\w+\]\]", text))
