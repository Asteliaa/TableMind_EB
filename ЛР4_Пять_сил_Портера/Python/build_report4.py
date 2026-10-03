import re
from pathlib import Path

import rep_common as C
import rep4

ROOT = Path(__file__).resolve().parent.parent
rep4.run()
text = C.resolve("\n".join(C.OUT)) + "\n"
text = re.sub(r"\n{3,}", "\n\n", text)
(ROOT / "ОТЧЕТ.md").write_text(text, encoding="utf-8")
print(f"ОТЧЕТ.md: {len(text)} символов, таблиц {len(C.TABLES)}, рисунков {len(C.FIGS)}, ё: {text.count('ё') + text.count('Ё')}, тире: {text.count('—')}")
print("неразрешенных ссылок:", re.findall(r"\[\[[tf]:\w+\]\]", text))
