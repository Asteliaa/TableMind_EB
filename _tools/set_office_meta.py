"""Проставляет автора в метаданных .xlsx/.docx без пересохранения содержимого.

Меняются только docProps/core.xml (dc:creator, cp:lastModifiedBy), имя темы
оформления в theme*.xml и имя автора примечаний; все остальные части архива
копируются байт-в-байт (диаграммы, формулы и форматирование не затрагиваются).

Запуск:
    python _tools/set_office_meta.py <файл или папка> [...]
"""
import re
import sys
import zipfile
from pathlib import Path

AUTHOR = "Р. В. Земляник"
THEME_NAME = "Office"
OLD_THEME_NAMES = ("ChatGPT",)
OLD_COMMENT_AUTHORS = ()


def fix_core(xml: str) -> str:
    for tag, ns in (("creator", "dc"), ("lastModifiedBy", "cp")):
        full = f"{ns}:{tag}"
        if re.search(rf"<{full}\s*/>|<{full}>.*?</{full}>", xml, flags=re.S):
            xml = re.sub(rf"<{full}\s*/>|<{full}>.*?</{full}>", f"<{full}>{AUTHOR}</{full}>", xml, flags=re.S)
        else:
            xml = xml.replace("</cp:coreProperties>", f"<{full}>{AUTHOR}</{full}></cp:coreProperties>")
    return xml


def fix_theme(xml: str) -> str:
    for old in OLD_THEME_NAMES:
        xml = xml.replace(f'name="{old}"', f'name="{THEME_NAME}"')
    return xml


def fix_comments(xml: str) -> str:
    for old in OLD_COMMENT_AUTHORS:
        xml = xml.replace(f"<author>{old}</author>", f"<author>{AUTHOR}</author>")
        xml = xml.replace(f"{old}:", f"{AUTHOR}:")
    return xml


def process(path: Path) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            name = info.filename
            if name == "docProps/core.xml":
                data = fix_core(data.decode("utf-8")).encode("utf-8")
            elif re.match(r"(xl|word|ppt)/theme/theme\d*\.xml$", name):
                data = fix_theme(data.decode("utf-8")).encode("utf-8")
            elif re.match(r"xl/comments\d*\.xml$", name):
                data = fix_comments(data.decode("utf-8")).encode("utf-8")
            zout.writestr(info, data, compress_type=info.compress_type)
    tmp.replace(path)
    print("OK", path.name)


if __name__ == "__main__":
    for arg in sys.argv[1:]:
        p = Path(arg)
        files = sorted(p.glob("*.xlsx")) + sorted(p.glob("*.docx")) if p.is_dir() else [p]
        for f in files:
            if not f.name.startswith("~$"):
                process(f)
