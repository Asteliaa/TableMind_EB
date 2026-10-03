"""Декодер выгрузок GT, сохраненных в сжатом виде (base62 + RLE нулей)."""
import re, sys
from pathlib import Path
B = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'

def decode(s, k):
    # RLE: '~'+len  (длина нулей), '_'+c  -> значение 62+idx(c); '_~n' -> '_0' + (n-1) нулей
    vals = []
    i = 0
    while i < len(s):
        c = s[i]
        if c == '~':
            vals += [0] * B.index(s[i + 1]); i += 2
        elif c == '_':
            nxt = s[i + 1]
            if nxt == '~':
                vals.append(62); vals += [0] * (B.index(s[i + 2]) - 1); i += 3
            else:
                vals.append(62 + B.index(nxt)); i += 2
        else:
            vals.append(B.index(c)); i += 1
    return [vals[j:j + k] for j in range(0, len(vals), k)]

def load(path):
    meta = {}
    for line in Path(path).read_text(encoding='utf-8').splitlines():
        key, _, val = line.partition(' ')
        meta[key] = val
    kws = [x.strip() for x in meta['keywords'].split(';')]
    rows = decode(meta['data'].strip(), len(kws))
    return meta, kws, rows

if __name__ == '__main__':
    for p in sys.argv[1:]:
        meta, kws, rows = load(p)
        cs = [sum(r[j] for r in rows) for j in range(len(kws))]
        ok = (len(rows) == int(meta['n'])) and (' '.join(map(str, cs)) == meta['checksum'])
        print(Path(p).name, len(rows), cs, 'OK' if ok else 'MISMATCH ' + meta['checksum'])
