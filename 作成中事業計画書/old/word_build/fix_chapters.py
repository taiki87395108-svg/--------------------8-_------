# -*- coding: utf-8 -*-
"""章データJSONの文字列に old→new の置換を当てる（件数を必ず表示。0件はエラー表示）。
使い方: python fix_chapters.py fixes.json   （fixes.json = [{"file":"row6.json","old":"…","new":"…","count":1}, ...]）"""
import json, sys, os
sys.stdout.reconfigure(encoding='utf-8')
HERE = os.path.dirname(os.path.abspath(__file__))
fixes = json.load(open(sys.argv[1], encoding='utf-8'))

def walk(o, old, new, cnt):
    if isinstance(o, str):
        n = o.count(old)
        if n:
            cnt[0] += n
            return o.replace(old, new)
        return o
    if isinstance(o, list):
        return [walk(x, old, new, cnt) for x in o]
    if isinstance(o, dict):
        return {k: walk(v, old, new, cnt) for k, v in o.items()}
    return o

by_file = {}
for f in fixes:
    by_file.setdefault(f['file'], []).append(f)
bad = 0
for fn, lst in by_file.items():
    p = os.path.join(HERE, 'chapters', fn)
    d = json.load(open(p, encoding='utf-8'))
    for f in lst:
        cnt = [0]
        d = walk(d, f['old'], f['new'], cnt)
        exp = f.get('count')
        flag = '' if (cnt[0] > 0 and (exp is None or exp == cnt[0])) else '  <<< 想定と違う'
        if flag: bad += 1
        print(f'{fn}: {cnt[0]}件  {f["old"][:40]} → {f["new"][:40]}{flag}')
    json.dump(d, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('NG:', bad)
