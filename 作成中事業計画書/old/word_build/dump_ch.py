# -*- coding: utf-8 -*-
"""章データJSONを読みやすいテキストにして表示する（確認用）"""
import json, sys
sys.stdout.reconfigure(encoding='utf-8')

def dump(blocks, ind=''):
    for i, b in enumerate(blocks):
        t = b['t']
        if t in ('h2', 'h3', 'p', 'pn', 'note'):
            print(f'{ind}[{i}:{t}] {b["text"]}')
        elif t == 'bullets':
            print(f'{ind}[{i}:bullets {b.get("mark","")}]')
            for it in b['items']:
                print(f'{ind}    - {it}')
        elif t == 'table':
            print(f'{ind}[{i}:table] widths={b.get("widths")}')
            if b.get('header'):
                print(f'{ind}    H: ' + ' | '.join(b['header']))
            for r in b['rows']:
                print(f'{ind}    R: ' + ' | '.join(str(c).replace(chr(10), " / ") for c in r))
        elif t == 'img':
            print(f'{ind}[{i}:img] {b["src"]} w={b.get("width")} cap={b.get("caption")}')
        elif t == 'imgrow':
            print(f'{ind}[{i}:imgrow] ' + ', '.join(x['src'] for x in b['items']))
        elif t == 'labeled':
            print(f'{ind}[{i}:labeled] lw={b.get("label_width")}')
            for it in b['items']:
                print(f'{ind}  <{it["label"]}>')
                dump(it['blocks'], ind + '    ')
        elif t == 'space':
            print(f'{ind}[{i}:space]')
        else:
            print(f'{ind}[{i}:{t}] ??? {b}')

for f in sys.argv[1:]:
    d = json.load(open(f, encoding='utf-8'))
    print(f'===== {f}  row={d["row"]}  {d["title"]}')
    dump(d['blocks'])
