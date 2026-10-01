# -*- coding: utf-8 -*-
"""docx → LibreOffice PDF → ページ数と各ページの先頭/末尾行を表示（見出しの取り残し確認）"""
import sys, os, subprocess, fitz, re
sys.stdout.reconfigure(encoding='utf-8')
src = sys.argv[1]; dpi = int(sys.argv[2]) if len(sys.argv) > 2 else 0
out = os.path.dirname(os.path.abspath(src))
subprocess.run([r'C:\Program Files\LibreOffice\program\soffice.exe', '--headless', '--convert-to', 'pdf', '--outdir', out, src], capture_output=True)
pdf = os.path.splitext(src)[0] + '.pdf'
d = fitz.open(pdf); print('pages', d.page_count)
base = os.path.splitext(os.path.basename(src))[0]
HEAD = re.compile(r'^(〔|【|■|[①-⑳]|[0-9０-９]+[．.]|[１-６]‐|[1-6]-[1-3]|２-|３‐|１‐)')
for i, p in enumerate(d):
    if dpi:
        p.get_pixmap(dpi=dpi).save(os.path.join(out, f'{base}_p{i+1:02d}.png'))
    lines = []
    for b in p.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            t = ''.join(s['text'] for s in l['spans']).strip()
            if t: lines.append((l['bbox'][1], l['bbox'][3], t))
    lines.sort()
    if not lines: continue
    last = lines[-1]
    flag = '  <<< 見出しが末尾' if HEAD.match(last[2]) and len(last[2]) < 40 else ''
    print(f'p{i+1:02d} 先頭: {lines[0][2][:30]} | 末尾({last[1]:.0f}): {last[2][:36]}{flag}')
