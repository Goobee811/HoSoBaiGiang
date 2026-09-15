# -*- coding: utf-8 -*-
"""Đo số trang của từng heading từ PDF -> toc.json (danh sách theo thứ tự)."""
import subprocess, json, re, sys, os, docx, unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, '..'))
WORK = os.environ.get('HSBG_WORK', os.path.join(HERE, '..', 'build'))

DOCX = os.path.join(REPO, 'Ho_so_bai_giang_lai_xe_2026_v1.0.docx')
PDF  = os.path.join(WORK, 'pdf', 'Ho_so_bai_giang_lai_xe_2026_v1.0.pdf')

d = docx.Document(DOCX)
heads = []
for p in d.paragraphs:
    sn, t = p.style.name, p.text.strip()
    if not t:
        continue
    if sn == 'Heading 1':
        heads.append((1, t))
    elif sn == 'Heading 2':
        heads.append((2, t))

txt = subprocess.run(['pdftotext', '-layout', PDF, '-'],
                     capture_output=True).stdout.decode('utf8', 'ignore')
pages = txt.split('\f')

def squash(s):
    """Bỏ hết khoảng trắng + hạ chữ thường -> khớp được cả khi PDF ngắt dòng."""
    s = unicodedata.normalize('NFC', s).lower()
    return re.sub(r'\s+', '', s)

sq_pages = [squash(p) for p in pages]

# trang mục lục: chứa 'MỤC LỤC' và nằm ở đầu tài liệu -> bỏ qua khi dò
toc_pages = set()
for i, pg in enumerate(pages[:10]):
    if 'MỤC LỤC' in pg:
        toc_pages.add(i)
        j = i + 1
        while j < len(pages) and 'PHẦN 1' not in pages[j]:
            toc_pages.add(j); j += 1
        break

out, start = [], 0
miss = 0
for lv, t in heads:
    key = squash(t)
    found = None
    for i in range(start, len(pages)):
        if i in toc_pages:
            continue
        if key and key in sq_pages[i]:
            found = i + 1
            start = i
            break
    if found is None:
        miss += 1
    out.append({'level': lv, 'text': t, 'page': found or ''})

json.dump(out, open(os.path.join(WORK, 'toc.json'), 'w', encoding='utf8'), ensure_ascii=False, indent=0)
print('Khớp %d/%d heading (thiếu %d)' % (len(heads) - miss, len(heads), miss), file=sys.stderr)
for h in out:
    if not h['page']:
        print('   THIẾU:', h['text'][:60], file=sys.stderr)
