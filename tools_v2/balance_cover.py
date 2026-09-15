# -*- coding: utf-8 -*-
"""Giãn đều chiều dọc trang bìa để nội dung lấp cân đối trong khung.
Không thêm, không bớt, không đổi chữ hay ảnh — chỉ chỉnh khoảng cách đoạn."""
import sys, docx
from docx.shared import Pt
from docx.oxml.ns import qn

SRC, DST = sys.argv[1], sys.argv[2]
BUMP = {  # nhận diện đoạn theo nội dung -> số điểm cộng thêm vào space_before
    'logo':      int(sys.argv[3]) if len(sys.argv) > 3 else 16,
    'hang':      int(sys.argv[4]) if len(sys.argv) > 4 else 12,
    'hang_after':int(sys.argv[5]) if len(sys.argv) > 5 else 10,
    'legal':     int(sys.argv[6]) if len(sys.argv) > 6 else 14,
}

d = docx.Document(SRC)


def bump(par, pts, after=False):
    pf = par.paragraph_format
    if after:
        cur = pf.space_after.pt if pf.space_after is not None else 0
        pf.space_after = Pt(cur + pts)
    else:
        cur = pf.space_before.pt if pf.space_before is not None else 0
        pf.space_before = Pt(cur + pts)


for p in d.paragraphs:
    t = p.text.strip()
    has_img = p._p.find('.//' + qn('a:blip')) is not None
    if has_img and not t:
        bump(p, BUMP['logo'])
    elif t.startswith('CÁC HẠNG'):
        bump(p, BUMP['hang'])
        bump(p, BUMP['hang_after'], after=True)
    elif t.startswith('Biên soạn theo'):
        bump(p, BUMP['legal'])

d.save(DST)
print('Đã giãn:', ', '.join('%s +%dpt' % kv for kv in BUMP.items()))
