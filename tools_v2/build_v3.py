# -*- coding: utf-8 -*-
"""Dựng Ho_so_bai_giang_lai_xe_oto_2026_v2.docx — đen trắng, định dạng theo file gốc."""
import os, re, sys, json, glob
import docx
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.opc.constants import RELATIONSHIP_TYPE as RT

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, '..'))
WORK = os.environ.get('HSBG_WORK', os.path.join(REPO, 'build'))
MEDIA = os.path.join(WORK, 'x', 'word', 'media')
EXTENTS = json.load(open(os.path.join(WORK, 'extents.json'), encoding='utf8'))
SRC = os.path.join(REPO, 'drafts_v3')
OUT = os.path.join(REPO, 'Ho_so_bai_giang_lai_xe_oto_2026_v3.docx')

FONT = 'Times New Roman'
BODY = 13          # pt
H1 = 16            # pt
LINE = 1.5
SPACE_BEFORE = 6   # pt
MAXW = Cm(16.5)    # 21 - 3.0 - 1.5

# ───────── tiện ích ─────────
def _el(tag, **a):
    e = OxmlElement(tag)
    for k, v in a.items():
        e.set(qn(k), str(v))
    return e

RED = RGBColor(0xC0, 0x00, 0x00)

def setfont(run, size=BODY, bold=False, italic=False, red=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RED if red else RGBColor(0, 0, 0)
    rPr = run._r.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rPr.insert(0, rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rf.set(qn(a), FONT)

def borders_black(tbl):
    b = OxmlElement('w:tblBorders')
    for e in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b.append(_el('w:' + e, **{'w:val': 'single', 'w:sz': 6,
                                  'w:space': '0', 'w:color': '000000'}))
    tbl._tbl.tblPr.append(b)

def autofit(tbl, pct=5000):
    p = tbl._tbl.tblPr
    for t in ('w:tblW', 'w:tblLayout'):
        for e in p.findall(qn(t)):
            p.remove(e)
    p.append(_el('w:tblW', **{'w:w': pct, 'w:type': 'pct'}))
    p.append(_el('w:tblLayout', **{'w:type': 'autofit'}))
    for row in tbl.rows:
        for c in row.cells:
            tcPr = c._tc.get_or_add_tcPr()
            for e in tcPr.findall(qn('w:tcW')):
                tcPr.remove(e)

def cellmar(tbl):
    m = OxmlElement('w:tblCellMar')
    for n, v in (('top', 60), ('start', 100), ('bottom', 60), ('end', 100)):
        m.append(_el('w:' + n, **{'w:w': v, 'w:type': 'dxa'}))
    tbl._tbl.tblPr.append(m)

def repeat_hdr(row):
    row._tr.get_or_add_trPr().append(_el('w:tblHeader', **{'w:val': 'true'}))

_pid = [900]
_IMG_CT = {'.png': 'image/png', '.jpeg': 'image/jpeg', '.jpg': 'image/jpeg',
           '.emf': 'image/x-emf', '.wmf': 'image/x-wmf', '.gif': 'image/gif'}

def add_img(doc, par, fname):
    path = os.path.join(MEDIA, fname)
    if not os.path.exists(path):
        return
    cx, cy = EXTENTS.get(fname, (Inches(4), Inches(3)))
    if cx > MAXW:
        cy = int(cy * MAXW / cx); cx = int(MAXW)
    ext = os.path.splitext(fname)[1].lower()
    blob = open(path, 'rb').read()
    pn = PackURI('/word/media/%s' % fname)
    pkg = doc.part.package
    part = next((p for p in pkg.iter_parts() if str(p.partname) == str(pn)), None)
    if part is None:
        part = Part(pn, _IMG_CT.get(ext, 'image/png'), blob, pkg)
    rId = doc.part.relate_to(part, RT.IMAGE)
    _pid[0] += 1; i = _pid[0]
    xml = ('<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<wp:inline distT="0" distB="0" distL="0" distR="0">'
           '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
           '<wp:docPr id="%d" name="Hinh %d"/><wp:cNvGraphicFramePr>'
           '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'noChangeAspect="1"/></wp:cNvGraphicFramePr>'
           '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
           '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
           '<pic:nvPicPr><pic:cNvPr id="%d" name="Hinh %d"/><pic:cNvPicPr/></pic:nvPicPr>'
           '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
           '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
           '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
           '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
           ) % (cx, cy, i, i, i, i, rId, cx, cy)
    par.add_run()._r.append(parse_xml(xml))

# ───────── inline markdown ─────────
def runs(par, text, bold=False, italic=False, size=BODY, red=False):
    text = re.sub(r'<br\s*/?>', '\n', text).replace('\\*', '\x00')
    b, it, rd = bold, italic, red
    buf = []
    def flush():
        if not buf: return
        s = ''.join(buf).replace('\x00', '*'); buf.clear()
        for n, piece in enumerate(s.split('\n')):
            if n: par.add_run().add_break()
            if piece:
                setfont(par.add_run(piece), size, b, it, rd)
    i, n = 0, len(text)
    while i < n:
        if text.startswith('{{', i): flush(); rd = True;  i += 2; continue
        if text.startswith('}}', i): flush(); rd = red;   i += 2; continue
        if text.startswith('**', i): flush(); b = not b;  i += 2; continue
        if text[i] == '*':           flush(); it = not it; i += 1; continue
        if text[i] == '`':           i += 1; continue
        buf.append(text[i]); i += 1
    flush()

# ───────── tài liệu ─────────
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin = sec.bottom_margin = Cm(2.0)
sec.left_margin, sec.right_margin = Cm(3.0), Cm(1.5)
sec.header_distance, sec.footer_distance = Cm(1.25), Cm(1.25)

st = doc.styles['Normal']
st.font.name = FONT; st.font.size = Pt(BODY)
st.font.color.rgb = RGBColor(0, 0, 0)
st._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
st.paragraph_format.line_spacing = LINE
st.paragraph_format.space_before = Pt(SPACE_BEFORE)
st.paragraph_format.space_after = Pt(0)
st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

h1 = doc.styles['Heading 1']
h1.font.name = FONT; h1.font.size = Pt(H1); h1.font.bold = True
h1.font.color.rgb = RGBColor(0, 0, 0)
h1._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
h1.paragraph_format.line_spacing = LINE
h1.paragraph_format.space_before = Pt(SPACE_BEFORE)
h1.paragraph_format.space_after = Pt(SPACE_BEFORE)
h1.paragraph_format.keep_with_next = True

for _nm in ('Heading 2', 'Heading 3'):
    _h = doc.styles[_nm]
    _h.font.name = FONT; _h.font.size = Pt(BODY); _h.font.bold = True
    _h.font.italic = False
    _h.font.color.rgb = RGBColor(0, 0, 0)
    _h._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    _h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _h.paragraph_format.line_spacing = LINE
    _h.paragraph_format.keep_with_next = True


def para(align=None, before=SPACE_BEFORE, indent=None, keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = LINE; pf.space_before = Pt(before); pf.space_after = Pt(0)
    pf.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent is not None: pf.left_indent = Cm(indent)
    pf.keep_with_next = keep
    return p

def page_break():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    p.add_run().add_break(WD_BREAK.PAGE)

def render_table(rows):
    def cells(r):
        r = r.strip()
        if r.startswith('|'): r = r[1:]
        if r.endswith('|'):  r = r[:-1]
        return [c.strip() for c in r.split('|')]
    hdr = cells(rows[0]); sep = cells(rows[1])
    aligns = ['center' if c.startswith(':') and c.endswith(':') else 'left' for c in sep]
    body = [cells(r) for r in rows[2:]]
    ncol = len(hdr)
    body = [(b + [''] * ncol)[:ncol] for b in body]
    t = doc.add_table(rows=1, cols=ncol)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    borders_black(t); autofit(t); cellmar(t)
    repeat_hdr(t.rows[0])
    AL = {'center': WD_ALIGN_PARAGRAPH.CENTER, 'left': WD_ALIGN_PARAGRAPH.LEFT}
    for k, h in enumerate(hdr):
        c = t.rows[0].cells[k]; p = c.paragraphs[0]
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        runs(p, h, bold=True, size=BODY - 1)
    for row in body:
        rr = t.add_row()
        for k, v in enumerate(row):
            c = rr.cells[k]; p = c.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
            p.alignment = AL.get(aligns[k], WD_ALIGN_PARAGRAPH.LEFT)
            v = re.sub(r'(?<!^)\s+-\s+', '\n- ', v)
            runs(p, v, size=BODY - 1)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0); sp.paragraph_format.space_after = Pt(0)
    setfont(sp.add_run(''), 4)

# ───────── MỤC LỤC ─────────
TOCLIST = []
_tocf = os.path.join(WORK, 'toc_v3.json')
if os.path.exists(_tocf):
    TOCLIST = json.load(open(_tocf, encoding='utf8'))

_t = doc.add_paragraph(style='Heading 1')
runs(_t, 'MỤC LỤC', bold=True, size=H1)

_toc_anchor = doc.add_paragraph()
_r1 = _toc_anchor.add_run(); _r1._r.append(_el('w:fldChar', **{'w:fldCharType': 'begin'}))
_r2 = _toc_anchor.add_run()
_itx = OxmlElement('w:instrText'); _itx.set(qn('xml:space'), 'preserve')
_itx.text = r' TOC \o "1-2" \h \z \u '
_r2._r.append(_itx)
_r3 = _toc_anchor.add_run(); _r3._r.append(_el('w:fldChar', **{'w:fldCharType': 'separate'}))


def toc_par(level, text, page):
    tp = doc.add_paragraph()
    pf = tp.paragraph_format
    pf.line_spacing = 1.3
    pf.space_before = Pt(6 if level == 1 else 0); pf.space_after = Pt(0)
    pf.left_indent = Cm(0 if level == 1 else 0.8)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pr = tp._p.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    tb = _el('w:tab', **{'w:val': 'right', 'w:pos': '9355'})
    tb.set(qn('w:leader'), 'dot')
    tabs.append(tb); pr.append(tabs)
    setfont(tp.add_run(text), BODY if level == 1 else BODY - 1, bold=(level == 1))
    tp.add_run('\t')
    setfont(tp.add_run(str(page) if page else ''), BODY if level == 1 else BODY - 1,
            bold=(level == 1))
    return tp


# ───────── vòng dựng ─────────
FILES = sorted(glob.glob(os.path.join(SRC, '*.md')))
first = True
for f in FILES:
    lines = open(f, encoding='utf8').read().split('\n')
    page_break()
    first = False
    i = 0
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1; continue
        if s.startswith('@@LABEL'):
            p = para(WD_ALIGN_PARAGRAPH.CENTER, keep=True)
            runs(p, s[7:].strip(), bold=True)
            i += 1; continue
        if s.startswith('@@CENTER'):
            p = para(WD_ALIGN_PARAGRAPH.CENTER, before=10, keep=True)
            runs(p, s[8:].strip(), bold=True)
            i += 1; continue
        if s.startswith('@@image'):
            p = para(WD_ALIGN_PARAGRAPH.CENTER, before=4)
            add_img(doc, p, s[2:].strip())
            i += 1; continue
        if s.startswith('# '):
            p = doc.add_paragraph(style='Heading 1')
            runs(p, s[2:].strip(), bold=True, size=H1)
            i += 1; continue
        if s.startswith('### '):
            p = doc.add_paragraph(style='Heading 3')
            pf = p.paragraph_format
            pf.line_spacing = LINE; pf.space_before = Pt(8); pf.space_after = Pt(0)
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT; pf.keep_with_next = True
            runs(p, s[4:].strip(), bold=True)
            i += 1; continue
        if s.startswith('## '):
            p = doc.add_paragraph(style='Heading 2')
            pf = p.paragraph_format
            pf.line_spacing = LINE; pf.space_before = Pt(10); pf.space_after = Pt(0)
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT; pf.keep_with_next = True
            runs(p, s[3:].strip(), bold=True)
            i += 1; continue
        if s.startswith('|') and i + 1 < len(lines) and re.match(r'^\|\s*:?-{2,}', lines[i+1].strip()):
            j = i; buf = []
            while j < len(lines) and lines[j].strip().startswith('|'):
                buf.append(lines[j].strip()); j += 1
            render_table(buf); i = j; continue
        if s.startswith('> '):
            j = i; buf = []
            while j < len(lines) and lines[j].strip().startswith('>'):
                buf.append(re.sub(r'^>\s?', '', lines[j].strip())); j += 1
            p = para(indent=0.5, before=8)
            pPr = p._p.get_or_add_pPr()
            bd = OxmlElement('w:pBdr')
            for e in ('top', 'left', 'bottom', 'right'):
                bd.append(_el('w:' + e, **{'w:val': 'single', 'w:sz': 4,
                                           'w:space': '6', 'w:color': '000000'}))
            pPr.append(bd)
            p.paragraph_format.right_indent = Cm(0.5)
            runs(p, ' '.join(buf))
            sp = doc.add_paragraph(); sp.paragraph_format.space_before = Pt(0)
            sp.paragraph_format.space_after = Pt(0); setfont(sp.add_run(''), 4)
            i = j; continue
        if re.match(r'^- ', s):
            p = para(indent=0.75, before=2)
            p.paragraph_format.first_line_indent = Cm(-0.4)
            runs(p, '- ' + s[2:])
            i += 1; continue
        if re.match(r'^\+ ', s):
            p = para(indent=0.5, before=2)
            runs(p, s)
            i += 1; continue
        if re.match(r'^\d+[.)] ', s):
            p = para(indent=0.5, before=4)
            p.paragraph_format.first_line_indent = Cm(-0.5)
            runs(p, s)
            i += 1; continue
        # đoạn văn thường
        j = i; buf = []
        while j < len(lines) and lines[j].strip() and not re.match(
                r'^(#{1,4} |\||>|@@|- |\+ |\d+[.)] )', lines[j].strip()):
            buf.append(lines[j].strip()); j += 1
        p = para()
        runs(p, ' '.join(buf))
        i = j

# ───────── điền mục lục ─────────
_heads = []
for _p in doc.paragraphs:
    _sn = _p.style.name; _t2 = _p.text.strip()
    if not _t2 or _t2 == 'MỤC LỤC':
        continue
    if _sn == 'Heading 1': _heads.append((1, _t2))
    elif _sn == 'Heading 2': _heads.append((2, _t2))

_made = []
for _k, (_lv, _t2) in enumerate(_heads):
    _pg = ''
    if _k < len(TOCLIST) and TOCLIST[_k].get('text') == _t2:
        _pg = TOCLIST[_k].get('page', '')
    _made.append(toc_par(_lv, _t2, _pg))
_endp = doc.add_paragraph()
_re2 = _endp.add_run(); _re2._r.append(_el('w:fldChar', **{'w:fldCharType': 'end'}))
_prev = _toc_anchor._p
for _tp in _made + [_endp]:
    _prev.addnext(_tp._p); _prev = _tp._p

# ───────── footer "Trang N" ─────────
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fp.paragraph_format.space_before = Pt(0); fp.paragraph_format.space_after = Pt(0)
setfont(fp.add_run('Trang '), BODY)
r1 = fp.add_run(); r1._r.append(_el('w:fldChar', **{'w:fldCharType': 'begin'}))
r2 = fp.add_run()
it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve')
it.text = r' PAGE   \* MERGEFORMAT '
r2._r.append(it)
r3 = fp.add_run(); r3._r.append(_el('w:fldChar', **{'w:fldCharType': 'separate'}))
r4 = fp.add_run('1')
r5 = fp.add_run(); r5._r.append(_el('w:fldChar', **{'w:fldCharType': 'end'}))
for r in (r1, r2, r3, r4, r5):
    setfont(r, BODY)

doc.settings.element.append(_el('w:updateFields', **{'w:val': 'true'}))
doc.save(OUT)
print('ĐÃ TẠO:', os.path.basename(OUT), os.path.getsize(OUT) // 1024, 'KB')
