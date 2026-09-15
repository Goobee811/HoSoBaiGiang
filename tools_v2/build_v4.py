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
OUT = os.path.join(REPO, 'Ho_so_bai_giang_lai_xe_oto_2026_v4.docx')

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
sec.header_distance, sec.footer_distance = Cm(1.1), Cm(1.0)

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


# ───────── TRANG BÌA ─────────
INDIGO = (0x2B, 0x16, 0x63)
BLUE   = (0x0A, 0x6F, 0xC2)
TINT   = 'EDF2FA'
GREY2  = '595959'

SCHOOL_MINISTRY = 'BỘ GIAO THÔNG VẬN TẢI'
SCHOOL_NAME     = 'TRƯỜNG CAO ĐẲNG NGHỀ GIAO THÔNG VẬN TẢI TRUNG ƯƠNG III'
DOC_TITLE       = 'HỒ SƠ BÀI GIẢNG'
DOC_SUBTITLE    = 'THỰC HÀNH LÁI XE Ô TÔ'
CLASS_LINE      = 'CÁC HẠNG  B  ·  C1  ·  C  ·  D  ·  E'
TEACHER         = 'Thầy TRẦN THANH HẢI'
PHONE           = '0903.785.483 (Zalo)'
YEAR            = '2026'
LEGAL_LINE      = ('Biên soạn theo Luật 36/2024/QH15  ·  Thông tư 17/2026/TT-BXD  ·  '
                   'Thông tư 108/2026/TT-BCA')


def _colorrun(run, size, rgb, bold=False, italic=False, spacing=None, caps=False):
    setfont(run, size, bold=bold, italic=italic)
    run.font.color.rgb = RGBColor(*rgb) if isinstance(rgb, tuple) else RGBColor.from_string(rgb)
    rPr = run._r.get_or_add_rPr()
    if spacing:
        rPr.append(_el('w:spacing', **{'w:val': spacing}))
    if caps:
        rPr.append(_el('w:caps', **{'w:val': 'true'}))


def _band(fill, pts):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    autofit(t)
    b = OxmlElement('w:tblBorders')
    for e in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b.append(_el('w:' + e, **{'w:val': 'none', 'w:sz': 0, 'w:space': '0', 'w:color': 'auto'}))
    t._tbl.tblPr.append(b)
    m = OxmlElement('w:tblCellMar')
    for n in ('top', 'start', 'bottom', 'end'):
        m.append(_el('w:' + n, **{'w:w': 0, 'w:type': 'dxa'}))
    t._tbl.tblPr.append(m)
    c = t.cell(0, 0)
    c._tc.get_or_add_tcPr().append(
        _el('w:shd', **{'w:val': 'clear', 'w:color': 'auto', 'w:fill': fill}))
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    setfont(p.add_run(' '), pts)
    return t


def _line(text, size, rgb, bold=False, italic=False, before=0, after=0,
          spacing=None, align=WD_ALIGN_PARAGRAPH.CENTER):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(before); pf.space_after = Pt(after)
    pf.line_spacing = 1.0
    _colorrun(p.add_run(text), size, rgb, bold=bold, italic=italic, spacing=spacing)
    return p


def build_cover():
    _band('2B1663', 7)
    _line(SCHOOL_MINISTRY, 11, GREY2, bold=True, before=20, after=4, spacing=40)
    _line(SCHOOL_NAME, 12.5, INDIGO, bold=True, after=6)

    rp = doc.add_paragraph()
    rp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rp.paragraph_format.space_before = Pt(0); rp.paragraph_format.space_after = Pt(0)
    rp.paragraph_format.left_indent = Cm(6.0); rp.paragraph_format.right_indent = Cm(6.0)
    _pPr = rp._p.get_or_add_pPr()
    _bb = OxmlElement('w:pBdr')
    _bb.append(_el('w:bottom', **{'w:val': 'single', 'w:sz': 10,
                                  'w:space': '0', 'w:color': '0A6FC2'}))
    _pPr.append(_bb)
    setfont(rp.add_run(' '), 2)

    lp = doc.add_paragraph()
    lp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lp.paragraph_format.space_before = Pt(44); lp.paragraph_format.space_after = Pt(0)
    _logo = os.path.join(MEDIA, 'logo_truong.jpeg')
    if os.path.exists(_logo):
        _s = int(Cm(6.18))
        add_img(doc, lp, 'logo_truong.jpeg')
        for _ext in lp._p.iter(qn('wp:extent')):
            _ext.set('cx', str(_s)); _ext.set('cy', str(_s))
        for _ext in lp._p.iter(qn('a:ext')):
            _ext.set('cx', str(_s)); _ext.set('cy', str(_s))

    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    autofit(t)
    bd = OxmlElement('w:tblBorders')
    bd.append(_el('w:top', **{'w:val': 'single', 'w:sz': 18, 'w:space': '0', 'w:color': '0A6FC2'}))
    bd.append(_el('w:bottom', **{'w:val': 'single', 'w:sz': 18, 'w:space': '0', 'w:color': '0A6FC2'}))
    for e in ('left', 'right', 'insideH', 'insideV'):
        bd.append(_el('w:' + e, **{'w:val': 'none', 'w:sz': 0, 'w:space': '0', 'w:color': 'auto'}))
    t._tbl.tblPr.append(bd)
    cellmar(t)
    c = t.cell(0, 0)
    c._tc.get_or_add_tcPr().append(
        _el('w:shd', **{'w:val': 'clear', 'w:color': 'auto', 'w:fill': TINT}))
    p0 = c.paragraphs[0]
    p0.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(26); p0.paragraph_format.space_after = Pt(2)
    p0.paragraph_format.line_spacing = 1.0
    _colorrun(p0.add_run(DOC_TITLE), 30, INDIGO, bold=True, spacing=30)
    p1 = c.add_paragraph()
    p1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(0); p1.paragraph_format.space_after = Pt(26)
    p1.paragraph_format.line_spacing = 1.0
    _colorrun(p1.add_run(DOC_SUBTITLE), 19, BLUE, bold=True, spacing=16)

    _line(CLASS_LINE, 18, INDIGO, bold=True, before=32, after=24, spacing=24)

    it = doc.add_table(rows=3, cols=2)
    it.alignment = WD_TABLE_ALIGNMENT.CENTER
    bd2 = OxmlElement('w:tblBorders')
    for e in ('top', 'left', 'bottom', 'right', 'insideV'):
        bd2.append(_el('w:' + e, **{'w:val': 'none', 'w:sz': 0, 'w:space': '0', 'w:color': 'auto'}))
    bd2.append(_el('w:insideH', **{'w:val': 'single', 'w:sz': 4, 'w:space': '0', 'w:color': 'D5DEEA'}))
    it._tbl.tblPr.append(bd2)
    cellmar(it)
    rows = [('Môn học', 'Thực hành lái xe ô tô'),
            ('Giáo viên thực hành', '%s\n%s' % (TEACHER, PHONE)),
            ('Năm biên soạn', YEAR)]
    for i, (k, v) in enumerate(rows):
        c0, c1 = it.cell(i, 0), it.cell(i, 1)
        p = c0.paragraphs[0]
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(7); p.paragraph_format.space_after = Pt(7)
        p.paragraph_format.line_spacing = 1.0
        _colorrun(p.add_run(k), 14, GREY2)
        p = c1.paragraphs[0]
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(7); p.paragraph_format.space_after = Pt(7)
        p.paragraph_format.line_spacing = 1.1
        for _n, _piece in enumerate(v.split('\n')):
            if _n:
                p.add_run().add_break()
            _colorrun(p.add_run(_piece), 16, INDIGO, bold=True)
    for r in it.rows:
        r.cells[0].width = Cm(5.6); r.cells[1].width = Cm(9.9)

    _line(LEGAL_LINE, 10.5, GREY2, italic=True, before=32, after=12)
    _band('0A6FC2', 7)


COVER_ONLY = os.environ.get('COVER_ONLY') == '1'
WITH_COVER = os.environ.get('NO_COVER') != '1'

if WITH_COVER:
    # section 1 = trang bìa: lề đối xứng + khung viền trang
    sec.left_margin = sec.right_margin = Cm(2.25)
    _sp = sec._sectPr
    _pb = OxmlElement('w:pgBorders')
    _pb.set(qn('w:offsetFrom'), 'page')
    for _e in ('top', 'left', 'bottom', 'right'):
        _pb.append(_el('w:' + _e, **{'w:val': 'thickThinSmallGap', 'w:sz': 24,
                                     'w:space': '24', 'w:color': '2B1663'}))
    _mar = _sp.find(qn('w:pgMar'))
    (_mar.addnext(_pb) if _mar is not None else _sp.append(_pb))
    build_cover()
    if COVER_ONLY:
        OUT = os.path.join(REPO, 'HSBG_Bia_v3.docx')

# ───────── MỤC LỤC ─────────
TOCLIST = []
_tocf = os.path.join(WORK, 'toc_v4.json')
if os.path.exists(_tocf):
    TOCLIST = json.load(open(_tocf, encoding='utf8'))

if COVER_ONLY:
    doc.save(OUT)
    print('ĐÃ TẠO:', os.path.basename(OUT), os.path.getsize(OUT) // 1024, 'KB')
    raise SystemExit(0)

if WITH_COVER:
    from docx.enum.section import WD_SECTION
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    sec.left_margin, sec.right_margin = Cm(3.0), Cm(1.5)
    sec.header_distance, sec.footer_distance = Cm(1.1), Cm(1.0)
    sec.header.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False
    # section thân bài kế thừa sectPr của bìa -> gỡ khung viền trang
    for _old in sec._sectPr.findall(qn('w:pgBorders')):
        sec._sectPr.remove(_old)
else:
    page_break()

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

# ───────── header / footer (bố cục như bản v1) ─────────
GREY = '595959'
HDR_TEXT = 'HỒ SƠ BÀI GIẢNG THỰC HÀNH LÁI XE Ô TÔ  |  CÁC HẠNG B, C1, C, D, E'
VERSION = 'Bản v4'
UPDATED = 'Tháng 9/2026'
UNIT = 'Đào tạo lái xe ô tô'


def _rule(par, edge, color=GREY, sz=6, space='2'):
    pPr = par._p.get_or_add_pPr()
    b = OxmlElement('w:pBdr')
    b.append(_el('w:' + edge, **{'w:val': 'single', 'w:sz': sz,
                                 'w:space': space, 'w:color': color}))
    pPr.append(b)


def _grey(run, size, bold=False):
    setfont(run, size, bold=bold)
    run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)


def _field(par, instr):
    a = par.add_run(); a._r.append(_el('w:fldChar', **{'w:fldCharType': 'begin'}))
    b = par.add_run()
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = instr
    b._r.append(it)
    c = par.add_run(); c._r.append(_el('w:fldChar', **{'w:fldCharType': 'separate'}))
    d = par.add_run('1')
    e = par.add_run(); e._r.append(_el('w:fldChar', **{'w:fldCharType': 'end'}))
    return [a, b, c, d, e]


hp = sec.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
hp.paragraph_format.space_before = Pt(0)
hp.paragraph_format.space_after = Pt(2)
hp.paragraph_format.line_spacing = 1.0
_grey(hp.add_run(HDR_TEXT), 9, bold=True)
_rule(hp, 'bottom')

fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
fp.paragraph_format.space_before = Pt(0)
fp.paragraph_format.space_after = Pt(0)
fp.paragraph_format.line_spacing = 1.0
_rule(fp, 'top', space='4')
pPr = fp._p.get_or_add_pPr()
tabs = OxmlElement('w:tabs')
tabs.append(_el('w:tab', **{'w:val': 'center', 'w:pos': '4536'}))
tabs.append(_el('w:tab', **{'w:val': 'right', 'w:pos': '9072'}))
pPr.append(tabs)

_grey(fp.add_run('%s — %s' % (VERSION, UPDATED)), 9)
fp.add_run('\t')
setfont(fp.add_run('Trang '), 10, bold=True)
for _r in _field(fp, r' PAGE \* MERGEFORMAT '):
    setfont(_r, 10, bold=True)
_grey(fp.add_run(' / '), 10)
for _r in _field(fp, r' NUMPAGES \* MERGEFORMAT '):
    _grey(_r, 10)
fp.add_run('\t')
_grey(fp.add_run(UNIT), 9)

if WITH_COVER:
    _cover_sec = doc.sections[0]
    _cover_sec.header.is_linked_to_previous = False
    _cover_sec.footer.is_linked_to_previous = False
    for _pp in list(_cover_sec.header.paragraphs) + list(_cover_sec.footer.paragraphs):
        for _rr in list(_pp.runs):
            _rr._r.getparent().remove(_rr._r)

doc.settings.element.append(_el('w:updateFields', **{'w:val': 'true'}))
doc.save(OUT)
print('ĐÃ TẠO:', os.path.basename(OUT), os.path.getsize(OUT) // 1024, 'KB')
