# -*- coding: utf-8 -*-
"""Hệ thiết kế + tiện ích docx cho giáo trình lái xe."""
import re
from docx.shared import Pt, Cm, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# ───────────────────────── BẢNG MÀU ─────────────────────────
NAVY      = '0B2C4D'   # tiêu đề chương
BLUE      = '10508C'   # tiêu đề cấp 2
BLUE_MID  = '1565C0'   # tiêu đề cấp 3
BLUE_PALE = 'E7EFF7'   # nền bảng / hộp xanh
GREY_BG   = 'F2F4F6'   # nền phụ
GREY_LINE = 'C9D2DA'
GREY_TXT  = '5A6672'
INK       = '1B2733'   # chữ chính
WHITE     = 'FFFFFF'
AMBER     = 'A35A00'
AMBER_BG  = 'FDF3E3'
RED       = 'B3261E'
RED_BG    = 'FBEBE9'
GREEN     = '2A6A43'
GREEN_BG  = 'E9F3EC'
PURPLE    = '5F3B86'
PURPLE_BG = 'F1EBF7'
TEAL      = '15686E'
TEAL_BG   = 'E5F1F2'

FONT = 'Arial'

# ───────────────────── ĐỊNH NGHĨA HỘP THÔNG TIN ─────────────────────
BOXES = {
    'GHI NHỚ':            ('📖', BLUE_MID, BLUE_PALE),
    'LƯU Ý':              ('⚠️', AMBER,    AMBER_BG),
    'LƯU Ý AN TOÀN':      ('⚠️', RED,      RED_BG),
    'LƯU Ý CHUYỂN TIẾP':  ('⚠️', AMBER,    AMBER_BG),
    'MẸO QUAN SÁT':       ('💡', GREEN,    GREEN_BG),
    'THỰC HÀNH':          ('🔧', TEAL,     TEAL_BG),
    'KIỂM TRA':           ('📋', PURPLE,   PURPLE_BG),
    'LỖI THƯỜNG GẶP':     ('⚠️', RED,      RED_BG),
    'CẦN KIỂM CHỨNG':     ('🔍', RED,      RED_BG),
    'CẦN RÀ SOÁT':        ('🔍', AMBER,    AMBER_BG),
}
DEFAULT_BOX = ('•', GREY_TXT, GREY_BG)


def box_kind(title):
    """Nhận diện loại hộp từ tiêu đề, ưu tiên nhãn dài nhất."""
    t = title.upper()
    best = None
    for k in BOXES:
        if '[' + k + ']' in t or k in t:
            if best is None or len(k) > len(best):
                best = k
    return best


# ───────────────────────── TIỆN ÍCH XML ─────────────────────────
def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def shade(cell_or_par, hexcolor):
    pr = cell_or_par._tc.get_or_add_tcPr() if hasattr(cell_or_par, '_tc') \
        else cell_or_par._p.get_or_add_pPr()
    pr.append(_el('w:shd', **{'w:val': 'clear', 'w:color': 'auto', 'w:fill': hexcolor}))


def set_borders(tbl, color=GREY_LINE, sz=4, inside=True):
    tblPr = tbl._tbl.tblPr
    b = OxmlElement('w:tblBorders')
    edges = ['top', 'left', 'bottom', 'right'] + (['insideH', 'insideV'] if inside else [])
    for e in edges:
        b.append(_el('w:' + e, **{'w:val': 'single', 'w:sz': sz,
                                  'w:space': '0', 'w:color': color}))
    tblPr.append(b)


def no_borders(tbl):
    b = OxmlElement('w:tblBorders')
    for e in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b.append(_el('w:' + e, **{'w:val': 'none', 'w:sz': 0, 'w:space': '0', 'w:color': 'auto'}))
    tbl._tbl.tblPr.append(b)


def left_accent(tbl, color, sz=18):
    """Viền trái dày làm dải nhấn cho hộp thông tin."""
    b = OxmlElement('w:tblBorders')
    b.append(_el('w:left', **{'w:val': 'single', 'w:sz': sz, 'w:space': '0', 'w:color': color}))
    for e in ['top', 'bottom', 'right']:
        b.append(_el('w:' + e, **{'w:val': 'single', 'w:sz': 2, 'w:space': '0', 'w:color': color}))
    tbl._tbl.tblPr.append(b)


def cell_margins(tbl, top=80, start=120, bottom=80, end=120):
    mar = OxmlElement('w:tblCellMar')
    for name, val in (('top', top), ('start', start), ('bottom', bottom), ('end', end)):
        mar.append(_el('w:' + name, **{'w:w': val, 'w:type': 'dxa'}))
    tbl._tbl.tblPr.append(mar)


def keep_with_next(par):
    par.paragraph_format.keep_with_next = True


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(_el('w:tblHeader', **{'w:val': 'true'}))


def cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(_el('w:cantSplit', **{'w:val': 'true'}))


def page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(6)  # WD_BREAK.PAGE


def hrule(doc, color=GREY_LINE, sz=6, space_before=6, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    pPr = p._p.get_or_add_pPr()
    b = OxmlElement('w:pBdr')
    b.append(_el('w:bottom', **{'w:val': 'single', 'w:sz': sz, 'w:space': '1', 'w:color': color}))
    pPr.append(b)
    return p


def field(par, instr):
    """Chèn field code (PAGE, TOC, NUMPAGES...)."""
    r1 = par.add_run(); r1._r.append(_el('w:fldChar', **{'w:fldCharType': 'begin'}))
    r2 = par.add_run()
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = instr
    r2._r.append(it)
    r3 = par.add_run(); r3._r.append(_el('w:fldChar', **{'w:fldCharType': 'separate'}))
    r4 = par.add_run('…')
    r5 = par.add_run(); r5._r.append(_el('w:fldChar', **{'w:fldCharType': 'end'}))
    return [r1, r2, r3, r4, r5]


# ───────────────────── CHÈN ẢNH (kể cả EMF) ─────────────────────
_IMG_CT = {'.png': 'image/png', '.jpeg': 'image/jpeg', '.jpg': 'image/jpeg',
           '.gif': 'image/gif', '.emf': 'image/x-emf', '.wmf': 'image/x-wmf'}
_pic_id = [1000]


def add_picture_any(doc, par, path, cx, cy):
    """Chèn ảnh bất kỳ định dạng, kể cả EMF mà python-docx không hỗ trợ sẵn."""
    import os
    ext = os.path.splitext(path)[1].lower()
    ct = _IMG_CT.get(ext, 'image/png')
    with open(path, 'rb') as f:
        blob = f.read()
    name = os.path.basename(path)
    partname = PackURI('/word/media/%s' % name)
    package = doc.part.package
    part = None
    for p in package.iter_parts():
        if str(p.partname) == str(partname):
            part = p
            break
    if part is None:
        part = Part(partname, ct, blob, package)
    rId = doc.part.relate_to(part, RT.IMAGE)
    _pic_id[0] += 1
    pid = _pic_id[0]
    xml = (
        '<w:drawing %s>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%d" cy="%d"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="%d" name="Hinh %d"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="%d" name="Hinh %d"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
    ) % (
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
        cx, cy, pid, pid, pid, pid, rId, cx, cy)
    from docx.oxml import parse_xml
    run = par.add_run()
    run._r.append(parse_xml(xml))
    return run


def autofit_pct(tbl, pct=5000):
    """Bảng rộng 100 % vùng chữ, cột tự co theo nội dung."""
    tblPr = tbl._tbl.tblPr
    for tag in ('w:tblW', 'w:tblLayout'):
        for e in tblPr.findall(qn(tag)):
            tblPr.remove(e)
    tblPr.append(_el('w:tblW', **{'w:w': pct, 'w:type': 'pct'}))
    tblPr.append(_el('w:tblLayout', **{'w:type': 'autofit'}))
    for row in tbl.rows:
        for c in row.cells:
            tcPr = c._tc.get_or_add_tcPr()
            for e in tcPr.findall(qn('w:tcW')):
                tcPr.remove(e)
