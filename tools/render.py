# -*- coding: utf-8 -*-
"""Bộ render Markdown -> Word cho giáo trình lái xe."""
import re, os, json
from docx.shared import Pt, Cm, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from style import *
from style import autofit_pct, _el

WORK = os.environ.get('HSBG_WORK', os.path.dirname(os.path.abspath(__file__)) + '/../build')
MEDIA = os.path.join(WORK, 'x/word/media')
EXTENTS = json.load(open(os.path.join(WORK, 'extents.json')))
MAXW = Inches(6.1)

# ───────────────── INLINE FORMATTING ─────────────────
def emit_runs(par, text, base_bold=False, base_color=None, size=None, italic=False):
    """Quét inline markdown theo kiểu stack — xử lý được nhấn mạnh lồng nhau."""
    text = re.sub(r'<br\s*/?>', '\n', text)
    bold, ital, mono, strike = base_bold, italic, False, False
    buf = []

    def flush():
        if not buf:
            return
        chunk = ''.join(buf); buf.clear()
        for n, piece in enumerate(chunk.split('\n')):
            if n:
                par.add_run().add_break()
            if not piece:
                continue
            r = par.add_run(piece)
            r.bold = bold; r.italic = ital; r.font.strike = strike
            r.font.name = 'Consolas' if mono else FONT
            r.font.size = Pt(size if size else (8.5 if mono else 10.5))
            col = base_color or (AMBER if mono else None)
            if col:
                r.font.color.rgb = RGBColor.from_string(col)

    i, n = 0, len(text)
    while i < n:
        if not mono and text.startswith('**', i):
            flush(); bold = not bold; i += 2; continue
        if not mono and text.startswith('~~', i):
            flush(); strike = not strike; i += 2; continue
        if text[i] == '`':
            flush(); mono = not mono; i += 1; continue
        if not mono and text[i] == '*':
            flush(); ital = not ital; i += 1; continue
        buf.append(text[i]); i += 1
    flush()
    return par


# ───────────────── PARSER ─────────────────
def parse(md):
    """Trả về danh sách block: (kind, payload)."""
    lines = md.split('\n')
    blocks, i = [], 0
    while i < len(lines):
        ln = lines[i]
        s = ln.strip()
        # fenced code
        if s.startswith('```'):
            j = i + 1; buf = []
            while j < len(lines) and not lines[j].strip().startswith('```'):
                buf.append(lines[j]); j += 1
            blocks.append(('code', buf)); i = j + 1; continue
        # blockquote
        if s.startswith('>'):
            j = i; buf = []
            while j < len(lines) and (lines[j].strip().startswith('>') or
                                      (lines[j].strip() == '' and j + 1 < len(lines)
                                       and lines[j + 1].strip().startswith('>'))):
                t = lines[j].strip()
                buf.append(re.sub(r'^>\s?', '', t) if t.startswith('>') else '')
                j += 1
            blocks.append(('quote', buf)); i = j; continue
        # table
        if s.startswith('|'):
            j = i; buf = []
            while j < len(lines) and lines[j].strip().startswith('|'):
                buf.append(lines[j].strip()); j += 1
            blocks.append(('table', buf)); i = j; continue
        # heading
        m = re.match(r'^(#{1,6})\s+(.*)$', s)
        if m:
            blocks.append(('h%d' % len(m.group(1)), m.group(2))); i += 1; continue
        # hr
        if re.match(r'^-{3,}$', s):
            blocks.append(('hr', None)); i += 1; continue
        # list
        if re.match(r'^([-*]|\d+\.)\s+', s):
            j = i; buf = []
            while j < len(lines) and re.match(r'^([-*]|\d+\.)\s+', lines[j].strip()):
                t = lines[j].strip()
                ordered = bool(re.match(r'^\d+\.', t))
                buf.append((ordered, re.sub(r'^([-*]|\d+\.)\s+', '', t)))
                j += 1
            blocks.append(('list', buf)); i = j; continue
        if s:
            j = i; buf = []
            while j < len(lines) and lines[j].strip() and not re.match(
                    r'^(#{1,6}\s|\||>|```|-{3,}$|[-*]\s|\d+\.\s)', lines[j].strip()):
                buf.append(lines[j].strip()); j += 1
            blocks.append(('p', ' '.join(buf))); i = j; continue
        i += 1
    return blocks


def split_table(rows):
    """Tách bảng markdown -> (header, aligns, body)."""
    def cells(r):
        r = r.strip()
        if r.startswith('|'): r = r[1:]
        if r.endswith('|'): r = r[:-1]
        return [c.strip() for c in r.split('|')]
    if len(rows) < 2:
        return None
    head = cells(rows[0])
    sep = cells(rows[1])
    if not all(re.match(r'^:?-{2,}:?$', c) for c in sep):
        return None
    aligns = []
    for c in sep:
        aligns.append('center' if c.startswith(':') and c.endswith(':')
                      else 'right' if c.endswith(':') else 'left')
    body = [cells(r) for r in rows[2:]]
    n = len(head)
    body = [(b + [''] * n)[:n] for b in body]
    return head, aligns, body


# ───────────────── RENDER ─────────────────
IMG_RE = re.compile(r'`(image\d+\.(?:jpeg|jpg|png|emf))`')
ALIGN = {'left': WD_ALIGN_PARAGRAPH.LEFT, 'center': WD_ALIGN_PARAGRAPH.CENTER,
         'right': WD_ALIGN_PARAGRAPH.RIGHT}


class Renderer:
    def __init__(self, doc):
        self.doc = doc
        self.fig = 0
        self.ctx = ''

    # ---------- headings ----------
    def h1(self, text, page_break_before=True):
        if page_break_before:
            p = self.doc.add_paragraph(); p.add_run().add_break(WD_BREAK.PAGE)
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.style = self.doc.styles['Heading 1']
        emit_runs(p, text, base_bold=True, base_color=NAVY, size=19)
        hrule(self.doc, NAVY, 12, 2, 10)

    def h2(self, text):
        self.ctx = text
        p = self.doc.add_paragraph(); p.style = self.doc.styles['Heading 2']
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(5)
        keep_with_next(p)
        emit_runs(p, text, base_bold=True, base_color=BLUE, size=13.5)

    def h3(self, text):
        self.ctx = text
        # bước thực hiện -> kiểu riêng
        if text.startswith('▌'):
            return self.step(text[1:].strip())
        p = self.doc.add_paragraph(); p.style = self.doc.styles['Heading 3']
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        keep_with_next(p)
        emit_runs(p, text, base_bold=True, base_color=BLUE_MID, size=11.5)

    def h4(self, text):
        self.ctx = text
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.left_indent = Cm(0.2)
        keep_with_next(p)
        emit_runs(p, text, base_bold=True, base_color=INK, size=11)

    def step(self, text):
        """Tiêu đề bước: dải xanh đậm, số bước nổi bật."""
        self.ctx = text
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.LEFT
        no_borders(t); cell_margins(t, 60, 140, 60, 140)
        c = t.cell(0, 0); shade(c, BLUE_PALE)
        c.width = MAXW
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        emit_runs(p, text, base_bold=True, base_color=NAVY, size=11.5)
        cant_split(t.rows[0]); keep_with_next(p)
        sp = self.doc.add_paragraph(); sp.paragraph_format.space_after = Pt(2)
        sp.paragraph_format.space_before = Pt(0)
        for r in sp.runs: r.font.size = Pt(2)

    # ---------- body ----------
    def para(self, text, indent=0.0):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.18
        if indent: p.paragraph_format.left_indent = Cm(indent)
        emit_runs(p, text)
        return p

    def bullets(self, items, indent=0.4):
        for ordered, txt in items:
            p = self.doc.add_paragraph(
                style='List Number' if ordered else 'List Bullet')
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.left_indent = Cm(0.75 + indent)
            p.paragraph_format.line_spacing = 1.12
            emit_runs(p, txt)

    def code(self, lines):
        t = self.doc.add_table(rows=1, cols=1)
        no_borders(t); cell_margins(t, 100, 140, 100, 140)
        c = t.cell(0, 0); shade(c, GREY_BG); c.width = MAXW
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        for n, ln in enumerate(lines):
            if n: p.add_run().add_break()
            r = p.add_run(ln)
            r.font.name = 'Consolas'; r.font.size = Pt(8)
            r.font.color.rgb = RGBColor.from_string(NAVY)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def table(self, rows, inner_width=MAXW):
        parsed = split_table(rows)
        if not parsed:
            for r in rows: self.para(r)
            return
        head, aligns, body = parsed
        t = self.doc.add_table(rows=1, cols=len(head))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = True
        autofit_pct(t)
        set_borders(t, GREY_LINE, 4)
        cell_margins(t)
        # header
        hr_ = t.rows[0]
        repeat_header(hr_); cant_split(hr_)
        blank_head = not any(h.strip() for h in head)
        for k, h in enumerate(head):
            c = hr_.cells[k]
            if not blank_head: shade(c, NAVY)
            p = c.paragraphs[0]
            p.alignment = ALIGN[aligns[k]]
            p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
            emit_runs(p, h, base_bold=True,
                      base_color=(None if blank_head else WHITE), size=9.5)
        # body
        for n, row in enumerate(body):
            cells_ = t.add_row()
            cant_split(cells_)
            for k, v in enumerate(row):
                c = cells_.cells[k]
                if n % 2 == 1: shade(c, GREY_BG)
                p = c.paragraphs[0]
                p.alignment = ALIGN[aligns[k]]
                p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.1
                # bullet trong ô: tách bằng "• " hoặc "- "
                v2 = re.sub(r'<br\s*/?>', '\n', v)
                v2 = re.sub(r'(?<![\n^])\s*•\s*', '\n• ', v2)
                v2 = re.sub(r'\n\s*\n+', '\n', v2)
                emit_runs(p, v2, size=9.5)
        sp = self.doc.add_paragraph(); sp.paragraph_format.space_after = Pt(6)
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.line_spacing = 1.0
        sp.add_run().font.size = Pt(2)

    def image(self, fname, caption):
        path = os.path.join(MEDIA, fname)
        if not os.path.exists(path):
            return
        cx, cy = EXTENTS.get(fname, (Inches(4), Inches(3)))
        if cx > MAXW:
            cy = int(cy * MAXW / cx); cx = int(MAXW)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(2)
        keep_with_next(p)
        add_picture_any(self.doc, p, path, int(cx), int(cy))
        self.fig += 1
        cp = self.doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_after = Pt(10)
        r = cp.add_run('Hình %d. ' % self.fig)
        r.bold = True; r.font.size = Pt(8.5); r.font.name = FONT
        r.font.color.rgb = RGBColor.from_string(BLUE)
        r2 = cp.add_run(caption)
        r2.italic = True; r2.font.size = Pt(8.5); r2.font.name = FONT
        r2.font.color.rgb = RGBColor.from_string(GREY_TXT)

    # ---------- hộp thông tin ----------
    def box(self, lines):
        raw = '\n'.join(lines)
        # ảnh?
        if 'Hình minh hoạ' in raw or 'Hình minh họa' in raw:
            files = IMG_RE.findall(raw)
            desc = re.sub(r'.*?Hình minh ho[aạ]\s*:\s*', '', raw, flags=re.S)
            desc = IMG_RE.sub('', desc)
            desc = re.sub(r'[`*()]|\bcủa bản gốc\b|\bdùng lại\b', '', desc)
            desc = re.sub(r'[,;—–-]+', ' ', desc)
            desc = re.sub(r'\s+', ' ', desc).strip(' .,;—–-')
            ctx = re.sub(r'\s+', ' ', re.sub(r'[`*]', '', self.ctx or '')).strip()
            ctx = re.sub(r'^\d+\.\d*\.?\s*', '', ctx)
            for k, fn in enumerate(files):
                if desc:
                    cap = desc[0].upper() + desc[1:]
                    if len(files) > 1:
                        cap = '%s (%d/%d)' % (cap, k + 1, len(files))
                elif ctx:
                    cap = ctx if len(files) == 1 else '%s (%d/%d)' % (ctx, k + 1, len(files))
                else:
                    cap = 'Minh hoạ cho nội dung bên trên'
                self.image(fn, cap)
            return
        blocks = parse(raw)
        title = None
        if blocks and blocks[0][0] in ('h1', 'h2', 'h3', 'h4'):
            title = blocks[0][1]; blocks = blocks[1:]
        kind = box_kind(title or raw[:120])
        icon, col, bg = BOXES.get(kind, DEFAULT_BOX)
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        left_accent(t, col); cell_margins(t, 120, 180, 120, 180)
        c = t.cell(0, 0); shade(c, bg); c.width = MAXW
        first = c.paragraphs[0]
        first._p.getparent().remove(first._p)
        if title:
            p = c.add_paragraph()
            p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(4)
            clean = re.sub(r'^\s*[^\w\[]*\s*', '', title).strip()
            ic = p.add_run(icon + '  ')
            ic.font.size = Pt(10.5); ic.font.name = 'Segoe UI Emoji'
            emit_runs(p, clean, base_bold=True, base_color=col, size=10.5)
        sub = Renderer.__new__(Renderer)
        sub.doc = c; sub.fig = self.fig; sub.ctx = self.ctx
        sub._in_box = True
        for kd, pl in blocks:
            sub.dispatch(kd, pl, in_box=True)
        self.fig = sub.fig
        # dọn đoạn rỗng cuối
        while len(c.paragraphs) > 1 and not c.paragraphs[-1].text.strip() \
                and not c.paragraphs[-1].runs:
            c.paragraphs[-1]._p.getparent().remove(c.paragraphs[-1]._p)
        cant_split(t.rows[0])
        sp = self.doc.add_paragraph(); sp.paragraph_format.space_after = Pt(6)
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.line_spacing = 1.0
        sp.add_run().font.size = Pt(2)

    # ---------- dispatch ----------
    def dispatch(self, kind, payload, in_box=False):
        if kind == 'h1':
            (self.h4 if in_box else self.h1)(payload)
        elif kind == 'h2':
            (self.h4 if in_box else self.h2)(payload)
        elif kind == 'h3':
            (self.h4 if in_box else self.h3)(payload)
        elif kind in ('h4', 'h5', 'h6'):
            self.h4(payload)
        elif kind == 'p':
            self.para(payload)
        elif kind == 'list':
            self.bullets(payload)
        elif kind == 'table':
            self.table(payload)
        elif kind == 'code':
            self.code(payload)
        elif kind == 'quote':
            self.box(payload)
        elif kind == 'hr':
            pass

    def render(self, md, first_h1_break=True):
        blocks = parse(md)
        seen_h1 = False
        n = 0
        while n < len(blocks):
            kind, payload = blocks[n]
            if kind == 'h1':
                self.h1(payload, page_break_before=(first_h1_break or seen_h1))
                seen_h1 = True
                # phụ đề: h3 ngay sau h1
                if n + 1 < len(blocks) and blocks[n + 1][0] == 'h3':
                    st = blocks[n + 1][1]
                    if not st.startswith('▌'):
                        p = self.doc.add_paragraph()
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(10)
                        emit_runs(p, st, base_color=GREY_TXT, size=11, italic=True)
                        n += 1
                n += 1; continue
            self.dispatch(kind, payload)
            n += 1
