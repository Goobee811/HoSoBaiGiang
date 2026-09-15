import sys, os
import docx, sys
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

d = docx.Document((sys.argv[1] if len(sys.argv)>1 else os.path.join(os.path.dirname(os.path.abspath(__file__)),'..','HSBG.docx')))

def iter_block(parent):
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)

out = []
pi = 0
ti = 0
for b in iter_block(d):
    if isinstance(b, Paragraph):
        pi += 1
        txt = b.text.strip()
        style = b.style.name if b.style is not None else '?'
        # detect images
        imgs = b._p.findall('.//'+qn('a:blip'))
        n_img = len(imgs)
        # also pict/imagedata
        n_img += len(b._p.findall('.//{urn:schemas-microsoft-com:vml}imagedata'))
        marks = []
        if n_img: marks.append(f'[IMG x{n_img}]')
        # numbering
        if b._p.find('.//'+qn('w:numPr')) is not None: marks.append('[LIST]')
        if not txt and not n_img:
            continue
        out.append(f'§P{pi} <{style}> {" ".join(marks)} {txt}')
    else:
        ti += 1
        rows = len(b.rows); cols = len(b.columns)
        out.append(f'§TBL{ti} ({rows}x{cols})')
        for ri, row in enumerate(b.rows):
            cells = []
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen: continue
                seen.add(id(c._tc))
                cells.append(' '.join(c.text.split()))
            out.append('   | ' + ' | '.join(cells))

print('\n'.join(out))
print(f'\n=== TOTAL paragraphs={pi} tables={ti} ===', file=sys.stderr)
