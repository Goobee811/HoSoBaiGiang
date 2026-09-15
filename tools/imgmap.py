import sys, os
import docx, re
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table
d = docx.Document((sys.argv[1] if len(sys.argv)>1 else os.path.join(os.path.dirname(os.path.abspath(__file__)),'..','HSBG.docx')))
rels = {r.rId: r.target_ref for r in d.part.rels.values()}
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
cur_bai='(đầu tài liệu)'
pi=0
def walk(el, parent):
    global pi, cur_bai
    for child in el.iterchildren():
        if child.tag==qn('w:p'):
            pi+=1
            p=Paragraph(child,parent)
            t=p.text.strip()
            if t and (t.upper().startswith('BÀI ') or t.upper().startswith('CHƯƠNG') or t.upper().startswith('QUY ')):
                cur_bai=t[:60]
            ids=[]
            for b in child.iter(qn('a:blip')):
                ids.append(b.get(R+'embed') or b.get(R+'link'))
            for im in child.iter('{urn:schemas-microsoft-com:vml}imagedata'):
                ids.append(im.get(R+'id'))
            for rid in ids:
                tgt=rels.get(rid,'?')
                print(f'{tgt.split("/")[-1]:<15} §P{pi:<4} [{cur_bai}]  ctx="{t[:70]}"')
        elif child.tag==qn('w:tbl'):
            walk(child, parent)
        elif child.tag==qn('w:tr') or child.tag==qn('w:tc'):
            walk(child, parent)
walk(d.element.body, d)
