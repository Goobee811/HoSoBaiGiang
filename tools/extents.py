import sys, os
import docx, json
from docx.oxml.ns import qn
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
WP='{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
d=docx.Document((sys.argv[1] if len(sys.argv)>1 else os.path.join(os.path.dirname(os.path.abspath(__file__)),'..','HSBG.docx')))
rels={r.rId:r.target_ref for r in d.part.rels.values()}
out={}
for p in d.element.body.iter(qn('w:p')):
    for anchor in list(p.iter(WP+'inline'))+list(p.iter(WP+'anchor')):
        ext=anchor.find(WP+'extent')
        blip=None
        for b in anchor.iter(qn('a:blip')): blip=b; break
        if ext is None or blip is None: continue
        rid=blip.get(R+'embed') or blip.get(R+'link')
        name=rels.get(rid,'').split('/')[-1]
        cx,cy=int(ext.get('cx')),int(ext.get('cy'))
        if name and name not in out: out[name]=(cx,cy)
# vml
for sh in d.element.body.iter('{urn:schemas-microsoft-com:vml}shape'):
    style=sh.get('style') or ''
    imd=sh.find('{urn:schemas-microsoft-com:vml}imagedata')
    if imd is None: continue
    name=rels.get(imd.get(R+'id'),'').split('/')[-1]
    import re
    w=re.search(r'width:([\d.]+)pt',style); h=re.search(r'height:([\d.]+)pt',style)
    if name and w and h and name not in out:
        out[name]=(int(float(w.group(1))*12700), int(float(h.group(1))*12700))
print(json.dumps(out,indent=0))
