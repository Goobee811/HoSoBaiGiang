# -*- coding: utf-8 -*-
"""Thêm khung viền trang vào file bìa, giữ nguyên toàn bộ nội dung."""
import re, shutil, sys, zipfile, os

SRC = sys.argv[1]
DST = sys.argv[2]
STYLE = sys.argv[3] if len(sys.argv) > 3 else 'thickThinSmallGap'
SZ = sys.argv[4] if len(sys.argv) > 4 else '24'
COLOR = sys.argv[5] if len(sys.argv) > 5 else '2B1663'

PGMAR_NEW = ('<w:pgMar w:top="1134" w:right="1276" w:bottom="1134" w:left="1276" '
             'w:header="624" w:footer="567" w:gutter="0"/>')

borders = '<w:pgBorders w:offsetFrom="page">' + ''.join(
    '<w:%s w:val="%s" w:sz="%s" w:space="24" w:color="%s"/>' % (e, STYLE, SZ, COLOR)
    for e in ('top', 'left', 'bottom', 'right')) + '</w:pgBorders>'

zin = zipfile.ZipFile(SRC)
names = zin.namelist()
data = {n: zin.read(n) for n in names}
zin.close()

xml = data['word/document.xml'].decode('utf8')
xml = re.sub(r'<w:pgMar[^/]*/>', PGMAR_NEW, xml)
# pgBorders phải đứng ngay sau pgMar theo lược đồ OOXML
xml = xml.replace(PGMAR_NEW, PGMAR_NEW + borders)
data['word/document.xml'] = xml.encode('utf8')

zout = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for n in names:
    zout.writestr(n, data[n])
zout.close()
print('ĐÃ TẠO:', os.path.basename(DST), os.path.getsize(DST) // 1024, 'KB',
      '| khung:', STYLE, 'sz', SZ, 'màu', COLOR)
