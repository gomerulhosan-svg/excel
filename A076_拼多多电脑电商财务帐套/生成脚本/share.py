# -*- coding: utf-8 -*-
"""把一整列「同一个公式往下拖」的格子改成 Excel 自己用的「共享公式」写法：只在第一格存公式全文，
   下面的格子只写「跟第一格一样」。文件小很多，Excel / WPS 打开、存盘都快。

   判断「同一个公式」：把公式里相对的行号换成「比本行差几行」（字符串里的内容不动、$ 锁定的行号不动），
   换完一样的就是同一个公式——Excel 就是这么往下推共享公式的。
   含跨工作簿链接（[1]…）的公式不共享（保守起见）。"""
import re
import sys
import zipfile
import shutil
from lxml import etree

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REF = re.compile(r'(?<![A-Za-z0-9_\.一-鿿])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![0-9A-Za-z_\(一-鿿])')
CELL = re.compile(r'([A-Z]+)(\d+)$')


def normalize(f, row):
    parts = f.split('"')
    for i in range(0, len(parts), 2):          # 双引号外面的
        parts[i] = REF.sub(lambda m: m.group(0) if m.group(3) else f'{m.group(1)}{m.group(2)}R{int(m.group(4)) - row}', parts[i])
    return '"'.join(parts)


def sheet_parts(z):
    wbx = etree.fromstring(z.read('xl/workbook.xml'))
    rels = etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid2t = {r.get('Id'): r.get('Target') for r in rels}
    out = {}
    for s in wbx.iter(f'{{{NS}}}sheet'):
        t = rid2t[s.get(f'{{{RNS}}}id')].lstrip('/')
        out[t if t.startswith('xl/') else 'xl/' + t] = s.get('name')
    return out


def share(path, min_run=4):
    tmp = path + '.share'
    zin = zipfile.ZipFile(path)
    parts = sheet_parts(zin)
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    stats = {}
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename in parts:
            root = etree.fromstring(data)
            cols = {}
            for c in root.iter(f'{{{NS}}}c'):
                f = c.find(f'{{{NS}}}f')
                if f is None or f.get('t') or not f.text or '[' in f.text:
                    continue
                m = CELL.match(c.get('r'))
                col, row = m.group(1), int(m.group(2))
                cols.setdefault(col, []).append((row, c, f))
            si = 0
            nshared = 0
            for col, cells in cols.items():
                cells.sort(key=lambda x: x[0])
                i = 0
                while i < len(cells):
                    r0, c0, f0 = cells[i]
                    key = normalize(f0.text, r0)
                    j = i + 1
                    while (j < len(cells) and cells[j][0] == cells[j - 1][0] + 1
                           and normalize(cells[j][2].text, cells[j][0]) == key):
                        j += 1
                    if j - i >= min_run:
                        r1 = cells[j - 1][0]
                        f0.set('t', 'shared')
                        f0.set('ref', f'{col}{r0}:{col}{r1}')
                        f0.set('si', str(si))
                        for _, _, fk in cells[i + 1:j]:
                            fk.text = None
                            fk.set('t', 'shared')
                            fk.set('si', str(si))
                        si += 1
                        nshared += j - i
                    i = j
            stats[parts[item.filename]] = nshared
            data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(item, data)
    zout.close()
    zin.close()
    shutil.move(tmp, path)
    return stats


if __name__ == '__main__':
    print(share(sys.argv[1]))
