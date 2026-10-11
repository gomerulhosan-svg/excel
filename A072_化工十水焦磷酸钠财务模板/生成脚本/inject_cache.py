# -*- coding: utf-8 -*-
"""把 LibreOffice 在「副本」上算出来的值，写回成品里每个公式格子的缓存值。

为什么不直接交 LibreOffice 存过的那份：它会把整本册子按自己的方式重写一遍
（样式、条件格式、数据验证、定义名称都可能被改动），原表的格式就不再是原样了。
所以成品始终是 openpyxl 生成的那份，只是把算好的数塞进 <v>，
这样手机/微信预览、不自动重算的查看器打开也能看到数；Excel/WPS 打开时照样全部重算一遍。

跑法：python3 inject_cache.py <成品.xlsx> <LibreOffice算过的副本.xlsx>
"""
import sys
import zipfile
import shutil
import datetime as dt
from lxml import etree
import openpyxl
from openpyxl.utils.datetime import to_excel

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ERRORS = {'#N/A', '#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'}


def sheet_parts(z):
    wbx = etree.fromstring(z.read('xl/workbook.xml'))
    rels = etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid2t = {r.get('Id'): r.get('Target') for r in rels}
    out = {}
    for s in wbx.iter(f'{{{NS}}}sheet'):
        t = rid2t[s.get(f'{{{RNS}}}id')]
        t = t.lstrip('/')
        out[s.get('name')] = t if t.startswith('xl/') else 'xl/' + t
    return out


def inject(target, valsrc, names=None):
    """names：{成品里的表名: 算数副本里的表名}，不给就同名对同名"""
    names = names or {}
    vals = openpyxl.load_workbook(valsrc, data_only=True)
    zv = zipfile.ZipFile(valsrc)
    vparts = sheet_parts(zv)

    def empty_str_cells(sheet):
        """副本里公式结果是空串的格子（t="str"、<v> 空）：openpyxl 读成 None，要按空串写回，不然预览当成 0"""
        out = set()
        if sheet not in vparts:
            return out
        for _, c in etree.iterparse(zv.open(vparts[sheet]), tag=f'{{{NS}}}c'):
            if c.get('t') == 'str':
                v = c.find(f'{{{NS}}}v')
                if v is None or not v.text:
                    out.add(c.get('r'))
            c.clear()
        return out
    tmp = target + '.tmp'
    zin = zipfile.ZipFile(target)
    parts = sheet_parts(zin)
    part2sheet = {v: k for k, v in parts.items()}
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    n_set = n_skip = 0
    for item in zin.infolist():
        data = zin.read(item.filename)
        sh = part2sheet.get(item.filename)
        src = names.get(sh, sh)
        if sh and src in vals.sheetnames:
            wsv = vals[src]
            empties = empty_str_cells(src)
            root = etree.fromstring(data)
            for c in root.iter(f'{{{NS}}}c'):
                f = c.find(f'{{{NS}}}f')
                if f is None:
                    continue
                v = c.find(f'{{{NS}}}v')
                if v is not None:
                    c.remove(v)
                val = wsv[c.get('r')].value
                if val is None:
                    if c.get('r') in empties:
                        c.set('t', 'str')
                        etree.SubElement(c, f'{{{NS}}}v')
                        n_set += 1
                    else:
                        c.attrib.pop('t', None)
                        n_skip += 1
                    continue
                v = etree.SubElement(c, f'{{{NS}}}v')
                if isinstance(val, bool):
                    c.set('t', 'b'); v.text = '1' if val else '0'
                elif isinstance(val, (int, float)):
                    c.attrib.pop('t', None); v.text = repr(float(val)) if isinstance(val, float) else str(val)
                elif isinstance(val, (dt.datetime, dt.date)):
                    c.attrib.pop('t', None); v.text = repr(float(to_excel(val)))
                elif isinstance(val, dt.time):
                    c.attrib.pop('t', None); v.text = repr(float(to_excel(val)))
                elif isinstance(val, str) and val in ERRORS:
                    c.set('t', 'e'); v.text = val
                else:
                    c.set('t', 'str'); v.text = str(val)
                n_set += 1
            data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(item, data)
    zout.close()
    zin.close()
    zv.close()
    shutil.move(tmp, target)
    return n_set, n_skip


if __name__ == '__main__':
    n, s = inject(sys.argv[1], sys.argv[2])
    print(f'✓ 写入缓存值 {n} 格（空值 {s} 格）')
