# -*- coding: utf-8 -*-
"""把报表本里「跨文件链接的缓存值」按最新的录入本重写一遍（A065 refresh_link_cache.py 搬过来，改了三处，见下）。

Excel / WPS 打开报表本时会问要不要更新链接，更新了就是最新的；但没更新之前（或者录入本不在同一个文件夹），
用的是报表本文件里存着的那份缓存。生成完不把缓存写好，单独打开报表本第一眼就是空的。本脚本把这份缓存按录入本重新写进去。

跟 A065 版的区别：
  1. 名称里的外部区域（MIRROR_MODE='name'：录_流水_A＝[1]现金流水总表!$A$6:$A$3965）也展开成格子写缓存；
  2. 空格子不写（缓存只跟有数的格子一样大；LibreOffice 实测缓存里没有的格子当空格，取数公式 IF(x="","",x) 两样都是空）；
  3. 扫报表本的公式不用 openpyxl 整本载入（报表本 90 万个公式，载入要几分钟），直接流式读 sheet XML；
     公式里的表名带引号的（'[1]表名'!A6）也认。

跑法：python3 -I refresh_link_cache.py <报表本.xlsx> <录入本.xlsx>
      源文件按 externalLink1、2… 的顺序给，或者用「文件名=路径」显式指定。
"""
import datetime as dt
import os
import re
import shutil
import sys
import zipfile

import openpyxl
from lxml import etree
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL, range_boundaries

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
CELL = r'\$?[A-Z]{1,3}\$?\d+'
REF = re.compile(r"(?:'\[(\d+)\]((?:[^']|'')+)'|\[(\d+)\]([^!'\[\]\s,()+\-*/&=<>\"]+))!(" + CELL + r")(?::(" + CELL + r"))?")
NAME_REF = re.compile(r"^(?:'\[(\d+)\]((?:[^']|'')+)'|\[(\d+)\]([^!']+))!(" + CELL + r"(?::" + CELL + r")?)$")


def xml_escape(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _add(out, idx, sheet, a, b=None):
    st = out.setdefault(int(idx), {}).setdefault(sheet.replace("''", "'"), set())
    c0, r0, c1, r1 = range_boundaries((a + (':' + b if b else '')).replace('$', ''))
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            st.add(f'{CL(c)}{r}')


def collect_refs(path):
    """报表本里引用的外部格子：{链接序号: {表名: {格子}}}（公式里的 [n]表!格 / 区域，名称里的 [n]表!区域）"""
    out = {}
    z = zipfile.ZipFile(path)
    for part in z.namelist():
        if not (part.startswith('xl/worksheets/sheet') and part.endswith('.xml')):
            continue
        for _, el in etree.iterparse(z.open(part), tag=NS + 'f', huge_tree=True):
            t = el.text
            if t and '[' in t:
                body = ''.join(t.split('"')[0::2])
                for m in REF.finditer(body):
                    _add(out, m.group(1) or m.group(3), m.group(2) or m.group(4), m.group(5), m.group(6))
            el.clear()
    wbx = etree.fromstring(z.read('xl/workbook.xml'))
    for dn in wbx.iter(NS + 'definedName'):
        m = NAME_REF.match((dn.text or '').strip())
        if m:
            a, _, b = m.group(5).partition(':')
            _add(out, m.group(1) or m.group(3), m.group(2) or m.group(4), a, b or None)
    z.close()
    return out


def link_targets(path):
    """externalLink 序号 → (part 文件名, 被链接的文件名)"""
    z = zipfile.ZipFile(path)
    wbxml = z.read('xl/workbook.xml').decode('utf8')
    m = re.search(r'<externalReferences>(.*?)</externalReferences>', wbxml, re.S)
    ids = re.findall(r'r:id="(rId\d+)"', m.group(1)) if m else []
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rel = {}
    for tag in re.findall(r'<Relationship\b[^>]*>', rels):      # 属性顺序不固定（openpyxl 存的是 Type/Target/Id）
        i, t = re.search(r'\bId="([^"]+)"', tag), re.search(r'\bTarget="([^"]+)"', tag)
        if i and t:
            rel[i.group(1)] = t.group(1)
    out = {}
    for k, rid in enumerate(ids, 1):
        part = rel[rid].replace('../', '')
        nm = part.split('/')[-1]
        tgt = re.search(rb'Target="([^"]+)"', z.read('xl/externalLinks/_rels/' + nm + '.rels')).group(1).decode()
        out[k] = (nm, tgt)
    z.close()
    return out


def build_sheetdata(src_path, sheets_cells):
    """从源文件（录入本，带缓存值）读值，拼出 <sheetNames> 与 <sheetDataSet>；空格子不写"""
    wb = openpyxl.load_workbook(src_path, data_only=True)
    names = wb.sheetnames
    parts = []
    n = 0
    for sid, nm in enumerate(names):
        cells = sheets_cells.get(nm)
        if not cells:
            parts.append(f'<sheetData sheetId="{sid}" refreshError="1"/>')
            continue
        ws = wb[nm]
        byrow = {}
        for a in cells:
            m = re.match(r'([A-Z]+)(\d+)$', a)
            byrow.setdefault(int(m.group(2)), []).append((CI(m.group(1)), a))
        rows = []
        for r in sorted(byrow):
            cs = []
            for _, a in sorted(byrow[r]):
                v = ws[a].value
                if v is None or v == '':
                    continue
                if isinstance(v, bool):
                    cs.append(f'<cell r="{a}" t="b"><v>{1 if v else 0}</v></cell>')
                elif isinstance(v, (int, float)):
                    cs.append(f'<cell r="{a}"><v>{v!r}</v></cell>')
                elif isinstance(v, (dt.datetime, dt.date)):
                    d0 = dt.datetime(1899, 12, 30)
                    d = v if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day)
                    cs.append(f'<cell r="{a}"><v>{(d - d0).days + (d - d0).seconds / 86400.0!r}</v></cell>')
                elif isinstance(v, str) and v.startswith('#') and v in ('#N/A', '#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'):
                    cs.append(f'<cell r="{a}" t="e"><v>{v}</v></cell>')
                else:
                    cs.append(f'<cell r="{a}" t="str"><v>{xml_escape(v)}</v></cell>')
            if cs:
                n += len(cs)
                rows.append(f'<row r="{r}">{"".join(cs)}</row>')
        parts.append(f'<sheetData sheetId="{sid}">{"".join(rows)}</sheetData>')
    wb.close()
    sn = ''.join(f'<sheetName val="{xml_escape(x)}"/>' for x in names)
    return f'<sheetNames>{sn}</sheetNames>', f'<sheetDataSet>{"".join(parts)}</sheetDataSet>', n


def refresh(target, srcmap, log=print):
    """srcmap：{链接里写的文件名: 源文件路径}。返回写进缓存的格子数"""
    refs = collect_refs(target)
    links = link_targets(target)
    newparts = {}
    tot = 0
    for idx, (partname, tgtfile) in links.items():
        src = srcmap.get(tgtfile) or srcmap.get(os.path.basename(tgtfile))
        if not src:
            log(f'   [{idx}] {tgtfile} 没给源文件，跳过')
            continue
        sn, sd, n = build_sheetdata(src, refs.get(idx, {}))
        newparts['xl/externalLinks/' + partname] = (sn, sd)
        ncell = sum(len(v) for v in refs.get(idx, {}).values())
        tot += n
        log(f'   外部链接 [{idx}] {tgtfile} ← {os.path.basename(src)}：引用 {ncell:,} 格，有数的 {n:,} 格写进缓存')
    tmp = target + '.tmp'
    zin = zipfile.ZipFile(target)
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename in newparts:
            t = data.decode('utf8')
            sn, sd = newparts[it.filename]
            t = re.sub(r'<sheetNames>.*?</sheetNames>|<sheetNames/>', lambda m: sn, t, flags=re.S)
            if '<sheetDataSet>' in t or '<sheetDataSet/>' in t:
                t = re.sub(r'<sheetDataSet>.*?</sheetDataSet>|<sheetDataSet/>', lambda m: sd, t, flags=re.S)
            else:
                t = t.replace('</externalBook>', sd + '</externalBook>')
            data = t.encode('utf8')
        zout.writestr(it, data)
    zin.close()
    zout.close()
    shutil.move(tmp, target)
    return tot


if __name__ == '__main__':
    target = sys.argv[1]
    srcmap = {}
    for a in sys.argv[2:]:
        if '=' in a:
            k, v = a.split('=', 1)
            srcmap[k] = v
        else:
            srcmap[os.path.basename(a)] = a
    print('刷新', os.path.basename(target))
    refresh(target, srcmap)
    print('   完成')
