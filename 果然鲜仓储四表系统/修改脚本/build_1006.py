# -*- coding: utf-8 -*-
"""10/06 这一轮：以你 10 月 6 日发来的四本（存档在 ../参考/1006上传/）为底，改出成品。

   《01》m1006_01 → 存盘 → 副本上 LibreOffice 整册重算 → 数写回成品
   《02》m1006_02 → 存盘 → 外链缓存按新《01》刷新 → 重算 → 写回
   《03》m1006_03a（账务骨架）、m1006_03b（报表）→ 存盘 → 外链缓存按新《01》《02》刷新 → 重算 → 写回
   《04》m1006_04 → 存盘 → 外链缓存按新《01》《02》刷新 → 重算 → 写回
   成品本身不经 LibreOffice 存盘（字体、格式、批注、下拉保持原样），只把算好的数塞进公式格的缓存。

跑法：python3 build_1006.py [--keep]      （--keep 保留 ../_calc 里的重算副本，便于核对）
"""
import os, sys, shutil, zipfile, re, json
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
SRC = os.path.join(HERE, '..', '参考', '1006上传')
OUT = os.path.join(HERE, '..')
CALC = os.path.join(OUT, '_calc')
NAMES = {'01': '01_水果进销存台账模板.xlsx', '02': '02_物料与周转物台账模板.xlsx',
         '03': '03_财务账套与报表模板.xlsx', '04': '04_综合查询对账单模板.xlsx'}
src = lambda k: os.path.join(SRC, NAMES[k])
out = lambda k: os.path.join(OUT, NAMES[k])
calc = lambda k: os.path.join(CALC, NAMES[k])

from build_0928 import fonts, recalc, hidden_in_merges, cell_present
from inject_cache import inject
from refresh_link_cache import refresh, link_targets
import m1006_01, m1006_02, m1006_03a, m1006_03b, m1006_04, m1006_99


def ext_links(path):
    z = zipfile.ZipFile(path)
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    n = rels.count('relationships/externalLink"')
    z.close()
    return n


def keep_hidden(before, after):
    for sh, ref in sorted(hidden_in_merges(before)):
        try:
            ok = cell_present(after, sh, ref)
        except KeyError:
            ok = True                      # 那张表这一轮删掉了（果然鲜销售汇总）
        if not ok:
            raise SystemExit(f'合并格里藏的值丢了：{sh}!{ref}')


def unhide_cols(path):
    """只给重算副本用：LibreOffice 存盘时会把「隐藏列」尾部的公式格截掉（少算 12 万多格），先把列都显示出来再算"""
    tmp = path + '.tmp'
    zi = zipfile.ZipFile(path)
    zo = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for it in zi.infolist():
        d = zi.read(it.filename)
        if it.filename.startswith('xl/worksheets/sheet') and it.filename.endswith('.xml'):
            d = re.sub(rb'(<col\b[^>]*?)\s+hidden="(?:1|true)"', rb'\1', d)
        zo.writestr(it, d)
    zo.close(); zi.close()
    shutil.move(tmp, path)


def count_formulas(path):
    """成品里每个公式格（含共享公式的子格）都数一遍"""
    from lxml import etree
    NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    z = zipfile.ZipFile(path)
    n = 0
    for nm in z.namelist():
        if nm.startswith('xl/worksheets/sheet') and nm.endswith('.xml'):
            for _, el in etree.iterparse(z.open(nm), tag=NS + 'f'):
                n += 1
                el.clear()
    z.close()
    return n


def lint(path):
    """版面体检：<col> 不能重叠、不能超过 XFD；同一 sheetView 里 selection 不能重复；冻结不能冻住一大片"""
    from lxml import etree
    NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    z = zipfile.ZipFile(path)
    wbx = etree.fromstring(z.read('xl/workbook.xml'))
    rels = {r.get('Id'): r.get('Target') for r in etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    probs = []
    for sh in wbx.iter(NS + 'sheet'):
        t = rels[sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')].lstrip('/')
        t = t if t.startswith('xl/') else 'xl/' + t
        head = z.read(t)
        head = head[:head.find(b'<sheetData')] + b'</worksheet>' if b'<sheetData' in head else head
        root = etree.fromstring(head)
        name = sh.get('name')
        spans = sorted((int(c.get('min')), int(c.get('max'))) for c in root.iter(NS + 'col'))
        for (a1, b1), (a2, b2) in zip(spans, spans[1:]):
            if a2 <= b1:
                probs.append(f'{name}: <col> 重叠 {a1}-{b1} / {a2}-{b2}')
        if any(b > 16384 for _, b in spans):
            probs.append(f'{name}: <col> 超过 XFD')
        for sv in root.iter(NS + 'sheetView'):
            panes = [s.get('pane') for s in sv.iter(NS + 'selection')]
            if len(panes) != len(set(panes)) or len(panes) > 4:
                probs.append(f'{name}: selection 重复 {panes}')
            for p in sv.iter(NS + 'pane'):
                if p.get('state') in ('frozen', 'frozenSplit') and float(p.get('ySplit') or 0) > 15:
                    probs.append(f'{name}: 冻结了 {p.get("ySplit")} 行')
    z.close()
    return probs


def refresh_links(k):
    """把成品 k 的外链缓存按已经算好的新《01》《02》重写（按文件名对上，绝对路径也认）"""
    srcmap = {}
    for idx, (part, tgt) in link_targets(out(k)).items():
        base = tgt.replace('\\', '/').split('/')[-1]
        for kk, nm in NAMES.items():
            if nm == base:
                srcmap[tgt] = out(kk)
    refresh(out(k), srcmap)


def step(k, mods):
    print(f'== 《{k}》')
    shutil.copy(src(k), out(k))
    n0 = ext_links(src(k))
    wb = openpyxl.load_workbook(src(k))
    for m, fn in mods:
        getattr(m, fn)(wb)
    wb.save(out(k))
    for m, fn in mods:
        if hasattr(m, 'post'):
            m.post(out(k))
    if ext_links(out(k)) != n0:
        raise SystemExit(f'外链关系数变了：{NAMES[k]} {n0} → {ext_links(out(k))}')
    keep_hidden(src(k), out(k))
    fonts(out(k))
    if n0:
        refresh_links(k)
    probs = lint(out(k))
    for p in probs:
        print('   ⚠ 版面：', p)
    if any('冻结了' in p or 'selection 重复' in p for p in probs):
        raise SystemExit(f'{NAMES[k]} 版面有问题（冻结/选择区），先修模块')
    shutil.copy(out(k), calc(k))
    unhide_cols(calc(k))
    j = recalc(calc(k))
    nf = count_formulas(out(k))
    if nf != j.get('total_formulas'):
        raise SystemExit(f'{NAMES[k]} 公式数对不上：成品 {nf}，LibreOffice 算了 {j.get("total_formulas")}')
    n, s = inject(out(k), calc(k))
    print(f'   公式 {nf} 个全部算到；写回缓存 {n} 格（空 {s} 格）')
    return j


if __name__ == '__main__':
    os.makedirs(CALC, exist_ok=True)
    for k in NAMES:                       # 重算时四本都要在同一个文件夹（先放底稿，改完一本换一本）
        shutil.copy(src(k), calc(k))
    res = {}
    res['01'] = step('01', [(m1006_01, 'apply'), (m1006_99, 'apply_01')])
    res['02'] = step('02', [(m1006_02, 'apply')])
    res['03'] = step('03', [(m1006_03a, 'apply'), (m1006_03b, 'apply'), (m1006_99, 'apply_03')])
    res['04'] = step('04', [(m1006_04, 'apply'), (m1006_99, 'apply_04')])
    json.dump({k: {x: v.get(x) for x in ('status', 'total_formulas', 'total_errors', 'error_summary')}
               for k, v in res.items()}, open(os.path.join(CALC, 'recalc_summary.json'), 'w'), ensure_ascii=False, indent=1)
    if '--keep' not in sys.argv:
        shutil.rmtree(CALC, ignore_errors=True)
    print('\n完成，成品在', os.path.abspath(OUT))
