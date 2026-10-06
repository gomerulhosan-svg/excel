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
import m1006_01, m1006_02, m1006_03a, m1006_03b, m1006_04


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
    shutil.copy(out(k), calc(k))
    j = recalc(calc(k))
    n, s = inject(out(k), calc(k))
    print(f'   写回缓存 {n} 格（空 {s} 格）')
    return j


if __name__ == '__main__':
    os.makedirs(CALC, exist_ok=True)
    for k in NAMES:                       # 重算时四本都要在同一个文件夹（先放底稿，改完一本换一本）
        shutil.copy(src(k), calc(k))
    res = {}
    res['01'] = step('01', [(m1006_01, 'apply')])
    res['02'] = step('02', [(m1006_02, 'apply')])
    res['03'] = step('03', [(m1006_03a, 'apply'), (m1006_03b, 'apply')])
    res['04'] = step('04', [(m1006_04, 'apply')])
    json.dump({k: {x: v.get(x) for x in ('status', 'total_formulas', 'total_errors', 'error_summary')}
               for k, v in res.items()}, open(os.path.join(CALC, 'recalc_summary.json'), 'w'), ensure_ascii=False, indent=1)
    if '--keep' not in sys.argv:
        shutil.rmtree(CALC, ignore_errors=True)
    print('\n完成，成品在', os.path.abspath(OUT))
