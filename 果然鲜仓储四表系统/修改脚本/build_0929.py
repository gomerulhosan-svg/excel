# -*- coding: utf-8 -*-
"""0929 这一轮：在 0928 成品上改《01》《03》（见 m0929.py），《02》《04》不动。

   ① 《01》库存等级接口加「单箱净重KG」→ 副本上 LibreOffice 重算 → 数写回成品；
   ② 《03》借款汇总、库存价值与欠款比对 → 按新的《01》刷新跨文件链接缓存 → 副本重算 → 数写回成品。
   成品本身不经 LibreOffice 存盘（字体、格式、批注、下拉保持原样）。

跑法：python3 build_0929.py
"""
import os, sys, shutil, zipfile, re
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
OUT = os.path.join(HERE, '..')
CALC = os.path.join(OUT, '_calc')
NAMES = {'01': '01_水果进销存台账模板.xlsx', '02': '02_物料与周转物台账模板.xlsx',
         '03': '03_财务账套与报表模板.xlsx', '04': '04_综合查询对账单模板.xlsx'}
out = lambda k: os.path.join(OUT, NAMES[k])

from build_0928 import check_links, fonts, recalc, hidden_in_merges, cell_present
from inject_cache import inject
from refresh_link_cache import refresh, link_targets
import m0929


def keep_hidden(before, after):
    for sh, ref in sorted(hidden_in_merges(before)):
        if not cell_present(after, sh, ref):
            raise SystemExit(f'合并格里藏的值丢了：{sh}!{ref}')


if __name__ == '__main__':
    os.makedirs(CALC, exist_ok=True)
    print('① 改《01》库存等级接口')
    bak01 = os.path.join(CALC, 'before_01.xlsx'); shutil.copy(out('01'), bak01)
    wb = openpyxl.load_workbook(out('01'))
    m0929.apply_01(wb)
    wb.save(out('01'))
    keep_hidden(bak01, out('01'))
    fonts(out('01'))
    c01 = os.path.join(CALC, NAMES['01']); shutil.copy(out('01'), c01)
    recalc(c01)
    n, s = inject(out('01'), c01)
    print(f'   《01》写回缓存 {n} 格')

    print('② 改《03》借款汇总、库存价值与欠款比对')
    bak03 = os.path.join(CALC, 'before_03.xlsx'); shutil.copy(out('03'), bak03)
    wb = openpyxl.load_workbook(out('03'))
    m0929.apply_03(wb)
    wb.save(out('03'))
    check_links(out('03'))
    keep_hidden(bak03, out('03'))
    fonts(out('03'))
    # 跨文件链接缓存按新的《01》《02》刷新（新加的 I 列也进缓存）
    srcmap = {}
    for idx, (part, tgt) in link_targets(out('03')).items():
        base = tgt.replace('\\', '/').split('/')[-1]
        for k, nm in NAMES.items():
            if nm == base:
                srcmap[tgt] = c01 if k == '01' else out(k)
    refresh(out('03'), srcmap)
    for k in NAMES:
        if k != '01':
            shutil.copy(out(k), os.path.join(CALC, NAMES[k]))
    recalc(os.path.join(CALC, NAMES['03']))
    n, s = inject(out('03'), os.path.join(CALC, NAMES['03']))
    print(f'   《03》写回缓存 {n} 格')
    if '--keep' not in sys.argv:
        shutil.rmtree(CALC, ignore_errors=True)
    print('完成')
