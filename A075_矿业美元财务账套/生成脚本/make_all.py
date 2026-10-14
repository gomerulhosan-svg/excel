# -*- coding: utf-8 -*-
"""一条龙（照 A062 修改脚本 build_1008.py 的流程）：
   生成 → 同一列往下拖的公式压成「共享公式」→ 复制一份、把隐藏列都显示出来 → LibreOffice（关多线程）整本重算、
   数报错、核对公式个数 → 把算好的值写回成品当缓存。成品本身不经 LibreOffice 存盘。

    python3 -I make_all.py [输出.xlsx] [--keep]     （--keep 保留 _calc 里的重算副本）

默认输出 ../A075_矿业美元财务账套（提速版）.xlsx（成品；LibreOffice 重算用的副本在 ../_calc/，_calc/ 不进 git）。
"""
import os
import shutil
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import build
from share import share
from inject_cache import inject
from lo_recalc import recalc, count, unhide_cols

ROOT = os.path.dirname(HERE)
CALC = os.path.join(ROOT, '_calc')
# 原件里本来有 3 处 #VALUE!（资金账户余额表 K6、K7、K25：基础资料 J5:J7「对应科目」空着，J 取到空、K＝空−数）。
# 提速版把 K 改成「J 是空就空着」（s_ledger.build_acct），重算应该 0 报错；这里不再放行任何报错。
KNOWN_ERRORS = set()
OUT = os.path.join(ROOT, 'A075_矿业美元财务账套（提速版）.xlsx')


def main(out=OUT, keep=False, log=print):
    t0 = time.time()
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    build.build(out, log=log)
    s = share(out)
    nf, _, _ = count(out)
    log(f'① 生成并压共享公式：公式 {nf} 个（共享公式 {sum(s.values())} 格），{os.path.getsize(out) / 1e6:.1f} MB，{time.time() - t0:.0f} 秒')
    os.makedirs(CALC, exist_ok=True)
    calc = os.path.join(CALC, os.path.splitext(os.path.basename(out))[0] + '_算数.xlsx')
    shutil.copy(out, calc)
    unhide_cols(calc)
    st = recalc(calc)
    if 'error' in st:
        log('✗ 重算没跑成：' + st['error'])
        return 2
    log(f"② LibreOffice 重算（关多线程）：公式 {st['formulas']} 个，报错 {st['errors']} 处，{st['wall_seconds']} 秒")
    for k, v in st['error_cells'].items():
        log(f'    {k} {v[:15]}')
    errs = {x for v in st['error_cells'].values() for x in v}
    if errs and errs <= KNOWN_ERRORS and st['errors'] == len(errs):
        log(f'    （这 {len(errs)} 处是已知的）')
    ok = (st['errors'] == 0 or (errs <= KNOWN_ERRORS and st['errors'] == len(errs))) and st['formulas'] == nf
    if st['formulas'] != nf:
        log(f"✗ 公式个数对不上：成品 {nf}，LibreOffice 存回来 {st['formulas']}")
    n, _ = inject(out, calc)
    log(f'③ 写回缓存 {n} 格；{os.path.getsize(out) / 1e6:.1f} MB；共 {time.time() - t0:.0f} 秒')
    if not keep:
        os.remove(calc)
    return 0 if ok else 1


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    sys.exit(main(args[0] if args else OUT, keep='--keep' in sys.argv))
