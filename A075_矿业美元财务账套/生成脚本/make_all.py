# -*- coding: utf-8 -*-
"""一条龙（两本联动版）：
   一本版（内存）→ 存一本版（算数用）、报表本、录入本（build.build_split）→ 三本都压共享公式（含「[」的公式不压）
   → 两本 WPS 体检（WPS 存后会多出来的 XML 超过 1 MB 就停）→ LibreOffice（关多线程）算一本版（报表本的数从它来）、算录入本
   → 数报错、核对公式个数 → 算好的值写回两本当缓存 → 报表本的外部链接缓存按录入本重写（refresh_link_cache）。
   成品不经 LibreOffice 存盘。

    python3 -I make_all.py [输出目录] [--keep]     （--keep 保留 _calc 里的一本版和重算副本）

默认输出 ../A075-1_录入本.xlsx 和 ../A075-2_报表本.xlsx（同一个文件夹；报表本的外部链接写的是相对文件名）。
一本版和 LibreOffice 重算副本在 <输出目录>/_calc/（_calc/ 不进 git）。
"""
import os
import shutil
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import build
import p_wps
from common import WB1_FILE, WB2_FILE
from share import share
from inject_cache import inject
from lo_recalc import recalc, count, unhide_cols
from refresh_link_cache import refresh

ROOT = os.path.dirname(HERE)
# 不放行任何报错（提速版起一本版 0 报错）
KNOWN_ERRORS = set()


def _calc(path, calc_dir, tag, log):
    """副本显示隐藏列 → LibreOffice 关多线程整本重算 → (副本路径, 状态)"""
    calc = os.path.join(calc_dir, os.path.splitext(os.path.basename(path))[0] + '_算数.xlsx')
    shutil.copy(path, calc)
    unhide_cols(calc)
    st = recalc(calc)
    if 'error' in st:
        log(f'✗ {tag}重算没跑成：' + st['error'])
        return calc, st
    log(f"   {tag}：LibreOffice 重算公式 {st['formulas']:,} 个，报错 {st['errors']} 处，{st['wall_seconds']} 秒")
    for k, v in st['error_cells'].items():
        log(f'    {k} {v[:15]}')
    return calc, st


def _ok(st, nf, tag, log):
    if 'error' in st:
        return False
    errs = {x for v in st['error_cells'].values() for x in v}
    ok = st['errors'] == 0 or (errs <= KNOWN_ERRORS and st['errors'] == len(errs))
    if st['formulas'] != nf:
        log(f"✗ {tag}公式个数对不上：成品 {nf:,}，LibreOffice 存回来 {st['formulas']:,}")
        ok = False
    return ok


def main(out_dir=ROOT, keep=False, log=print, tweak=None):
    t0 = time.time()
    out_dir = os.path.abspath(out_dir)
    calc_dir = os.path.join(out_dir, '_calc')
    os.makedirs(calc_dir, exist_ok=True)
    one = os.path.join(calc_dir, 'A075_一本版.xlsx')
    out1, out2, _ = build.build_split(out_dir, log=log, tweak=tweak, onebook=one)
    nf = {}
    for p, tag in ((one, '一本版'), (out2, '报表本'), (out1, '录入本')):
        s = share(p)
        nf[p], _, _ = count(p)
        log(f'① {tag}压共享公式：公式 {nf[p]:,} 个（共享公式 {sum(s.values()):,} 格），{os.path.getsize(p) / 1e6:.1f} MB')
    ok_w = True
    for p, tag in ((out1, '录入本'), (out2, '报表本')):
        log(f'   【{tag}】')
        w, _, _ = p_wps.report(p, log=log)          # WPS 体检：成列的公式里不准有「表名!」、往下撑的区域
        ok_w = ok_w and w
    if not ok_w:
        return 3
    log('② LibreOffice 重算（关多线程）')
    calc1, st1 = _calc(one, calc_dir, '一本版（报表本的数从它来）', log)
    calc2, st2 = _calc(out1, calc_dir, '录入本', log)
    ok = _ok(st1, nf[one], '一本版', log) and _ok(st2, nf[out1], '录入本', log)
    if 'error' in st1 or 'error' in st2:
        return 2
    n2, _ = inject(out2, calc1)
    n1, _ = inject(out1, calc2)
    log(f'③ 写回缓存：报表本 {n2:,} 格、录入本 {n1:,} 格')
    refresh(out2, {WB1_FILE: out1}, log=log)
    log(f'④ {WB1_FILE} {os.path.getsize(out1) / 1e6:.1f} MB、{WB2_FILE} {os.path.getsize(out2) / 1e6:.1f} MB；共 {time.time() - t0:.0f} 秒')
    if not keep:
        for p in (one, calc1, calc2):
            if os.path.exists(p):
                os.remove(p)
    return 0 if ok else 1


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if args and args[0].lower().endswith('.xlsx'):
        sys.exit('参数是输出目录（两本都放进去），不是 .xlsx 文件名')
    sys.exit(main(args[0] if args else ROOT, keep='--keep' in sys.argv))
