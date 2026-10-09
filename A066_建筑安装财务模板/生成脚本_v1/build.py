# -*- coding: utf-8 -*-
"""A066 建筑安装财务模板 · 生成脚本

跑法：python3 build.py
输入：../参考/ 里你发来的两本原表（只读，当演示数据装进去）
输出：../A066_建筑安装财务模板.xlsx

做法：openpyxl 生成 → 一整列同样的公式改成共享公式（文件小、打开快）→ 复制一份让 LibreOffice 整本重算、数报错
      → 把算出来的数写回成品当缓存（手机、微信预览也能看到数；Excel/WPS 改了哪格就重算哪格）。成品不经 LibreOffice 存盘。
"""
import glob
import json
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from share import share
from inject_cache import inject
import make

ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, 'A066_建筑安装财务模板.xlsx')
CALC = os.path.join(HERE, '_calc')


def find_recalc():
    cands = [os.environ.get('RECALC_PY', ''), '/mnt/skills/public/xlsx/scripts/recalc.py']
    cands += glob.glob(os.path.expanduser('~/.claude/skills/**/xlsx/scripts/recalc.py'), recursive=True)
    for c in cands:
        if c and os.path.exists(c):
            return c


def lo_recalc(path, timeout=3000):
    rc = find_recalc()
    if not rc:
        print('✗ 没找到 recalc.py（LibreOffice 重算脚本），成品里没有缓存值——Excel/WPS 打开时会自己算')
        return None
    r = subprocess.run([sys.executable, rc, path, str(timeout), '--force'], capture_output=True, text=True,
                       timeout=timeout + 300)
    out = r.stdout
    i, j = out.find('{'), out.rfind('}')
    return json.loads(out[i:j + 1]) if i >= 0 else {'error': out or r.stderr}


def main(out=OUT, scenario=None):
    wb = make.build(scenario=scenario)
    wb.save(out)
    s = share(out)
    print(f'① 生成 {os.path.basename(out)}：共享公式 {sum(s.values())} 格')
    os.makedirs(CALC, exist_ok=True)
    calc = os.path.join(CALC, os.path.basename(out).replace('.xlsx', '_算数.xlsx'))
    shutil.copy(out, calc)
    st = lo_recalc(calc)
    if st is None:
        return 0
    if 'error' in st:
        print('✗ 重算没跑成：', str(st['error'])[:800])
        return 2
    print(f"② LibreOffice 重算：公式 {st['total_formulas']} 个，报错 {st['total_errors']} 处")
    for k, v in (st.get('error_summary') or {}).items():
        print(f'    {k} ×{v["count"]}  {v["locations"][:12]}')
    n, _ = inject(out, calc)
    print(f'③ 写回缓存 {n} 格；{os.path.getsize(out) / 1e6:.1f} MB')
    return 0 if st['total_errors'] == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
