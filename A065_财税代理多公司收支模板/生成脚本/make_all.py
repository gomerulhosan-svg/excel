# -*- coding: utf-8 -*-
"""A065 一条龙：生成成品 → 在副本上用 LibreOffice 全部重算、数报错 → 把算好的值写回成品当缓存。

跑法：python3 make_all.py
成品始终是 openpyxl 生成的那份（LibreOffice 只碰副本，见 inject_cache.py 开头的说明）。
"""
import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build
import inject_cache


def find_recalc():
    cands = [os.environ.get('RECALC_PY', ''), '/mnt/skills/public/xlsx/scripts/recalc.py']
    cands += glob.glob(os.path.expanduser('~/.claude/skills/**/xlsx/scripts/recalc.py'), recursive=True)
    for c in cands:
        if c and os.path.exists(c):
            return c
    return None


def lo_recalc(path, timeout=1200):
    rc = find_recalc()
    if not rc:
        print('✗ 没找到 recalc.py（LibreOffice 重算脚本），跳过重算，成品里没有缓存值——Excel/WPS 打开时会自己算')
        return None
    r = subprocess.run([sys.executable, rc, path, str(timeout), "--force"], capture_output=True, text=True, timeout=timeout + 120)
    out = r.stdout
    i, j = out.find('{'), out.rfind('}')
    return json.loads(out[i:j + 1]) if i >= 0 else {'error': out or r.stderr}


def main():
    out = build.OUT
    build.main(out)
    tmpd = tempfile.mkdtemp(prefix='a065_')
    copy = os.path.join(tmpd, 'lo_copy.xlsx')
    shutil.copy(out, copy)
    st = lo_recalc(copy)
    if st is None:
        return 0
    if 'error' in st:
        print('✗ 重算没跑成：', st['error'][:500])
        return 2
    print(f"  LibreOffice 重算：公式 {st['total_formulas']} 个，报错 {st['total_errors']} 处")
    for k, v in (st.get('error_summary') or {}).items():
        print(f'    {k} ×{v["count"]}  {v["locations"][:8]}')
    n, s = inject_cache.inject(out, copy)
    print(f'  缓存值写回成品 {n} 格')
    keep = os.path.join(HERE, '_lo_copy.xlsx')
    shutil.copy(copy, keep)           # 留一份算好的副本给 verify 用（.gitignore 忽略）
    shutil.rmtree(tmpd, ignore_errors=True)
    return 0 if st['total_errors'] == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
