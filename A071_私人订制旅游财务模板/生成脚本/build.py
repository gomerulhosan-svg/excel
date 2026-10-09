# -*- coding: utf-8 -*-
"""A071 私人订制旅游财务模板 · 生成脚本

跑法：python3 build.py [--only s_base,s_regs] [--out 路径] [--keep]
输出：../A071_私人订制旅游财务模板.xlsx（预置清单和示例行在 data.py）

做法：openpyxl 生成 → 一整列同样的公式改成共享公式（文件小、打开快）→ 复制一份（把隐藏列都显示出来，LibreOffice 才不会漏算隐藏列尾巴上的公式）
      让 LibreOffice 整本重算、数报错、核对公式个数 → 把算出来的数写回成品当缓存（手机、微信预览也能看到数；Excel/WPS 打开会再算）。
      成品不经 LibreOffice 存盘。
"""
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from share import share
from inject_cache import inject
import make

ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, 'A071_私人订制旅游财务模板.xlsx')
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


def unhide_cols(path):
    """只给重算副本用：LibreOffice 存盘时可能把「隐藏列」尾部的公式格截掉，先把列都显示出来再算"""
    tmp = path + '.tmp'
    zi = zipfile.ZipFile(path)
    zo = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for it in zi.infolist():
        d = zi.read(it.filename)
        if it.filename.startswith('xl/worksheets/sheet') and it.filename.endswith('.xml'):
            d = re.sub(rb'(<col\b[^>]*?)\s+hidden="(?:1|true)"', rb'\1', d)
        zo.writestr(it, d)
    zo.close()
    zi.close()
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


def main(out=OUT, only=None, scenario=None, keep=False, timeout=3000):
    wb = make.build(only=only, scenario=scenario)
    wb.save(out)
    s = share(out)
    nf = count_formulas(out)
    print(f'① 生成 {os.path.basename(out)}：公式 {nf} 个（共享公式 {sum(s.values())} 格）')
    os.makedirs(CALC, exist_ok=True)
    calc = os.path.join(CALC, os.path.basename(out).replace('.xlsx', '_算数.xlsx'))
    shutil.copy(out, calc)
    unhide_cols(calc)
    st = lo_recalc(calc, timeout)
    if st is None:
        return 0, None
    if 'error' in st:
        print('✗ 重算没跑成：', str(st['error'])[:800])
        return 2, st
    print(f"② LibreOffice 重算：公式 {st['total_formulas']} 个，报错 {st['total_errors']} 处")
    if st['total_formulas'] != nf:
        print(f"✗ 公式个数对不上：成品 {nf}，LibreOffice 算了 {st['total_formulas']}（有公式被漏算）")
    for k, v in (st.get('error_summary') or {}).items():
        print(f'    {k} ×{v["count"]}  {v["locations"][:15]}')
    n, _ = inject(out, calc)
    print(f'③ 写回缓存 {n} 格；{os.path.getsize(out) / 1e6:.1f} MB')
    if not keep:
        os.remove(calc)
    ok = st['total_errors'] == 0 and st['total_formulas'] == nf
    return (0 if ok else 1), st


if __name__ == '__main__':
    args = sys.argv[1:]
    only = None
    out = OUT
    if '--only' in args:
        only = args[args.index('--only') + 1].split(',')
    if '--out' in args:
        out = os.path.abspath(args[args.index('--out') + 1])
    rc, _ = main(out=out, only=only, keep='--keep' in args)
    sys.exit(rc)
