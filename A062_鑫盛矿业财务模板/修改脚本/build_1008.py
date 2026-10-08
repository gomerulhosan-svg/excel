# -*- coding: utf-8 -*-
"""10/08 这一轮：以你 10 月 5 日的版本（../参考/1008上传/A062_用户10.05版.xlsx）为底，改出成品
   ../A062_鑫盛矿业财务模板.xlsx。改了什么见 m1008.py 开头。

   做法：openpyxl 打开你那本 → m1008.apply 改 → 存盘 → 同一列往下拖的公式改回「共享公式」（文件小、打开快）
        → 复制一份、把隐藏列都显示出来（LibreOffice 才不会漏算隐藏列尾巴上的公式）→ LibreOffice 整本重算、
        数报错、核对公式个数 → 把算好的数写回成品当缓存（Excel / WPS 打开照样全部重算）。成品本身不经 LibreOffice 存盘。

跑法：python3 build_1008.py [--keep]     （--keep 保留 ../_calc 里的重算副本，verify_1008.py 要用）
"""
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import openpyxl
import m1008
from share import share
from inject_cache import inject

ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, '参考', '1008上传', 'A062_用户10.05版.xlsx')
OUT = os.path.join(ROOT, 'A062_鑫盛矿业财务模板.xlsx')
CALC = os.path.join(ROOT, '_calc')


def find_recalc():
    cands = [os.environ.get('RECALC_PY', ''), '/mnt/skills/public/xlsx/scripts/recalc.py']
    cands += glob.glob(os.path.expanduser('~/.claude/skills/**/xlsx/scripts/recalc.py'), recursive=True)
    for c in cands:
        if c and os.path.exists(c):
            return c


def lo_recalc(path, timeout=3600):
    rc = find_recalc()
    r = subprocess.run([sys.executable, rc, path, str(timeout), '--force'], capture_output=True, text=True,
                       timeout=timeout + 300)
    out = r.stdout
    i, j = out.find('{'), out.rfind('}')
    return json.loads(out[i:j + 1]) if i >= 0 else {'error': out or r.stderr}


def unhide_cols(path):
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


def main(keep=False):
    t0 = time.time()
    wb = openpyxl.load_workbook(SRC)
    for line in m1008.apply(wb):
        print('   ·', line)
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    s = share(OUT)
    nf = count_formulas(OUT)
    print(f'① 改好存盘：公式 {nf} 个（共享公式 {sum(s.values())} 格），{os.path.getsize(OUT) / 1e6:.1f} MB，{time.time() - t0:.0f} 秒')
    os.makedirs(CALC, exist_ok=True)
    calc = os.path.join(CALC, 'A062_算数.xlsx')
    shutil.copy(OUT, calc)
    unhide_cols(calc)
    t1 = time.time()
    st = lo_recalc(calc)
    if 'error' in st:
        print('✗ 重算没跑成：', str(st['error'])[:800])
        return 2
    print(f"② LibreOffice 重算：公式 {st['total_formulas']} 个，报错 {st['total_errors']} 处，{time.time() - t1:.0f} 秒")
    if st['total_formulas'] != nf:
        print(f"✗ 公式个数对不上：成品 {nf}，LibreOffice 算了 {st['total_formulas']}")
    for k, v in (st.get('error_summary') or {}).items():
        print(f'    {k} ×{v["count"]}  {v["locations"][:15]}')
    n, _ = inject(OUT, calc)
    print(f'③ 写回缓存 {n} 格；{os.path.getsize(OUT) / 1e6:.1f} MB')
    if not keep:
        os.remove(calc)
    return 0 if (st['total_errors'] == 0 and st['total_formulas'] == nf) else 1


if __name__ == '__main__':
    sys.exit(main(keep='--keep' in sys.argv))
