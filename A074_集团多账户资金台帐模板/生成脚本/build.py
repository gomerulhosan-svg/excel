# -*- coding: utf-8 -*-
"""A074 集团多账户资金台帐模板 · 生成脚本

跑法：python3 build.py [--only s_input,s_mirror,…] [--out 路径] [--keep] [--no-time] [--blank]
输出：../A074_集团多账户资金台帐模板.xlsx（基础资料和导入的流水在 data.py，读 导入/out/*.json）
      --blank：另出一本空白模板（基础资料照预置，资金台帐、数据整理记录是空的；默认文件名加「_空白」）
测试小容量：A074_SMALL=1（只能 --out 另存，不准写正式成品）
全容量时最后一步用 lo_time.py 量「改台帐一格要几秒」（LibreOffice），超过 TIME_WARN 秒报警（--no-time 跳过）

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
OUT = os.path.join(ROOT, 'A074_集团多账户资金台帐模板.xlsx')
CALC = os.path.join(HERE, '_calc')
TIME_WARN = 8                      # LibreOffice 改一格超过 8 秒报警（Excel 约快 3～5 倍；WPS 要真机量）


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


def count_formulas(path, per_sheet=False):
    """成品里每个公式格（含共享公式的子格）都数一遍；per_sheet=True 时另给 {表名: 个数}"""
    from lxml import etree
    NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    z = zipfile.ZipFile(path)
    names = {}
    if per_sheet:
        rel = {}
        for _, el in etree.iterparse(z.open('xl/_rels/workbook.xml.rels')):
            if el.tag.endswith('Relationship'):
                rel[el.get('Id')] = el.get('Target').split('/')[-1]
        for _, el in etree.iterparse(z.open('xl/workbook.xml'), tag=NS + 'sheet'):
            names['xl/worksheets/' + rel[el.get(RNS + 'id')]] = el.get('name')
    n, per = 0, {}
    for nm in z.namelist():
        if nm.startswith('xl/worksheets/sheet') and nm.endswith('.xml'):
            k = 0
            for _, el in etree.iterparse(z.open(nm), tag=NS + 'f'):
                k += 1
                el.clear()
            n += k
            per[names.get(nm, nm)] = k
    z.close()
    return (n, per) if per_sheet else n


def sheet_xml_sizes(path):
    """每张表 XML 的大小（MB）：超过 3 MB 的报出来（一列公式没压成共享公式、逐格设样式都会让它变大）"""
    from lxml import etree
    NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    z = zipfile.ZipFile(path)
    rel = {}
    for _, el in etree.iterparse(z.open('xl/_rels/workbook.xml.rels')):
        if el.tag.endswith('Relationship'):
            rel[el.get('Id')] = el.get('Target').split('/')[-1]
    names = {}
    for _, el in etree.iterparse(z.open('xl/workbook.xml'), tag=NS + 'sheet'):
        names['xl/worksheets/' + rel[el.get(RNS + 'id')]] = el.get('name')
    out = {names.get(i.filename, i.filename): i.file_size / 1e6 for i in z.infolist() if i.filename in names}
    z.close()
    return out


def lo_time(calc):
    """LibreOffice 里改台帐一格的秒数（lo_time.py；要 python3-uno）"""
    r = subprocess.run([sys.executable, '-I', os.path.join(HERE, 'lo_time.py'), calc], capture_output=True, text=True,
                       timeout=3600)
    line = (r.stdout.strip().splitlines() or [''])[-1]
    try:
        return json.loads(line)
    except ValueError:
        return {'error': (r.stdout + r.stderr)[-600:]}


def blank_scenario(ctx):
    """空白模板：基础资料照预置，流水、整理记录清空"""
    ctx['entries'] = []
    ctx['cleanlog'] = []
    ctx['pending'] = []
    ctx['acct_check'] = []
    ctx['source'] = '空白模板'


def main(out=OUT, only=None, scenario=None, keep=False, timeout=3000, timing=True):
    from layout import SMALL, FORMULA_BUDGET
    if SMALL and os.path.abspath(out) == os.path.abspath(OUT):
        print('✗ 现在是小容量（A074_SMALL=1），不准写正式成品；用 --out 另存')
        return 3, None
    wb = make.build(only=only, scenario=scenario)
    wb.save(out)
    s = share(out)
    nf, per = count_formulas(out, per_sheet=True)
    print(f'① 生成 {os.path.basename(out)}：公式 {nf} 个（共享公式 {sum(s.values())} 格）')
    top = sorted(per.items(), key=lambda kv: -kv[1])[:12]
    print('   公式最多的表：' + '、'.join(f'{k} {v}' for k, v in top))
    big = {k: v for k, v in sheet_xml_sizes(out).items() if v > 3}
    if big:
        print('   △ 表的 XML 超过 3 MB：' + '、'.join(f'{k} {v:.1f} MB' for k, v in sorted(big.items(), key=lambda kv: -kv[1])))
    over = (not SMALL) and nf > FORMULA_BUDGET
    if over:
        print(f'✗ 超过预算 {FORMULA_BUDGET}')
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
    slow = False
    if timing and not SMALL:
        tm = lo_time(calc)
        if 'error' in tm:
            print('④ 计时没跑成（要 LibreOffice 的 python3-uno）：', tm['error'][-300:])
        else:
            slow = tm['max'] > TIME_WARN
            print(f"④ LibreOffice 改台帐一格：{tm['edits']} 秒（整本 {tm['full']} 秒）" +
                  (f'  ✗ 超过 {TIME_WARN} 秒：先看设计 §8.4 的退路' if slow else '  ✓'))
    if not keep:
        os.remove(calc)
    ok = st['total_errors'] == 0 and st['total_formulas'] == nf and not over      # 慢只报警，不算失败（总负责人看 §8.4 定）
    return (0 if ok else 1), st


if __name__ == '__main__':
    args = sys.argv[1:]
    only = None
    out = OUT
    if '--only' in args:
        only = args[args.index('--only') + 1].split(',')
    if '--out' in args:
        out = os.path.abspath(args[args.index('--out') + 1])
    scen = None
    if '--blank' in args:
        scen = blank_scenario
        if '--out' not in args:
            out = OUT.replace('.xlsx', '_空白.xlsx')
    rc, _ = main(out=out, only=only, scenario=scen, keep='--keep' in args, timing='--no-time' not in args)
    sys.exit(rc)
