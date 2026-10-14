# -*- coding: utf-8 -*-
"""A074 集团多账户资金台帐模板 · 生成脚本

跑法：python3 build.py [--only s_input,s_mirror,…] [--out 路径] [--keep] [--no-time] [--blank] [--time-ref 参照.xlsx] [--time-calc]
输出：../A074_集团多账户资金台帐模板.xlsx（基础资料和导入的流水在 data.py，读 导入/out/*.json）
      ../A074_数据整理记录.xlsx（v3.1：修改明细挪成单独文件，静态，设计 §15.5；--out 另存时放在旁边，文件名加「_数据整理记录」）
      --blank：另出一本空白模板（基础资料照预置，资金台帐、数据整理记录是空的；默认文件名加「_空白」）
测试小容量：A074_SMALL=1（只能 --out 另存，不准写正式成品）

做法：openpyxl 生成（make：公式检查、WPS 存盘体检、录入模式检查、查环）→ 一整列同样的公式改成共享公式（share）
      → 复制一份（把隐藏列都显示出来，LibreOffice 才不会漏算隐藏列尾巴上的公式）让 LibreOffice 整本重算、数报错、核对公式个数
      → 把算出来的数写回成品当缓存（手机、微信预览也能看到数；Excel/WPS 打开会再算）。成品不经 LibreOffice 存盘。
      → 打印各表 XML：「共享保住」（Excel 存、或 WPS 照规矩存）和「WPS 存一次以后估计」两个数（设计 §15.1.9）
      → 全容量时用 lo_time.py 在**成品**上量改台帐一格的 CPU 秒，报「占 v2 成品（git 4575e4b）的百分比」（复核 compat-2：
        以前量的是 LibreOffice 另存过的 _calc 副本，公式全展开了，量出来比成品快一倍）；超了报警（不拦）。
        --time-calc：顺便也量 _calc 副本（对照用，多花几分钟）。
"""
import glob
import inspect
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from share import share, wps_estimate
from inject_cache import inject
import make
import data

ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, 'A074_集团多账户资金台帐模板.xlsx')
LOG_OUT = os.path.join(ROOT, 'A074_数据整理记录.xlsx')     # v3.1：修改明细单独一本（s_check.build_log(ws, ctx, detail=True)）
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
    """每张表 XML 的大小（MB）"""
    return {k: v[0] / 1e6 for k, v in wps_estimate(path).items()}


def xml_report(path):
    """打印各表 XML：共享保住（现在这本）／WPS 存一次以后估计（设计 §15.1.9、§15.6）→ (合计, WPS 后合计) 字节"""
    import layout
    XML_WARN_MB = getattr(layout, 'XML_WARN_MB', 25)
    est = wps_estimate(path)
    tot = sum(v[0] for v in est.values())
    totw = sum(v[1] for v in est.values())
    print(f'   XML：共享保住 {tot / 1e6:.1f} MB；WPS 存一次以后估计 {totw / 1e6:.1f} MB（会被拆的子格 {sum(v[2] for v in est.values()):,} 个）'
          f'；文件 {os.path.getsize(path) / 1e6:.1f} MB')
    top = sorted(est.items(), key=lambda kv: -kv[1][0])[:8]
    print('     各表（共享保住／WPS 存后，MB）：' + '、'.join(f'{k} {a / 1e6:.1f}/{b / 1e6:.1f}' for k, (a, b, _n) in top))
    grow = sorted(((k, b - a) for k, (a, b, _n) in est.items() if b - a > 1e5), key=lambda kv: -kv[1])[:8]
    if grow:
        print('     WPS 会拆胖的表：' + '、'.join(f'{k} +{d / 1e6:.1f}' for k, d in grow))
    if tot > XML_WARN_MB * 1e6:
        print(f'   △ XML 合计超过 {XML_WARN_MB} MB（设计 §15.6 估约 46 MB，不拦）')
    if totw > make.WPS_WARN_MB * 1e6:
        print(f'   ✗ WPS 存后估计超过 {make.WPS_WARN_MB} MB：看上面哪张表会被拆胖（直接跨表、往下撑的区域）')
    return tot, totw


def entry_report(wb):
    """录入模式下还会算的公式（make.entry_stats；设计 §15.2.3）"""
    es = getattr(wb, '_a074_entry', None) or {}
    if not es:
        return
    tot = sum(v[0] for v in es.values())
    live = sum(v[1] for v in es.values())
    shell = sum(v[2] for v in es.values())
    top = sorted(((k, v[1]) for k, v in es.items() if v[1]), key=lambda kv: -kv[1])[:10]
    print(f'   录入模式下照算的公式 {live:,} 个（共 {tot:,}；子表逐笔空壳 {shell:,}）：' + '、'.join(f'{k} {n:,}' for k, n in top))


def log_book_path(out):
    return LOG_OUT if os.path.abspath(out) == os.path.abspath(OUT) else out[:-5] + '_数据整理记录.xlsx'


def write_log_book(ctx, out):
    """修改明细单独一本（静态，全是值；设计 §15.5）：s_check.build_log(ws, ctx, detail=True)。没有导入记录（空白模板）不出"""
    import openpyxl
    from layout import SH_LOG
    if not (ctx.get('cleanlog') or ctx.get('pending')):
        return None
    try:
        import s_check
    except ModuleNotFoundError:
        return None
    if 'detail' not in inspect.signature(s_check.build_log).parameters:
        print('   （s_check.build_log 还没有 detail 参数：修改明细还在主本里，没单独出文件）')
        return None
    path = log_book_path(out)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SH_LOG
    n = s_check.build_log(ws, ctx, detail=True)
    make._no_emoji(wb)
    nf = sum(1 for w in wb.worksheets for c in w._cells.values() if c.data_type == 'f')
    wb.save(path)
    print(f'   数据整理记录（单独一本）：{os.path.basename(path)}，{n if isinstance(n, int) else "?"} 行，'
          f'{os.path.getsize(path) / 1e6:.1f} MB' + (f'  ✗ 里面有 {nf} 个公式（应该全是值）' if nf else ''))
    return path


def lo_time(path, ref=None, timeout=7200, extra=()):
    """在成品上量（lo_time.py：关多线程、CPU 秒、跟参照本同一个 LO 进程）→ 结果 dict；要 python3-uno。
       path 可以是一本或几本（list）；extra：传给 lo_time.py 的别的参数（如 ['--no-ref']）"""
    paths = [path] if isinstance(path, str) else list(path)
    cmd = [sys.executable, '-I', os.path.join(HERE, 'lo_time.py')] + paths + (['--ref', ref] if ref else []) + list(extra)
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        return {'error': f'超过 {timeout} 秒没量完'}
    lines = r.stdout.strip().splitlines()
    try:
        res = json.loads(lines[-1])
    except (ValueError, IndexError):
        return {'error': (r.stdout + r.stderr)[-600:]}
    res['_text'] = '\n'.join(lines[:-1])
    return res


def blank_scenario(ctx):
    """空白模板：基础资料照预置，流水、整理记录清空"""
    ctx['entries'] = []
    ctx['cleanlog'] = []
    ctx['pending'] = []
    ctx['acct_check'] = []
    ctx['source'] = '空白模板'


def main(out=OUT, only=None, scenario=None, keep=False, timeout=3000, timing=True, time_ref=None, time_calc=False):
    from layout import SMALL, FORMULA_BUDGET
    if SMALL and os.path.abspath(out) == os.path.abspath(OUT):
        print('✗ 现在是小容量（A074_SMALL=1），不准写正式成品；用 --out 另存')
        return 3, None
    ctx = data.build_ctx()
    if scenario:
        scenario(ctx)
    wb = make.build(only=only, ctx=ctx)
    wb.save(out)
    s = share(out)
    nf, per = count_formulas(out, per_sheet=True)
    print(f'① 生成 {os.path.basename(out)}：公式 {nf:,} 个（共享公式 {sum(s.values()):,} 格）；名称 {len(wb.defined_names):,} 个'
          f'（隐藏 {sum(1 for d in wb.defined_names.values() if d.hidden):,}）')
    top = sorted(per.items(), key=lambda kv: -kv[1])[:12]
    print('   公式最多的表：' + '、'.join(f'{k} {v:,}' for k, v in top))
    entry_report(wb)
    over = (not SMALL) and nf > FORMULA_BUDGET
    if over:
        print(f'✗ 超过预算 {FORMULA_BUDGET:,}')
    write_log_book(ctx, out)
    del wb
    os.makedirs(CALC, exist_ok=True)
    calc = os.path.join(CALC, os.path.basename(out).replace('.xlsx', '_算数.xlsx'))
    shutil.copy(out, calc)
    unhide_cols(calc)
    st = lo_recalc(calc, timeout)
    if st is None:
        xml_report(out)
        return 0, None
    if 'error' in st:
        print('✗ 重算没跑成：', str(st['error'])[:800])
        return 2, st
    print(f"② LibreOffice 重算：公式 {st['total_formulas']:,} 个，报错 {st['total_errors']} 处")
    if st['total_formulas'] != nf:
        print(f"✗ 公式个数对不上：成品 {nf}，LibreOffice 算了 {st['total_formulas']}（有公式被漏算）")
    for k, v in (st.get('error_summary') or {}).items():
        print(f'    {k} ×{v["count"]}  {v["locations"][:15]}')
    n, _ = inject(out, calc)
    print(f'③ 写回缓存 {n:,} 格；{os.path.getsize(out) / 1e6:.1f} MB')
    xml_report(out)
    if timing and not SMALL:
        tm = lo_time([out, calc] if time_calc else out, time_ref)
        if 'error' in tm:
            print('④ 计时没跑成（要 LibreOffice 的 python3-uno）：', tm['error'][-300:])
        else:
            print('④ ' + tm['_text'].replace('\n', '\n   '))
            print('   ' + ('✓ 没超报警线' if not tm.get('报警') else f"△ {len(tm['报警'])} 处超报警线（只报警；先看设计 §15.6、§15.9-5）"))
    if not keep:
        os.remove(calc)
    ok = st['total_errors'] == 0 and st['total_formulas'] == nf and not over      # 慢只报警，不算失败
    return (0 if ok else 1), st


if __name__ == '__main__':
    args = sys.argv[1:]
    only = None
    out = OUT
    if '--only' in args:
        only = args[args.index('--only') + 1].split(',')
    if '--out' in args:
        out = os.path.abspath(args[args.index('--out') + 1])
    tref = os.path.abspath(args[args.index('--time-ref') + 1]) if '--time-ref' in args else None
    scen = None
    if '--blank' in args:
        scen = blank_scenario
        if '--out' not in args:
            out = OUT.replace('.xlsx', '_空白.xlsx')
    rc, _ = main(out=out, only=only, scenario=scen, keep='--keep' in args, timing='--no-time' not in args, time_ref=tref,
                 time_calc='--time-calc' in args)
    sys.exit(rc)
