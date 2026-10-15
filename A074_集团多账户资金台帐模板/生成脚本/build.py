# -*- coding: utf-8 -*-
"""A074 集团多账户资金台帐模板 · 生成脚本

跑法：python3 build.py [--only s_input,s_mirror,…] [--out 路径] [--keep] [--no-time] [--blank] [--time-ref 参照.xlsx] [--time-calc]
      python3 build.py --split [--out-dir 目录] [--keep]     a74split 两本联动（录入本 A074-1、查询本 A074-2，见 main_split、split.py）
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


def patch_cache(path, sheet, values):
    """成品里某张表几个公式格的缓存值改成给定的文字（zip 级；查询本首页 K3 这种一本版里没有的公式，inject 拿不到它的数）"""
    from lxml import etree
    from inject_cache import sheet_parts, NS
    tmp = path + '.tmp'
    zin = zipfile.ZipFile(path)
    part = sheet_parts(zin)[sheet]
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    n = 0
    for it in zin.infolist():
        d = zin.read(it.filename)
        if it.filename == part:
            root = etree.fromstring(d)
            for c in root.iter(f'{{{NS}}}c'):
                if c.get('r') in values and c.find(f'{{{NS}}}f') is not None:
                    v = c.find(f'{{{NS}}}v')
                    if v is None:
                        v = etree.SubElement(c, f'{{{NS}}}v')
                    c.set('t', 'str')
                    v.text = str(values[c.get('r')])
                    n += 1
            d = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(it, d)
    zout.close()
    zin.close()
    shutil.move(tmp, path)
    return n


def main_split(out_dir=ROOT, keep=False, timeout=3000):
    """a74split 两本联动：一本版（内存，make 全套检查）→ 存一本版（只给 LibreOffice 算数）→ 改成查询本、查、存、退回 → 拆成录入本、查、存
       → A074_数据整理记录.xlsx 照旧单独出 → 三本压共享公式（含「[」的不压）→ 打印 XML／WPS 存后估计
       → LibreOffice（关多线程）算一本版、算录入本，数报错 → 一本版的数写进查询本、录入本的数写进录入本 → 查询本外部链接缓存按录入本重写。
       原来那本一本版成品（A074_集团多账户资金台帐模板.xlsx）不动。"""
    import time
    import split
    from layout import SMALL, WB1_FILE, WB2_FILE, MIRROR_MODE
    from refresh_link_cache import refresh
    t0 = time.time()
    out_dir = os.path.abspath(out_dir)
    if SMALL and out_dir == os.path.abspath(ROOT):
        print('✗ 现在是小容量（A074_SMALL=1），不准写正式成品；用 --out-dir 另存')
        return 3
    out1, out2 = os.path.join(out_dir, WB1_FILE), os.path.join(out_dir, WB2_FILE)
    os.makedirs(CALC, exist_ok=True)
    one = os.path.join(CALC, 'A074_一本版_拆本用.xlsx')
    ctx = data.build_ctx()
    wb = make.build(ctx=ctx)
    print(f'① 一本版（内存）{time.time() - t0:.0f} 秒；make 检查、查环都过')
    wb.save(one)
    undo = split.make_query(wb, ctx)
    winfo_q = make.check_book(wb, 'query')
    wb.save(out2)
    undo()
    print(f'   查询本：make 检查（WPS 体检、录入模式、_流 规则）、查环都过；存 {os.path.basename(out2)}')
    split.make_entry(wb, ctx)
    winfo_e = make.check_book(wb, 'entry')
    wb.save(out1)
    print(f'   录入本：make 检查（WPS 体检）、查环都过；存 {os.path.basename(out1)}（{time.time() - t0:.0f} 秒）')
    del wb
    write_log_book(ctx, OUT if out_dir == os.path.abspath(ROOT) else os.path.join(out_dir, os.path.basename(OUT)))
    nf = {}
    for p, tag in ((one, '一本版'), (out2, '查询本'), (out1, '录入本')):
        s_ = share(p)
        nf[p] = count_formulas(p)
        print(f'② {tag}：公式 {nf[p]:,} 个（共享 {sum(s_.values()):,} 格），{os.path.getsize(p) / 1e6:.1f} MB')
        if p != one:
            xml_report(p)
    calc_one = os.path.join(CALC, 'A074_一本版_拆本用_算数.xlsx')
    calc_e = os.path.join(CALC, os.path.basename(out1).replace('.xlsx', '_算数.xlsx'))
    ok = True
    for src, calc, tag in ((one, calc_one, '一本版（查询本的数从它来）'), (out1, calc_e, '录入本')):
        shutil.copy(src, calc)
        unhide_cols(calc)
        st = lo_recalc(calc, timeout)
        if st is None or 'error' in st:
            print(f'✗ {tag} 重算没跑成：', str((st or {}).get('error'))[:500])
            return 2
        print(f"③ LibreOffice 算{tag}：公式 {st['total_formulas']:,} 个，报错 {st['total_errors']} 处")
        for k, v in (st.get('error_summary') or {}).items():
            print(f'    {k} ×{v["count"]}  {v["locations"][:15]}')
        ok = ok and st['total_errors'] == 0 and st['total_formulas'] == nf[src]
        if st['total_formulas'] != nf[src]:
            print(f"✗ 公式个数对不上：{nf[src]} vs LibreOffice {st['total_formulas']}")
    n2, _ = inject(out2, calc_one)
    n1, _ = inject(out1, calc_e)
    k3 = _k3_text(ctx)
    patch_cache(out2, '首页', {'K3': k3})
    print(f'④ 写回缓存：查询本 {n2:,} 格（一本版算的；首页 K3＝「{k3}」）、录入本 {n1:,} 格')
    refresh(out2, {WB1_FILE: out1})
    print(f'⑤ {WB1_FILE} {os.path.getsize(out1) / 1e6:.1f} MB、{WB2_FILE} {os.path.getsize(out2) / 1e6:.1f} MB（取数写法 {MIRROR_MODE}）；'
          f'共 {time.time() - t0:.0f} 秒')
    if not keep:
        for p in (one, calc_one, calc_e):
            if os.path.exists(p):
                os.remove(p)
    return 0 if ok else 1


def _k3_text(ctx):
    """查询本首页 K3 的缓存值（跟公式 split.home_k3_formula 一样算）：流水笔数＝日期或账户有字的行，最后一笔＝认得出的最晚日期"""
    import datetime as dt
    from layout import N_LEDGER
    es = [e for e in (ctx.get('entries') or [])[:N_LEDGER] if e]
    n = max(sum(1 for e in es if e.get('日期') not in (None, '')), sum(1 for e in es if e.get('账户') not in (None, '')))
    ds = [e['日期'] for e in es if isinstance(e.get('日期'), (dt.date, dt.datetime))]
    last = max(ds) if ds else None
    return (f'数据来自录入本：流水 {n:,} 笔，最后一笔 ' + (f'{last.year}/{last.month}/{last.day}' if last else '（没有）')
            + '（跟录入本首页对一下）')


if __name__ == '__main__':
    if '--split' in sys.argv:
        a = sys.argv[1:]
        od = os.path.abspath(a[a.index('--out-dir') + 1]) if '--out-dir' in a else ROOT
        sys.exit(main_split(od, keep='--keep' in a))
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
