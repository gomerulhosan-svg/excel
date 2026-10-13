# -*- coding: utf-8 -*-
"""A074 集团多账户资金台帐：把各张表拼成一本：先按顺序建好所有空表（模块之间互相引用的表一定存在），再让各模块往里写；
最后定义名称、检查公式（Excel 2010/WPS 规定、防环）、保护查看表。"""
import importlib
import re
import openpyxl
from openpyxl.styles import Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.protection import SheetProtection
from layout import *
from openpyxl.utils import column_index_from_string as CI
from common import print_setup, F_HELP
import data

# 模块顺序：录入/基础资料 → 隐藏计算 → 子表 → 报表 → 首页/年末余额 → 核对/记录/说明（每个 s_*.py 提供 build(wb, ctx)）
MODULES = ['s_input', 's_mirror', 's_sub', 's_rep1', 's_rep2', 's_home', 's_check']
KEEP_OPEN = {RGB_IN, RGB_SEL}                     # 淡黄手填格、黄色选择格：保护时不锁

BANNED = re.compile(r'(?<![A-Z0-9_.])(XLOOKUP|XMATCH|FILTER|UNIQUE|SORT|SORTBY|SEQUENCE|LET|LAMBDA|MAXIFS|MINIFS|IFS|SWITCH|'
                    r'TEXTJOIN|CONCAT|INDIRECT|OFFSET|TODAY|NOW|RAND|RANDBETWEEN|AGGREGATE|TEXTSPLIT|PIVOTBY|TRIMRANGE|'
                    r'SHEETS|SHEET|ISFORMULA|FORMULATEXT|IFNA|DAYS)\(')
LOOKUP_RE = re.compile(r'(?<![A-Z0-9_.])LOOKUP\(')        # 只准 _流 的规则列用 LOOKUP(1,0/(…),…)
LOOKUP_OK = {H_FLOW}


def _balanced(f):
    depth, q_ = 0, False
    for ch in f:
        if ch == '"':
            q_ = not q_
        elif not q_:
            depth += (ch == '(') - (ch == ')')
            if depth < 0:
                return False
    return depth == 0 and not q_


def _depth(f):
    """括号最大深度（近似函数嵌套层数）"""
    d = m = 0
    q_ = False
    for ch in f:
        if ch == '"':
            q_ = not q_
        elif not q_:
            if ch == '(':
                d += 1
                m = max(m, d)
            elif ch == ')':
                d -= 1
    return m


_SUMIF_RE = re.compile(r'(?<![A-Z0-9_.])(SUMIFS?)\(')
_KEYNAME = re.compile(r'流_(\w*键|月码|有效)')
_RAWCOLS = None


def _args(f, i):
    """f[i] 是左括号：返回 (顶层参数列表, 右括号位置)"""
    depth, q_, cur, out = 0, False, '', []
    for j in range(i, len(f)):
        ch = f[j]
        if ch == '"':
            q_ = not q_
        if not q_:
            if ch == '(':
                depth += 1
                if depth == 1:
                    continue
            elif ch == ')':
                depth -= 1
                if depth == 0:
                    out.append(cur)
                    return out, j
            elif ch == ',' and depth == 1:
                out.append(cur)
                cur = ''
                continue
        cur += ch
    return out, len(f)


def _sumif_needs_key(f):
    """用 流_收/流_支 求和的 SUMIF(S) 条件里必须有键列、流_月码 或 流_有效（无效行这几列是 0）→ 返回违规的片段"""
    bad = []
    for m in _SUMIF_RE.finditer(f):
        args, _e = _args(f, m.end() - 1)
        if m.group(1) == 'SUMIF':
            if len(args) < 3:
                continue
            summ, crit_ranges = args[2], [args[0]]
        else:
            summ, crit_ranges = args[0], args[1::2]
        if re.search(r'流_(收|支)(?![\w])', summ) and not any(_KEYNAME.search(r) for r in crit_ranges):
            bad.append(f[m.start():m.start() + 80])
    return bad


def _flow_rules(ws):
    """_流 原始列（A..月码）只准引用 PAR_SAFE 里的 P_，不准引用块表；正常日 不准写 INDEX(块_…)（要直接写本块那一格）"""
    bad = []
    raw_last = CI(FLOW_RAW_LAST)
    for row in ws.iter_rows(min_row=2, max_col=CI(FLOW_LAST)):
        for c in row:
            v = c.value
            if not (isinstance(v, str) and v.startswith('=')):
                continue
            outside = ''.join(v.split('"')[0::2])
            if c.column <= raw_last:
                ps = {p for p in re.findall(r'(?<![\w.])P_(\w+)', outside)} - set(PAR_SAFE)
                if ps or '块_' in outside:
                    bad.append(f'{ws.title}!{c.coordinate}: 原始列引用了 {sorted(ps) or "块表"}')
            if c.column_letter == FLOW['正常日'] and '块_' in outside:
                bad.append(f'{ws.title}!{c.coordinate}: 正常日 要直接写本块「前最晚」那一格，不能 INDEX(块_…)（会成环）')
    return bad


def check_formulas(wb):
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith('='):
                    outside = ''.join(v.split('"')[0::2]).upper()
                    hit = BANNED.findall(outside) + (['_xlfn'] if '_XLFN.' in outside else [])
                    if ws.title not in LOOKUP_OK and LOOKUP_RE.search(outside):
                        hit.append('LOOKUP 只准在 _流 的摘要规则列用')
                    long_lit = [t for t in v.split('"')[1::2] if len(t) > 255]       # 公式里的文字常量 Excel 最多 255 字
                    dv_use = [k for k in DV_NAMES if re.search(r'(?<![\w.])' + re.escape(k) + r'(?![\w])', outside)]
                    depth = _depth(v)
                    if dv_use:
                        hit = hit + [f'引用了下拉名称 {dv_use}（OFFSET 易失）']
                    if depth > 50:
                        hit = hit + [f'嵌套 {depth} 层']
                    nokey = _sumif_needs_key(v) if ('流_收' in v or '流_支' in v) else []
                    if nokey:
                        hit = hit + [f'流_收/流_支 求和没有键列条件：{nokey[0]}']
                    if ws.title in PAUSE_SHEETS and c.row > 1 and not v.startswith(PAUSE_GUARD) \
                            and c.coordinate not in getattr(ws, '_pause_exempt', ()):
                        hit = hit + ['报表公式要套 layout.rpt()（报表暂停）']
                    if hit or long_lit or not _balanced(v) or len(v) > 8000:
                        bad.append(f'{ws.title}!{c.coordinate}: {hit} len={len(v)} 长文字={[len(t) for t in long_lit]} {v[:160]}')
    if H_FLOW in wb.sheetnames:
        bad += _flow_rules(wb[H_FLOW])
    assert not bad, '\n'.join(bad[:40])
    for ws in wb.worksheets:                         # 有效性的提示、出错文字 Excel 最多 255 字
        for dv in ws.data_validations.dataValidation:
            for t in (dv.prompt, dv.error, dv.promptTitle, dv.errorTitle):
                assert t is None or len(t) <= 255, f'{ws.title} 有效性文字超过 255 字：{t[:60]}…（{len(t)}）'
            f1 = dv.formula1 or ''
            assert not f1.startswith('='), f'{ws.title} 有效性 formula1 不要写开头的 =：{f1[:60]}'
            if dv.type == 'list' and f1.startswith('"'):
                assert len(f1) <= 257, f'{ws.title} 下拉直接写的清单超过 255 字'


def _no_emoji(wb):
    """💡 在老电脑的 Excel / WPS 上可能显示成方框：统一换成 ★"""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and '💡' in v:
                    c.value = v.replace('💡 ', '★ ').replace('💡', '★ ')
        for dv in ws.data_validations.dataValidation:
            for a in ('prompt', 'error', 'promptTitle', 'errorTitle'):
                t = getattr(dv, a)
                if t and '💡' in t:
                    setattr(dv, a, t.replace('💡 ', '★ ').replace('💡', '★ '))


def define_names(wb):
    for nm, ref in NAMES.items():
        wb.defined_names[nm] = DefinedName(nm, attr_text=ref)
    for nm, ref in DV_NAMES.items():
        wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def _fit_widths(wb):
    for ws in wb.worksheets:
        if ws.sheet_state != 'visible':
            continue
        need = {}
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400)):
            for c in row:
                if c.value is None:
                    continue
                fmt = c.number_format or ''
                sz = (c.font.sz or 10) if c.font is not None else 10
                w = 12.5 if ('yy' in fmt or 'm"' in fmt) else (13.5 if '#,##0.00' in fmt else (11 if '#,##0' in fmt else 0))
                if w:
                    w = w * max(1.0, sz / 10.0) + (1.5 if (c.font is not None and c.font.b) else 0)
                    need[c.column_letter] = max(need.get(c.column_letter, 0), w)
        for col, w in need.items():
            cd = ws.column_dimensions[col]
            if not cd.hidden and (cd.width or 8.43) < w:
                cd.width = w


def build(only=None, scenario=None, ctx=None):
    errs = self_check()
    assert not errs, 'layout 自检没过：\n' + '\n'.join(errs)
    ctx = ctx if ctx is not None else data.build_ctx()
    if scenario:
        scenario(ctx)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for _g, color, names in GROUPS:
        for n in names:
            ws = wb.create_sheet(n)
            ws.sheet_properties.tabColor = color[2:]
    for n in HIDDEN:
        wb.create_sheet(n)
    define_names(wb)
    for m in MODULES:
        if only and m not in only:
            continue
        try:
            mod = importlib.import_module(m)
        except ModuleNotFoundError as e:
            if e.name == m:
                print(f'   （模块 {m} 还没写，跳过）')
                continue
            raise
        mod.build(wb, ctx)
    for n in HIDDEN:
        wb[n].sheet_state = 'hidden'
        for row in wb[n].iter_rows(min_row=1, max_row=1):
            for c in row:
                c.font = F_HELP
    for grp, color, names in GROUPS:
        for n in names:
            ws = wb[n]
            ws.sheet_properties.tabColor = color[2:]
            if grp in VIEW_GROUPS and not getattr(ws, '_no_protect', False):
                for row in ws.iter_rows():
                    for c in row:
                        if c.fill is not None and c.fill.fgColor is not None and c.fill.fgColor.rgb in KEEP_OPEN:
                            c.protection = Protection(locked=False)
                ws.protection = SheetProtection(sheet=True, autoFilter=False, sort=False, formatColumns=False, formatRows=False,
                                                formatCells=False, selectLockedCells=False, selectUnlockedCells=False)
            if not ws.print_title_rows and grp not in ('首页',) and not getattr(ws, '_keep_print', False):
                print_setup(ws, None, landscape=True)
    _no_emoji(wb)
    bad = [nm for nm in wb.defined_names if not re.fullmatch(r'[A-Za-z_\\一-鿿][\w.一-鿿]*', nm)]
    assert not bad, f'名称里有 Excel 不认的字符：{bad[:10]}'
    check_formulas(wb)
    import cyclecheck                                  # 静态环检测（Excel 把 INDEX/SUMIF 的整片区域都当前驱）
    cyc = cyclecheck.find_cycles(wb)
    if cyc:
        print(f'✗ 公式有环 {len(cyc)} 处（Excel 打开会报循环引用），前 10 个：\n' + cyclecheck.fmt(cyc[:10]))
    assert not cyc, f'公式有环 {len(cyc)} 处（见上面打印的前 10 个）'
    _fit_widths(wb)
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = False
    wb.active = 0
    wb.worksheets[0].sheet_view.tabSelected = True
    wb.calculation.fullCalcOnLoad = False
    wb.calculation.calcId = 191029
    return wb
