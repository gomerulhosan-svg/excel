# -*- coding: utf-8 -*-
"""A070：把各张表拼成一本：先按顺序建好所有空表（模块之间互相引用的表一定存在），再让各模块往里写；最后定义名称、检查公式、保护报表。"""
import importlib
import re
import openpyxl
from openpyxl.styles import Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.protection import SheetProtection
from layout import *
from common import print_setup, link, F_HELP
import data

# 模块顺序：录入/基本信息 → 隐藏计算 → 各查看/报表（后面几个由各自的 s_*.py 提供 build(wb, ctx)）
MODULES = ['s_input', 's_mirror', 's_query', 's_seg', 's_fund', 's_follow', 's_stmt', 's_check', 's_home']
KEEP_OPEN = {'FFFFFF00', 'FFFFF7E0', 'FFEAF1FB'}        # 亮黄选择格、淡黄手填格、淡蓝粘贴格：保护时不锁

BANNED = re.compile(r'(?<![A-Z0-9_.])(XLOOKUP|XMATCH|FILTER|UNIQUE|SORT|SORTBY|SEQUENCE|LET|LAMBDA|MAXIFS|MINIFS|IFS|SWITCH|'
                    r'TEXTJOIN|CONCAT|INDIRECT|OFFSET|TODAY|NOW|RAND|RANDBETWEEN|AGGREGATE|TEXTSPLIT|PIVOTBY|TRIMRANGE)\(')


def _balanced(f):
    depth, q = 0, False
    for ch in f:
        if ch == '"':
            q = not q
        elif not q:
            depth += (ch == '(') - (ch == ')')
            if depth < 0:
                return False
    return depth == 0 and not q


def check_formulas(wb):
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith('='):
                    outside = ''.join(v.split('"')[0::2]).upper()
                    hit = BANNED.findall(outside) + (['_xlfn'] if '_XLFN.' in outside else [])
                    if hit or not _balanced(v) or len(v) > 8000:
                        bad.append(f'{ws.title}!{c.coordinate}: {hit} len={len(v)} {v[:160]}')
    assert not bad, '\n'.join(bad[:40])


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
            if grp in VIEW_GROUPS:
                for row in ws.iter_rows():
                    for c in row:
                        if c.fill is not None and c.fill.fgColor is not None and c.fill.fgColor.rgb in KEEP_OPEN:
                            c.protection = Protection(locked=False)
                ws.protection = SheetProtection(sheet=True, autoFilter=False, sort=False, formatColumns=False, formatRows=False,
                                                formatCells=False, selectLockedCells=False, selectUnlockedCells=False)
            if not ws.print_title_rows and grp not in ('首页',):
                print_setup(ws, None, landscape=True)
    check_formulas(wb)
    _fit_widths(wb)
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = False
    wb.active = 0
    wb.worksheets[0].sheet_view.tabSelected = True
    wb.calculation.fullCalcOnLoad = False
    wb.calculation.calcId = 191029
    return wb
