# -*- coding: utf-8 -*-
"""把各张表拼成一本：首页 → 录入（蓝）→ 成本查看（绿）→ 往来资金查看（绿）→ 报表（深红）→ 数据校验 → 档案（灰）；隐藏的取数表放最后。
   先按顺序把所有表建好（空表），再让各模块往里写——模块之间互相引用的表一定存在。"""
import importlib
import re
from copy import copy
import openpyxl
from openpyxl.styles import Protection, Font, PatternFill, Border, Alignment
from openpyxl.worksheet.protection import SheetProtection
from common import print_setup, link, CL
from layout import *
import demo

GROUPS = [
    ('首页', C_HOME, [SH_HOME]),
    ('录入', C_IN, [SH_CASH, SH_DN, SH_OUT, SH_ORD, SH_WAGE, SH_MJ]),
    ('查看', C_VIEW, [SH_SPL, SH_OSUM, SH_OQ, SH_ALLOC, SH_MSUM, SH_LATE,
                     SH_APS, SH_SST, SH_CST, SH_ACCB, SH_FEE, SH_WSUM]),
    ('报表', C_RPT, [SH_IS, SH_BS, SH_CF, SH_TB, SH_GL, SH_JE]),
    ('校验', C_CHK, [SH_CHK]),
    ('档案', C_ARC, [SH_BASE, SH_UNIT, SH_STY, SH_MAT, SH_COA]),
]
HIDDEN = [SH_SM, SH_AUX]
MODULES = ['s_base', 's_cash', 's_regs', 's_cost', 's_ledger', 's_views', 's_home']
KEEP_OPEN = {'FFFFFF00', 'FFFFF7E0', 'FFEAF1FB'}      # 黄色选择格、淡黄手填格、淡蓝粘贴格：保护时不锁


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


def _fit_widths(wb):
    """日期、金额列放宽，免得 Excel/WPS 里显示成 ####（微软雅黑的数字比默认字体宽）"""
    for ws in wb.worksheets:
        need = {}
        anchors = {(m.min_row, m.min_col) for m in ws.merged_cells.ranges if m.max_col > m.min_col}
        for row in ws.iter_rows():
            for c in row:
                if (c.row, c.column) in anchors:
                    continue
                fmt = c.number_format or ''
                sz = (c.font.sz or 10) if c.font is not None else 10
                if c.value is None or ws.column_dimensions[c.column_letter].hidden:
                    continue
                w = 0
                if 'yy' in fmt:
                    w = 12.5 if 'm/d' not in fmt or 'yyyy' in fmt else 8
                elif '#,##0.00' in fmt:
                    w = 13.5
                elif '#,##0' in fmt:
                    w = 11
                if w:
                    w = w * max(1.0, sz / 10.0) + (1.5 if (c.font is not None and c.font.b) else 0)
                    need[c.column_letter] = max(need.get(c.column_letter, 0), w)
        for col, w in need.items():
            cd = ws.column_dimensions[col]
            if (cd.width or 8.43) < w:
                cd.width = w


def _home_corner(ws):
    """每张表（首页除外）左上角 A1 放「←首页」（#23）：第 1 行标题合并区改成从 B1 开始（内容、样式挪到 B1），
       A1 白字、跟标题同底色、表内链接到首页 A1；各模块原来放在标题右边的「← 回首页」去掉；冻结窗格保证 A 列一直看得见。"""
    for c in ws[1]:                                   # 原来标题右边的回首页链接
        if c.column > 1 and c.hyperlink is not None and SH_HOME in (c.hyperlink.location or ''):
            c.hyperlink = None
            c.value = None
            c.font, c.fill, c.border = Font(), PatternFill(fill_type=None), Border()
    a1 = ws['A1']
    mg = next((m for m in ws.merged_cells.ranges if m.min_row == 1 and m.max_row == 1 and m.min_col == 1), None)
    val, fnt, fil, aln = a1.value, copy(a1.font), copy(a1.fill), copy(a1.alignment)
    last = mg.max_col if mg is not None else 2
    if mg is not None:
        ws.unmerge_cells(str(mg))
    if val is not None or mg is not None:
        b1 = ws['B1']
        b1.value, b1.font, b1.fill, b1.alignment = val, fnt, fil, aln
        if last > 2:
            ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=last)
    cd = ws.column_dimensions['A']
    cd.width = max(cd.width or 8.43, 6)
    a1.value = '←首页'
    a1.font = Font(name='微软雅黑', sz=9 if cd.width < 8 else 10, bold=True, color='FFFFFFFF', underline='single')
    a1.fill = fil if (fil is not None and fil.fill_type) else PatternFill('solid', fgColor=C_HOME)
    a1.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    a1.border = Border()
    link(a1, SH_HOME, 'A1')
    fp = ws.freeze_panes                              # 冻结了行、没冻结列（如 A5）：改成连 A 列一起冻住，往右拉也看得见「←首页」
    if fp:
        col, row = re.match(r'([A-Z]+)(\d+)', fp).groups()
        if col == 'A':
            ws.freeze_panes = f'B{row}'


BANNED = re.compile(r'(?<![A-Z0-9_.])(XLOOKUP|XMATCH|FILTER|UNIQUE|SORT|SORTBY|SEQUENCE|LET|LAMBDA|MAXIFS|MINIFS|IFS|SWITCH|'
                    r'TEXTJOIN|CONCAT|INDIRECT|OFFSET|TODAY|NOW|RAND|RANDBETWEEN)\(')


def _check_formulas(wb):
    """只用 Excel 2010 / 老版 WPS 都有的函数；不用易变函数；括号、引号要配对"""
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith('='):
                    outside = ''.join(v.split('"')[0::2]).upper()
                    hit = BANNED.findall(outside) + (['_xlfn'] if '_XLFN.' in outside else [])
                    if hit or not _balanced(v):
                        bad.append(f'{ws.title}!{c.coordinate}: {hit} {v[:120]}')
    assert not bad, '\n'.join(bad[:30])


def build(only=None, scenario=None):
    """only：只跑这几个模块（调试用）；scenario(ctx)：改演示数据（场景测试用）"""
    ctx = demo.build_ctx()
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
    for grp, color, names in GROUPS:
        for n in names:
            ws = wb[n]
            ws.sheet_properties.tabColor = color[2:]
            if grp != '首页':
                _home_corner(ws)
            if grp in ('查看', '报表', '校验', '首页'):
                for row in ws.iter_rows():
                    for c in row:
                        if c.fill is not None and c.fill.fgColor is not None and c.fill.fgColor.rgb in KEEP_OPEN:
                            c.protection = Protection(locked=False)
                ws.protection = SheetProtection(sheet=True, autoFilter=False, sort=False, formatColumns=False, formatRows=False,
                                                formatCells=False, selectLockedCells=False, selectUnlockedCells=False)
            if not ws.print_title_rows and grp != '首页':
                hdr = {'录入': '4:4' if n != SH_CASH else f'{J_HDR}:{J_HDR}', '档案': '4:4'}.get(grp)
                print_setup(ws, hdr, landscape=True)
    _check_formulas(wb)
    _fit_widths(wb)
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = False
    wb.active = 0
    wb.worksheets[0].sheet_view.tabSelected = True
    wb.calculation.fullCalcOnLoad = False
    wb.calculation.calcId = 191029
    return wb
