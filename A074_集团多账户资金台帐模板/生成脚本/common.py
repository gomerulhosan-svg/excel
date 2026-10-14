# -*- coding: utf-8 -*-
"""A074 集团多账户资金台帐模板 · 通用样式和小工具（配色照参考模板：深藏青标题条、琥珀色说明行、淡黄要填、灰色公式、※/△ 两级提示；字体统一微软雅黑）。
   各表的行列地址在 layout.py；公式片段（inref/esc/norm/vdate/chain/键…）也在 layout.py。"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from layout import (FONT, C_TITLE, HDR_FILL, RGB_IN, RGB_SEL, RGB_AUTO, RGB_KPI, RGB_TIP, RGB_TIP_FONT, RGB_ERR_FILL,
                    RGB_ERR_FONT, RGB_WARN_FONT, RGB_GUESS_FONT, FMT_MONEY, FMT_ACC, FMT_DATE, FMT_MONTH, FMT_INT, FMT_PCT,
                    FMT_YEAR, FMT_MONTH_SEL, SH_HOME)

YH = FONT
F_TITLE = Font(name=YH, sz=15, bold=True, color='FFFFFFFF')
F_TIP = Font(name=YH, sz=9, color=RGB_TIP_FONT)
F_HDR = Font(name=YH, sz=10, bold=True, color='FFFFFFFF')
F_TXT = Font(name=YH, sz=10, color='FF000000')
F_TXTB = Font(name=YH, sz=10, bold=True, color='FF000000')
F_IN = Font(name=YH, sz=10, color='FF000000')              # 录入格
F_AUTO = Font(name=YH, sz=10, color='FF7F7F7F')            # 公式格（灰字）
F_AUTOB = Font(name=YH, sz=10, bold=True, color='FF404040')
F_NOTE = Font(name=YH, sz=9, color='FF808080')
F_HELP = Font(name=YH, sz=8, color='FFBFBFBF')             # 隐藏帮助列
F_RED = Font(name=YH, sz=10, bold=True, color='FFC00000')
F_KPI_L = Font(name=YH, sz=10, bold=True, color='FF1F3864')
F_KPI_V = Font(name=YH, sz=12, bold=True, color='FFC00000')
F_SEC = Font(name=YH, sz=11, bold=True, color='FFFFFFFF')
F_SEL = Font(name=YH, sz=11, bold=True, color='FFC00000')  # 选择格：红字加粗 11 号
F_LINK = Font(name=YH, sz=10, bold=True, color='FF0563C1', underline='single')


def fill(rgb):
    return PatternFill('solid', fgColor=rgb)


def cf_bg(rgb):
    """条件格式用的底色：Excel/WPS 认 bgColor，前景、背景都填上"""
    return PatternFill('solid', fgColor=rgb, bgColor=rgb)


FILL_TITLE = fill(C_TITLE)
FILL_TIP = fill(RGB_TIP)
FILL_IN = fill(RGB_IN)        # 要人填：淡黄
FILL_SEL = fill(RGB_SEL)      # 报表选择格：黄
FILL_AUTO = fill(RGB_AUTO)    # 公式：灰
FILL_KPI = fill(RGB_KPI)      # 合计、重点数
FILL_SUB = fill('FFDDEBF7')   # 小计（大类、公司）
FILL_SECT = fill('FFE2EFDA')  # 段标题（收入/支出/往来/内部转账）
FILL_OK = fill('FFE2EFDA')
FILL_WARN = fill('FFFFC7CE')
FILL_NONE = PatternFill(fill_type=None)

thin = Side(style='thin', color='FFBFBFBF')
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
NOBD = Border()
AC = Alignment(horizontal='center', vertical='center')
ACW = Alignment(horizontal='center', vertical='center', wrap_text=True)
AL = Alignment(horizontal='left', vertical='center')
ALW = Alignment(horizontal='left', vertical='center', wrap_text=True)
AR = Alignment(horizontal='right', vertical='center')

MONEY, ACC, DATE, MONTH, INT, PCT = FMT_MONEY, FMT_ACC, FMT_DATE, FMT_MONTH, FMT_INT, FMT_PCT


def put(ws, coord, value=None, font=None, fill_=None, fmt=None, align=None, border=True):
    c = ws[coord]
    if value is not None:
        c.value = value
    if font is not None:
        c.font = font
    if fill_ is not None:
        c.fill = fill_
    if fmt is not None:
        c.number_format = fmt
    if align is not None:
        c.alignment = align
    if border:
        c.border = BD
    return c


def title(ws, text, last_col, tip=None, h1=30, h2=None, tab=None):
    """第 1 行标题（深藏青底白字 15 号）＋第 2 行说明（琥珀色）；tab＝工作表标签色"""
    if h2 is None:
        width = sum((ws.column_dimensions[CL(i)].width or 9) for i in range(1, CI(last_col) + 1))
        per_line = max(20, int(width / 1.9))
        lines = -(-len(tip or '') // per_line)
        h2 = max(24, 14 * lines + 8)
    if tip:
        tip = tip.replace('💡 ', '★ ').replace('💡', '★ ')
    ws.merge_cells(f'A1:{last_col}1')
    put(ws, 'A1', text, F_TITLE, FILL_TITLE, align=AC, border=False)
    ws.row_dimensions[1].height = h1
    if tip:
        ws.merge_cells(f'A2:{last_col}2')
        put(ws, 'A2', tip, F_TIP, FILL_TIP, align=ALW, border=False)
        ws.row_dimensions[2].height = h2
    if tab:
        ws.sheet_properties.tabColor = tab[2:] if len(tab) == 8 else tab
    ws.sheet_view.showGridLines = False


def header(ws, row, cols_texts, group, font=F_HDR, height=30):
    """表头：白字加粗、居中换行；底色按组（layout.HDR_FILL：录入/报表/子表/基础/核对…）"""
    color = HDR_FILL.get(group, group)
    for col, t in cols_texts:
        put(ws, f'{col}{row}', t, font, fill(color), align=ACW)
    ws.row_dimensions[row].height = height


def section(ws, row, c1, c2, text, group='报表'):
    color = HDR_FILL.get(group, group)
    ws.merge_cells(f'{c1}{row}:{c2}{row}')
    put(ws, f'{c1}{row}', text, F_SEC, fill(color), align=AL)
    for i in range(CI(c1) + 1, CI(c2) + 1):
        ws.cell(row=row, column=i).border = BD


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def hide(ws, *cols):
    for c in cols:
        ws.column_dimensions[c].hidden = True


def link(cell, sheet, ref='A1'):
    """表内跳转：写成 location，WPS / Excel 都能点过去（表名带 - （）的要加引号，这里统一加）"""
    v = cell.value
    sh = "'" + sheet.replace("'", "''") + "'"
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"{sh}!{ref}",
                               display=None if v is None or str(v).startswith('=') else str(v))
    return cell


def kpi(ws, lbl_cell, val_cell, label, formula, fmt=ACC):
    """汇总条：标签格（浅黄底深蓝字）＋数值格（浅黄底红字加粗）"""
    put(ws, lbl_cell, label, F_KPI_L, FILL_KPI, align=ACW)
    put(ws, val_cell, formula, F_KPI_V, FILL_KPI, fmt=fmt, align=AC)


def style_rows(ws, r0, r1, cols, auto=(), fmts=None, aligns=None, fills=None, bold=()):
    """数据区：白底细灰框；公式列灰底灰字；fills 可给手工列指定底色（如 FILL_IN）"""
    fmts, aligns, fills = fmts or {}, aligns or {}, fills or {}
    for r in range(r0, r1 + 1):
        for col in cols:
            c = ws[f'{col}{r}']
            is_auto = col in auto
            c.font = F_AUTOB if col in bold else (F_AUTO if is_auto else F_IN)
            c.fill = FILL_AUTO if is_auto else fills.get(col, FILL_NONE)
            c.border = BD
            c.alignment = aligns.get(col, AC)
            if col in fmts:
                c.number_format = fmts[col]


def cf_flags(ws, rng, first_cell):
    """※ 开头：浅红底红字加粗；△ 开头：琥珀色字（first_cell＝区域左上角那格，相对引用）"""
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first_cell},1)="※"'],
                                                   font=Font(name=YH, bold=True, color=RGB_ERR_FONT), fill=cf_bg(RGB_ERR_FILL)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first_cell},1)="△"'],
                                                   font=Font(name=YH, color=RGB_WARN_FONT)))


def cf_neg_red(ws, rng, first_cell):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({first_cell}),{first_cell}<0)'],
                                                   font=Font(name=YH, bold=True, color=RGB_ERR_FONT)))


def dv_list(ws, sqref, formula, prompt=None, stop=True, blank=True, error=None):
    dv = DataValidation(type='list', formula1=str(formula).lstrip('='), allow_blank=blank, showErrorMessage=True,
                        errorStyle='stop' if stop else 'warning')
    if prompt:
        dv.promptTitle, dv.prompt = '提示', prompt
        dv.showInputMessage = True
    dv.errorTitle = '不在清单里'
    dv.error = error or ('请从下拉里选。账户、公司、收费项、收付人都在【基础资料】里加：先去那边加一行再回来选'
                         if stop else '不在【基础资料】的清单里。要用新的，先去【基础资料】加一行；点「是」照样填进去（这一行会提示 ※/△）')
    ws.add_data_validation(dv)
    dv.add(sqref)
    return dv


def dv_date(ws, sqref, stop=True):
    dv = DataValidation(type='date', operator='between', formula1='36526', formula2='73050', allow_blank=True,
                        showErrorMessage=True, errorStyle='stop' if stop else 'warning', errorTitle='日期',
                        error='要填真正的日期，像 2026/9/1 或 2026-9-1。「2026.9.1」是文字，不算日期')
    ws.add_data_validation(dv)
    dv.add(sqref)
    return dv


def cf_row_flags(ws, rng, flag_cell):
    """整行标色（资金台帐）：flag_cell（如 $L5）以 ※ 开头 → 整行浅红底；以 △ 开头 → 琥珀字"""
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({flag_cell},1)="※"'], fill=cf_bg(RGB_ERR_FILL)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({flag_cell},1)="△"'], font=Font(name=YH, color=RGB_WARN_FONT)))


def cf_guess(ws, rng, formula):
    """「按摘要认出来的」收费项/收付人：琥珀色斜体（formula 例：AND($C5="",$N5<>"")）"""
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], font=Font(name=YH, italic=True, color=RGB_GUESS_FONT)))


_SEL_KIND = {  # 选择格的有效性（stop：乱填填不进去；普通粘贴会冲掉有效性，公式那边还有 dsel/yr_eff 兜底）
    'date': ('日期', '要填真正的日期，像 2026/10/1（2000～2099 年）。「2026.10.1」是文字，填不进去', FMT_DATE),
    'year': ('年份', '填年份（如 2026，也可以填 26）或某一天的日期', FMT_YEAR),
    'month': ('月份', '填 1～12、202609 这样的年月，或那个月的某一天', FMT_MONTH_SEL),   # 1～12 是哪一年的月，各表不一样：出错提示照各表的输入提示说（selector）
}


def selector(ws, cell_lbl, lbl, cell_in, value=None, dv_formula=None, fmt=None, prompt=None, kind=None, err=None):
    """报表顶上的选择格：标签 ＋ 黄底红字加粗的输入格（保护时不锁）。
       kind：'date'（日期 stop）/'year'（年份或日期 stop，格式「2026年」）/'month'（1～12、202609 或日期 stop）/None（下拉 dv_formula 或不限）。
       生效值的写法用 layout.dsel / yr_eff / ym_eff；选择格旁边要显示「按 … 算」（看不懂 →「※ 看不懂，按 … 算」）。
       err＝出错提示（不给：date/year 用 _SEL_KIND；month 有输入提示就照输入提示说——复核 F10：「1～12」按的是旁边年份格或
       ② 那天的年份，各表不一样，不能统一写「管理年度」）"""
    put(ws, cell_lbl, lbl, F_KPI_L, fill('FFD9E1F2'), align=AC)
    c = cell_in.replace('$', '')
    if kind in _SEL_KIND and fmt is None:
        fmt = _SEL_KIND[kind][2]
    put(ws, cell_in, value, F_SEL, FILL_SEL, fmt=fmt, align=AC)
    dv = None
    if kind == 'date':
        dv = DataValidation(type='date', operator='between', formula1='36526', formula2='73050', allow_blank=True)
    elif kind == 'year':
        dv = DataValidation(type='custom', allow_blank=True,
                            formula1=f'AND(ISNUMBER({c}),OR(AND({c}>=2000,{c}<=2100),AND({c}>=36526,{c}<=73050),AND({c}>=0,{c}<=99)))')
    elif kind == 'month':
        dv = DataValidation(type='custom', allow_blank=True,
                            formula1=f'AND(ISNUMBER({c}),OR(AND({c}>=1,{c}<=12),AND({c}>=200001,{c}<=210012),AND({c}>=36526,{c}<=73050)))')
    elif dv_formula:
        dv = DataValidation(type='list', formula1=str(dv_formula).lstrip('='), allow_blank=True, showErrorMessage=True,
                            errorStyle='warning', errorTitle='不在清单里', error='不在清单里：点「是」照样填（会按认不出处理）')
    elif prompt:
        dv = DataValidation(type='custom', formula1='TRUE', allow_blank=True, showErrorMessage=False)
    if dv is not None:
        if kind in _SEL_KIND:
            dv.showErrorMessage, dv.errorStyle = True, 'stop'
            e = err or _SEL_KIND[kind][1]
            if err is None and kind == 'month' and prompt:
                e = ('看不懂这个月份。' + prompt)[:255]
            dv.errorTitle, dv.error = _SEL_KIND[kind][0], e
        if prompt:
            dv.promptTitle, dv.prompt, dv.showInputMessage = '提示', prompt, True
        ws.add_data_validation(dv)
        dv.add(c)


def home_link(ws, coord, home=SH_HOME):
    c = put(ws, coord, '← 回首页', F_LINK, fill('FFFFFFFF'), align=AC, border=False)
    link(c, home, 'A1')
    return c


def print_setup(ws, rows=None, landscape=True, fit_width=True):
    """打印：横向、一页宽、每页重复表头"""
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    if fit_width:
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    if rows:
        ws.print_title_rows = rows
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.75
    ws.page_margins.header = ws.page_margins.footer = 0.3


# ─────────── 隐藏计数列（小清单用；大清单用 layout.next_pos 链式 MATCH） ───────────
def counter(ws, col, r0, n, cond, block=50):
    """隐藏的计数列：第 r0 行起 n 行，第 i 行＝前 i 行里满足条件的个数（cond(i) 给出第 i 行的条件公式）。
       条件先放在本表右边一列（自动从 CA 往右找），计数按 50 行一段求和（链不深、改一格不用整列重算）。取第 k 个：kth()。"""
    idx = getattr(ws, '_ccn', CI('CA'))
    ws._ccn = idx + 1
    cc = CL(idx)
    for i in range(n):
        r = r0 + i
        ws[f'{cc}{r}'] = f'=IF({cond(i)},1,0)'
        bs = r0 + (i // block) * block
        base = f'{col}{bs - 1}+' if bs > r0 else ''
        ws[f'{col}{r}'] = f'={base}SUM({cc}{bs}:{cc}{r})'
        ws[f'{col}{r}'].font = ws[f'{cc}{r}'].font = F_HELP
    ws.column_dimensions[cc].hidden = True


def kth(k, col, r0, n):
    """第 k 个满足条件的是第几行（1 起；没有就 0）"""
    last = f'${col}${r0 + n - 1}'
    return f'IF({k}>{last},0,IFERROR(MATCH({k},${col}${r0}:{last},0),0))'
