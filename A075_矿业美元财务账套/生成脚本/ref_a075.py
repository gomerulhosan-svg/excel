# -*- coding: utf-8 -*-
"""从用户上传的 A075 原件（../参考/A075_用户上传原件-2026.10.14.xlsx）里读「用户的东西」。

分工（写给以后改这份脚本的人）：
- 公式、下拉、条件格式、名称、容量、表的结构 —— 全部由生成脚本（common.py + s_*.py + p1008.py）按代码写出来；
- 用户填的 / 用户自己排的 —— 从原件按位置读：参数、资金账户、项目、往来单位、物料、记账规则各张对照表、科目表、
  月度汇率、流水（现在 0 行）、明细账默认科目；【首页】【使用说明】两张是用户自己重排过版面的，整张照搬（见 copy_sheet）；
  现金流水总表 T 列的旧批注也照搬（A062 鑫盛那几笔的说明，A075 用不上，见汇报「挡住使用的问题」第 2 条）。
原件只读，不改。读之前不执行原件里的任何东西（openpyxl 只解析 XML）。
"""
import datetime as dt
import os
from copy import copy

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

from common import *

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
REF = os.path.join(ROOT, '参考', 'A075_用户上传原件-2026.10.14.xlsx')


def load(path=REF):
    return openpyxl.load_workbook(path)


def _v(ws, ref):
    v = ws[ref].value
    if isinstance(v, str) and v.startswith('='):
        return None            # 输入区里不该有公式；有的话不当数据
    return v


def _rows(ws, cols, r0, r1):
    return [tuple(_v(ws, f'{c}{r}') for c in cols) for r in range(r0, r1 + 1)]


class Fixes:
    """顶替 A062 的 fixes.py：同名属性，内容全部来自原件"""


def read(ref):
    b, ru, ra, ac = ref[SH_BASE], ref[SH_RULE], ref[SH_RATE], ref[SH_ACC]
    params = {k: _v(b, f'B{P_ROW[k]}') for k, _, _ in PARAMS}
    fx = Fixes()
    fx.COMPANY, fx.GOLIVE, fx.VAT = params['company'], params['golive'], params['vat']
    accounts = []
    for r in range(ACC_R0, ACC_R1 + 1):
        accounts.append(dict(name=_v(b, f'{A_NAME}{r}'), bank=_v(b, f'{A_BANK}{r}'), cur=_v(b, f'{A_CUR}{r}'),
                             open=_v(b, f'{A_OPEN}{r}'), openb=_v(b, f'{A_OPENB}{r}'), code=_v(b, f'{A_CODE}{r}'),
                             note=_v(b, f'{A_NOTE}{r}')))
    fx.ACCOUNT_CODE = {a['name']: a['code'] for a in accounts if a['name']}
    projects = [_v(b, f'{PJ_NAME}{r}') for r in range(PRJ_R0, PRJ_R1 + 1)]
    cps = [dict(name=_v(b, f'{CP_NAME}{r}'), type=_v(b, f'{CP_TYPE}{r}'), cap=_v(b, f'{CP_CAP}{r}'), note=_v(b, f'{CP_NOTE}{r}'))
           for r in range(CP_R0, CP_R1 + 1)]
    fx.MATERIALS = [(_v(b, f'{M_NAME}{r}'), _v(b, f'{M_SPEC}{r}'), _v(b, f'{M_UNIT}{r}'), _v(b, f'{M_CAT}{r}'),
                     _v(b, f'{M_Q0}{r}'), _v(b, f'{M_P0}{r}'), _v(b, f'{M_NOTE}{r}')) for r in range(MAT_R0, MAT_R1 + 1)]
    rules = _rows(ru, [RC_NAME, RC_ATTR, RC_IN, RC_OUT, RC_HIST, RC_CFIN, RC_CFOUT, RC_NOTE], RC_R0, RC_R1)
    fx.MAT_CATS = _rows(ru, [MC_NAME, MC_ACC, MC_STOCK, MC_PROD, MC_NOTE], MC_R0, MC_R1)
    fx.USES = _rows(ru, [US_NAME, US_ACC, US_NOTE], US_R0, US_R1)
    fx.DEPTS = _rows(ru, [DP_NAME, DP_WAGE, DP_DEP, DP_NOTE], DP_R0, DP_R1)
    fx.PERSON_TYPES = _rows(ru, [PT_NAME, PT_ACC], PT_R0, PT_R1)
    fx.FA_CATS = _rows(ru, [FC_NAME, FC_ACC, FC_YEARS, FC_SALV, FC_NOTE], FC_R0, FC_R1)
    fx.CF_ITEMS = _rows(ru, [CFI_NO, CFI_NAME, CFI_SIDE], CFI_R0, CFI_R1)
    chart = []
    for r in range(AC_R0, AC_R1 + 1):
        code = _v(ac, f'{AC_CODE}{r}')
        a = dict(code=None if code is None else str(code), name=_v(ac, f'{AC_NAME}{r}'), cls=_v(ac, f'{AC_CLASS}{r}'),
                 dir=_v(ac, f'{AC_DIR}{r}'), aux=_v(ac, f'{AC_AUX}{r}'), fx=_v(ac, f'{AC_FX}{r}'),
                 cashf=_v(ac, f'{AC_CASHF}{r}'), stat=_v(ac, f'{AC_STAT}{r}'), para=_v(ac, f'{AC_PARA}{r}'),
                 open=_v(ac, f'{AC_OPEN}{r}'),
                 added=bool(ac[f'{AC_NAME}{r}'].font and ac[f'{AC_NAME}{r}'].font.b))
        chart.append(a if code is not None else {})
    rates = []
    for r in range(RATE_R0, RATE_R1 + 1):
        rates.append({c: _v(ra, f'{CL(i + 2)}{r}') for i, c in enumerate(RATE_CURS)})
    # 流水（A075 现在 0 行；有的话照搬录入列）
    cs = ref[SH_CASH]
    ledger = []
    inp = dict(date=K_DATE, account=K_ACC, category=K_CAT, memo=K_MEMO, inc=K_IN, exp=K_OUT, xrate=K_XR,
               project=K_PRJ, cp=K_CP, note=K_NOTE, ovr=K_OVR, merge=K_MERGE, att=K_ATT)
    for r in range(CASH_R0, CASH_R1 + 1):
        d = {k: _v(cs, f'{c}{r}') for k, c in inp.items()}
        if any(v is not None for v in d.values()):
            d['row'] = r
            ledger.append(d)
    start = params['start']
    if isinstance(start, (int, float)):        # WPS 存的日期格式号（57）openpyxl 不认，读出来是序列号
        from openpyxl.utils.datetime import from_excel
        start = from_excel(start)
    months = [dt.datetime(start.year + (start.month - 1 + i) // 12, (start.month - 1 + i) % 12 + 1, 1) for i in range(N_MONTHS)]
    return dict(params=params, fx=fx, accounts=accounts, projects=projects, counterparties=cps, rules=rules, chart=chart,
                rates=rates, ledger=ledger, months=months)


# ───────────────────────── 整张照搬（首页、使用说明：用户自己重排过的版面） ─────────────────────────
def copy_sheet(ref, wb, name, index=None):
    src = ref[name]
    if name in wb.sheetnames:
        old = wb[name]
        index = wb.sheetnames.index(name) if index is None else index
        wb.remove(old)
    ws = wb.create_sheet(name, index)
    for row in src.iter_rows():
        for c in row:
            if c.value is None and not c.has_style:
                continue
            d = ws.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                d.font, d.fill, d.border = copy(c.font), copy(c.fill), copy(c.border)
                d.alignment, d.protection = copy(c.alignment), copy(c.protection)
                d.number_format = c.number_format
            if c.hyperlink is not None:
                d.hyperlink = copy(c.hyperlink)
                d.hyperlink.ref = d.coordinate
    for mr in src.merged_cells.ranges:
        ws.merge_cells(str(mr))
    for k, cd in src.column_dimensions.items():
        n = ws.column_dimensions[k]
        n.width, n.hidden = cd.width, cd.hidden
        n.min, n.max = cd.min, cd.max
        if cd.outlineLevel:
            n.outlineLevel = cd.outlineLevel
    for k, rd in src.row_dimensions.items():
        n = ws.row_dimensions[k]
        if rd.height is not None:
            n.height = rd.height
        n.hidden = rd.hidden
    for dv in src.data_validations.dataValidation:
        nd = DataValidation(type=dv.type, operator=dv.operator, formula1=dv.formula1, formula2=dv.formula2,
                            allow_blank=dv.allow_blank, showErrorMessage=dv.showErrorMessage,
                            showInputMessage=dv.showInputMessage, errorStyle=dv.errorStyle,
                            errorTitle=dv.errorTitle, error=dv.error, promptTitle=dv.promptTitle, prompt=dv.prompt)
        ws.add_data_validation(nd)
        nd.sqref = dv.sqref
    for rng_ in src.conditional_formatting:
        for rule in rng_.rules:
            ws.conditional_formatting.add(str(rng_.sqref), rule)
    ws.sheet_properties.tabColor = copy(src.sheet_properties.tabColor)
    sv, dv_ = src.sheet_view, ws.sheet_view
    dv_.showGridLines, dv_.zoomScale = sv.showGridLines, sv.zoomScale
    if src.freeze_panes:
        ws.freeze_panes = src.freeze_panes
    ws.page_setup.orientation, ws.page_setup.paperSize = src.page_setup.orientation, src.page_setup.paperSize
    return ws


def copy_comments(ref, wb, name, col):
    """把原件某一列的批注照搬过来（作者、文字、框大小）"""
    from openpyxl.comments import Comment
    src, ws = ref[name], wb[name]
    n = 0
    for row in src.iter_rows(min_col=CI(col), max_col=CI(col)):
        for c in row:
            if c.comment is not None:
                cm = Comment(c.comment.text, c.comment.author)
                cm.width, cm.height = c.comment.width, c.comment.height
                ws[c.coordinate].comment = cm
                n += 1
    return n
