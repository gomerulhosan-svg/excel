# -*- coding: utf-8 -*-
"""补充文件（费用报销明细）：
- 【全量费用明细总表】原样搬进来（表头、列宽、字体都不动），当费用录入表：
  文字日期「2026.6.4」换成真日期、显示格式还是 2026.6.4；右边加一列「校验」；公式和样式铺到 2000 行。
  这里登记的费用按「支出类别」进【费用汇总】→ 利润表；钱付出去（报销、付款）时在资金台帐记「报销付款」冲掉，
  没付的部分在资产负债表上是「应付费用」。
- 【收款码到账对账】＝同一文件里的 Sheet1，原样搬进来，右边加说明。
"""
import re
import datetime as dt
from copy import copy
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from common import *
from s_baosun import _copy_sheet


def _parse(v):
    if isinstance(v, str):
        m = re.fullmatch(r'\s*(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})\s*', v)
        if m:
            return dt.datetime(int(m[1]), int(m[2]), int(m[3]))
    return None


def build_fee(wb, ctx, src3, changes):
    src = src3['全量费用明细总表']
    ws = wb.create_sheet(SH_FEE)
    _copy_sheet(src, ws)
    ws.sheet_properties.tabColor = 'C65911'
    last = max(r for r in range(FEE_R0, src.max_row + 1)
               if any(src.cell(r, c).value not in (None, '') for c in range(2, 9)))
    n_date = 0
    for r in range(FEE_R0, last + 1):
        d = _parse(ws[f'{F_DATE}{r}'].value)
        if d:
            ws[f'{F_DATE}{r}'].value = d
            n_date += 1
        ws[f'{F_DATE}{r}'].number_format = 'yyyy.m.d'          # 显示跟原来一样：2026.6.4
    changes.append((SH_FEE, f'B{FEE_R0}:B{last}', '日期', f'{n_date} 格文字日期', '改成真日期（显示不变）'))
    ctx['fee_last'] = last
    # 公式、样式铺到 2000 行
    for r in range(last + 1, FEE_R1 + 1):
        for c in range(1, 9):
            ws.cell(r, c)._style = copy(ws.cell(last, c)._style)
        ws[f'A{r}'] = f'=IF(AND(B{r}="",C{r}=""),"",ROW()-1)'
        ws[f'{F_DATE}{r}'].number_format = 'yyyy.m.d'
    # 校验列 I
    ws['I1']._style = copy(ws['H1']._style)
    ws['I1'] = '校验'
    ws.column_dimensions['I'].width = 24
    for r in range(FEE_R0, FEE_R1 + 1):
        blank = f'AND(B{r}="",C{r}="",D{r}="",E{r}="",F{r}="")'
        ws[f'I{r}'] = (f'=IF({blank},"",IF(NOT(ISNUMBER(B{r})),"✗ 日期不是真日期",IF(OR(B{r}<年初日,B{r}>=年末日+1),"✗ 日期不在本年度",'
                       f'IF(C{r}="","✗ 没填支出类别",IF(COUNTIF({EXP_NAMES},C{r})=0,"✗ 支出类别不在基础资料·费用项目里",'
                       f'IF(F{r}="","⚠ 没填支出金额",IF(NOT(ISNUMBER(F{r})),"✗ 金额不是数字","√")))))))')
        ws[f'I{r}'].font = F_AUTO
        ws[f'I{r}'].fill = FILL_AUTO
    rng = f'I{FEE_R0}:I{FEE_R1}'
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT($I{FEE_R0},1)="✗"'], fill=FILL_WARN,
                                                   font=Font(name=YH, sz=10, bold=True, color='FFC00000')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT($I{FEE_R0},1)="⚠"'], fill=fill('FFFFEB9C')))
    add_date_dv(ws, f'{F_DATE}{FEE_R0}:{F_DATE}{FEE_R1}')
    add_list_dv(ws, f'{F_CAT}{FEE_R0}:{F_CAT}{FEE_R1}', '=费用项目列表', '从下拉选支出类别；新类别先去【基础资料·费用项目】加')
    ws.auto_filter.ref = f'A{FEE_HDR}:I{FEE_R1}'
    ws.freeze_panes = 'A2'
    # 右边的说明
    ws.merge_cells('K1:Q6')
    put(ws, 'K1', '💡 费用照原来这样一笔一行记（支出类别从下拉选）。这里记的费用直接进【费用汇总】和【利润表】。\n'
                  '钱实际付出去的时候（报销给员工、付之前「未付」的），在【资金台帐】记一笔，收支项目选「报销付款」——'
                  '只冲掉应付，不会再算一次费用。没付的部分在【资产负债表】的「应付费用」里。\n'
                  '同一笔费用只在一个地方记：记了这张表，就别在资金台帐再选「费用支出」。',
        F_TIP, FILL_TIP, align=ALW, border=False)
    for c in 'KLMNOPQ':
        ws.column_dimensions[c].width = 12
    return ws


def build_qr(wb, ctx, src3):
    ws = wb.create_sheet(SH_QR)
    _copy_sheet(src3['Sheet1'], ws)
    ws.sheet_properties.tabColor = 'C65911'
    ws.sheet_view.topLeftCell = 'A1'
    ws.merge_cells('I1:P8')
    put(ws, 'I1', '💡 这张是你原文件里的 Sheet1，原样保留：\n'
                  '· B 列＝收款码当天收款，C 列＝农业银行实际到账（几笔相加），D 列＝差额，也就是收款码手续费（约 0.3%）；\n'
                  '· E/F/G 三列是另一路收款的同样对账（约 0.25%）；第 32 行 D32、第 34 行 G34 是 6 月合计。\n'
                  '手续费进账：每月把手续费合计在【资金台帐】记一笔——账户选「收款码」，收支项目选「费用支出」，费用项目选「手续费」。'
                  '这样资金余额才等于银行实际到账，利润表的财务费用也有了。',
        F_TIP, FILL_TIP, align=ALW, border=False)
    put(ws, 'I10', '6 月手续费合计（D32＋G34）', F_TXTB, FILL_SUBH, align=AL)
    ws.merge_cells('I10:L10')
    put(ws, 'M10', '=N(D32)+N(G34)', F_RED, fmt=MONEY2, align=AR_)
    put(ws, 'I11', '资金台帐里 6 月已记的手续费', F_TXTB, FILL_SUBH, align=AL)
    ws.merge_cells('I11:L11')
    put(ws, 'M11', (f'=-SUMIFS({cash(K_NET)},{cash(K_TO)},"{TO_EXP}",{cash(K_EXP)},"手续费",'
                    f'{cash(K_DATE)},">="&DATE(年度,6,1),{cash(K_DATE)},"<"&DATE(年度,7,1))'
                    f'+SUMIFS({fee(F_AMT)},{fee(F_CAT)},"手续费",{fee(F_DATE)},">="&DATE(年度,6,1),{fee(F_DATE)},"<"&DATE(年度,7,1))'),
        F_TXT, fmt=MONEY2, align=AR_)
    put(ws, 'I12', '=IF(ROUND(M10-M11,2)=0,"√ 6 月手续费已经记进账","⚠ 还差 "&TEXT(M10-M11,"#,##0.00")&" 元手续费没记到资金台帐")',
        F_TXTB, align=AL, border=False)
    ws.merge_cells('I12:P12')
    for c in 'IJKLMNOP':
        ws.column_dimensions[c].width = 11
    return ws
