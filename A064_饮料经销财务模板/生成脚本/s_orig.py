# -*- coding: utf-8 -*-
"""原表 1 的几张：出库明细 / 存条明细 / 总览汇总 / 商品库存 / 客户查询 —— 样子不动，只补公式、修错；
外加新表【采购进货】和隐藏的【_辅助】（商品库存自动列组合、客户查询取数用）。"""
from copy import copy
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill

from common import *


def _copy_style(src, dst):
    dst._style = copy(src._style)


def _copy_row_style(ws, src_row, dst_row, c0=1, c1=8):
    for c in range(c0, c1 + 1):
        _copy_style(ws.cell(src_row, c), ws.cell(dst_row, c))
    if ws.row_dimensions[src_row].height:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


# ───────────────────────── 出库明细 / 存条明细 ─────────────────────────
def _fix_ledger(ws, r0, r1, last_data, ctx, changes, sheet):
    """补客户/商品名规范化、金额公式铺满、下拉、筛选"""
    tmpl = last_data
    for r in range(r0, r1 + 1):
        if r > last_data:
            _copy_row_style(ws, tmpl, r, 1, 8)
            for c in range(1, 9):
                ws.cell(r, c).value = None
                ws.cell(r, c).font = Font(name=YH, sz=10, color='FF000000')
        # 名称规范化（fixes.py 里写明的）
        for col, mp, kind in [('C', ctx['cus_merge'], '客户'), ('D', ctx['goods_merge'], '商品')]:
            v = ws[f'{col}{r}'].value
            if isinstance(v, str):
                s = v.strip()
                new = mp.get(s, s)
                if new != v:
                    changes.append((sheet, f'{col}{r}', kind, v, new))
                    ws[f'{col}{r}'].value = new
        g = ws[f'G{r}'].value
        if g is None or (isinstance(g, str) and g.startswith('=')):
            ws[f'G{r}'].value = f'=IF(AND(E{r}<>"",F{r}<>""),E{r}*F{r},"")'
        ws[f'B{r}'].number_format = DATE
        ws[f'G{r}'].number_format = MONEY
    add_date_dv(ws, f'B{r0}:B{r1}')
    add_list_dv(ws, f'C{r0}:C{r1}', '=客户列表', '从下拉选客户；新客户先去【总览汇总】B 列加')
    add_list_dv(ws, f'D{r0}:D{r1}', '=商品列表', '从下拉选商品；新商品先去【基础资料·商品档案】加')
    ws.auto_filter.ref = f'A3:H{r1}'
    ws.freeze_panes = 'A4'


def fix_out(wb, ctx, changes):
    ws = wb[SH_OUT]
    _copy_style(ws['B3'], ws['C3'])
    ws['C3'].value = '客户名称'          # 原表这一格是空的
    ws['A2'].value = '💡 每次送货/拣货在此登记（客户、商品从下拉选），总览表、库存表、利润表、对账单自动更新'
    _fix_ledger(ws, OUT_R0, OUT_R1, ctx['out_last'], ctx, changes, SH_OUT)
    import fixes
    for r, d in fixes.OUT_DATE_FIX.items():
        old = ws[f'B{r}'].value
        if isinstance(old, str):
            ws[f'B{r}'].value = d
            changes.append((SH_OUT, f'B{r}', '日期', old, d.strftime('%Y-%m-%d')))
    widths(ws, {'D': 13})


def fix_cun(wb, ctx, changes):
    ws = wb[SH_CUN]
    ws['A2'].value = '💡 每次客户存条在此登记，一个商品一行，库存表按「客户+商品」自动汇总累加'
    _fix_ledger(ws, CUN_R0, CUN_R1, ctx['cun_last'], ctx, changes, SH_CUN)


# ───────────────────────── 采购进货（新） ─────────────────────────
def build_buy(wb, ctx):
    ws = wb.create_sheet(SH_BUY)
    ws.sheet_properties.tabColor = C_BUY[2:]
    title(ws, '采购进货明细台账', 'J', C_BUY,
          '💡 从厂家/供应商进货在这里登记，一个商品一行：数量×进价自动出金额；厂家搭赠的货填「赠品数量」，只加库存不加货款；'
          '退货数量填负数。付货款去【资金台帐】记一笔并选供应商。公司库存、成本、利润表、供应商对账单都从这里取数。')
    header(ws, BUY_HDR, [(B_SEQ, '序号'), (B_DATE, '日期'), (B_SUP, '供应商'), (B_GOODS, '商品品类'),
                         (B_QTY, '数量(件)'), (B_PRICE, '进价(元/件)'), (B_AMT, '金额(元)'),
                         (B_GIFT, '赠品数量(件)'), (B_NOTE, '备注'), (B_CHK, '校验')], C_BUY)
    rows = ctx.get('buy_rows', [])
    for i in range(BUY_R1 - BUY_R0 + 1):
        r = BUY_R0 + i
        d = rows[i] if i < len(rows) else {}
        put(ws, f'{B_SEQ}{r}', f'=IF(AND({B_SUP}{r}="",{B_GOODS}{r}=""),"",ROW()-{BUY_R0 - 1})', F_AUTO, FILL_AUTO, align=AC)
        put(ws, f'{B_DATE}{r}', d.get('date'), F_TXT, fmt=DATE, align=AC)
        put(ws, f'{B_SUP}{r}', d.get('supplier'), F_TXT, align=AC)
        put(ws, f'{B_GOODS}{r}', d.get('goods'), F_TXT, align=AC)
        put(ws, f'{B_QTY}{r}', d.get('qty'), F_TXT, fmt=QTY, align=AC)
        put(ws, f'{B_PRICE}{r}', d.get('price'), F_TXT, fmt=PRICE, align=AC)
        put(ws, f'{B_AMT}{r}', f'=IF(AND({B_QTY}{r}<>"",{B_PRICE}{r}<>""),{B_QTY}{r}*{B_PRICE}{r},"")', F_AUTO, FILL_AUTO, MONEY, AC)
        put(ws, f'{B_GIFT}{r}', d.get('gift'), F_TXT, fmt=QTY, align=AC)
        put(ws, f'{B_NOTE}{r}', d.get('note'), F_TXT, align=AL)
        blank = f'AND({B_DATE}{r}="",{B_SUP}{r}="",{B_GOODS}{r}="",{B_QTY}{r}="",{B_PRICE}{r}="",{B_GIFT}{r}="")'
        put(ws, f'{B_CHK}{r}',
            f'=IF({blank},"",IF(NOT(ISNUMBER({B_DATE}{r})),"✗ 日期不是真日期",IF(OR({B_DATE}{r}<年初日,{B_DATE}{r}>=年末日+1),"✗ 日期不在本年度",'
            f'IF({B_SUP}{r}="","✗ 没选供应商",IF(COUNTIF({SUP_NAMES},{B_SUP}{r})=0,"✗ 供应商不在基础资料里",'
            f'IF({B_GOODS}{r}="","✗ 没选商品",IF(COUNTIF({GOODS_NAMES},{B_GOODS}{r})=0,"✗ 商品不在商品档案里",'
            f'IF(AND({B_QTY}{r}="",{B_GIFT}{r}=""),"✗ 没填数量",IF(AND(N({B_QTY}{r})<>0,{B_PRICE}{r}=""),"✗ 没填进价","√")))))))))',
            F_AUTO, FILL_AUTO, align=AL)
    ws.conditional_formatting.add(f'{B_CHK}{BUY_R0}:{B_CHK}{BUY_R1}',
                                  FormulaRule(formula=[f'LEFT(${B_CHK}{BUY_R0},1)="✗"'], fill=FILL_WARN))
    add_date_dv(ws, f'{B_DATE}{BUY_R0}:{B_DATE}{BUY_R1}')
    add_list_dv(ws, f'{B_SUP}{BUY_R0}:{B_SUP}{BUY_R1}', '=供应商列表', '从下拉选供应商')
    add_list_dv(ws, f'{B_GOODS}{BUY_R0}:{B_GOODS}{BUY_R1}', '=商品列表', '从下拉选商品')
    widths(ws, {B_SEQ: 7, B_DATE: 13, B_SUP: 15, B_GOODS: 15, B_QTY: 11, B_PRICE: 12, B_AMT: 15,
                B_GIFT: 12, B_NOTE: 22, B_CHK: 22})
    ws.auto_filter.ref = f'A{BUY_HDR}:{B_CHK}{BUY_R1}'
    ws.freeze_panes = f'A{BUY_R0}'
    return ws


# ───────────────────────── 总览汇总（客户名单） ─────────────────────────
def fix_ov(wb, ctx):
    ws = wb[SH_OV]
    # 记住原来每个客户第一次出现那一行的样式（用户自己标的底色要跟着人走）
    orig_rows = {}
    for r in range(5, 155):
        b = ws.cell(r, 2).value
        if isinstance(b, str) and b.strip():
            n = ctx['cus_merge'].get(b.strip(), b.strip())
            orig_rows.setdefault(n, r)
    styles = {}
    for n, r in orig_rows.items():
        styles[n] = [copy(ws.cell(r, c)._style) for c in range(1, 11)]
    plain = [copy(ws.cell(100, c)._style) for c in range(1, 11)]
    tot_style = [copy(ws.cell(155, c)._style) for c in range(1, 11)]
    k_hdr_style = copy(ws['J3']._style)
    top_style = copy(ws['J4']._style)
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= 5:
            ws.unmerge_cells(str(mr))
    for r in range(5, max(ws.max_row, OV_TOT) + 1):
        for c in range(1, 21):
            ws.cell(r, c).value = None

    ws['A2'].value = '💡 一篮子查看所有客户情况：新增客户在 B 列往下加（C 填初始存条），领用和收款自动更新'
    ws['K3']._style = k_hdr_style
    ws['K3'].value = '期初欠款(元)'
    ws['K4']._style = top_style
    ws.column_dimensions['K'].width = 15
    ws.column_dimensions['J'].width = 19

    cust = ctx['customers']       # [(name, deposit, src_rows)]
    for i in range(OV_R1 - OV_R0 + 1):
        r = OV_R0 + i
        name, dep = (cust[i][0], cust[i][1]) if i < len(cust) else (None, None)
        st = styles.get(name, plain) if name else plain
        for c in range(1, 11):
            ws.cell(r, c)._style = copy(st[c - 1])
        ws.cell(r, 11)._style = copy(st[2])
        ws.row_dimensions[r].height = 15
        ws[f'A{r}'] = f'=IF(B{r}="","",COUNTA($B${OV_R0}:B{r}))'
        ws[f'B{r}'] = name
        ws[f'C{r}'] = dep
        ws[f'D{r}'] = f'=IF(B{r}<>"",SUMIF({SH_OUT}!C:C,B{r},{SH_OUT}!G:G),"")'
        ws[f'E{r}'] = f'=IF(AND(C{r}<>"",D{r}<>""),C{r}-D{r},"")'
        ws[f'F{r}'] = (f'=IF(B{r}<>"",SUMIFS({SH_CASH}!${K_NET}:${K_NET},{SH_CASH}!${K_CUS}:${K_CUS},B{r},'
                       f'{SH_CASH}!${K_TO}:${K_TO},"{TO_AR}"),"")')
        ws[f'G{r}'] = f'=IF(B{r}="","",N(K{r})+N(D{r})-N(F{r}))'
        ws[f'H{r}'] = f'=IF(B{r}="","",IF(ROUND(G{r},2)=0,"已结清",IF(G{r}>0,"有欠款","预存款")))'
        ws[f'I{r}'] = f'=IF(B{r}="","",IFERROR(SUMIF({SH_INV}!B:B,B{r},{SH_INV}!I:I),0))'
        ws[f'J{r}'] = f'=IF(B{r}="","",IFERROR(SUMIF({SH_INV}!B:B,B{r},{SH_INV}!J:J),0))'
        ws[f'K{r}'].number_format = MONEY
        ws[f'I{r}'].number_format = 'General'       # 原表这一列是数量（件），却套了 ¥ 金额格式
    r = OV_TOT
    for c in range(1, 12):
        ws.cell(r, c)._style = copy(tot_style[min(c, 10) - 1] if c != 2 else tot_style[0])
    ws.merge_cells(f'A{r}:B{r}')
    ws[f'A{r}'] = '合计'
    for col in 'CDEFGIJK':
        ws[f'{col}{r}'] = f'=SUM({col}{OV_R0}:{col}{OV_R1})'
    for col in 'CDEFGIJK':
        ws[f'{col}{OV_TOP}'] = f'={col}{OV_TOT}'
    ws[f'I{OV_TOT}'].number_format = 'General'
    ws[f'I{OV_TOP}'].number_format = 'General'
    ws[f'H{OV_TOP}'] = None
    # 重名提醒：同一个客户写了两行，SUMIF 会把他的领用/收款算两遍
    ws.conditional_formatting.add(f'B{OV_R0}:B{OV_R1}',
                                  FormulaRule(formula=[f'AND(B{OV_R0}<>"",COUNTIF($B${OV_R0}:$B${OV_R1},B{OV_R0})>1)'],
                                              fill=FILL_WARN))
    ws.auto_filter.ref = f'A{OV_TOP}:K{OV_R1}'
    ws.freeze_panes = 'A5'


# ───────────────────────── _辅助 + 商品库存（自动列出组合） ─────────────────────────
AUX_COMBO_R0, AUX_COMBO_R1 = 4, 4003      # A 列：存条 1000 行 + 出库 3000 行
AUX_Q_COL = 'D'                           # 客户查询取数


def build_aux_combo(wb, ctx):
    ws = wb[SH_AUX]
    put(ws, 'A3', '商品库存组合键', F_NOTE, border=False)
    put(ws, 'B3', '第k个组合在A列的位置', F_NOTE, border=False)
    put(ws, 'D3', '客户查询：第几个', F_NOTE, border=False)
    put(ws, 'A2', f'=COUNT($A${AUX_COMBO_R0}:$A${AUX_COMBO_R1})', F_NOTE, border=False)   # 组合个数
    cus = f'{SH_OV}!$B${OV_R0}:$B${OV_R1}'
    # 注意：这里全部「按位置」取明细表的第 k 行（INDEX），不直接写 存条明细!C100 这种地址——
    # 用户在明细表里删一行/插一行，区域会跟着伸缩，位置始终对得上，不会出 #REF! 也不会串行
    for sh, r0, r1, a0 in ((SH_CUN, CUN_R0, CUN_R1, AUX_COMBO_R0), (SH_OUT, OUT_R0, OUT_R1, AUX_COMBO_R0 + CUN_R1 - CUN_R0 + 1)):
        C = f'{sh}!$C${r0}:$C${r1}'
        Dg = f'{sh}!$D${r0}:$D${r1}'
        for k in range(1, r1 - r0 + 2):
            r = a0 + k - 1
            c, d = at(C, k), at(Dg, k)
            key = (f'IFERROR(MATCH({c},{cus},0),900)*10000000+IFERROR(MATCH({d},{GOODS_NAMES},0),900)*10000+ROW()')
            dup = f'COUNTIFS({upto(C, k)},{c},{upto(Dg, k)},{d})>1'
            if sh == SH_OUT:        # 出库里的组合，存条里已经有的就不重复列
                dup = f'OR(COUNTIFS({cun("C")},{c},{cun("D")},{d})>0,{dup})'
            ws[f'A{r}'] = f'=IF(OR({c}="",{d}=""),"",IF({dup},"",{key}))'
    for k in range(1, INV_R1 - INV_R0 + 2):
        r = AUX_COMBO_R0 + k - 1
        ws[f'B{r}'] = (f'=IF({k}>$A$2,"",IFERROR(MATCH(SMALL($A${AUX_COMBO_R0}:$A${AUX_COMBO_R1},{k}),'
                       f'$A${AUX_COMBO_R0}:$A${AUX_COMBO_R1},0),""))')
    # 公司库存「客户存条未提」：商品库存第 k 行这个组合，截至【公司库存】所选月底还没提走的数量（多提的按 0，不能抵别人的）
    put(ws, 'R3', '存条未提(公司库存月底)', F_NOTE, border=False)
    end = f'DATE(年度,{SH_STK}!$C$3+1,1)'
    IB, IC = f'{SH_INV}!$B${INV_R0}:$B${INV_R1}', f'{SH_INV}!$C${INV_R0}:$C${INV_R1}'
    for k in range(1, INV_R1 - INV_R0 + 2):
        r = AUX_COMBO_R0 + k - 1
        cb, cc = at(IB, k), at(IC, k)
        cond = lambda sh: f'{sh}!$C:$C,{cb},{sh}!$D:$D,{cc},{sh}!$B:$B,"<"&{end}'
        ws[f'R{r}'] = (f'=IF({cb}="",0,MAX(0,SUMIFS({SH_CUN}!$E:$E,{cond(SH_CUN)})'
                       f'-SUMIFS({SH_OUT}!$E:$E,{cond(SH_OUT)})))')
        # 客户查询：商品库存第 k 行是不是查询的客户，是的话是第几个
        ws[f'{AUX_Q_COL}{r}'] = (f'=IF(AND({cb}<>"",{cb}={SH_Q}!$B$4),COUNTIF({upto(IB, k)},{SH_Q}!$B$4),"")')


def fix_inv(wb, ctx):
    ws = wb[SH_INV]
    even = [copy(ws.cell(18, c)._style) for c in range(1, 12)]     # 原表偶数行：整行淡蓝，剩余数量红色加粗
    odd = [copy(ws.cell(19, c)._style) for c in range(1, 12)]      # 原表奇数行
    tot = [copy(ws.cell(298, c)._style) for c in range(1, 12)]
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= 4:
            ws.unmerge_cells(str(mr))
    for r in range(4, max(ws.max_row, INV_TOT) + 1):
        for c in range(1, 21):
            ws.cell(r, c).value = None
            ws.cell(r, c)._style = copy(ws.cell(300, 1)._style) if r > INV_TOT else ws.cell(r, c)._style
    ws['A2'].value = '💡 每个客户每个商品的库存：全自动，存条/出库里出现新的「客户+商品」这里自动多一行'
    pos = lambda r: f'{SH_AUX}!$B{r}'
    for r in range(INV_R0, INV_R1 + 1):
        for c in range(1, 12):
            ws.cell(r, c)._style = copy((even if r % 2 == 0 else odd)[c - 1])
        ws.row_dimensions[r].height = 15
        p = pos(r)
        ws[f'A{r}'] = f'=IF(B{r}="","",ROW()-{INV_R0 - 1})'
        ws[f'B{r}'] = (f'=IF({p}="","",IF({p}<={CUN_R1 - CUN_R0 + 1},INDEX({cun("C")},{p}),'
                       f'INDEX({out("C")},{p}-{CUN_R1 - CUN_R0 + 1})))')
        ws[f'C{r}'] = (f'=IF({p}="","",IF({p}<={CUN_R1 - CUN_R0 + 1},INDEX({cun("D")},{p}),'
                       f'INDEX({out("D")},{p}-{CUN_R1 - CUN_R0 + 1})))')
        cond = lambda sh: f'{sh}!C:C,$B{r},{sh}!D:D,$C{r}'
        ws[f'D{r}'] = f'=IF($B{r}="","",IFERROR(SUMIFS({SH_CUN}!E:E,{cond(SH_CUN)}),0))'
        ws[f'E{r}'] = (f'=IF(OR($B{r}="",N(D{r})=0),"",IFERROR(ROUND(SUMIFS({SH_CUN}!G:G,{cond(SH_CUN)})'
                       f'/SUMIFS({SH_CUN}!E:E,{cond(SH_CUN)},{SH_CUN}!F:F,">0"),2),""))')
        ws[f'F{r}'] = f'=IF($B{r}="","",IFERROR(SUMIFS({SH_CUN}!G:G,{cond(SH_CUN)}),0))'
        ws[f'G{r}'] = f'=IF($B{r}="","",IFERROR(SUMIFS({SH_OUT}!E:E,{cond(SH_OUT)}),0))'
        ws[f'H{r}'] = f'=IF($B{r}="","",IFERROR(SUMIFS({SH_OUT}!G:G,{cond(SH_OUT)}),0))'
        ws[f'I{r}'] = f'=IF($B{r}="","",D{r}-G{r})'
        ws[f'J{r}'] = f'=IF($B{r}="","",F{r}-H{r})'
        agg = lambda sh, r0, r1: (f'IFERROR(_xlfn.AGGREGATE(14,6,{sh}!$B${r0}:$B${r1}/(({sh}!$C${r0}:$C${r1}=$B{r})'
                                  f'*({sh}!$D${r0}:$D${r1}=$C{r})),1),0)')     # 最后一次进/出的日期（AGGREGATE：Excel 2010 起都有）
        ws[f'K{r}'] = f'=IF($B{r}="","",MAX({agg(SH_CUN, CUN_R0, CUN_R1)},{agg(SH_OUT, OUT_R0, OUT_R1)}))'
        ws[f'K{r}'].number_format = 'yyyy/mm/dd;;'
        ws[f'E{r}'].number_format = MONEY
    r = INV_TOT
    for c in range(1, 12):
        ws.cell(r, c)._style = copy(tot[c - 1])
    ws.merge_cells(f'A{r}:C{r}')
    ws[f'A{r}'] = '合计'
    for col in 'DFGHIJ':
        ws[f'{col}{r}'] = f'=SUM({col}{INV_R0}:{col}{INV_R1})'
    # 剩余数量为负：领用比存条多，标红
    ws.conditional_formatting.add(f'I{INV_R0}:J{INV_R1}',
                                  FormulaRule(formula=[f'AND(ISNUMBER($I{INV_R0}),$I{INV_R0}<0)'],
                                              font=Font(name=YH, sz=10, bold=True, color='FFC00000')))
    ws.auto_filter.ref = f'A{INV_HDR}:K{INV_R1}'
    ws.freeze_panes = 'A4'


# ───────────────────────── 客户查询（原表，修好） ─────────────────────────
Q_R0, Q_R1 = 12, 51                      # 一个客户最多列 40 个商品


def fix_q(wb, ctx):
    ws = wb[SH_Q]
    ws['A2'].value = '💡 在黄色格子选客户名称，自动显示该客户的所有库存和货款情况（打印对账单去【客户对账单】）'
    add_list_dv(ws, 'B4', '=客户列表', '选一个客户')
    for i, col in enumerate('ABCDEF'):
        v = f'VLOOKUP($B$4,{SH_OV}!$B${OV_R0}:$H${OV_R1},{i + 2},FALSE)'
        ws[f'{col}8'] = f'=IFERROR(IF({v}="","",{v}),"")'
    s12 = [copy(ws.cell(12, c)._style) for c in range(1, 8)]
    s13 = [copy(ws.cell(13, c)._style) for c in range(1, 8)]
    for r in range(Q_R0, max(ws.max_row, Q_R1) + 1):
        for c in range(1, 10):
            ws.cell(r, c).value = None
    src_cols = ['C', 'D', 'F', 'G', 'H', 'I', 'J']   # 商品 / 入库数 / 入库额 / 领用数 / 领用额 / 剩余数 / 剩余额
    for r in range(Q_R0, Q_R1 + 1):
        k = r - Q_R0 + 1
        st = s12 if (r - Q_R0) % 2 == 0 else s13
        for c in range(1, 8):
            ws.cell(r, c)._style = copy(st[c - 1])
        ws[f'I{r}'] = f'=IFERROR(MATCH({k},{SH_AUX}!${AUX_Q_COL}${AUX_COMBO_R0}:${AUX_Q_COL}${AUX_COMBO_R0 + INV_R1 - INV_R0},0),"")'
        for j, sc in enumerate(src_cols):
            col = 'ABCDEFG'[j]
            ws[f'{col}{r}'] = f'=IF($I{r}="","",INDEX({SH_INV}!${sc}${INV_R0}:${sc}${INV_R1},$I{r}))'
    ws.column_dimensions['I'].hidden = True
    put(ws, 'E4', '← 下拉选客户', F_NOTE, border=False)
