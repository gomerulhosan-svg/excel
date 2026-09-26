# -*- coding: utf-8 -*-
"""报表：成本计算（全月一次加权平均）/ 公司库存 / 费用汇总 / 利润表 / 资产负债表。

口径（写给以后改这份脚本的人）：
- 收入 = 出库明细的领用金额，按出库日期归月；
- 成本 = 出库数量 × 当月加权平均进价；当月没进过货、也没有期初库存时，用商品档案的「参考进价」估算，
  估算的部分在利润表上单列一行提示；
- 报损 = 各月报损表的报损金额，全额进利润表「商品报损损失」，同时冲减存货；
- 费用 / 返利 / 其他收入 / 投入 / 提取 / 借款 都从资金台帐按「去向」取；
- 资产负债表由上面这些流量滚出来，结构上保证 资产 = 负债 + 所有者权益（有对不上的都显式列成「待查」行）。
"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from common import *
from s_baosun import AUX_BS_P0, AUX_BS_TOT, AUX_BS_UNM, bsx_amt, bsx_pcs

F_ITAL = Font(name=YH, sz=9, italic=True, color='FF7F7F7F')
LAST_COST_COL = CL(COST_FIX + 12 * COST_W)          # EU


def _date_rng(sheet_col, m):
    return f'{sheet_col},">="&DATE(年度,{m},1),{sheet_col},"<"&DATE(年度,{m}+1,1)'


# ───────────────────────── 成本计算 ─────────────────────────
def build_cost(wb, ctx):
    ws = wb.create_sheet(SH_COST)
    ws.sheet_properties.tabColor = C_COST[2:]
    title(ws, '成本计算（全月一次加权平均 · 全自动，不用填）', 'N', C_COST,
          '💡 每个商品每个月：可用＝上月结存＋本月进货；平均单价＝可用金额÷可用数量；出库成本＝出库数量×平均单价；报损按报损表金额冲减。'
          '当月没货可算（没录进货、也没期初）就先用【基础资料】的参考进价估算，「估算成本」一栏会标出来。')
    ws.row_dimensions[4].hidden = True
    fixed = ['序号', '商品品类', '报损表品名', '每件数量', '参考进价', '期初数量', '期初金额']
    for i, t in enumerate(fixed):
        c = CL(i + 1)
        ws.merge_cells(f'{c}5:{c}6')
        put(ws, f'{c}5', t, F_HDR, fill(C_COST), align=ACW)
        put(ws, f'{c}6', None, F_HDR, fill(C_COST))
    month_colors = ['FF2F75B5', 'FF548235']
    for m in range(1, 13):
        c0 = cost_col(m, COST_FIELDS[0])
        c1 = cost_col(m, COST_FIELDS[-1])
        ws.merge_cells(f'{c0}5:{c1}5')
        col = month_colors[m % 2]
        put(ws, f'{c0}5', f'{m}月', F_HDR, fill(col), align=AC)
        for f in COST_FIELDS:
            cc = cost_col(m, f)
            put(ws, f'{cc}6', f, F_HDR, fill(col), align=ACW)
            ws[f'{cc}4'] = m
            ws.column_dimensions[cc].width = 13 if ('金额' in f or '成本' in f) else 10
        ws.column_dimensions[cost_col(m, '估算成本')].width = 9
    ws.row_dimensions[6].height = 30
    put(ws, 'A7', '单位', F_NOTE, FILL_AUTO, align=AC)
    for m in range(1, 13):
        for f in COST_FIELDS:
            put(ws, f'{cost_col(m, f)}7', '件' if '数量' in f else ('元/件' if f == '平均单价' else '元'), F_NOTE, FILL_AUTO, align=AC)

    for i in range(COST_R1 - COST_R0 + 1):
        r = COST_R0 + i
        br = BASE_R0 + i
        ar = AUX_BS_P0 + i
        b = f'$B{r}'
        put(ws, f'A{r}', f'=IF(B{r}="","",ROW()-{COST_R0 - 1})', F_AUTO, align=AC)
        g = lambda col: at(base_rng(col, BASE_R0, BASE_R1), i + 1)     # 按位置取商品档案第 i+1 个
        put(ws, f'B{r}', f'=IF({g(G_NAME)}="","",{g(G_NAME)}&"")', F_TXT, align=AC)
        put(ws, f'C{r}', f'=IF({g(G_BSNAME)}="","",{g(G_BSNAME)}&"")', F_NOTE, align=AC)
        put(ws, f'D{r}', f'=IF({g(G_PACK)}="","",{g(G_PACK)})', F_AUTO, align=AC)
        put(ws, f'E{r}', f'=IF({g(G_COST)}="","",{g(G_COST)})', F_AUTO, fmt=PRICE, align=AC)
        put(ws, f'F{r}', f'=N({g(G_Q0)})', F_AUTO, fmt=QTY, align=AC)
        put(ws, f'G{r}', f'=IF({g(G_A0)}="",0,{g(G_A0)})', F_AUTO, fmt=MONEY2, align=AC)
        for m in range(1, 13):
            C = lambda f: f'{cost_col(m, f)}{r}'
            prevQ = f'F{r}' if m == 1 else f'{cost_col(m - 1, "结存数量")}{r}'
            prevA = f'G{r}' if m == 1 else f'{cost_col(m - 1, "结存金额")}{r}'
            put(ws, C('采购数量'),
                f'=IF({b}="",0,SUMIFS({buy(B_QTY)},{buy(B_GOODS)},{b},{_date_rng(buy(B_DATE), m)})'
                f'+SUMIFS({buy(B_GIFT)},{buy(B_GOODS)},{b},{_date_rng(buy(B_DATE), m)}))', F_TXT, fmt=QTY, align=AR_)
            put(ws, C('采购金额'),
                f'=IF({b}="",0,SUMIFS({buy(B_AMT)},{buy(B_GOODS)},{b},{_date_rng(buy(B_DATE), m)}))', F_TXT, fmt=MONEY2, align=AR_)
            put(ws, C('可用数量'), f'={prevQ}+{C("采购数量")}', F_AUTO, fmt=QTY, align=AR_)
            put(ws, C('可用金额'), f'={prevA}+{C("采购金额")}', F_AUTO, fmt=MONEY2, align=AR_)
            put(ws, C('平均单价'),
                f'=IF(AND({C("可用数量")}>0,{C("可用金额")}>0),{C("可用金额")}/{C("可用数量")},'
                f'IF(AND({C("采购数量")}>0,{C("采购金额")}>0),{C("采购金额")}/{C("采购数量")},N($E{r})))',
                F_AUTO, fmt=PRICE, align=AR_)
            put(ws, C('出库数量'),
                f'=IF({b}="",0,SUMIFS({out(O_QTY)},{out(O_GOODS)},{b},{_date_rng(out(O_DATE), m)}))', F_TXT, fmt=QTY, align=AR_)
            put(ws, C('出库金额'),
                f'=IF({b}="",0,SUMIFS({out(O_AMT)},{out(O_GOODS)},{b},{_date_rng(out(O_DATE), m)}))', F_TXT, fmt=MONEY2, align=AR_)
            put(ws, C('出库成本'), f'={C("出库数量")}*{C("平均单价")}', F_TXT, fmt=MONEY2, align=AR_)
            put(ws, C('报损数量'), f'={SH_AUX}!{bsx_pcs(m)}{ar}', F_TXT, fmt=QTY, align=AR_)
            put(ws, C('报损金额'), f'={SH_AUX}!{bsx_amt(m)}{ar}', F_TXT, fmt=MONEY2, align=AR_)
            put(ws, C('结存数量'), f'={C("可用数量")}-{C("出库数量")}-{C("报损数量")}', F_TXTB, fmt=QTY, align=AR_)
            put(ws, C('结存金额'), f'={C("可用金额")}-{C("出库成本")}-{C("报损金额")}', F_TXTB, fmt=MONEY2, align=AR_)
            put(ws, C('估算成本'),
                f'=IF(OR(AND({C("可用数量")}>0,{C("可用金额")}>0),AND({C("采购数量")}>0,{C("采购金额")}>0)),0,{C("出库成本")})',
                F_ITAL, fmt=MONEY2, align=AR_)
    r = COST_TOT
    put(ws, f'B{r}', '合计', F_TXTB, FILL_TOT, align=AC)
    for c in ['F', 'G'] + [cost_col(m, f) for m in range(1, 13) for f in COST_FIELDS]:
        if c in [cost_col(m, '平均单价') for m in range(1, 13)]:
            put(ws, f'{c}{r}', None, F_TXTB, FILL_TOT)
            continue
        put(ws, f'{c}{r}', f'=SUM({c}{COST_R0}:{c}{COST_R1})', F_TXTB, FILL_TOT, MONEY2 if ('金额' in ws[f'{c}6'].value if ws[f'{c}6'].value else False) or c == 'G' or '成本' in (ws[f'{c}6'].value or '') else QTY, AR_)
    widths(ws, {'A': 5, 'B': 14, 'C': 15, 'D': 6, 'E': 9, 'F': 9, 'G': 11})
    ws.freeze_panes = f'H{COST_R0}'
    # 结存为负标红
    for m in range(1, 13):
        c = cost_col(m, '结存数量')
        ws.conditional_formatting.add(f'{c}{COST_R0}:{c}{COST_R1}',
                                      FormulaRule(formula=[f'{c}{COST_R0}<0'], font=Font(name=YH, sz=10, bold=True, color='FFC00000')))
    return ws


def cost_ytd(field, row, mcell):
    """成本计算第 row 行某字段 1..mcell 月累计"""
    return (f'SUMIFS({SH_COST}!$H{row}:${LAST_COST_COL}{row},{SH_COST}!$H$4:${LAST_COST_COL}$4,"<="&{mcell},'
            f'{SH_COST}!$H$6:${LAST_COST_COL}$6,"{field}")')


def cost_at(field, row, mcell):
    """成本计算第 row 行某字段在第 mcell 月的值"""
    return f'INDEX({SH_COST}!$H{row}:${LAST_COST_COL}{row},({mcell}-1)*{COST_W}+{COST_FIELDS.index(field) + 1})'


# ───────────────────────── 公司库存 ─────────────────────────
STK_HDR, STK_R0 = 5, 6
STK_R1 = STK_R0 + (BASE_R1 - BASE_R0)
STK_TOT = STK_R1 + 1


def build_stock(wb, ctx):
    ws = wb.create_sheet(SH_STK)
    ws.sheet_properties.tabColor = C_RPT[2:]
    title(ws, '公司库存（进销存汇总 · 按商品）', 'Q', C_RPT,
          '💡 公司自己仓库里的货：期初＋采购进货－出库领用－报损＝结存。右边对比客户存条还没提走的数量——'
          '「扣除存条后可用」为负，说明仓库里的货不够兑现客户已付款的存条。选月份看截止到那个月底的数。')
    put(ws, 'A3', '截止月份：', F_KPI_L, align=AR_, border=False)
    ws.merge_cells('A3:B3')
    put(ws, 'C3', ctx['kpi_month'], Font(name=YH, sz=13, bold=True, color='FF1F4E79'), FILL_SEL, '0"月"', AC)
    add_list_dv(ws, 'C3', '"1,2,3,4,5,6,7,8,9,10,11,12"', '看截止到哪个月底')
    put(ws, 'D3', '=TEXT(DATE(年度,C3+1,0),"yyyy年m月d日")&" 止"', F_NOTE, align=AL, border=False)
    ws.merge_cells('D3:F3')
    hdr = ['序号', '商品品类', '供应商', '期初数量', '期初金额', '采购数量\n(含赠品)', '采购金额', '出库数量', '出库成本',
           '报损数量\n(折件)', '报损金额', '结存数量', '结存金额', '平均单价', '客户存条\n未提(件)', '扣除存条后\n可用(件)', '提示']
    for i, t in enumerate(hdr):
        put(ws, f'{CL(i + 1)}{STK_HDR}', t, F_HDR, fill(C_RPT), align=ACW)
    ws.row_dimensions[STK_HDR].height = 34
    M = '$C$3'
    for i in range(STK_R1 - STK_R0 + 1):
        r = STK_R0 + i
        cr = COST_R0 + i
        br = BASE_R0 + i
        put(ws, f'A{r}', f'=IF(B{r}="","",ROW()-{STK_R0 - 1})', F_AUTO, align=AC)
        put(ws, f'B{r}', f'={SH_COST}!B{cr}', F_TXT, align=AC)
        put(ws, f'C{r}', f'=IF(B{r}="","",{at(base_rng(G_SUP, BASE_R0, BASE_R1), i + 1)}&"")', F_TXT, align=AC)
        put(ws, f'D{r}', f'=IF(B{r}="","",{SH_COST}!F{cr})', F_TXT, fmt=QTY, align=AR_)
        put(ws, f'E{r}', f'=IF(B{r}="","",{SH_COST}!G{cr})', F_TXT, fmt=MONEY2, align=AR_)
        for col, f in [('F', '采购数量'), ('G', '采购金额'), ('H', '出库数量'), ('I', '出库成本'), ('J', '报损数量'), ('K', '报损金额')]:
            put(ws, f'{col}{r}', f'=IF(B{r}="","",{cost_ytd(f, cr, M)})', F_TXT,
                fmt=MONEY2 if ('金额' in f or '成本' in f) else QTY, align=AR_)
        put(ws, f'L{r}', f'=IF(B{r}="","",{cost_at("结存数量", cr, M)})', F_TXTB, fmt=QTY, align=AR_)
        put(ws, f'M{r}', f'=IF(B{r}="","",{cost_at("结存金额", cr, M)})', F_TXTB, fmt=MONEY2, align=AR_)
        put(ws, f'N{r}', f'=IF(OR(B{r}="",N(L{r})<=0),"",M{r}/L{r})', F_AUTO, fmt=PRICE, align=AR_)
        put(ws, f'O{r}', f'=IF(B{r}="","",SUMIF({SH_INV}!$C${INV_R0}:$C${INV_R1},B{r},{SH_AUX}!$R$4:$R${4 + INV_R1 - INV_R0}))',
            F_TXT, fmt=QTY, align=AR_)
        put(ws, f'P{r}', f'=IF(B{r}="","",L{r}-O{r})', F_TXTB, fmt=QTY, align=AR_)
        put(ws, f'Q{r}',
            f'=IF(B{r}="","",IF(L{r}<0,"✗ 结存为负：进货没录全；","")'
            f'&IF(AND(N(H{r})>0,N(I{r})=0),"⚠ 没有进价，成本按 0 算；","")'
            f'&IF(AND(L{r}>=0,P{r}<0),"⚠ 不够兑现客户存条",""))', F_AUTO, align=AL)
    r = STK_TOT
    put(ws, f'B{r}', '合计', F_TXTB, FILL_TOT, align=AC)
    for col in 'DEFGHIJKLMOP':
        put(ws, f'{col}{r}', f'=SUM({col}{STK_R0}:{col}{STK_R1})', F_TXTB, FILL_TOT,
            MONEY2 if col in 'EGIKM' else QTY, AR_)
    for col in 'ACNQ':
        put(ws, f'{col}{r}', None, F_TXTB, FILL_TOT)
    ws.conditional_formatting.add(f'Q{STK_R0}:Q{STK_R1}', FormulaRule(formula=[f'LEFT($Q{STK_R0},1)="✗"'], fill=FILL_WARN))
    ws.conditional_formatting.add(f'Q{STK_R0}:Q{STK_R1}', FormulaRule(formula=[f'LEFT($Q{STK_R0},1)="⚠"'], fill=fill('FFFFEB9C')))
    ws.conditional_formatting.add(f'L{STK_R0}:L{STK_R1}', FormulaRule(formula=[f'AND(ISNUMBER(L{STK_R0}),L{STK_R0}<0)'],
                                                                      font=Font(name=YH, sz=10, bold=True, color='FFC00000')))
    widths(ws, {'A': 5, 'B': 15, 'C': 12, 'D': 9, 'E': 11, 'F': 10, 'G': 12, 'H': 10, 'I': 12, 'J': 9,
                'K': 11, 'L': 10, 'M': 14, 'N': 9, 'O': 10, 'P': 11, 'Q': 30})
    ws.freeze_panes = f'C{STK_R0}'
    ws.auto_filter.ref = f'A{STK_HDR}:Q{STK_R1}'
    return ws


# ───────────────────────── 费用汇总 ─────────────────────────
EXP_HDR, EXP_ROW0 = 4, 5
EXP_ROW1 = EXP_ROW0 + (EXP_R1 - EXP_R0)          # 54
EXP_UNASSIGNED = EXP_ROW1 + 1                    # 55
EXP_TOTAL = EXP_ROW1 + 2                         # 56
EXP_CLASS0 = EXP_ROW1 + 4                        # 58..63 按 EC_ALL 顺序
EXP_CLASS_ROW = {c: EXP_CLASS0 + i for i, c in enumerate(EC_ALL)}
EXP_WHO_HDR = EXP_CLASS0 + len(EC_ALL) + 2       # 按报销人表头
EXP_WHO_N = 40
EXP_MCOL = {m: CL(3 + m) for m in range(1, 13)}  # D..O


def build_exp(wb, ctx):
    ws = wb.create_sheet(SH_EXP)
    ws.sheet_properties.tabColor = C_RPT[2:]
    title(ws, '费用汇总（按费用项目 × 月份 · 自动）', 'Q', C_RPT,
          '💡 全部直接从【资金台帐】的「费用支出」统计（按费用项目、按日期归月）；底下按利润表归类小计，直接进利润表；'
          '再往下是按报销人汇总。没选或没对上费用项目的单列一行，并入管理费用。')
    hdr = ['序号', '费用项目', '利润表归类'] + [f'{m}月' for m in range(1, 13)] + ['全年合计', '占比']
    for i, t in enumerate(hdr):
        put(ws, f'{CL(i + 1)}{EXP_HDR}', t, F_HDR, fill(C_RPT), align=ACW)
    ws.row_dimensions[EXP_HDR].height = 30
    for m in range(1, 13):
        ws[f'{EXP_MCOL[m]}3'] = m
    ws.row_dimensions[3].hidden = True

    def exp_sum(m, item=None, who=None):
        c_extra = (f',{cash(K_EXP)},{item}' if item else '') + (f',{cash(K_WHO)},{who}' if who else '')
        return (f'-SUMIFS({cash(K_NET)},{cash(K_TO)},"{TO_EXP}"{c_extra},'
                f'{cash(K_DATE)},">="&DATE(年度,{m},1),{cash(K_DATE)},"<"&DATE(年度,{m}+1,1))')

    for i in range(EXP_ROW1 - EXP_ROW0 + 1):
        r = EXP_ROW0 + i
        br = EXP_R0 + i
        put(ws, f'A{r}', f'=IF(B{r}="","",ROW()-{EXP_ROW0 - 1})', F_AUTO, align=AC)
        put(ws, f'B{r}', f'=IF({at(EXP_NAMES, i + 1)}="","",{at(EXP_NAMES, i + 1)}&"")', F_TXT, align=AC)
        put(ws, f'C{r}', f'=IF(B{r}="","",{at(base_rng(E_CLASS, EXP_R0, EXP_R1), i + 1)}&"")', F_TXT, align=AC)
        for m in range(1, 13):
            put(ws, f'{EXP_MCOL[m]}{r}', f'=IF($B{r}="","",{exp_sum(m, f"$B{r}")})', F_TXT, fmt=MONEY2, align=AR_)
        put(ws, f'P{r}', f'=IF(B{r}="","",SUM(D{r}:O{r}))', F_TXTB, fmt=MONEY2, align=AR_)
        put(ws, f'Q{r}', f'=IF(OR(B{r}="",N($P${EXP_TOTAL})=0),"",P{r}/$P${EXP_TOTAL})', F_AUTO, fmt=PCT, align=AR_)
    r = EXP_UNASSIGNED
    put(ws, f'B{r}', '没选/没对上费用项目的', F_RED, align=AC)
    put(ws, f'C{r}', EC_ADMIN, F_TXT, align=AC)
    for m in range(1, 13):
        c = EXP_MCOL[m]
        put(ws, f'{c}{r}', f'={c}{EXP_TOTAL}-SUM({c}{EXP_ROW0}:{c}{EXP_ROW1})', F_TXT, fmt=MONEY2, align=AR_)
    put(ws, f'P{r}', f'=SUM(D{r}:O{r})', F_TXTB, fmt=MONEY2, align=AR_)
    r = EXP_TOTAL
    put(ws, f'B{r}', '费用合计', F_TXTB, FILL_TOT, align=AC)
    for m in range(1, 13):
        put(ws, f'{EXP_MCOL[m]}{r}', f'={exp_sum(m)}', F_TXTB, FILL_TOT, MONEY2, AR_)
    put(ws, f'P{r}', f'=SUM(D{r}:O{r})', F_TXTB, FILL_TOT, MONEY2, AR_)
    put(ws, f'B{EXP_CLASS0 - 1}', '按利润表归类', F_SEC, border=False)
    for cls, rr in EXP_CLASS_ROW.items():
        put(ws, f'B{rr}', cls, F_TXTB, FILL_SUBH, align=AC)
        for m in range(1, 13):
            c = EXP_MCOL[m]
            if cls == EC_ADMIN:     # 管理费用 = 合计 − 其他各类（没归类、归错类的都落到这里，保证分类加总 = 合计）
                others = '-'.join(f'{c}{EXP_CLASS_ROW[o]}' for o in EC_ALL if o != EC_ADMIN)
                f = f'={c}{EXP_TOTAL}-{others}'
            else:
                f = f'=SUMIF($C${EXP_ROW0}:$C${EXP_ROW1},"{cls}",{c}${EXP_ROW0}:{c}${EXP_ROW1})'
            put(ws, f'{c}{rr}', f, F_TXT, FILL_SUBH, MONEY2, AR_)
        put(ws, f'P{rr}', f'=SUM(D{rr}:O{rr})', F_TXTB, FILL_SUBH, MONEY2, AR_)
    # ── 按报销人（资金台帐「报销人」列，名字自动列出）──
    r0 = EXP_WHO_HDR
    put(ws, f'B{r0 - 1}', '按报销人（资金台帐里填了报销人的费用）', F_SEC, border=False)
    for i, t in enumerate(['序号', '报销人', ''] + [f'{m}月' for m in range(1, 13)] + ['全年合计', '占比']):
        put(ws, f'{CL(i + 1)}{r0}', t, F_HDR, fill(C_RPT), align=ACW)
    who_tot = f'$P${EXP_TOTAL}'
    for k in range(1, EXP_WHO_N + 1):
        r = r0 + k
        put(ws, f'A{r}', f'=IF(B{r}="","",{k})', F_AUTO, align=AC)
        put(ws, f'B{r}', f'=IFERROR(INDEX({cash(K_WHO)},SMALL({SH_AUX}!$S$4:$S${CASH_R1 - CASH_R0 + 4},{k})-3)&"","")', F_TXT, align=AC)
        put(ws, f'C{r}', None)
        for m in range(1, 13):
            put(ws, f'{EXP_MCOL[m]}{r}', f'=IF($B{r}="","",{exp_sum(m, who=f"$B{r}")})', F_TXT, fmt=MONEY2, align=AR_)
        put(ws, f'P{r}', f'=IF(B{r}="","",SUM(D{r}:O{r}))', F_TXTB, fmt=MONEY2, align=AR_)
        put(ws, f'Q{r}', f'=IF(OR(B{r}="",N({who_tot})=0),"",P{r}/{who_tot})', F_AUTO, fmt=PCT, align=AR_)
    r = r0 + EXP_WHO_N + 1
    put(ws, f'B{r}', '没填报销人的', F_RED, align=AC)
    put(ws, f'C{r}', None)
    for m in range(1, 13):
        c = EXP_MCOL[m]
        put(ws, f'{c}{r}', f'={c}{EXP_TOTAL}-SUM({c}{r0 + 1}:{c}{r0 + EXP_WHO_N})', F_TXT, fmt=MONEY2, align=AR_)
    put(ws, f'P{r}', f'=SUM(D{r}:O{r})', F_TXTB, fmt=MONEY2, align=AR_)
    put(ws, f'Q{r}', f'=IF(N({who_tot})=0,"",P{r}/{who_tot})', F_AUTO, fmt=PCT, align=AR_)
    widths(ws, {'A': 5, 'B': 16, 'C': 11, **{EXP_MCOL[m]: 10.5 for m in range(1, 13)}, 'P': 12, 'Q': 7})
    ws.freeze_panes = f'D{EXP_ROW0}'
    return ws


def build_aux_fee(wb, ctx):
    """_辅助 S 列：资金台帐第 k 行是费用、且报销人第一次出现时，记下行号（给费用汇总按报销人自动列名字）"""
    ws = wb[SH_AUX]
    put(ws, 'S3', '报销人首次出现', F_NOTE, border=False)
    for k in range(1, CASH_R1 - CASH_R0 + 2):
        r = k + 3
        e, t = at(cash(K_WHO), k), at(cash(K_TO), k)
        ws[f'S{r}'] = (f'=IF(OR({e}="",{t}<>"{TO_EXP}"),"",IF(COUNTIFS({upto(cash(K_WHO), k)},{e},'
                       f'{upto(cash(K_TO), k)},"{TO_EXP}")>1,"",ROW()))')


# ───────────────────────── 利润表 ─────────────────────────
PL_M0 = 5                               # E 列 = 1 月
PL_MCOL = {m: CL(PL_M0 + m - 1) for m in range(1, 13)}  # E..P
PL = {}                                 # 行名 → 行号（给资产负债表、首页用）


def build_pl(wb, ctx):
    ws = wb.create_sheet(SH_PL)
    ws.sheet_properties.tabColor = C_RPT[2:]
    title(ws, '利润表（按月 · 自动）', 'Q', C_RPT,
          '💡 不用填：收入取【出库明细】，成本取【成本计算】，费用/返利/营业外收支取【资金台帐】，报损取各月报损表。'
          '选报表月份看「本月」和「本年累计」，右边是 1～12 月逐月数。')
    put(ws, 'A3', '报表月份：', F_KPI_L, align=AR_, border=False)
    put(ws, 'B3', ctx['kpi_month'], Font(name=YH, sz=13, bold=True, color='FF1F4E79'), FILL_SEL, '0"月"', AC)
    add_list_dv(ws, 'B3', '"1,2,3,4,5,6,7,8,9,10,11,12"', '选月份')
    put(ws, 'C3', '=年度&"年"&B3&"月"', F_KPI_L, align=AL, border=False)
    put(ws, 'E3', '=IF(基础资料!B6="","",基础资料!B6)', F_KPI_L, align=AL, border=False)
    hdr = [('A', '项  目'), ('B', '行次'), ('C', '本月金额'), ('D', '本年累计')] + \
          [(PL_MCOL[m], f'{m}月') for m in range(1, 13)] + [('Q', '全年合计')]
    for col, t in hdr:
        put(ws, f'{col}5', t, F_HDR, fill(C_RPT if col in 'ABQ' else ('FFC65911' if col in 'CD' else 'FF548235')), align=ACW)
    ws.row_dimensions[5].height = 26
    for m in range(1, 13):
        ws[f'{PL_MCOL[m]}4'] = m
    ws.row_dimensions[4].hidden = True

    cash_to = lambda to, m: (f'SUMIFS({cash(K_NET)},{cash(K_TO)},"{to}",{cash(K_DATE)},">="&DATE(年度,{m},1),'
                             f'{cash(K_DATE)},"<"&DATE(年度,{m}+1,1))')
    exp_cls = lambda cls, m: f'INDEX({SH_EXP}!$D${EXP_CLASS_ROW[cls]}:$O${EXP_CLASS_ROW[cls]},{m})'
    cost_tot = lambda f, m: f'INDEX({SH_COST}!$H${COST_TOT}:${LAST_COST_COL}${COST_TOT},({m}-1)*{COST_W}+{COST_FIELDS.index(f) + 1})'

    lines = [
        # (key, 标题, 月公式生成器 or ('calc', 表达式模板), 样式)
        ('rev', '一、营业收入', ('sum', ['rev_main', 'rev_reb']), 'L1'),
        ('rev_main', '    主营业务收入（出库领用）', lambda m: f'SUMIFS({out(O_AMT)},{_date_rng(out(O_DATE), m)})', 'L2'),
        ('rev_reb', '    其他业务收入（厂家返利/补贴）', lambda m: cash_to(TO_REB, m), 'L2'),
        ('cogs', '减：营业成本', ('sum', ['cogs_main']), 'L1'),
        ('cogs_main', '    主营业务成本（出库成本）', lambda m: cost_tot('出库成本', m), 'L2'),
        ('cogs_est', '      其中：按参考进价估算的', lambda m: cost_tot('估算成本', m), 'MEMO'),
        ('rev_zero', '      另：没有进价、成本按 0 算的销售额', lambda m: (
            f'SUMIFS({SH_COST}!${cost_col(m, "出库金额")}${COST_R0}:${cost_col(m, "出库金额")}${COST_R1},'
            f'{SH_COST}!${cost_col(m, "出库成本")}${COST_R0}:${cost_col(m, "出库成本")}${COST_R1},0,'
            f'{SH_COST}!${cost_col(m, "出库数量")}${COST_R0}:${cost_col(m, "出库数量")}${COST_R1},">0")'), 'MEMO'),
        ('gp', '二、毛利（主营收入－主营成本）', ('diff', 'rev_main', ['cogs_main']), 'L1'),
        ('gpm', '    毛利率', ('ratio', 'gp', 'rev_main'), 'PCT'),
        ('tax', '减：税金及附加', lambda m: exp_cls(EC_TAX, m), 'L2'),
        ('sell', '    销售费用', lambda m: exp_cls(EC_SELL, m), 'L2'),
        ('admin', '    管理费用', lambda m: exp_cls(EC_ADMIN, m), 'L2'),
        ('fin', '    财务费用', lambda m: exp_cls(EC_FIN, m), 'L2'),
        ('loss', '    商品报损损失', lambda m: f'INDEX({SH_AUX}!${bsx_amt(1)}${AUX_BS_TOT}:${bsx_amt(12)}${AUX_BS_TOT},{m})', 'L2'),
        ('op', '三、营业利润', ('diff', 'rev', ['cogs', 'tax', 'sell', 'admin', 'fin', 'loss']), 'L1'),
        ('noi', '加：营业外收入', lambda m: cash_to(TO_OI, m), 'L2'),
        ('noe', '减：营业外支出', lambda m: exp_cls(EC_NOI, m), 'L2'),
        ('ebt', '四、利润总额', ('expr', '{op}+{noi}-{noe}'), 'L1'),
        ('it', '减：所得税费用', lambda m: exp_cls(EC_IT, m), 'L2'),
        ('np', '五、净利润', ('diff', 'ebt', ['it']), 'NP'),
        ('npm', '    净利率', ('ratio', 'np', 'rev'), 'PCT'),
        (None, '', None, 'BLANK'),
        (None, '参考指标', None, 'SEC'),
        ('qty', '    出库数量（件）', lambda m: f'SUMIFS({out(O_QTY)},{_date_rng(out(O_DATE), m)})', 'QTY'),
        ('recv', '    客户回款', lambda m: cash_to(TO_AR, m), 'L2'),
        ('paid', '    付供应商货款', lambda m: f'-{cash_to(TO_AP, m)}', 'L2'),
        ('buy', '    采购进货金额', lambda m: f'SUMIFS({buy(B_AMT)},{_date_rng(buy(B_DATE), m)})', 'L2'),
        ('cashend', '    月末资金余额', lambda m: (f'SUM({SH_BASE}!${A_BAL0}${ACC_R0}:${A_BAL0}${ACC_R1})+SUMIFS({cash(K_NET)},'
                                             f'{cash(K_DATE)},">="&年初日,{cash(K_DATE)},"<"&DATE(年度,{m}+1,1))'), 'BAL'),
    ]
    r = 6
    rows = {}
    for key, _, _, _ in lines:
        if key:
            rows[key] = None
    # 先排行号
    rr = 6
    for key, t, fn, st in lines:
        if key:
            rows[key] = rr
        rr += 1
    PL.update(rows)
    MC = '$B$3'
    for key, t, fn, st in lines:
        font = {'L1': F_TXTB, 'NP': Font(name=YH, sz=11, bold=True, color='FFC00000'), 'MEMO': F_ITAL,
                'SEC': F_SEC}.get(st, F_TXT)
        fl = {'L1': FILL_SUBH, 'NP': FILL_TOT}.get(st)
        if st == 'BLANK':
            r += 1
            continue
        if st == 'SEC':
            put(ws, f'A{r}', t, F_SEC, border=False)
            r += 1
            continue
        put(ws, f'A{r}', t, font, fl, align=AL)
        put(ws, f'B{r}', r - 5, F_NOTE, fl, align=AC)
        fmt = PCT if st == 'PCT' else (QTY if st == 'QTY' else MONEY2)
        for m in range(1, 13):
            c = PL_MCOL[m]
            if callable(fn):
                f = '=' + fn(m)
            elif fn[0] == 'sum':
                f = '=' + '+'.join(f'{c}{rows[k]}' for k in fn[1])
            elif fn[0] == 'diff':
                f = f'={c}{rows[fn[1]]}-' + '-'.join(f'{c}{rows[k]}' for k in fn[2])
            elif fn[0] == 'ratio':
                f = f'=IF(N({c}{rows[fn[2]]})=0,"",{c}{rows[fn[1]]}/{c}{rows[fn[2]]})'
            elif fn[0] == 'expr':
                f = '=' + fn[1].format(**{k: f'{c}{rows[k]}' for k in rows})
            put(ws, f'{c}{r}', f, font, fl, fmt, AR_)
        # 本月 / 累计 / 全年
        if st == 'PCT':
            for col in ('C', 'D', 'Q'):
                put(ws, f'{col}{r}', f'=IF(N({col}{rows[fn[2]]})=0,"",{col}{rows[fn[1]]}/{col}{rows[fn[2]]})', font, fl, PCT, AR_)
        elif st == 'BAL':
            put(ws, f'C{r}', f'=INDEX(E{r}:P{r},{MC})', font, fl, fmt, AR_)
            put(ws, f'D{r}', f'=INDEX(E{r}:P{r},{MC})', font, fl, fmt, AR_)
            put(ws, f'Q{r}', f'=P{r}', font, fl, fmt, AR_)
        else:
            put(ws, f'C{r}', f'=INDEX(E{r}:P{r},{MC})', font, fill(fl.fgColor.rgb) if fl else fill('FFFFF2CC'), fmt, AR_)
            put(ws, f'D{r}', f'=SUMPRODUCT(($E$4:$P$4<={MC})*E{r}:P{r})', font, fill(fl.fgColor.rgb) if fl else fill('FFFFF2CC'), fmt, AR_)
            put(ws, f'Q{r}', f'=SUM(E{r}:P{r})', font, fl, fmt, AR_)
        r += 1
    widths(ws, {'A': 32, 'B': 5, 'C': 14, 'D': 14, **{PL_MCOL[m]: 12.5 for m in range(1, 13)}, 'Q': 14})
    ws.freeze_panes = 'C6'
    put(ws, f'A{r + 1}', '说明：没录进货时，成本先按【基础资料】商品档案的「参考进价」估算（上面「其中」那一行）；'
                         '参考进价也没填的商品成本按 0 算，毛利会虚高——【公司库存】的提示列会标出来。', F_NOTE, align=ALW, border=False)
    ws.merge_cells(f'A{r + 1}:Q{r + 2}')
    return ws


# ───────────────────────── 资产负债表 ─────────────────────────
BAL = {}


def build_bal(wb, ctx):
    ws = wb.create_sheet(SH_BAL)
    ws.sheet_properties.tabColor = C_RPT[2:]
    title(ws, '资产负债表（简易 · 自动）', 'I', C_RPT,
          '💡 不用填：由期初数（基础资料、总览汇总 K 列）加上本年各张明细滚出来。「平衡检查」必须是 √；'
          '底下「附注」列出了几项要留意的数（没登记的客户、没分类的收支等），不为 0 就去对应的表里查。')
    put(ws, 'A3', '报表月份：', F_KPI_L, align=AR_, border=False)
    put(ws, 'B3', ctx['kpi_month'], Font(name=YH, sz=13, bold=True, color='FF1F4E79'), FILL_SEL, '0"月"', AC)
    add_list_dv(ws, 'B3', '"1,2,3,4,5,6,7,8,9,10,11,12"', '选月份：看这个月底的数')
    put(ws, 'C3', '截止日期：', F_KPI_L, align=AR_, border=False)
    put(ws, 'D3', '=DATE(年度,B3+1,0)', F_KPI_L, fmt='yyyy"年"m"月"d"日"', align=AL, border=False)
    ws.merge_cells('D3:E3')
    D = '$D$3'
    M = '$B$3'
    hdr = [('A', '资  产'), ('B', '行次'), ('C', '期末余额'), ('D', '年初余额'),
           ('F', '负债和所有者权益'), ('G', '行次'), ('H', '期末余额'), ('I', '年初余额')]
    for col, t in hdr:
        put(ws, f'{col}5', t, F_HDR, fill(C_RPT), align=ACW)
    ws.row_dimensions[5].height = 26

    rng = lambda extra='': (f'{cash(K_DATE)},">="&年初日,{cash(K_DATE)},"<"&({D}+1){extra}')
    s_to = lambda to: f'SUMIFS({cash(K_NET)},{cash(K_TO)},"{to}",{rng()})'
    acc0 = f'SUM({SH_BASE}!${A_BAL0}${ACC_R0}:${A_BAL0}${ACC_R1})'
    cus_bal = f'{SH_AUX}!$J${OV_R0}:$J${OV_R1}'
    sup_bal = f'{SH_AUX}!$K${SUP_R0}:$K${SUP_R1}'

    # 行号布局：负债权益 6..17，资产合计对齐到 17
    L = ['ap', 'adv', 'loan', 'oap', 'ltot', 'eq0', 'inj', 'draw', 'np', 'etot', 'letot']
    for i, k in enumerate(L):
        BAL[k] = 6 + i
    TOT = BAL['letot']                 # 17
    BAL['atot'] = TOT
    DIFF = TOT + 2                     # 19 平衡检查
    OTHR = TOT + 4                     # 其他往来净额（辅助）
    N0 = TOT + 6                       # 附注标题
    BAL['diff'] = DIFF
    note_keys = ['cus_net', 'cus_res', 'sup_net', 'sup_res', 'unk', 'xfer', 'bs_unm', 'buy_unm']
    note_rows = {k: N0 + 1 + i for i, k in enumerate(note_keys)}
    BAL.update({f'n_{k}': v for k, v in note_rows.items()})
    NR = lambda k: f'$C${note_rows[k]}'
    H = lambda k: f'H{BAL[k]}'
    I = lambda k: f'I{BAL[k]}'

    assets = [
        ('cash', '货币资金', f'={acc0}+SUMIFS({cash(K_NET)},{rng()})', f'={acc0}'),
        ('ar', '应收账款', f'=SUMIF({cus_bal},">0")+MAX(0,{NR("cus_res")})', f'=SUMIF({CUS_OPEN},">0")'),
        ('prepay', '预付账款', f'=-SUMIF({sup_bal},"<0")+MAX(0,-{NR("sup_res")})', f'=-SUMIF({base_rng(S_AP0, SUP_R0, SUP_R1)},"<0")'),
        ('oar', '其他应收款', f'=MAX(0,$C${OTHR})', f'=MAX(0,{P_OTH0})'),
        ('inv', '存货', f'={cost_at("结存金额", COST_TOT, M)}-{NR("bs_unm")}+{NR("buy_unm")}',
         f'=SUM({base_rng(G_A0, BASE_R0, BASE_R1)})'),
    ]
    r = 6
    for key, t, fc, fd in assets:
        BAL[key] = r
        put(ws, f'A{r}', t, F_TXT, align=AL)
        put(ws, f'B{r}', r - 5, F_NOTE, align=AC)
        put(ws, f'C{r}', fc, F_TXT, fmt=MONEY2, align=AR_)
        put(ws, f'D{r}', fd, F_TXT, fmt=MONEY2, align=AR_)
        r += 1
    for rr in range(r, TOT):
        for col in 'ABCD':
            put(ws, f'{col}{rr}', None)

    liabs = {
        'ap': ('应付账款', f'=SUMIF({sup_bal},">0")+MAX(0,{NR("sup_res")})', f'=SUMIF({base_rng(S_AP0, SUP_R0, SUP_R1)},">0")'),
        'adv': ('预收账款（客户存条/预存款）', f'=-SUMIF({cus_bal},"<0")+MAX(0,-{NR("cus_res")})', f'=-SUMIF({CUS_OPEN},"<0")'),
        'loan': ('短期借款', f'={P_LOAN0}+{s_to(TO_LOAN)}', f'={P_LOAN0}'),
        'oap': ('其他应付款', f'=MAX(0,-$C${OTHR})', f'=MAX(0,-{P_OTH0})'),
        'ltot': ('负债合计', f'=SUM(H{BAL["ap"]}:H{BAL["oap"]})', f'=SUM(I{BAL["ap"]}:I{BAL["oap"]})'),
        'eq0': ('年初净资产（期初权益）', f'={I("eq0")}', f'=D{TOT}-{I("ltot")}'),
        'inj': ('加：本年股东投入', f'={s_to(TO_INV)}', None),
        'draw': ('减：本年股东提取', f'=-{s_to(TO_DRAW)}', None),
        'np': ('加：本年利润（截至报表月）', f'=SUMPRODUCT(({SH_PL}!$E$4:$P$4<={M})*{SH_PL}!$E${PL["np"]}:$P${PL["np"]})', None),
        'etot': ('所有者权益合计', f'={H("eq0")}+{H("inj")}-{H("draw")}+{H("np")}', f'={I("eq0")}'),
        'letot': ('负债和所有者权益合计', f'={H("ltot")}+{H("etot")}', f'={I("ltot")}+{I("etot")}'),
    }
    for key in L:
        t, fc, fd = liabs[key]
        r = BAL[key]
        bold = key in ('ltot', 'etot', 'letot')
        fl = FILL_TOT if bold else None
        put(ws, f'F{r}', t, F_TXTB if bold else F_TXT, fl, align=AC if bold else AL)
        put(ws, f'G{r}', r + 1, F_NOTE, fl, align=AC)
        put(ws, f'H{r}', fc, F_TXTB if bold else F_TXT, fl, MONEY2, AR_)
        put(ws, f'I{r}', fd, F_TXTB if bold else F_TXT, fl, MONEY2, AR_)
    put(ws, f'A{TOT}', '资产总计', F_TXTB, FILL_TOT, align=AC)
    put(ws, f'B{TOT}', 6, F_NOTE, FILL_TOT, align=AC)
    put(ws, f'C{TOT}', f'=SUM(C6:C{TOT - 1})', F_TXTB, FILL_TOT, MONEY2, AR_)
    put(ws, f'D{TOT}', f'=SUM(D6:D{TOT - 1})', F_TXTB, FILL_TOT, MONEY2, AR_)
    # 平衡检查
    put(ws, f'A{DIFF}', '平衡检查', F_SEC, border=False)
    put(ws, f'C{DIFF}', f'=ROUND(C{TOT}-H{TOT},2)', F_TXTB, fmt=MONEY2, align=AR_)
    put(ws, f'D{DIFF}', f'=IF(ABS(C{DIFF})<0.01,"√ 资产＝负债＋所有者权益","✗ 差 "&TEXT(C{DIFF},"#,##0.00")&" 元，看附注")',
        F_TXTB, align=AL, border=False)
    ws.merge_cells(f'D{DIFF}:I{DIFF}')
    ws.conditional_formatting.add(f'D{DIFF}', FormulaRule(formula=[f'LEFT($D${DIFF},1)="√"'], font=Font(name=YH, sz=11, bold=True, color='FF00B050')))
    ws.conditional_formatting.add(f'D{DIFF}', FormulaRule(formula=[f'LEFT($D${DIFF},1)="✗"'], font=Font(name=YH, sz=11, bold=True, color='FFC00000')))
    # 其他往来净额（正=别人欠我，负=我欠别人）：年初 + 其他往来 + 没分类的 + 没配对的内部转账
    put(ws, f'A{OTHR}', '其他往来净额（辅助：正数进其他应收，负数进其他应付）', F_NOTE, align=AL)
    put(ws, f'C{OTHR}', f'={P_OTH0}-{s_to(TO_OTH)}-{NR("unk")}-{NR("xfer")}', F_NOTE, fmt=MONEY2, align=AR_)

    # 附注
    put(ws, f'A{N0}', '附注（不为 0 的要去查）', F_SEC, border=False)
    cus_total = (f'SUM({CUS_OPEN})+SUMIFS({out(O_AMT)},{out(O_DATE)},">="&年初日,{out(O_DATE)},"<"&({D}+1))'
                 f'-{s_to(TO_AR)}')
    sup_total = (f'SUM({base_rng(S_AP0, SUP_R0, SUP_R1)})+SUMIFS({buy(B_AMT)},{buy(B_DATE)},">="&年初日,{buy(B_DATE)},"<"&({D}+1))'
                 f'+{s_to(TO_AP)}')
    known = '+'.join(s_to(t) for t in TO_ALL)
    notes = {
        'cus_net': ('客户往来净额（正＝客户欠我们，负＝客户预存）', f'={cus_total}', ''),
        'cus_res': ('　其中：不在总览汇总名单里的客户（或资金台帐里没选客户的回款）', f'=C{note_rows["cus_net"]}-SUM({cus_bal})',
                    '→ 看【数据校验】客户那几项'),
        'sup_net': ('供应商往来净额（正＝我们欠厂家，负＝预付）', f'={sup_total}', ''),
        'sup_res': ('　其中：不在供应商档案里的供应商', f'=C{note_rows["sup_net"]}-SUM({sup_bal})', '→ 看【数据校验】'),
        'unk': ('资金台帐里没分类的收支净额（算进其他往来）', f'=SUMIFS({cash(K_NET)},{rng()})-({known})', '→ 资金台帐校验列'),
        'xfer': ('没配对的内部转账（转出≠转入，算进其他往来）', f'={s_to(TO_XFER)}', '→ 内部转账要记成一出一进两行'),
        'bs_unm': ('报损表里没对上商品档案的报损金额（已从存货里减）',
                   f'=SUMPRODUCT(({SH_AUX}!${bsx_amt(1)}$4:${bsx_amt(12)}$4<={M})*{SH_AUX}!${bsx_amt(1)}${AUX_BS_UNM}:${bsx_amt(12)}${AUX_BS_UNM})',
                   '→ 基础资料·商品档案「报损表品名」'),
        'buy_unm': ('采购进货里不在商品档案的商品金额（已加进存货）',
                    f'=SUMIFS({buy(B_AMT)},{buy(B_DATE)},">="&年初日,{buy(B_DATE)},"<"&({D}+1))-{cost_ytd("采购金额", COST_TOT, M)}',
                    '→ 采购进货校验列'),
    }
    for key in note_keys:
        t, f, hint = notes[key]
        rr = note_rows[key]
        put(ws, f'A{rr}', t, F_TXT, align=ALW)
        ws.merge_cells(f'A{rr}:B{rr}')
        put(ws, f'C{rr}', f, F_TXT, fmt=MONEY2, align=AR_)
        put(ws, f'D{rr}', hint, F_NOTE, align=AL, border=False)
        ws.row_dimensions[rr].height = 28
    widths(ws, {'A': 30, 'B': 5, 'C': 16, 'D': 16, 'E': 2, 'F': 30, 'G': 5, 'H': 16, 'I': 16})
    return ws


def build_aux_balances(wb, ctx):
    """_辅助 J 列：每个客户截至资产负债表日的余额；K 列：每个供应商"""
    ws = wb[SH_AUX]
    D = f'{SH_BAL}!$D$3'
    put(ws, 'J3', '客户余额(资产负债表日)', F_NOTE, border=False)
    put(ws, 'K3', '供应商余额(资产负债表日)', F_NOTE, border=False)
    for r in range(OV_R0, OV_R1 + 1):
        k = r - OV_R0 + 1
        nm, op = at(CUS_NAMES, k), at(CUS_OPEN, k)
        ws[f'J{r}'] = (f'=IF({nm}="","",N({op})+SUMIFS({out(O_AMT)},{out(O_CUS)},{nm},'
                       f'{out(O_DATE)},">="&年初日,{out(O_DATE)},"<"&({D}+1))-SUMIFS({cash(K_NET)},{cash(K_CUS)},{nm},'
                       f'{cash(K_TO)},"{TO_AR}",{cash(K_DATE)},">="&年初日,{cash(K_DATE)},"<"&({D}+1)))')
    for r in range(SUP_R0, SUP_R1 + 1):
        k = r - SUP_R0 + 1
        nm, op = at(SUP_NAMES, k), at(base_rng(S_AP0, SUP_R0, SUP_R1), k)
        ws[f'K{r}'] = (f'=IF({nm}="","",N({op})+SUMIFS({buy(B_AMT)},{buy(B_SUP)},{nm},'
                       f'{buy(B_DATE)},">="&年初日,{buy(B_DATE)},"<"&({D}+1))+SUMIFS({cash(K_NET)},{cash(K_SUP)},{nm},'
                       f'{cash(K_TO)},"{TO_AP}",{cash(K_DATE)},">="&年初日,{cash(K_DATE)},"<"&({D}+1)))')
