# -*- coding: utf-8 -*-
"""报损（原表 2）：
- 【报损明细台账】原来一个月一张（6 月「报损明细台账」、7 月、8 月），合成一年一张：列还是原来那些
  （所属公司/产品名称/规格型号/报损数量/单位/产品单价/报损金额/报损原因/备注），前面加一列「月份」。
  按月看：表头筛选选月份（上面「筛选后合计」跟着变），或者第 3 行选月份/公司直接出合计。
- 【报损汇总一览】保留原来的内容（分公司、分品项排行），按所选月份汇总。

报损数据的「中转站」放在 _辅助 表右边（每个商品每个月的报损金额/数量），
成本计算、利润表、报损汇总一览都从中转站取。原来要用 INDIRECT 去各月表取数，现在一张表直接 SUMIFS。"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from common import *

# ── _辅助 里报损中转站的位置 ──
AUX_BS_P0 = 6                                  # 商品块：第 6..105 行 ↔ 基础资料 5..104
AUX_BS_P1 = AUX_BS_P0 + (BASE_R1 - BASE_R0)
AUX_BS_S0 = 110                                # 供应商块：第 110..159 行 ↔ 基础资料 Q5..Q54
AUX_BS_S1 = AUX_BS_S0 + (SUP_R1 - SUP_R0)
AUX_BS_TOT = 162                               # 每月报损总金额
AUX_BS_TOTQ = 163                              # 每月报损总数量（原始数，瓶/罐）
AUX_BS_UNM = 164                               # 每月没对上商品档案的报损金额
AUX_BS_NOPRICE = 165                           # 每月「有数量没单价」的行数
AUX_BS_NOCOMP = 166                            # 每月「有产品没填所属公司」的行数
AUX_BS_NONNUM = 167                            # 每月数量/单价不是数字的格数
AUX_BS_NOH = 168                               # 每月「数量、单价都有，金额格却没公式」的行数（中间插了行）
BSX_NAME, BSX_PACK = 'AA', 'AB'
BSX_AMT0 = CI('AC')          # 金额 12 列 AC..AN
BSX_RAW0 = BSX_AMT0 + 12     # 原始数量 AO..AZ
BSX_BOX0 = BSX_AMT0 + 24     # 按箱/件录的数量 BA..BL
BSX_PCS0 = BSX_AMT0 + 36     # 折成件 BM..BX
BSX_SEL = CL(BSX_AMT0 + 48)  # BY 选定月份金额
BSX_KEY = CL(BSX_AMT0 + 49)  # BZ 排名键
BSX_SELQ = CL(BSX_AMT0 + 50) # CA 选定月份数量

# ── 报损汇总一览 ──
BSS_SEL = '$A$4'             # 统计月份（全年 / 1..12）
BSS_TOT = 'B4'               # 所选月份报损总金额
BSS_C0 = 8                   # 分公司：第 8..12 行 = 供应商前 5 家，13 = 其他，14 = 合计
BSS_CTOT = BSS_C0 + 6
BSS_P0 = 18                  # 分品项：第 18..37 行 = 前 20 名，38 = 其余，39 = 合计
BSS_PN = 20
BSS_PTOT = BSS_P0 + BSS_PN + 1


def bsx_amt(m):  return CL(BSX_AMT0 + m - 1)
def bsx_pcs(m):  return CL(BSX_PCS0 + m - 1)


def build(wb, ctx, src2=None):
    _build_ledger(wb, ctx)
    _build_staging(wb, ctx)
    _build_summary(wb, ctx)


# ───────────────────────── 报损明细台账（一年一张） ─────────────────────────
def _crit(base, month_cell, comp_cell, value_rng=None, count=False):
    """查询：月份（全年/1..12）× 所属公司（空＝全部）四种组合"""
    fn = 'COUNTIFS' if count else 'SUMIFS'
    head = f'{fn}(' + ('' if count else f'{value_rng},') + base
    m = f',{L_MON}${BSL_R0}:{L_MON}${BSL_R1},{month_cell}'
    c = f',{L_COMP}${BSL_R0}:{L_COMP}${BSL_R1},{comp_cell}'
    y = f',{L_MON}${BSL_R0}:{L_MON}${BSL_R1},">=1",{L_MON}${BSL_R0}:{L_MON}${BSL_R1},"<=12"'   # 全年＝1～12 月（跟汇总、利润表一致）
    return (f'IF({comp_cell}="",IF({month_cell}="全年",{head}{y}),{head}{m})),'
            f'IF({month_cell}="全年",{head}{y}{c}),{head}{m}{c})))')


def _build_ledger(wb, ctx):
    ws = wb.create_sheet(SH_BSL)
    ws.sheet_properties.tabColor = C_BS[2:]
    widths(ws, {L_SEQ: 6, L_MON: 7, L_COMP: 13, L_NAME: 19, L_SPEC: 9, L_QTY: 12, L_UNIT: 6, L_PRICE: 12,
                L_AMT: 13, L_WHY: 14, L_NOTE: 14, L_CHK: 26})
    title(ws, '产品过期报损明细台账', L_CHK, C_BS,
          '💡 一年一张：每条报损记一行，「月份」填 1～12（几月的报损）。同一个月同一个商品可以只记一行，数量像原来一样写 =5+27+16 累加；'
          '金额自动＝数量×单价。只看某个月：点「月份」表头的筛选按钮选那个月，第 4 行的合计跟着变；'
          '或者在第 3 行黄格子选月份、公司，右边直接出合计。新记录往下面空行填，别在第一条记录上面插行。'
          '报损汇总一览、利润表、公司库存都从这张表取数。')
    ws['A1'] = '=年度&"年 产品过期报损明细台账"'
    # 第 3 行：查询
    ws.row_dimensions[3].height = 28
    put(ws, 'A3', '查询：', F_KPI_L, align=AR_, border=False)
    put(ws, 'B3', '全年', Font(name=YH, sz=11, bold=True, color='FF1F4E79'), FILL_SEL, '0"月";;;@', AC)
    add_list_dv(ws, 'B3', '"全年,1,2,3,4,5,6,7,8,9,10,11,12"', '选「全年」或某个月')
    put(ws, 'C3', None, Font(name=YH, sz=11, bold=True, color='FF1F4E79'), FILL_SEL, align=AC)
    add_list_dv(ws, 'C3', '=供应商列表', '选一家公司；空着＝全部公司')
    base = f'{L_NAME}${BSL_R0}:{L_NAME}${BSL_R1},"<>"'
    n = _crit(base, '$B$3', '$C$3', count=True)
    q = _crit(base, '$B$3', '$C$3', f'{L_QTY}${BSL_R0}:{L_QTY}${BSL_R1}')
    a = _crit(base, '$B$3', '$C$3', f'{L_AMT}${BSL_R0}:{L_AMT}${BSL_R1}')
    put(ws, 'D3', (f'=IF($B$3="全年","全年",$B$3&"月")&IF($C$3="","·全部公司","·"&$C$3)&"：共 "&{n}&" 行，报损数量 "'
                   f'&IF({q}=INT({q}),TEXT({q},"#,##0"),TEXT({q},"#,##0.##"))&"，报损金额 "&TEXT({a},"#,##0.00")&" 元"'),
        Font(name=YH, sz=11, bold=True, color='FFC00000'), align=AL, border=False)
    ws.merge_cells(f'D3:{L_CHK}3')
    # 第 4 行：筛选后合计
    ws.row_dimensions[4].height = 22
    put(ws, 'A4', '↓ 表头筛选后的合计（没筛选＝所有行）', F_TXTB, FILL_TOT, align=AR_)
    ws.merge_cells(f'A4:{L_SPEC}4')
    put(ws, f'{L_QTY}4', f'=SUBTOTAL(109,{L_QTY}{BSL_R0}:{L_QTY}{BSL_R1})', F_TXTB, FILL_TOT, QTY, AR_)
    for col in (L_UNIT, L_PRICE):
        put(ws, f'{col}4', None, F_TXTB, FILL_TOT)
    put(ws, f'{L_AMT}4', f'=SUBTOTAL(109,{L_AMT}{BSL_R0}:{L_AMT}{BSL_R1})', F_TXTB, FILL_TOT, MONEY, AR_)
    for col in (L_WHY, L_NOTE, L_CHK):
        put(ws, f'{col}4', None, F_TXTB, FILL_TOT)
    header(ws, BSL_HDR, [(L_SEQ, '序号'), (L_MON, '月份'), (L_COMP, '所属公司'), (L_NAME, '产品名称'), (L_SPEC, '规格型号'),
                         (L_QTY, '报损数量'), (L_UNIT, '单位'), (L_PRICE, '产品单价(元)'), (L_AMT, '报损金额(元)'),
                         (L_WHY, '报损原因'), (L_NOTE, '备注'), (L_CHK, '校验')], C_BS, height=30)

    rows = ctx['baosun_rows']
    for i in range(BSL_R1 - BSL_R0 + 1):
        r = BSL_R0 + i
        d = rows[i] if i < len(rows) else {}
        ws[f'{L_SEQ}{r}'] = f'=IF({L_NAME}{r}="","",ROW()-{BSL_R0 - 1})'
        for col, key in ((L_MON, 'month'), (L_COMP, 'company'), (L_NAME, 'name'), (L_SPEC, 'spec'), (L_QTY, 'qty'),
                         (L_UNIT, 'unit'), (L_PRICE, 'price'), (L_WHY, 'why'), (L_NOTE, 'note')):
            v = d.get(key)
            if v is not None:
                ws[f'{col}{r}'] = v
        ws[f'{L_AMT}{r}'] = f'=IF(AND(ISNUMBER({L_QTY}{r}),ISNUMBER({L_PRICE}{r})),{L_QTY}{r}*{L_PRICE}{r},"")'
        blank = f'AND({L_MON}{r}="",{L_COMP}{r}="",{L_NAME}{r}="",{L_QTY}{r}="",{L_PRICE}{r}="")'
        ws[f'{L_CHK}{r}'] = (
            f'=IF({blank},"",IF({L_NAME}{r}="","✗ 没填产品名称",IF({L_MON}{r}="","✗ 没填月份",'
            f'IF(OR(NOT(ISNUMBER({L_MON}{r})),N({L_MON}{r})<1,N({L_MON}{r})>12),"✗ 月份只填 1～12 的数字",'
            f'IF({L_QTY}{r}="","⚠ 没填数量",IF(NOT(ISNUMBER({L_QTY}{r})),"✗ 数量不是数字（单位写在「单位」列）",'
            f'IF({L_PRICE}{r}="","⚠ 没填单价，金额按 0 算",IF(NOT(ISNUMBER({L_PRICE}{r})),"✗ 单价不是数字",'
            f'IF({L_COMP}{r}="","⚠ 没填所属公司",'
            f'IF(COUNTIF({base_rng(G_BSNAME, BASE_R0, BASE_R1)},{L_NAME}{r})=0,"⚠ 商品档案「报损表品名」里没有这个名字","√"))))))))))')
    style_rows(ws, BSL_R0, BSL_R1, 'ABCDEFGHIJKL', auto=(L_SEQ, L_AMT, L_CHK),
               fmts={L_MON: '0"月"', L_QTY: QTY, L_PRICE: PRICE, L_AMT: MONEY},
               aligns={L_NAME: AL, L_WHY: AL, L_NOTE: AL, L_CHK: AL, L_QTY: AR_, L_PRICE: AR_, L_AMT: AR_})
    rng = f'{L_CHK}{BSL_R0}:{L_CHK}{BSL_R1}'
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(${L_CHK}{BSL_R0},1)="✗"'], fill=FILL_WARN,
                                                   font=Font(name=YH, sz=10, bold=True, color='FFC00000')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(${L_CHK}{BSL_R0},1)="⚠"'], fill=fill('FFFFEB9C')))
    # 有数量没单价：金额会变成 0，单价格子标红
    ws.conditional_formatting.add(f'{L_PRICE}{BSL_R0}:{L_PRICE}{BSL_R1}',
                                  FormulaRule(formula=[f'AND(${L_NAME}{BSL_R0}<>"",${L_QTY}{BSL_R0}<>"",${L_PRICE}{BSL_R0}="")'],
                                              fill=FILL_WARN))
    add_list_dv(ws, f'{L_MON}{BSL_R0}:{L_MON}{BSL_R1}', '"1,2,3,4,5,6,7,8,9,10,11,12"', '几月的报损（1～12）')
    add_list_dv(ws, f'{L_COMP}{BSL_R0}:{L_COMP}{BSL_R1}', '=供应商列表', '所属公司（基础资料·供应商）', stop=False)
    add_list_dv(ws, f'{L_NAME}{BSL_R0}:{L_NAME}{BSL_R1}', '=报损品名列表',
                '从下拉选（基础资料·商品档案「报损表品名」）；新品名先去那里加上，库存和成本才对得上', stop=False)
    ws.auto_filter.ref = f'{L_SEQ}{BSL_HDR}:{L_CHK}{BSL_R1}'
    ws.freeze_panes = f'A{BSL_R0}'
    return ws


# ───────────────────────── _辅助：报损中转站 ─────────────────────────
def _build_staging(wb, ctx):
    ws = wb[SH_AUX]
    base = f'{SH_BASE}!'
    put(ws, f'{BSX_NAME}3', '报损中转站：每个商品每月报损（按报损表品名匹配【报损明细台账】）', F_NOTE, border=False)
    for g, t in [(BSX_AMT0, '金额'), (BSX_RAW0, '原始数量'), (BSX_BOX0, '按箱/件录的数量'), (BSX_PCS0, '折成件')]:
        put(ws, f'{CL(g)}3', t, F_NOTE, border=False)
        for m in range(1, 13):
            ws[f'{CL(g + m - 1)}4'] = m
    ws[f'{BSX_NAME}5'] = '报损品名'
    ws[f'{BSX_PACK}5'] = '每件数量'
    ws[f'{BSX_SEL}5'] = '所选月份金额'
    ws[f'{BSX_KEY}5'] = '排名键'
    ws[f'{BSX_SELQ}5'] = '所选月份数量'

    def safe(expr):                 # 表里有怪值 → 当 0，由【数据校验】报出来，不让整本报表跟着报错
        return f'IFERROR({expr},0)'

    Q, N, C, U, H, A, M = (bsl(c) for c in (L_QTY, L_NAME, L_COMP, L_UNIT, L_PRICE, L_AMT, L_MON))
    sel = f'{SH_BSS}!{BSS_SEL}'
    mask = f'((${bsx_amt(1)}$4:${bsx_amt(12)}$4={sel})+({sel}="全年"))'
    for i in range(BASE_R1 - BASE_R0 + 1):
        r = AUX_BS_P0 + i
        nm, pk = at(base_rng(G_BSNAME, BASE_R0, BASE_R1), i + 1), at(base_rng(G_PACK, BASE_R0, BASE_R1), i + 1)
        ws[f'{BSX_NAME}{r}'] = f'=IF({nm}="","",{nm}&"")'
        ws[f'{BSX_PACK}{r}'] = f'=IF(N({pk})<=0,1,{pk})'
        for m in range(1, 13):
            a, raw, box, pcs = (CL(BSX_AMT0 + m - 1), CL(BSX_RAW0 + m - 1), CL(BSX_BOX0 + m - 1), CL(BSX_PCS0 + m - 1))
            skip = f'${BSX_NAME}{r}=""'
            ws[f'{a}{r}'] = f'=IF({skip},0,' + safe(f'SUMIFS({A},{N},${BSX_NAME}{r},{M},{a}$4)') + ')'
            ws[f'{raw}{r}'] = f'=IF({skip},0,' + safe(f'SUMIFS({Q},{N},${BSX_NAME}{r},{M},{raw}$4)') + ')'
            ws[f'{box}{r}'] = (f'=IF({skip},0,' + safe(f'SUMIFS({Q},{N},${BSX_NAME}{r},{M},{box}$4,{U},"箱")'
                               f'+SUMIFS({Q},{N},${BSX_NAME}{r},{M},{box}$4,{U},"件")') + ')')
            ws[f'{pcs}{r}'] = f'=({raw}{r}-{box}{r})/${BSX_PACK}{r}+{box}{r}'
        ws[f'{BSX_SEL}{r}'] = f'=SUMPRODUCT({mask}*{bsx_amt(1)}{r}:{bsx_amt(12)}{r})'
        ws[f'{BSX_SELQ}{r}'] = f'=SUMPRODUCT({mask}*{CL(BSX_RAW0)}{r}:{CL(BSX_RAW0 + 11)}{r})'
        ws[f'{BSX_KEY}{r}'] = f'=IF({BSX_SEL}{r}>0,{BSX_SEL}{r}+(1000-ROW())/10000000,"")'
    # 供应商块
    put(ws, f'{BSX_NAME}{AUX_BS_S0 - 1}', '报损中转站：每家公司每月报损金额', F_NOTE, border=False)
    for i in range(SUP_R1 - SUP_R0 + 1):
        r = AUX_BS_S0 + i
        nm = at(SUP_NAMES, i + 1)
        ws[f'{BSX_NAME}{r}'] = f'=IF({nm}="","",{nm}&"")'
        for m in range(1, 13):
            a = bsx_amt(m)
            ws[f'{a}{r}'] = f'=IF(${BSX_NAME}{r}="",0,' + safe(f'SUMIFS({A},{C},${BSX_NAME}{r},{M},{a}$4,{N},"<>")') + ')'
        ws[f'{BSX_SEL}{r}'] = f'=SUMPRODUCT({mask}*{bsx_amt(1)}{r}:{bsx_amt(12)}{r})'
    # 每月总计 / 没对上商品档案的 / 各种问题行数
    for rr, t in ((AUX_BS_TOT, '每月报损总金额'), (AUX_BS_TOTQ, '每月报损总数量'), (AUX_BS_UNM, '没对上商品档案的金额'),
                  (AUX_BS_NOPRICE, '有数量没单价的行数'), (AUX_BS_NOCOMP, '有产品没填所属公司的行数'),
                  (AUX_BS_NONNUM, '数量/单价不是数字的格数'), (AUX_BS_NOH, '有数量单价没金额公式的行数')):
        ws[f'{BSX_NAME}{rr}'] = t
    for m in range(1, 13):
        a = bsx_amt(m)
        mm = f'{M},{a}$4'
        ws[f'{a}{AUX_BS_TOT}'] = '=' + safe(f'SUMIFS({A},{N},"<>",{mm})')
        ws[f'{a}{AUX_BS_TOTQ}'] = '=' + safe(f'SUMIFS({Q},{N},"<>",{mm})')
        ws[f'{a}{AUX_BS_UNM}'] = f'={a}{AUX_BS_TOT}-SUM({a}{AUX_BS_P0}:{a}{AUX_BS_P1})'
        ws[f'{a}{AUX_BS_NOPRICE}'] = '=' + safe(f'COUNTIFS({N},"<>",{Q},"<>",{H},"",{mm})')
        ws[f'{a}{AUX_BS_NOCOMP}'] = '=' + safe(f'COUNTIFS({N},"<>",{C},"",{mm})')
        ws[f'{a}{AUX_BS_NONNUM}'] = '=' + safe(f'SUMPRODUCT(ISTEXT({Q})*({M}={a}$4))+SUMPRODUCT(ISTEXT({H})*({M}={a}$4))')
        # 数量、单价都是数字，金额却是空：只有金额格没公式（插行插出来的）才会这样；数量/单价是文字的另有一项报
        ws[f'{a}{AUX_BS_NOH}'] = '=' + safe(f'SUMPRODUCT(ISNUMBER({Q})*ISNUMBER({H})*({A}="")*({M}={a}$4))')
    for rr in (AUX_BS_TOT, AUX_BS_TOTQ, AUX_BS_UNM):
        ws[f'{BSX_SEL}{rr}'] = f'=SUMPRODUCT({mask}*{bsx_amt(1)}{rr}:{bsx_amt(12)}{rr})'


# ───────────────────────── 报损汇总一览 ─────────────────────────
def _build_summary(wb, ctx):
    ws = wb.create_sheet(SH_BSS)
    ws.sheet_properties.tabColor = C_BS[2:]
    aux = f'{SH_AUX}!'
    widths(ws, {'A': 11, 'B': 24, 'C': 16, 'D': 16, 'E': 12, 'F': 6})
    title(ws, '产品过期报损汇总一览', 'E', C_BS,
          '💡 全自动：左上角黄格子选「全年」或某个月，下面按公司、按品项（金额从大到小）汇总【报损明细台账】。')
    header(ws, 3, [('A', '统计月份'), ('B', '总报损金额(元)'), ('C', '总报损数量'), ('D', '报损品项数'), ('E', '涉及公司数')],
           C_BS, height=26)
    ws.row_dimensions[4].height = 26
    put(ws, 'A4', '全年', Font(name=YH, sz=12, bold=True, color='FF1F4E79'), FILL_SEL, '0"月";;;@', AC)
    add_list_dv(ws, 'A4', '"全年,1,2,3,4,5,6,7,8,9,10,11,12"', '选「全年」或某个月')
    kf = Font(name=YH, sz=12, bold=True, color='FFC00000')
    put(ws, 'B4', f'={aux}{BSX_SEL}{AUX_BS_TOT}', kf, fmt=MONEY, align=AC)
    put(ws, 'C4', f'={aux}{BSX_SEL}{AUX_BS_TOTQ}', kf, fmt=QTY, align=AC)
    put(ws, 'D4', f'=COUNTIF({aux}{BSX_SELQ}{AUX_BS_P0}:{BSX_SELQ}{AUX_BS_P1},">0")', kf, fmt='0', align=AC)   # 有数量就算（没单价的也算）
    put(ws, 'E4', f'=COUNTIF({aux}{BSX_SEL}{AUX_BS_S0}:{BSX_SEL}{AUX_BS_S1},">0")+IF(ROUND(C{BSS_C0 + 5},2)>0,1,0)', kf,
        fmt='0', align=AC)

    # 分公司
    put(ws, f'A{BSS_C0 - 2}', '分公司报损汇总（按所选月份）', F_SEC, border=False)
    header(ws, BSS_C0 - 1, [('A', '序号'), ('B', '公司名称'), ('C', '报损总金额(元)'), ('D', '金额占比')], C_BS, height=24)
    ct = f'$C${BSS_CTOT}'
    for i in range(5):
        r = BSS_C0 + i
        ws[f'A{r}'] = i + 1
        ws[f'B{r}'] = f'=IF({at(SUP_NAMES, i + 1)}="","",{at(SUP_NAMES, i + 1)})'
        ws[f'C{r}'] = f'=IF(B{r}="","",{aux}{BSX_SEL}{AUX_BS_S0 + i})'
        ws[f'D{r}'] = f'=IF(OR(B{r}="",N({ct})=0),"",C{r}/{ct})'
    r = BSS_C0 + 5
    ws[f'A{r}'] = 6
    ws[f'B{r}'] = '其他/未登记公司'
    ws[f'C{r}'] = f'={BSS_TOT}-SUM(C{BSS_C0}:C{BSS_C0 + 4})'
    ws[f'D{r}'] = f'=IF(N({ct})=0,"",C{r}/{ct})'
    style_rows(ws, BSS_C0, BSS_C0 + 5, 'ABCD', fmts={'C': MONEY, 'D': '0.0%'}, aligns={'C': AR_, 'D': AR_})
    put(ws, f'A{BSS_CTOT}', '合计', F_TXTB, FILL_TOT, align=AC)
    put(ws, f'B{BSS_CTOT}', None, F_TXTB, FILL_TOT)
    put(ws, f'C{BSS_CTOT}', f'=SUM(C{BSS_C0}:C{BSS_C0 + 5})', F_TXTB, FILL_TOT, MONEY, AR_)
    put(ws, f'D{BSS_CTOT}', f'=IF(N(C{BSS_CTOT})=0,"",SUM(D{BSS_C0}:D{BSS_C0 + 5}))', F_TXTB, FILL_TOT, '0.0%', AR_)

    # 分品项
    put(ws, f'A{BSS_P0 - 2}', '分品项报损汇总（按金额从大到小 · 所选月份）', F_SEC, border=False)
    header(ws, BSS_P0 - 1, [('A', '排名'), ('B', '产品名称'), ('C', '所属公司'), ('D', '报损总金额(元)'), ('E', '金额占比')],
           C_BS, height=24)
    keys = f'{aux}${BSX_KEY}${AUX_BS_P0}:${BSX_KEY}${AUX_BS_P1}'
    pt = f'$D${BSS_PTOT}'
    for k in range(1, BSS_PN + 1):
        r = BSS_P0 + k - 1
        ws[f'A{r}'] = k
        ws[f'F{r}'] = f'=IFERROR(MATCH(LARGE({keys},{k}),{keys},0),"")'
        ws[f'B{r}'] = f'=IF($F{r}="","",INDEX({aux}${BSX_NAME}${AUX_BS_P0}:${BSX_NAME}${AUX_BS_P1},$F{r}))'
        ws[f'C{r}'] = f'=IF($F{r}="","",INDEX({SH_BASE}!${G_SUP}${BASE_R0}:${G_SUP}${BASE_R1},$F{r})&"")'
        ws[f'D{r}'] = f'=IF($F{r}="","",INDEX({aux}${BSX_SEL}${AUX_BS_P0}:${BSX_SEL}${AUX_BS_P1},$F{r}))'
        ws[f'E{r}'] = f'=IF(OR($F{r}="",N({pt})=0),"",D{r}/{pt})'
    r = BSS_P0 + BSS_PN
    ws[f'B{r}'] = '其余品项 / 没对上商品档案的'
    ws[f'D{r}'] = f'={pt}-SUM(D{BSS_P0}:D{r - 1})'
    ws[f'E{r}'] = f'=IF(N({pt})=0,"",D{r}/{pt})'
    style_rows(ws, BSS_P0, r, 'ABCDE', fmts={'D': MONEY, 'E': '0.0%'}, aligns={'B': AL, 'D': AR_, 'E': AR_})
    for rr in range(BSS_P0, r + 1):
        ws[f'F{rr}'].font = F_NOTE
    put(ws, f'A{BSS_PTOT}', '合计', F_TXTB, FILL_TOT, align=AC)
    for col in 'BC':
        put(ws, f'{col}{BSS_PTOT}', None, F_TXTB, FILL_TOT)
    put(ws, f'D{BSS_PTOT}', f'={BSS_TOT}', F_TXTB, FILL_TOT, MONEY, AR_)
    put(ws, f'E{BSS_PTOT}', f'=IF(N(D{BSS_PTOT})=0,"",1)', F_TXTB, FILL_TOT, '0%', AR_)
    ws.column_dimensions['F'].hidden = True
    ws.freeze_panes = 'A5'
    return ws
