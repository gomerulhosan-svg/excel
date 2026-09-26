# -*- coding: utf-8 -*-
"""【记账分录】（所有凭证的分录总表）→【_凭证索引】（每张凭证一行：有没有、日期、几条、几页、当月第几号）
→【凭证汇总】（报表月的凭证清单）→【记账凭证】（打印版，A5 横向，一页 8 条分录）。

记账分录按「来源 × 源行位置 × 第几条腿」的固定槽位排：现金流水每行 2 条（本账户一条、对方科目一条），
采购入库 3 条（存货、进项税、应付），领用 2 条，销售 3 条（应收、收入、销项税），工资 2 条，手工凭证 1 条；
月末结转每月一组。没有金额的槽位空着（「有金额」=0），不进凭证、不进账。"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.pagebreak import Break, RowBreak

from common import *
from s_mend import (ME_SUM, ME_DEP_HDR, ME_DEP0, DEP_COL, N_DEPT, fx_row, cc_row, ME_X0, FX_COLS, CC_COLS)

ME = q(SH_MEND)


def _dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# ─────────────────────────── 各来源的分录写法 ───────────────────────────
def _x(rg, r):
    """按本行 K 列（源位置）到来源表取数"""
    return f'INDEX({rg},${J_POS}{r})'


def _copy_up(ws, r, r1, cols):
    for c in cols:
        ws[f'{c}{r}'] = f'=${c}{r1}'


def w_cash(ws, r1, k):
    r2 = r1 + 1
    X = lambda col, r=r1: _x(cash(col), r)
    C = f'${J_VID}{r1}'
    ws[f'{J_VID}{r1}'] = f'=IFERROR({X(K_VID)},"")'
    ws[f'{J_DATE}{r1}'] = f'=IF({C}="","",{X(K_DATE)})'
    ws[f'{J_MEMO}{r1}'] = f'=IF({C}="","",{X(K_VMEMO)})'
    ws[f'{J_CODE}{r1}'] = f'=IF({C}="","",IF({X(K_DR)}>0,{X(K_ACODE)},{X(K_CCODE)}))'
    ws[f'{J_DR}{r1}'] = f'=IF({C}="",0,{X(K_DR)}+{X(K_CR)})'
    ws[f'{J_CR}{r1}'] = 0
    ws[f'{J_CP}{r1}'] = f'=IF({C}="","",{X(K_CP)}&"")'
    ws[f'{J_PRJ}{r1}'] = f'=IF({C}="","",{X(K_PRJ)}&"")'
    ws[f'{J_START}{r1}'] = f'=IF({C}="",0,(IF({X(K_GSEQ)}="",1,{X(K_GSEQ)})-1)*2)'
    _copy_up(ws, r2, r1, [J_VID, J_DATE, J_MEMO, J_CP, J_PRJ])
    ws[f'{J_CODE}{r2}'] = f'=IF({C}="","",IF({X(K_DR, r2)}>0,{X(K_CCODE, r2)},{X(K_ACODE, r2)}))'
    ws[f'{J_DR}{r2}'] = 0
    ws[f'{J_CR}{r2}'] = f'=${J_DR}{r1}'


def _w_mod(ws, r1, sh, vid, date, memo, legs, cp, prj, pre):
    """录入模块通用：legs = [(科目公式, 借方公式, 贷方公式)]，公式里用 {X} 表示按源位置取数"""
    X = lambda col, r=r1: _x(mod(sh, col), r)
    C = f'${J_VID}{r1}'
    ws[f'{J_VID}{r1}'] = f'=IFERROR({X(vid)},"")'
    ws[f'{J_DATE}{r1}'] = f'=IF({C}="","",{X(date)})'
    ws[f'{J_MEMO}{r1}'] = f'=IF({C}="","",{memo(X)})'
    ws[f'{J_CP}{r1}'] = f'=IF({C}="","",{X(cp)}&"")' if cp else '=""'
    ws[f'{J_PRJ}{r1}'] = f'=IF({C}="","",{X(prj)}&"")' if prj else '=""'
    ws[f'{J_START}{r1}'] = f'=IF({C}="",0,N({X(pre)}))'
    for j, (code, dr, crd) in enumerate(legs):
        r = r1 + j
        if j:
            _copy_up(ws, r, r1, [J_VID, J_DATE, J_MEMO, J_CP, J_PRJ])
        ws[f'{J_CODE}{r}'] = f'=IF({C}="","",{code(X)})'
        ws[f'{J_DR}{r}'] = f'=IF({C}="",0,{dr(X)})' if dr else 0
        ws[f'{J_CR}{r}'] = f'=IF({C}="",0,{crd(X)})' if crd else 0


def w_buy(ws, r1, k):
    _w_mod(ws, r1, SH_BUY, B_VID, B_DATE,
           lambda X: (f'"采购入库 "&{X(B_MAT)}&IF({X(B_QTY)}=""," 运杂关税"," "&{X(B_QTY)}&{X(B_UNIT)})'
                      f'&IF({X(B_SUP)}="",""," · "&{X(B_SUP)})'),
           [(lambda X: f'{X(B_ACC)}&""', lambda X: f'N({X(B_COST)})', None),
            (lambda X: '进项税科目&""', lambda X: f'N({X(B_VAT)})', None),
            (lambda X: '应付科目&""', None, lambda X: f'N({X(B_TOT)})'),
            (lambda X: '运杂税费科目&""', None, lambda X: f'N({X(B_FRT)})')],
           B_SUP, B_PRJ, B_GPRE)
    ws[f'{J_CP}{r1 + 3}'] = '=""'        # 运杂关税不是欠供应商的，不挂往来单位


def w_iss(ws, r1, k):
    _w_mod(ws, r1, SH_ISS, I_VID, I_DATE,
           lambda X: f'"领用 "&{X(I_MAT)}&" "&{X(I_QTY)}&{X(I_UNIT)}&" · "&{X(I_USE)}',
           [(lambda X: f'{X(I_DR)}&""', lambda X: f'N({X(I_AMT)})', None),
            (lambda X: f'{X(I_CR)}&""', None, lambda X: f'N({X(I_AMT)})')],
           None, I_PRJ, I_GPRE)


def w_sale(ws, r1, k):
    _w_mod(ws, r1, SH_SALE, S_VID, S_DATE,
           lambda X: (f'"销售"&IF({X(S_ITEM)}="","原矿",{X(S_ITEM)})&" "&FIXED({X(S_TON)},2)&"吨"'
                      f'&IF({X(S_CUS)}="",""," · "&{X(S_CUS)})'),
           [(lambda X: '应收科目&""', lambda X: f'N({X(S_TOT)})', None),
            (lambda X: '收入科目&""', None, lambda X: f'N({X(S_NET)})'),
            (lambda X: '销项税科目&""', None, lambda X: f'N({X(S_TAX)})')],
           S_CUS, S_PRJ, S_GPRE)


def w_pay(ws, r1, k):
    _w_mod(ws, r1, SH_PAY, W_VID, W_DATE,
           lambda X: (f'"计提"&MONTH({X(W_DATE)})&"月工资 · "&{X(W_DEPT)}&IF({X(W_TYPE)}="",""," · "&{X(W_TYPE)})'),
           [(lambda X: f'{X(W_DR)}&""', lambda X: f'N({X(W_AMT)})', None),
            (lambda X: f'{X(W_CR)}&""', None, lambda X: f'N({X(W_AMT)})')],
           None, None, W_GPRE)


def w_man(ws, r1, k):
    X = lambda col: _x(mod(SH_MAN, col), r1)
    C = f'${J_VID}{r1}'
    G = mod(SH_MAN, MN_GRP)
    ws[f'{J_VID}{r1}'] = f'=IFERROR({X(MN_VID)},"")'
    ws[f'{J_DATE}{r1}'] = f'=IF({C}="","",{X(MN_VDATE)})'
    ws[f'{J_MEMO}{r1}'] = (f'=IF({C}="","",IF({X(MN_MEMO)}<>"",{X(MN_MEMO)}&"",'
                           f'IFERROR(INDEX({mod(SH_MAN, MN_MEMO)},MATCH({X(MN_GKEY)},{mod(SH_MAN, MN_GKEY)},0))&"","")))')
    ws[f'{J_CODE}{r1}'] = f'=IF({C}="","",{X(MN_CODE2)})'
    ws[f'{J_DR}{r1}'] = f'=IF({C}="",0,N({X(MN_DR)}))'
    ws[f'{J_CR}{r1}'] = f'=IF({C}="",0,N({X(MN_CR)}))'
    ws[f'{J_CP}{r1}'] = f'=IF({C}="","",{X(MN_CP)}&"")'
    ws[f'{J_PRJ}{r1}'] = f'=IF({C}="","",{X(MN_PRJ)}&"")'
    ws[f'{J_START}{r1}'] = f'=IF({C}="",0,N({X(MN_SEQ)})-1)'


def _w_me(ws, r1, n, src, legs, memo):
    """月末结转：legs = [(科目公式, 借方公式, 贷方公式)]；没有金额的腿凭证ID 留空"""
    vid = SRC[src] * 100000 + n
    mrow = ME_M0 + n - 1
    for j, (code, dr, crd) in enumerate(legs):
        r = r1 + j
        ws[f'{J_DR}{r}'] = f'={dr}' if dr else 0
        ws[f'{J_CR}{r}'] = f'={crd}' if crd else 0
        C = f'${J_VID}{r}'
        ws[f'{J_VID}{r}'] = f'=IF(AND(ROUND(${J_DR}{r},2)=0,ROUND(${J_CR}{r},2)=0),"",{vid})'
        ws[f'{J_DATE}{r}'] = f'=IF({C}="","",{ME}!$B${mrow})'
        ws[f'{J_MEMO}{r}'] = f'=IF({C}="","",{memo(f"${J_DATE}{r}")})'
        ws[f'{J_CODE}{r}'] = f'=IF({C}="","",{code}&"")'
        ws[f'{J_CP}{r}'] = '=""'
        ws[f'{J_PRJ}{r}'] = '=""'
    ws[f'{J_START}{r1}'] = 0


ME_MEMO = {
    '折旧': lambda d: f'"计提"&MONTH({d})&"月固定资产折旧"',
    '调汇': lambda d: f'MONTH({d})&"月末外币账户按月末汇率调汇"',
    '成本': lambda d: f'"结转"&MONTH({d})&"月生产成本（转入原矿）"',
    '销售成本': lambda d: f'"结转"&MONTH({d})&"月原矿销售成本"',
    '兑换': lambda d: f'MONTH({d})&"月外币兑换差额转汇兑损益"',
}


def w_dep(ws, r1, n):
    mrow, drow = ME_M0 + n - 1, ME_DEP0 + n - 1
    end = f'{ME}!$B${mrow}'
    hdr = lambda j: f'{ME}!{DEP_COL(j)}${ME_DEP_HDR + 1}'
    # 启用日以前的月份：记生产成本的部门改记「启用日前成本科目」（跟那时候的流水口径一样，不堆进在产品）
    code = lambda j: (f'IF(AND({end}<启用日,OR(LEFT({hdr(j)},4)="4001",LEFT({hdr(j)},4)="4101")),启用日前成本科目,{hdr(j)})')
    legs = [(code(j), f'{ME}!{DEP_COL(j)}{drow}', None) for j in range(1, N_DEPT + 1)]
    legs.append(('累计折旧科目', None, f'{ME}!{ME_SUM["dep"]}{mrow}'))
    _w_me(ws, r1, n, '折旧', legs, ME_MEMO['折旧'])


def w_fx(ws, r1, n):
    mrow = ME_M0 + n - 1
    legs = []
    for i in range(1, N_ACC + 1):
        J = f'{ME}!{FX_COLS["diff"]}{fx_row(i, n)}'
        legs.append((f'{ME}!{FX_COLS["code"]}{fx_row(i, n)}', f'MAX({J},0)', f'MAX(-{J},0)'))
    net = f'{ME}!{ME_SUM["fx"]}{mrow}'
    legs.append(('汇兑损益科目', f'MAX(-{net},0)', f'MAX({net},0)'))
    _w_me(ws, r1, n, '调汇', legs, ME_MEMO['调汇'])


def w_cc(ws, r1, n, nleaf):
    mrow = ME_M0 + n - 1
    tot = f'{ME}!{ME_SUM["cost"]}{mrow}'
    legs = [('原矿科目', f'MAX({tot},0)', f'MAX(-{tot},0)')]
    for j in range(1, nleaf + 1):
        E = f'{ME}!{CC_COLS["carry"]}{cc_row(j, n)}'
        legs.append((f'{ME}!{CC_COLS["code"]}{cc_row(j, n)}', f'MAX(-{E},0)', f'MAX({E},0)'))
    _w_me(ws, r1, n, '成本', legs, ME_MEMO['成本'])


def w_cogs(ws, r1, n):
    a = f'{ME}!{ME_SUM["cogs"]}{ME_M0 + n - 1}'
    _w_me(ws, r1, n, '销售成本', [('销售成本科目', f'MAX({a},0)', f'MAX(-{a},0)'),
                                   ('原矿科目', f'MAX(-{a},0)', f'MAX({a},0)')], ME_MEMO['销售成本'])


def w_xchg(ws, r1, n):
    a = f'{ME}!D{ME_X0 + n - 1}'
    _w_me(ws, r1, n, '兑换', [('汇兑损益科目', f'MAX({a},0)', f'MAX(-{a},0)'),
                               (f'"{EXCH_CODE}"', f'MAX(-{a},0)', f'MAX({a},0)')], ME_MEMO['兑换'])


def blocks(ctx):
    nleaf = len(ctx['cost_leaves'])
    return [
        dict(src='流水', n=CASH_R1 - CASH_R0 + 1, legs=2, w=w_cash, slots=CASH_R1 - CASH_R0 + 1),
        dict(src='采购', n=BUY_R1 - MOD_R0 + 1, legs=4, w=w_buy, slots=BUY_R1 - MOD_R0 + 1),
        dict(src='领用', n=ISS_R1 - MOD_R0 + 1, legs=2, w=w_iss, slots=ISS_R1 - MOD_R0 + 1),
        dict(src='销售', n=SALE_R1 - MOD_R0 + 1, legs=3, w=w_sale, slots=SALE_R1 - MOD_R0 + 1),
        dict(src='工资', n=PAY_R1 - MOD_R0 + 1, legs=2, w=w_pay, slots=PAY_R1 - MOD_R0 + 1),
        dict(src='手工', n=MAN_R1 - MOD_R0 + 1, legs=1, w=w_man, slots=MAN_R1 - MOD_R0 + 1),
        dict(src='折旧', n=N_MONTHS, legs=N_DEPT + 1, w=w_dep, slots=N_MONTHS),
        dict(src='调汇', n=N_MONTHS, legs=N_ACC + 1, w=w_fx, slots=N_MONTHS),
        dict(src='成本', n=N_MONTHS, legs=nleaf + 1, w=lambda ws, r, n: w_cc(ws, r, n, nleaf), slots=N_MONTHS),
        dict(src='销售成本', n=N_MONTHS, legs=2, w=w_cogs, slots=N_MONTHS),
        dict(src='兑换', n=N_MONTHS, legs=2, w=w_xchg, slots=N_MONTHS),
    ]


def layout(ctx):
    """算出每个来源在【记账分录】【_凭证索引】里的起始行，存进 ctx"""
    bl = {}
    je, vx = JE_R0, VX_R0
    for b in blocks(ctx):
        b['je0'], b['vx0'] = je, vx
        je += b['n'] * b['legs']
        vx += b['slots']
        bl[b['src']] = b
    ctx['je_r1'], ctx['vx_r1'] = je - 1, vx - 1
    ctx['blocks'] = bl
    return bl


def je_rng(ctx, col):
    return rng(SH_JE, col, JE_R0, ctx['je_r1'])


def vx_rng(ctx, col):
    return rng(SH_VIDX, col, VX_R0, ctx['vx_r1'])


# ─────────────────────────── 记账分录 ───────────────────────────
def build_je(wb, ctx):
    bl = layout(ctx)
    ws = wb.create_sheet(SH_JE)
    R1 = ctx['je_r1']
    widths(ws, {J_DATE: 11, J_VNO: 7, J_MEMO: 40, J_CODE: 10, J_NAME: 34, J_DR: 17, J_CR: 17, J_CP: 12, J_PRJ: 10,
                J_SRC: 8, J_POS: 6, J_VID: 9, J_SEQ: 5, J_KEY: 12, J_GLK: 14, J_NZ: 5, J_START: 5})
    title(ws, '记账分录（所有凭证的分录，全自动）', J_SRC, C_JE,
          '💡 不用填（表锁着，防止排序把公式打乱）。现金流水、采购入库、领用出库、销售结算、工资计提、手工凭证、月末结转，每一笔自动变成这里的分录；'
          '科目余额表、明细账、报表、打印的凭证都从这里取。想看某个月、某个科目：用表头的筛选（只能筛选、不能排序；借方、贷方都是「-」的是空槽位，筛掉就行）。'
          '按日期顺序看请用【明细账】。凭证号＝当月第几号。')
    header(ws, JE_HDR, [(J_DATE, '日期'), (J_VNO, '凭证号'), (J_MEMO, '摘要'), (J_CODE, '科目编码'), (J_NAME, '科目全称'),
                        (J_DR, '借方金额'), (J_CR, '贷方金额'), (J_CP, '往来单位'), (J_PRJ, '项目'), (J_SRC, '来源'),
                        (J_POS, '源位置'), (J_VID, '凭证ID'), (J_SEQ, '条'), (J_KEY, '查找键'), (J_GLK, '明细账键'),
                        (J_NZ, '有额'), (J_START, '前条')], C_JE, height=30)
    codes, full = acc(AC_CODE), acc(AC_FULL)
    vno = vx_rng(ctx, VX_NO)
    for src, b in bl.items():
        off = b['vx0'] - VX_R0            # 这个来源在凭证索引里前面有几行
        for k in range(1, b['n'] + 1):
            r1 = b['je0'] + (k - 1) * b['legs']
            for j in range(b['legs']):
                r = r1 + j
                ws[f'{J_SRC}{r}'] = src
                ws[f'{J_POS}{r}'] = k
            b['w'](ws, r1, k)
            for j in range(b['legs']):
                r = r1 + j
                ws[f'{J_NZ}{r}'] = (f'=IF(AND(${J_VID}{r}<>"",OR(ROUND(${J_DR}{r},2)<>0,ROUND(${J_CR}{r},2)<>0)),1,0)')
                ws[f'{J_SEQ}{r}'] = f'=IF(${J_NZ}{r}=0,"",${J_START}${r1}+SUM(${J_NZ}${r1}:${J_NZ}{r}))'
                ws[f'{J_KEY}{r}'] = f'=IF(${J_SEQ}{r}="","",${J_VID}{r}*1000+${J_SEQ}{r})'
                ws[f'{J_NAME}{r}'] = (f'=IF(${J_NZ}{r}=0,"",IFERROR(INDEX({full},MATCH(${J_CODE}{r},{codes},0)),'
                                      f'"（科目表里没有）"))')
                ws[f'{J_VNO}{r}'] = (f'=IF(${J_NZ}{r}=0,"",IFERROR(INDEX({vno},${J_VID}{r}-{SRC[src] * 100000}+{off}),""))')
                ws[f'{J_GLK}{r}'] = (f'=IF(${J_NZ}{r}=0,"",IF(AND(明细账科目<>"",LEFT(${J_CODE}{r},LEN(明细账科目))=明细账科目,'
                                     f'${J_DATE}{r}>=明细账起,${J_DATE}{r}<=明细账止),'
                                     f'INT(${J_DATE}{r})*10^7+N(${J_VNO}{r})*1000+${J_SEQ}{r},""))')
    vis = [J_DATE, J_VNO, J_MEMO, J_CODE, J_NAME, J_DR, J_CR, J_CP, J_PRJ, J_SRC]
    hid = [J_POS, J_VID, J_SEQ, J_KEY, J_GLK, J_NZ, J_START]
    style_rows(ws, JE_R0, R1, vis, auto=tuple(vis),
               fmts={J_DATE: DATE, J_DR: MONEY, J_CR: MONEY, J_VNO: '0'},
               aligns={J_MEMO: AL, J_NAME: AL, J_DR: AR_, J_CR: AR_, J_CODE: AL})
    for r in range(JE_R0, R1 + 1):
        for c in hid:
            ws[f'{c}{r}'].font = F_NOTE
    hide(ws, hid)
    ws.auto_filter.ref = f'A{JE_HDR}:{J_SRC}{R1}'
    ws.freeze_panes = f'C{JE_R0}'
    lock_formula_sheet(ws)
    for nm, col in (('分录日期', J_DATE), ('分录凭证号', J_VNO), ('分录科目', J_CODE), ('分录借方', J_DR), ('分录贷方', J_CR),
                    ('分录往来', J_CP), ('分录来源', J_SRC), ('分录凭证ID', J_VID), ('分录键', J_KEY), ('分录明细键', J_GLK),
                    ('分录有额', J_NZ), ('分录摘要', J_MEMO), ('分录全称', J_NAME)):
        _dn(wb, nm, je_rng(ctx, col))
    return ws


# ─────────────────────────── _凭证索引 ───────────────────────────
def build_vindex(wb, ctx):
    bl = ctx['blocks']
    ws = wb.create_sheet(SH_VIDX)
    R1 = ctx['vx_r1']
    put(ws, 'A1', '凭证索引（每张可能出现的凭证一行：有没有、日期、几条分录、几页、当月第几号）—— 自动，不用动', F_NOTE, border=False)
    for c, t in ((VX_ID, '凭证ID'), (VX_SRC, '来源'), (VX_ON, '有'), (VX_DATE, '日期'), (VX_KEY, '排序键'), (VX_LINES, '条数'),
                 (VX_PAGES, '页数'), (VX_MKEY, '月份'), (VX_NO, '凭证号'), (VX_MEMO, '摘要'), (VX_ATT, '附单据'),
                 (VX_FIND, '查找键')):
        put(ws, f'{c}3', t, F_NOTE, border=False)
    je_memo = je_rng(ctx, J_MEMO)
    je_nz = lambda r0, r1: rng(SH_JE, J_NZ, r0, r1)
    MK = rng(SH_VIDX, VX_MKEY, VX_R0, R1)
    KY = rng(SH_VIDX, VX_KEY, VX_R0, R1)
    modcols = {'采购': (SH_BUY, B_VID, B_DATE, B_LEGS), '领用': (SH_ISS, I_VID, I_DATE, I_LEGS),
               '销售': (SH_SALE, S_VID, S_DATE, S_LEGS), '工资': (SH_PAY, W_VID, W_DATE, W_LEGS)}
    for src, b in bl.items():
        for k in range(1, b['slots'] + 1):
            r = b['vx0'] + k - 1
            vid = SRC[src] * 100000 + k
            A, Cc, D = f'${VX_ID}{r}', f'${VX_ON}{r}', f'${VX_DATE}{r}'
            ws[f'{VX_ID}{r}'] = vid
            ws[f'{VX_SRC}{r}'] = src
            att = '""'
            if src == '流水':
                ws[f'{VX_ON}{r}'] = f'=IFERROR(IF(INDEX({cash(K_VID)},{k})={A},1,0),0)'
                ws[f'{VX_DATE}{r}'] = f'=IF({Cc}=0,"",INDEX({cash(K_DATE)},{k}))'
                ws[f'{VX_LINES}{r}'] = (f'=IF({Cc}=0,0,IF(INDEX({cash(K_MERGE)},{k})="",2,2*COUNTIF({cash(K_VID)},{A})))')
                ws[f'{VX_MEMO}{r}'] = f'=IF({Cc}=0,"",INDEX({cash(K_VMEMO)},{k}))'
                att = (f'IF(INDEX({cash(K_MERGE)},{k})="",N(INDEX({cash(K_ATT)},{k})),SUMIF({cash(K_VID)},{A},{cash(K_ATT)}))')
            elif src in modcols:
                sh, cv, cd, cl = modcols[src]
                ws[f'{VX_ON}{r}'] = f'=IFERROR(IF(INDEX({mod(sh, cv)},{k})={A},1,0),0)'
                ws[f'{VX_DATE}{r}'] = f'=IF({Cc}=0,"",INDEX({mod(sh, cd)},{k}))'
                ws[f'{VX_LINES}{r}'] = f'=IF({Cc}=0,0,SUMIF({mod(sh, cv)},{A},{mod(sh, cl)}))'
                jr = b['je0'] - JE_R0 + (k - 1) * b['legs'] + 1
                ws[f'{VX_MEMO}{r}'] = f'=IF({Cc}=0,"",INDEX({je_memo},{jr}))'
            elif src == '手工':               # 第 k 行是这张凭证的第一行才有
                R = mod(SH_MAN, MN_VID)
                ws[f'{VX_ON}{r}'] = f'=IFERROR(IF(INDEX({R},{k})={A},1,0),0)'
                ws[f'{VX_DATE}{r}'] = f'=IF({Cc}=0,"",INDEX({mod(SH_MAN, MN_VDATE)},{k}))'
                ws[f'{VX_LINES}{r}'] = f'=IF({Cc}=0,0,COUNTIF({R},{A}))'
                ws[f'{VX_MEMO}{r}'] = f'=IF({Cc}=0,"",INDEX({je_memo},{b["je0"] - JE_R0 + k}))'
                att = f'SUMIF({mod(SH_MAN, MN_GKEY)},INDEX({mod(SH_MAN, MN_GKEY)},{k}),{mod(SH_MAN, MN_ATT)})'
            else:                                   # 月末结转：n = k
                r0 = b['je0'] + (k - 1) * b['legs']
                nz = je_nz(r0, r0 + b['legs'] - 1)
                ws[f'{VX_ON}{r}'] = f'=IF(COUNTIF({nz},1)>0,1,0)'
                ws[f'{VX_DATE}{r}'] = f'=IF({Cc}=0,"",{ME}!$B${ME_M0 + k - 1})'
                ws[f'{VX_LINES}{r}'] = f'=COUNTIF({nz},1)'
                ws[f'{VX_MEMO}{r}'] = f'=IF({Cc}=0,"",{ME_MEMO[src](D)})'
            ws[f'{VX_KEY}{r}'] = f'=IF({Cc}=0,"",INT({D})*10^7+{A})'
            ws[f'{VX_PAGES}{r}'] = f'=IF({Cc}=0,0,MAX(1,ROUNDUP(${VX_LINES}{r}/{VP_LINES},0)))'
            ws[f'{VX_MKEY}{r}'] = f'=IF({Cc}=0,"",YEAR({D})*100+MONTH({D}))'
            ws[f'{VX_NO}{r}'] = f'=IF({Cc}=0,"",COUNTIFS({MK},${VX_MKEY}{r},{KY},"<="&${VX_KEY}{r}))'
            ws[f'{VX_ATT}{r}'] = f'=IF({Cc}=0,"",{att})' if att != '""' else '=""'
            ws[f'{VX_FIND}{r}'] = f'=IF({Cc}=0,"",${VX_MKEY}{r}*10000+${VX_NO}{r})'
            ws[f'{VX_DATE}{r}'].number_format = DATE
    for c, w in ((VX_ID, 9), (VX_SRC, 8), (VX_ON, 4), (VX_DATE, 11), (VX_KEY, 15), (VX_LINES, 5), (VX_PAGES, 5),
                 (VX_MKEY, 8), (VX_NO, 6), (VX_MEMO, 40), (VX_ATT, 5), (VX_FIND, 12)):
        ws.column_dimensions[c].width = w
    for nm, col in (('凭证ID', VX_ID), ('凭证有', VX_ON), ('凭证日期', VX_DATE), ('凭证条数', VX_LINES), ('凭证页数', VX_PAGES),
                    ('凭证月份', VX_MKEY), ('凭证号', VX_NO), ('凭证摘要', VX_MEMO), ('凭证附单据', VX_ATT),
                    ('凭证查找键', VX_FIND), ('凭证来源', VX_SRC)):
        _dn(wb, nm, vx_rng(ctx, col))
    ws.freeze_panes = 'A4'
    ws.sheet_state = 'hidden'
    return ws


# ─────────────────────────── 凭证汇总 ───────────────────────────
VL = dict(no='A', date='B', word='C', memo='D', dr='E', cr='F', ok='G', lines='H', pages='I', page0='J', src='K',
          att='L', vid='M', idx='N')
VL_R1 = VL_R0 + VL_N - 1


def build_vlist(wb, ctx):
    ws = wb.create_sheet(SH_VLIST)
    c = VL
    widths(ws, {c['no']: 7, c['date']: 11, c['word']: 9, c['memo']: 44, c['dr']: 17, c['cr']: 17, c['ok']: 12,
                c['lines']: 7, c['pages']: 6, c['page0']: 8, c['src']: 8, c['att']: 7, c['vid']: 9, c['idx']: 6})
    title(ws, '凭证汇总（报表月的全部凭证 · 按日期编号）', c['att'], C_VCH,
          '💡 月份跟【首页】选的报表年度/月份走。凭证号＝当月第几号（按日期排，同一天按录入顺序）。'
          '「打印起始页」是这张凭证在【记账凭证】里从第几页开始——打印时页码范围填 1 到「总页数」。'
          '凭证号默认不印在凭证上（留空手写），要印就把【基础资料】「凭证号是否打印」改成「是」。')
    ws.row_dimensions[3].height = 26
    for i in range(1, CI(c['att']) + 1):
        put(ws, f'{CL(i)}3', None, fill_=FILL_KPI, border=False)
    kp = [('A', '月份：', 'B', f'={REP_M0}', MONTH), ('C', '凭证张数：', 'D', f'=COUNT({c["no"]}{VL_R0}:{c["no"]}{VL_R1})', '0" 张"'),
          ('E', '借方合计：', 'F', f'=SUM({c["dr"]}{VL_R0}:{c["dr"]}{VL_R1})', MONEY),
          ('G', '总页数：', 'H', f'=SUM({c["pages"]}{VL_R0}:{c["pages"]}{VL_R1})', '0" 页"')]
    for lc, lab, vc, f, fmt in kp:
        put(ws, f'{lc}3', lab, F_KPI_L, FILL_KPI, align=AR_, border=False)
        put(ws, f'{vc}3', f, Font(name=YH, sz=12, bold=True, color='FFC00000'), FILL_KPI, fmt, AL, border=False)
    put(ws, 'I3', f'=IF(COUNTIFS(凭证月份,报表年度*100+报表月份,凭证有,1)>{VL_N},"✗ 本月凭证超过 {VL_N} 张，找我加容量",'
                  f'IF(H3>{VP_PAGES},"✗ 总页数超过 {VP_PAGES} 页，找我加容量",""))', F_RED, FILL_KPI, align=AL, border=False)
    ws.merge_cells('I3:L3')
    header(ws, VL_HDR, [(c['no'], '凭证号'), (c['date'], '日期'), (c['word'], '字号'), (c['memo'], '摘要（第一条）'),
                        (c['dr'], '借方合计'), (c['cr'], '贷方合计'), (c['ok'], '借贷平衡'), (c['lines'], '分录\n条数'),
                        (c['pages'], '页数'), (c['page0'], '打印\n起始页'), (c['src'], '来源'), (c['att'], '附单\n据张'),
                        (c['vid'], '凭证ID'), (c['idx'], '索引行')], C_VCH, height=34)
    mk = '(报表年度*100+报表月份)'
    for k in range(1, VL_N + 1):
        r = VL_R0 + k - 1
        X = f'${c["idx"]}{r}'
        ws[f'{c["idx"]}{r}'] = f'=IFERROR(MATCH({mk}*10000+{k},凭证查找键,0),"")'
        ws[f'{c["no"]}{r}'] = f'=IF({X}="","",{k})'
        ws[f'{c["vid"]}{r}'] = f'=IF({X}="","",INDEX(凭证ID,{X}))'
        ws[f'{c["date"]}{r}'] = f'=IF({X}="","",INDEX(凭证日期,{X}))'
        ws[f'{c["word"]}{r}'] = f'=IF({X}="","","记-"&TEXT({k},"000"))'
        ws[f'{c["memo"]}{r}'] = f'=IF({X}="","",INDEX(凭证摘要,{X}))'
        V = f'${c["vid"]}{r}'
        ws[f'{c["dr"]}{r}'] = f'=IF({X}="","",SUMIF(分录凭证ID,{V},分录借方))'
        ws[f'{c["cr"]}{r}'] = f'=IF({X}="","",SUMIF(分录凭证ID,{V},分录贷方))'
        ws[f'{c["ok"]}{r}'] = (f'=IF({X}="","",IF(ROUND({c["dr"]}{r}-{c["cr"]}{r},2)=0,"√",'
                               f'"✗ 差"&FIXED({c["dr"]}{r}-{c["cr"]}{r},2)))')
        ws[f'{c["lines"]}{r}'] = f'=IF({X}="","",INDEX(凭证条数,{X}))'
        ws[f'{c["pages"]}{r}'] = f'=IF({X}="","",INDEX(凭证页数,{X}))'
        ws[f'{c["page0"]}{r}'] = ('=IF({X}="","",1)'.format(X=X) if k == 1 else
                                  f'=IF({X}="","",{c["page0"]}{r - 1}+{c["pages"]}{r - 1})')
        ws[f'{c["src"]}{r}'] = f'=IF({X}="","",INDEX(凭证来源,{X}))'
        ws[f'{c["att"]}{r}'] = f'=IF({X}="","",IF(N(INDEX(凭证附单据,{X}))=0,"",INDEX(凭证附单据,{X})))'
    cols = [c[x] for x in ('no', 'date', 'word', 'memo', 'dr', 'cr', 'ok', 'lines', 'pages', 'page0', 'src', 'att')]
    style_rows(ws, VL_R0, VL_R1, cols, auto=tuple(cols),
               fmts={c['date']: DATE, c['dr']: MONEY, c['cr']: MONEY, c['no']: '0'},
               aligns={c['memo']: AL, c['dr']: AR_, c['cr']: AR_})
    for r in range(VL_R0, VL_R1 + 1):
        ws[f'{c["vid"]}{r}'].font = F_NOTE
        ws[f'{c["idx"]}{r}'].font = F_NOTE
    hide(ws, [c['vid'], c['idx']])
    ws.conditional_formatting.add(f'{c["ok"]}{VL_R0}:{c["ok"]}{VL_R1}',
                                  FormulaRule(formula=[f'LEFT(${c["ok"]}{VL_R0},1)="✗"'], fill=FILL_WARN,
                                              font=Font(name=YH, sz=10, bold=True, color='FFC00000')))
    _dn(wb, '汇总起始页', rng(SH_VLIST, c['page0'], VL_R0, VL_R1))
    _dn(wb, '汇总页数', rng(SH_VLIST, c['pages'], VL_R0, VL_R1))
    _dn(wb, '汇总凭证ID', rng(SH_VLIST, c['vid'], VL_R0, VL_R1))
    _dn(wb, '汇总日期', rng(SH_VLIST, c['date'], VL_R0, VL_R1))
    _dn(wb, '汇总附单据', rng(SH_VLIST, c['att'], VL_R0, VL_R1))
    _dn(wb, '本月总页数', f"{q(SH_VLIST)}!$H$3")
    ws.freeze_panes = f'C{VL_R0}'
    ws.auto_filter.ref = f'A{VL_HDR}:{c["att"]}{VL_R1}'
    lock_formula_sheet(ws)
    return ws


# ─────────────────────────── 记账凭证（打印） ───────────────────────────
# 打印区 B:G；右边 I..N 是核对用的辅助列（印不出来），P..R 每条分录的查找键（隐藏）
VP = dict(memo='B', acc1='C', acc2='D', dr='E', cr='F', tick='G',
          page='I', no='J', word='K', pp='L', vid='M', k='N', key='P', jrow='Q', code='R')


def build_vprint(wb, ctx):
    ws = wb.create_sheet(SH_VPRT)
    c = VP
    widths(ws, {'A': 1.5, c['memo']: 36, c['acc1']: 13, c['acc2']: 26, c['dr']: 16, c['cr']: 16, c['tick']: 5, 'H': 2,
                c['page']: 9, c['no']: 8, c['word']: 26, c['pp']: 6, c['vid']: 9, c['k']: 6, 'O': 2,
                c['key']: 11, c['jrow']: 7, c['code']: 10})
    ws.sheet_properties.tabColor = C_VCH[2:]
    thin_b = Side(style='thin', color='FF000000')
    med = Side(style='medium', color='FF000000')
    dbl = Side(style='double', color='FF000000')
    F_T = Font(name=YH, sz=16, bold=True, color='FF000000')
    F_H = Font(name=YH, sz=10, bold=True, color='FF000000')
    F_L = Font(name=YH, sz=10, color='FF000000')
    F_S = Font(name=YH, sz=9, color='FF000000')
    F_NO = Font(name=YH, sz=10, bold=True, color='FFC00000')
    breaks = []
    for p in range(1, VP_PAGES + 1):
        R = (p - 1) * VP_ROWS + 1
        P0 = f'${c["page"]}${R}'
        K = f'${c["k"]}${R}'
        PP = f'${c["pp"]}${R}'
        VID = f'${c["vid"]}${R}'
        pages = f'INDEX(汇总页数,{K})'
        # 辅助列
        ws[f'{c["page"]}{R}'] = f'=IF({p}>本月总页数,"",{p})'
        ws[f'{c["k"]}{R}'] = f'=IF({P0}="","",COUNTIF(汇总起始页,"<="&{P0}))'
        ws[f'{c["vid"]}{R}'] = f'=IF({P0}="","",INDEX(汇总凭证ID,{K}))'
        ws[f'{c["pp"]}{R}'] = f'=IF({P0}="","",{P0}-INDEX(汇总起始页,{K})+1)'
        ws[f'{c["no"]}{R}'] = f'=IF({P0}="","",{K})'
        ws[f'{c["word"]}{R}'] = (f'=IF({P0}="","","记-"&TEXT({K},"000")&IF({pages}>1,"（"&{PP}&"/"&{pages}&"）",""))')
        for col in (c['page'], c['k'], c['vid'], c['pp'], c['word']):
            ws[f'{col}{R}'].font = F_NOTE
        ws[f'{c["no"]}{R}'].font = F_NO
        # 第 1 行：标题
        ws.merge_cells(f'{c["memo"]}{R}:{c["tick"]}{R}')
        put(ws, f'{c["memo"]}{R}', f'=IF({P0}="","","记  账  凭  证")', F_T, align=AC, border=False)
        ws.row_dimensions[R].height = 26
        # 第 2 行：单位 · 日期 · 字号 · 附单据
        r = R + 1
        put(ws, f'{c["memo"]}{r}', f'=IF({P0}="","","单位："&公司名称)', F_S, align=AL, border=False)
        ws.merge_cells(f'{c["acc1"]}{r}:{c["acc2"]}{r}')
        put(ws, f'{c["acc1"]}{r}', f'=IF({P0}="","",INDEX(汇总日期,{K}))', F_H, None, 'yyyy"年"m"月"d"日"', AC, border=False)
        ws.merge_cells(f'{c["dr"]}{r}:{c["cr"]}{r}')
        put(ws, f'{c["dr"]}{r}', (f'=IF({P0}="","","记 字第 "&IF(凭证号打印="是",TEXT({K},"000"),"　　　　")&" 号"'
                                  f'&IF({pages}>1,"（共"&{pages}&"页 第"&{PP}&"页）",""))'), F_H, align=AC, border=False)
        put(ws, f'{c["tick"]}{r}', f'=IF({P0}="","",IF(N(INDEX(汇总附单据,{K}))=0,"附　 张","附"&N(INDEX(汇总附单据,{K}))&"张"))',
            F_S, align=AC, border=False)
        ws.row_dimensions[r].height = 18
        # 第 3 行：表头
        r = R + 2
        for col, t in ((c['memo'], '摘  要'), (c['acc1'], '总账科目'), (c['acc2'], '明细科目'), (c['dr'], '借方金额'),
                       (c['cr'], '贷方金额'), (c['tick'], '√')):
            put(ws, f'{col}{r}', f'=IF({P0}="","","{t}")', F_H, align=AC, border=False)
        ws.row_dimensions[r].height = 20
        # 8 条分录
        for i in range(1, VP_LINES + 1):
            r = R + 2 + i
            key, jr, code = f'${c["key"]}{r}', f'${c["jrow"]}{r}', f'${c["code"]}{r}'
            ws[f'{c["key"]}{r}'] = f'=IF({P0}="","",{VID}*1000+({PP}-1)*{VP_LINES}+{i})'
            ws[f'{c["jrow"]}{r}'] = f'=IF({P0}="","",IFERROR(MATCH({key},分录键,0),""))'
            ws[f'{c["code"]}{r}'] = f'=IF({jr}="","",INDEX(分录科目,{jr}))'
            put(ws, f'{c["memo"]}{r}', f'=IF({jr}="","",INDEX(分录摘要,{jr}))', F_S, align=ALW, border=False)
            put(ws, f'{c["acc1"]}{r}', f'=IF({jr}="","",IFERROR(INDEX({acc(AC_NAME)},MATCH(LEFT({code},4),{acc(AC_CODE)},0)),""))',
                F_L, align=ALW, border=False)
            put(ws, f'{c["acc2"]}{r}', (f'=IF({jr}="","",IF(LEN({code})>4,MID(INDEX(分录全称,{jr}),LEN({c["acc1"]}{r})+2,200),"")'
                                        f'&"（"&{code}&"）")'), F_L, align=ALW, border=False)
            put(ws, f'{c["dr"]}{r}', f'=IF({jr}="","",IF(INDEX(分录借方,{jr})=0,"",INDEX(分录借方,{jr})))', F_L, None, MONEY0, AR_, border=False)
            put(ws, f'{c["cr"]}{r}', f'=IF({jr}="","",IF(INDEX(分录贷方,{jr})=0,"",INDEX(分录贷方,{jr})))', F_L, None, MONEY0, AR_, border=False)
            put(ws, f'{c["tick"]}{r}', None, F_L, border=False)
            for col in (c['key'], c['jrow'], c['code']):
                ws[f'{col}{r}'].font = F_NOTE
            ws.row_dimensions[r].height = 22
        # 合计
        r = R + 3 + VP_LINES
        last = f'IF({P0}="",FALSE,{PP}={pages})'
        put(ws, f'{c["memo"]}{r}', f'=IF({P0}="","",IF({last},"合    计","（接下页）"))', F_H, align=AC, border=False)
        ws.merge_cells(f'{c["acc1"]}{r}:{c["acc2"]}{r}')
        put(ws, f'{c["acc1"]}{r}', None, F_H, border=False)
        put(ws, f'{c["dr"]}{r}', f'=IF({last},SUMIF(分录凭证ID,{VID},分录借方),"")', F_H, None, MONEY0, AR_, border=False)
        put(ws, f'{c["cr"]}{r}', f'=IF({last},SUMIF(分录凭证ID,{VID},分录贷方),"")', F_H, None, MONEY0, AR_, border=False)
        put(ws, f'{c["tick"]}{r}', None, F_H, border=False)
        ws.row_dimensions[r].height = 22
        # 签字
        r = R + 4 + VP_LINES
        sig = [(c['memo'], '会计主管：', '会计主管'), (c['acc1'], '记账：', None), (c['acc2'], '出纳：', '出纳'),
               (c['dr'], '复核：', '复核人'), (c['cr'], '制单：', '制单人')]
        for col, t, nm in sig:
            f = f'=IF({P0}="","","{t}"&' + (f'IF({nm}="","",{nm})' if nm else '""') + ')'
            put(ws, f'{col}{r}', f, F_S, align=AL, border=False)
        ws.row_dimensions[r].height = 18
        ws.row_dimensions[R + 5 + VP_LINES].height = 8
        breaks.append(R + VP_ROWS - 1)
    # 边框只在有凭证的页上显示（空页整页空白）
    last_r = VP_PAGES * VP_ROWS
    box = Border(left=thin_b, right=thin_b, top=thin_b, bottom=thin_b)
    tot = Border(left=thin_b, right=thin_b, top=thin_b, bottom=thin_b)
    # 表头 + 分录行 + 合计行：按页内行号判断
    rel = f'MOD(ROW()-1,{VP_ROWS})'
    used = f'INDEX(${c["page"]}$1:${c["page"]}${last_r},ROW()-{rel})<>""'
    ws.conditional_formatting.add(f'{c["memo"]}1:{c["tick"]}{last_r}',
                                  FormulaRule(formula=[f'AND({used},{rel}>=2,{rel}<={2 + VP_LINES})'], border=box))
    ws.conditional_formatting.add(f'{c["memo"]}1:{c["tick"]}{last_r}',
                                  FormulaRule(formula=[f'AND({used},{rel}={3 + VP_LINES})'], border=tot))
    ws.conditional_formatting.add(f'{c["memo"]}1:{c["tick"]}{last_r}',
                                  FormulaRule(formula=[f'AND({used},{rel}=2)'], fill=fill('FFF2F2F2')))
    for b in breaks:
        ws.row_breaks.append(Break(id=b))
    ws.print_area = f'{c["memo"]}1:{c["tick"]}{last_r}'
    ws.page_setup.paperSize = 11          # A5
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top, ws.page_margins.bottom = 0.4, 0.3
    ws.print_options.horizontalCentered = True
    # 右边说明（打印区外）
    tips = ['💡 右边 I、J、K 三列是核对用的（打印区域只到 G 列，印不出来）：I＝打印第几页，J＝凭证号（红字，照着手写到凭证上），K＝字号。',
            '一张凭证超过 8 条分录会接着印下一页，两页是同一个号，别顺着页码写。',
            '月份跟【首页】的报表年度/月份走；打印时页码范围填 1 到【凭证汇总】的「总页数」，后面的空页不用印。']
    for i, t in enumerate(tips):
        ws[f'S{1 + i}'] = t
        ws[f'S{1 + i}'].font = F_TIP
        ws[f'S{1 + i}'].alignment = ALW
    ws.column_dimensions['S'].width = 70
    hide(ws, [c['pp'], c['vid'], c['k'], c['key'], c['jrow'], c['code']])
    ws.freeze_panes = None
    return ws
