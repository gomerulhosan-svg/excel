# -*- coding: utf-8 -*-
"""【基础资料】公司 / 资金账户 / 收支项目 / 摘要关键词 / 期初往来；【往来单位】；【_辅助】下拉清单。"""
import datetime as dt
from common import *


def build_base(wb, ctx):
    ws = wb.create_sheet(SH_BASE)
    widths(ws, {'A': 5, 'B': 12, 'C': 30, 'D': 21, 'E': 12, 'F': 4, 'G': 2,
                'H': 5, 'I': 13, 'J': 11, 'K': 8, 'L': 21, 'M': 20, 'N': 13, 'O': 11, 'P': 13, 'Q': 13, 'R': 14, 'S': 2,
                'T': 5, 'U': 12, 'V': 10, 'W': 8, 'X': 26, 'Y': 2,
                'Z': 5, 'AA': 12, 'AB': 7, 'AC': 12, 'AD': 2,
                'AE': 5, 'AF': 11, 'AG': 24, 'AH': 12, 'AI': 12, 'AJ': 14, 'AK': 16})
    title(ws, '基 础 资 料（公司 · 资金账户 · 收支项目 · 摘要关键词 · 期初往来）', 'AK', C_BASE,
          '💡 淡黄格子手填，灰格子自动。① 公司：填简称、全称、税号（发票靠税号认是哪家公司的）；'
          '② 资金账户：每个银行卡、微信、支付宝、现金各一行，选所属公司，填建账那天的期初余额；'
          '③ 收支项目：就是你截图那张「收支类型」，可以往下加，类别决定落到报表哪一块；'
          '④ 摘要关键词：流水摘要/对方户名里出现这个词就自动认成这个项目（比如「费用外收」→ 银行手续费）；'
          '⑤ 期初往来：建账前就有的应收、应付、借款余额，填一次。往来单位在下一张表【往来单位】。')
    # 第 3 行分块标题
    for c1, c2, t, col in (('A', 'E', '① 公司（5 家，最多 8 家）', C_HOME), ('H', 'R', '② 资金账户（15 个，最多 20 个）', C_CASH),
                           ('T', 'X', '③ 收支项目（收支类型）', C_RPT), ('Z', 'AC', '④ 摘要关键词 → 收支项目', C_INV),
                           ('AE', 'AK', '⑤ 期初往来（建账前的余额，填一次）', C_AR)):
        section(ws, 3, c1, c2, t, col)
    # ① 公司
    header(ws, BASE_HDR, [('A', '序号'), ('B', '公司简称'), ('C', '公司全称（跟营业执照一致）'), ('D', '纳税人识别号'), ('E', '备注')],
           C_HOME)
    for i, r in enumerate(range(CO_R0, CO_R1 + 1)):
        put(ws, f'A{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in 'BCDE':
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, align=AL if c in 'CE' else AC, fmt='@' if c == 'D' else None)
        ws[f'{CO_NORM}{r}'] = f'=IF({CO_FULL}{r}="","",{norm(CO_FULL + str(r))})'
        ws[f'{CO_NORM}{r}'].font = F_HELP
    ws[f'B{CO_R0}'], ws[f'C{CO_R0}'], ws[f'D{CO_R0}'] = ctx['my_co']
    ws[f'E{CO_R0}'] = '其余 4 家请补上'
    hide(ws, CO_NORM)
    # ② 资金账户
    header(ws, BASE_HDR, [('H', '序号'), ('I', '账户名称（简称）'), ('J', '所属公司'), ('K', '类型'), ('L', '账号'), ('M', '开户行'),
                          ('N', '期初余额'), ('O', '期初日期'), ('P', '当前余额（自动）'), ('Q', '最后一笔日期（自动）'), ('R', '备注')],
           C_CASH)
    for i, r in enumerate(range(AC_R0, AC_R1 + 1)):
        put(ws, f'H{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in 'IJKLMNOR':
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, align=AL if c in 'LMR' else AC,
                fmt=MONEY if c == 'N' else (DATE if c == 'O' else ('@' if c == 'L' else None)))
        put(ws, f'P{r}', f'=IF(I{r}="","",ROUND(N(N{r})+SUMIFS({jr(J_NET)},{jr(J_ACC)},I{r}),2))', F_AUTOB, FILL_AUTO, MONEY, AR)
        put(ws, f'Q{r}', f'=IF(I{r}="","",IF(COUNTIFS({jr(J_ACC)},I{r},{jr(J_DATE)},">0")=0,"",'
                         f'_xlfn.MAXIFS({jr(J_DATE)},{jr(J_ACC)},I{r})))', F_AUTO, FILL_AUTO, DATE, AC)
    acc, opening, _ = ctx['bank']
    r = AC_R0
    ws[f'I{r}'], ws[f'J{r}'], ws[f'K{r}'], ws[f'L{r}'], ws[f'M{r}'] = ctx['my_acc'], ctx['my_co'][0], '银行', acc, '中国农业银行宁波五乡支行'
    ws[f'N{r}'], ws[f'O{r}'] = opening, dt.datetime(2026, 9, 1)
    ws[f'R{r}'] = '期初＝9/1 第一笔之前的余额'
    dv_list(ws, f'J{AC_R0}:J{AC_R1}', f'={CO_NAMES}', '这个账户是哪家公司的')
    dv_list(ws, f'K{AC_R0}:K{AC_R1}', '"银行,微信,支付宝,现金,其他"')
    dv_date(ws, f'O{AC_R0}:O{AC_R1}')
    # ③ 收支项目
    header(ws, BASE_HDR, [('T', '序号'), ('U', '收支项目'), ('V', '类别'), ('W', '报表显示'), ('X', '说明')], C_RPT)
    notes = {'新增': '客户第一次来款自动算新增', '续费': '同一客户第二次起自动算续费', '其他收入': '报表上显示「其他」',
             '其他成本': '报表上显示「其他」', '往来': '借款、还款、押金等（其他应收/其他应付）',
             '内部划转': '5 家公司之间转钱（自动认：对方户名是自家公司）', '账户互转': '同一家公司两个账户之间倒钱（自动认）',
             '社保挂靠': '需要手工选', '工资挂靠': '需要手工选', '咨询服务': '截图里没有，你说的「改为咨询服务」加在这；不要可删', '实收资本': '收到为准', '投资': '支出为准'}
    r = IT_R0
    for cls, names in ITEMS:
        for n in names:
            put(ws, f'T{r}', r - IT_R0 + 1, F_AUTO, FILL_AUTO, align=AC)
            put(ws, f'U{r}', n, F_IN, FILL_IN, align=AC)
            put(ws, f'V{r}', cls, F_IN, FILL_IN, align=AC)
            put(ws, f'W{r}', SHOW_NAME.get(n, n), F_IN, FILL_IN, align=AC)
            put(ws, f'X{r}', notes.get(n), F_NOTE, FILL_IN, align=AL)
            r += 1
    for rr in range(r, IT_R1 + 1):
        put(ws, f'T{rr}', rr - IT_R0 + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in 'UVWX':
            put(ws, f'{c}{rr}', None, F_IN, FILL_IN, align=AC)
    dv_list(ws, f'V{IT_R0}:V{IT_R1}', '"' + ','.join(CLS_ALL) + '"')
    # ④ 摘要关键词
    header(ws, BASE_HDR, [('Z', '序号'), ('AA', '关键词'), ('AB', '适用'), ('AC', '收支项目')], C_INV)
    kws = [('费用外收', '支出', '银行手续费'), ('手续费', '支出', '银行手续费'), ('短信费', '支出', '银行手续费'),
           ('年费', '支出', '银行手续费'), ('账户管理费', '支出', '银行手续费'), ('国库', '支出', '税金'),
           ('扣税', '支出', '税金'), ('税款', '支出', '税金'), ('结息', '收入', '其他收入'), ('利息', '支出', '贷款利息'),
           ('房租', '支出', '店面租金'), ('租金', '支出', '店面租金'), ('刻章', '收入', '刻章'), ('刻章', '支出', '刻章费用'),
           ('返税', '收入', '园区返税'), ('加油', '支出', '汽车费用'), ('油费', '支出', '汽车费用'), ('停车', '支出', '汽车费用'),
           ('话费', '支出', '办公费用'), ('宽带', '支出', '办公费用'), ('工资', '支出', '人工-工资')]
    for i, r in enumerate(range(KW_R0, KW_R1 + 1)):
        put(ws, f'Z{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in ('AA', 'AB', 'AC'):
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, align=AC)
        if i < len(kws):
            ws[f'AA{r}'], ws[f'AB{r}'], ws[f'AC{r}'] = kws[i]
    dv_list(ws, f'AB{KW_R0}:AB{KW_R1}', '"收入,支出,全部"')
    dv_list(ws, f'AC{KW_R0}:AC{KW_R1}', f'={IT_NAMES}')
    # ⑤ 期初往来
    header(ws, BASE_HDR, [('AE', '序号'), ('AF', '所属公司'), ('AG', '往来单位'), ('AH', '期初应收\n（对方欠我方货款）'),
                          ('AI', '期初应付\n（我方欠对方货款）'), ('AJ', '期初其他往来\n借出/押金＋ 借入－'), ('AK', '备注')], C_AR)
    for i, r in enumerate(range(OP_R0, OP_R1 + 1)):
        put(ws, f'AE{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in ('AF', 'AG', 'AH', 'AI', 'AJ', 'AK'):
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, align=AC if c in ('AF',) else (AL if c in ('AG', 'AK') else AR),
                fmt=MONEY if c in ('AH', 'AI', 'AJ') else None)
    dv_list(ws, f'AF{OP_R0}:AF{OP_R1}', f'={CO_NAMES}')
    dv_list(ws, f'AG{OP_R0}:AG{OP_R1}', f'={SH_AUX}!$C$1:$C$620', '往来单位或另一家自家公司', stop=False)
    ws.freeze_panes = 'A5'
    return ws


def build_party(wb, ctx):
    ws = wb.create_sheet(SH_PARTY)
    widths(ws, {'A': 6, 'B': 30, 'C': 9, 'D': 12, 'E': 30, 'F': 22, 'G': 21, 'H': 21, 'I': 10, 'J': 10, 'K': 22})
    title(ws, '往 来 单 位（客户 · 供应商 · 个人 · 税务银行…… 绑定收支项目、流水户名）', 'K', C_BASE,
          '💡 流水里的「对方户名」跟这里的名称、别名 1、别名 2 任意一个对上（括号全角半角都认），或者「对方账号」对上，资金台帐就自动填好往来单位；'
          '「绑定收支项目」填了，这家的收付款就自动归到这个项目（比如房东 → 店面租金、刻章店 → 刻章费用）。客户不用绑：第一次来款自动算「新增」，'
          '以后算「续费」；建账前就合作的老客户在「建账前已合作」填 是，首笔就算续费。发票靠税号对单位。供应商的进项发票默认计应付；'
          '只是报销用的小票（吃饭、加油……）不用登记，不登记的进项发票不算应付。')
    header(ws, PT_HDR, [('A', '序号'), ('B', '名称（标准名）'), ('C', '类型'), ('D', '绑定收支项目'), ('E', '流水户名 / 别名 1'),
                        ('F', '别名 2'), ('G', '对方账号'), ('H', '纳税人识别号'), ('I', '建账前\n已合作'), ('J', '进项发票\n计应付'),
                        ('K', '备注')], C_BASE)
    for i, r in enumerate(range(PT_R0, PT_R1 + 1)):
        put(ws, f'A{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in 'BCDEFGHIJK':
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, align=AL if c in 'BEFK' else AC, fmt='@' if c in 'GH' else None)
        for c, src in ((PT_N0, 'B'), (PT_N1, 'E'), (PT_N2, 'F')):
            ws[f'{c}{r}'] = f'=IF({src}{r}="","",{norm(src + str(r))})'
            ws[f'{c}{r}'].font = F_HELP
    for i, p in enumerate(ctx['parties']):
        r = PT_R0 + i
        ws[f'B{r}'] = p['名称']
        ws[f'C{r}'] = p['类型']
        ws[f'D{r}'] = p.get('绑定')
        ws[f'E{r}'] = p.get('别名1')
        ws[f'H{r}'] = p.get('税号')
        ws[f'K{r}'] = p.get('备注')
    hide(ws, PT_N0, PT_N1, PT_N2)
    dv_list(ws, f'C{PT_R0}:C{PT_R1}', '"' + ','.join(PT_TYPES) + '"')
    dv_list(ws, f'D{PT_R0}:D{PT_R1}', f'={IT_NAMES}', '这家单位的收付款自动归到这个项目；客户一般不用填（自动新增/续费）')
    dv_list(ws, f'I{PT_R0}:I{PT_R1}', '"是"')
    dv_list(ws, f'J{PT_R0}:J{PT_R1}', '"是,否"', '供应商留空＝是；填「否」这家的进项发票不算应付')
    ws.auto_filter.ref = f'A{PT_HDR}:K{PT_R1}'
    ws.freeze_panes = f'C{PT_R0}'
    return ws


def build_aux(wb, ctx):
    ws = wb.create_sheet(SH_AUX)
    ws['A1'] = '全部'
    for i in range(CO_R1 - CO_R0 + 1):
        ws[f'A{i + 2}'] = f'=IF({SH_BASE}!$B${CO_R0 + i}="","",{SH_BASE}!$B${CO_R0 + i})'
    # C 列：往来单位 ＋ 自家公司（给期初往来、往来对账单的下拉）
    for i in range(PT_R1 - PT_R0 + 1):
        ws[f'C{i + 1}'] = f'=IF({SH_PARTY}!$B${PT_R0 + i}="","",{SH_PARTY}!$B${PT_R0 + i})'
    for i in range(CO_R1 - CO_R0 + 1):
        ws[f'C{PT_R1 - PT_R0 + 2 + i}'] = f'=IF({SH_BASE}!$B${CO_R0 + i}="","",{SH_BASE}!$B${CO_R0 + i})'
    ws.sheet_state = 'hidden'
    return ws
