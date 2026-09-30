# -*- coding: utf-8 -*-
"""【期初余额】建账日那天的数，只填一次：项目期初、应付期初（单位×项目）、其他期初。
   资金账户的期初在【基础资料】②，工人的期初欠薪在【往来单位】。"""
from common import *
from layout import *


def build_open(wb, ctx):
    ws = wb.create_sheet(SH_OPEN)
    widths(ws, {'A': 5, 'B': 16, 'C': 14, 'D': 14, 'E': 13, 'F': 13, 'G': 12, 'H': 12, 'I': 12, 'J': 13, 'K': 13, 'L': 12, 'M': 24,
                'N': 2, 'O': 5, 'P': 16, 'Q': 14, 'R': 13, 'S': 13, 'T': 12, 'U': 24, 'V': 10})
    title(ws, '期 初 余 额（建账日那天，以前的账一次性录进来）', 'V', C_BASE,
          '💡 建账日期在【基础资料】① 里（现在是 ' + ctx['open_date'].strftime('%Y-%m-%d') + '）。以后的收付款、应付、产值确认都按实际日期在各表录，'
          '建账日之前的只在这里填累计数。'
          '左边「项目期初」：每个老项目一行，建账前累计确认的产值、累计收到的工程款、累计开的票、人工、其他直接费、税金（实交或估算）、已摊管理费（材料、分包、机械从右边自动汇总）。'
          '右边「应付期初」：每家材料商、分包、机械，每个项目一行：累计应付、累计已付、累计收到的票（跟你原来的应付总表一样）。'
          '下面「其他期初」：贷款、保证金、欠的税、实收资本等。填完看【资产负债表】的年初数平不平。')
    section(ws, 3, 'A', 'M', '① 项目期初（建账前累计）', C_RPT)
    header(ws, OP_HDR, [(OPJ_SEQ, '序号'), (OPJ_PJ, '项目'), (OPJ_REV, '累计确认产值\n（含税）'), (OPJ_REC, '累计收到\n工程款'),
                        (OPJ_INV, '累计开票'), (OPJ_LAB, '人工费'), (OPJ_OTH, '其他直接费'), (OPJ_TAX, '税金\n（估算）'), (OPJ_ALLOC, '已摊管理费'),
                        (OPJ_MAT, '材料\n（自动）'), (OPJ_SUB, '分包\n（自动）'), (OPJ_MACH, '机械\n（自动）'), (OPJ_NOTE, '备注')], C_RPT)
    for i, r in enumerate(range(OPJ_R0, OPJ_R1 + 1)):
        put(ws, f'{OPJ_SEQ}{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in (OPJ_PJ, OPJ_REV, OPJ_REC, OPJ_INV, OPJ_LAB, OPJ_OTH, OPJ_TAX, OPJ_ALLOC, OPJ_NOTE):
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, None if c in (OPJ_PJ, OPJ_NOTE) else MONEY,
                AL if c in (OPJ_PJ, OPJ_NOTE) else AR)
        for c, t in ((OPJ_MAT, '材料供应商'), (OPJ_SUB, '分包'), (OPJ_MACH, '机械运输')):
            put(ws, f'{c}{r}', f'=IF({OPJ_PJ}{r}="","",SUMIFS({oapr(OAP_AMT)},{oapr(OAP_PJ)},{OPJ_PJ}{r},{oapr(OAP_TYPE)},"{t}"))',
                F_AUTO, FILL_AUTO, MONEY, AR)
    section(ws, 3, 'O', 'V', '② 应付期初（单位 × 项目，建账前累计）', C_AR)
    header(ws, OP_HDR, [(OAP_SEQ, '序号'), (OAP_UNIT, '单位'), (OAP_PJ, '项目'), (OAP_AMT, '累计应付'), (OAP_PAID, '累计已付'),
                        (OAP_INV, '累计收票'), (OAP_NOTE, '备注'), (OAP_TYPE, '类型\n（自动）')], C_AR)
    UT = rng(SH_UNIT, UN_TYPE, UN_R0, UN_R1)
    for i, r in enumerate(range(OAP_R0, OAP_R1 + 1)):
        put(ws, f'{OAP_SEQ}{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in (OAP_UNIT, OAP_PJ, OAP_AMT, OAP_PAID, OAP_INV, OAP_NOTE):
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, None if c in (OAP_UNIT, OAP_PJ, OAP_NOTE) else MONEY,
                AL if c in (OAP_UNIT, OAP_PJ, OAP_NOTE) else AR)
        put(ws, f'{OAP_TYPE}{r}', f'=IF({OAP_UNIT}{r}="","",IFERROR(INDEX({UT},MATCH({OAP_UNIT}{r},{UN_NAMES},0))&"","？"))',
            F_AUTO, FILL_AUTO, align=AC)
    # 合计行
    tr = OPJ_R1 + 1
    put(ws, f'{OPJ_PJ}{tr}', '合计', F_TXTB, FILL_TOT, align=AC)
    for c in (OPJ_REV, OPJ_REC, OPJ_INV, OPJ_LAB, OPJ_OTH, OPJ_TAX, OPJ_ALLOC, OPJ_MAT, OPJ_SUB, OPJ_MACH):
        put(ws, f'{c}{tr}', f'=SUM({c}{OPJ_R0}:{c}{OPJ_R1})', F_AUTOB, FILL_TOT, MONEY, AR)
    # 其他期初
    section(ws, OO_R0 - 2, 'A', 'M', '③ 其他期初', C_BASE)
    header(ws, OO_R0 - 1, [('A', '序号'), (OO_LBL, '项目'), (OO_VAL, '金额'), (OO_NOTE, '说明')], C_BASE)
    ws.merge_cells(f'{OO_NOTE}{OO_R0 - 1}:M{OO_R0 - 1}')
    notes = ['建账那天欠银行的本金', '投标、履约保证金、押金还没退回的', '上面没有的其他应收（员工借款请在【往来单位】填负的期初欠薪）',
             '建账那天还欠税务局的税（多交了填负数）', '上面没有的其他应付（老板个人借款请填在【基础资料】② 他的个人户期初里，填负数）',
             '股东实际缴进来的注册资本', '建账前赊购的车、设备还欠多少（车、设备本身在【基础资料】⑥ 登记）',
             '会计或老板认定的数；不填也行，【资产负债表】自动倒算并对比']
    for i, (lab, note) in enumerate(zip(OO_ITEMS, notes)):
        r = OO_R0 + i
        put(ws, f'A{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        put(ws, f'{OO_LBL}{r}', lab, F_TXT, align=AL)
        put(ws, f'{OO_VAL}{r}', None, F_IN, FILL_IN, MONEY, AR)
        ws.merge_cells(f'{OO_NOTE}{r}:M{r}')
        put(ws, f'{OO_NOTE}{r}', note, F_NOTE, align=AL)
        ws.row_dimensions[r].height = 22
    for i, v in enumerate(ctx.get('open_other', [])):
        if v not in (None, ''):
            ws[f'{OO_VAL}{OO_R0 + i}'] = v
    for i, p in enumerate(ctx['open_proj']):
        r = OPJ_R0 + i
        for c, k in ((OPJ_PJ, 'pj'), (OPJ_REV, 'rev'), (OPJ_REC, 'rec'), (OPJ_INV, 'inv'), (OPJ_LAB, 'lab'), (OPJ_OTH, 'oth'), (OPJ_TAX, 'tax'),
                     (OPJ_ALLOC, 'alloc'), (OPJ_NOTE, 'note')):
            v = p.get(k)
            if v not in (None, ''):
                ws[f'{c}{r}'] = v
    for i, a in enumerate(ctx['open_ap']):
        r = OAP_R0 + i
        for c, k in ((OAP_UNIT, 'unit'), (OAP_PJ, 'pj'), (OAP_AMT, 'amt'), (OAP_PAID, 'paid'), (OAP_INV, 'inv'), (OAP_NOTE, 'note')):
            v = a.get(k)
            if v not in (None, ''):
                ws[f'{c}{r}'] = v
    dv_list(ws, f'{OPJ_PJ}{OPJ_R0}:{OPJ_PJ}{OPJ_R1}', f'={PJ_NAMES}')
    dv_list(ws, f'{OAP_PJ}{OAP_R0}:{OAP_PJ}{OAP_R1}', f'={PJ_NAMES}', stop=False)
    dv_list(ws, f'{OAP_UNIT}{OAP_R0}:{OAP_UNIT}{OAP_R1}', f'={UN_NAMES}', stop=False)
    ws.freeze_panes = f'C{OPJ_R0}'
    return ws
