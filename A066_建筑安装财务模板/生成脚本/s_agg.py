# -*- coding: utf-8 -*-
"""隐藏的汇总层：【_月汇总】公司利润表项目 × 月；【_项目汇总】每个项目的各种数；【_余额】每个单位/人、每个账户的余额。
   各报表都从这三张取数，每个数只算一遍（不卡），口径也一致。"""
from common import *
from layout import *
import calc
import cats


def _open_proj(name, col):
    return f'SUMIFS({opr(col)},{opr(OPJ_PJ)},{name})'


def build_ms(wb, ctx):
    ws = wb.create_sheet(SH_MS)
    ws['A1'] = '公司利润表项目 × 月（取数用，别改）'
    ws['A2'], ws[f'{MS_PREV}2'] = '年月→', f'={AX_PYM}'
    for i, c in enumerate(MS_MCOLS):
        ws[f'{c}{MS_YMROW}'] = f'={AX_Y}*100+{i + 1}'
    ws[f'{MS_YTD}{MS_YMROW}'] = f'={AX_YM1}'
    ws[f'{MS_PREV}1'], ws[f'{MS_YTD}1'] = '建账起到上年底', '本年到截止月'
    for k, name in enumerate(cats.IS_NAMES):
        r = MS_R0 + k
        ws[f'{MS_NAME}{r}'] = name
        ws[f'{MS_PREV}{r}'] = f'={calc.line(name, AX_OYM, AX_PYM)}'
        for c in MS_MCOLS:
            ws[f'{c}{r}'] = f'=IF({c}${MS_YMROW}>{AX_YM1},0,{calc.line(name, f"MAX({c}${MS_YMROW},{AX_OYM})", f"{c}${MS_YMROW}")})'
        ws[f'{MS_YTD}{r}'] = f'=SUMPRODUCT(({MS_MCOLS[0]}${MS_YMROW}:{MS_MCOLS[-1]}${MS_YMROW}<={AX_YM1})*{MS_MCOLS[0]}{r}:{MS_MCOLS[-1]}{r})'
    ws.sheet_state = 'hidden'
    return ws


def ms(name, col):
    """_月汇总 里某个利润表项目某一列的格子"""
    return f"{SH_MS}!${col}${MS_R0 + cats.IS_NAMES.index(name)}"


def build_ps(wb, ctx):
    ws = wb.create_sheet(SH_PS)
    ws['A1'] = '每个项目的各种数（取数用，别改）：前＝建账起到上年底，本＝本年到截止月'
    for (m, p), c in PS_COL.items():
        ws[f'{c}2'] = f'{m}{"(前)" if p == "前" else "(本年)"}'
    for i, c in enumerate(PS_LABM):
        ws[f'{c}2'] = f'人工{i + 1}月'
    for i, c in enumerate(PS_DRV):
        ws[f'{c}2'] = f'={AX_Y0}+{i}'
    ws[f'{PS_AR_E}2'], ws[f'{PS_AR_B}2'], ws[f'{PS_OPEN}2'] = '应收(截止)', '应收(年初)', '期初应收'
    per = {'前': (AX_OYM, AX_PYM), '本': (f'MAX({AX_YM0},{AX_OYM})', AX_YM1)}
    OFS = lambda a, b, pj: f'SUMIFS({ofr(OF_AMT)},{ofr(OF_PJ)},{pj},{ofr(OF_YM)},">="&{a},{ofr(OF_YM)},"<="&{b})'
    INV = lambda a, b, pj: (f'SUMIFS({ivr(IV_TOTU)},{ivr(IV_DIR)},"销项",{ivr(IV_PJ)},{pj},'
                            f'{ivr(IV_YM)},">="&{a},{ivr(IV_YM)},"<="&{b})')
    for i in range(PJ_R1 - PJ_R0 + 1):
        r = PS_R0 + i
        pr_ = PJ_R0 + i
        ws[f'{PS_NAME}{r}'] = f'={SH_PROJ}!${PJ_NAME}${pr_}&""'
        pj = f'${PS_NAME}{r}'
        g = lambda f: f'=IF({pj}="",0,{f})'
        for (m, p), c in PS_COL.items():
            a, b = per[p]
            f = {'确认收入': calc.rev(a, b, pj), '现金回款': calc.jline('应收账款', a, b, pj, sign=1), '抵账': OFS(a, b, pj),
                 '开票': INV(a, b, pj), '销项税': calc.vat_out(a, b, pj), '进项税': calc.vat_in(a, b, pj),
                 '材料': calc.line('材料费', a, b, pj), '人工': calc.line('人工费', a, b, pj), '分包': calc.line('分包费', a, b, pj),
                 '机械': calc.line('机械运输费', a, b, pj),
                 '其他直接': f'({calc.aplog("其他直接费", a, b, pj)}+{calc.jline("其他直接费", a, b, pj)})',
                 '甲方扣款': calc.deduct(a, b, pj), '贴息': calc.jline('票据贴息', a, b, pj)}[m]
            ws[f'{c}{r}'] = g(f)
        for k, c in enumerate(PS_LABM):
            ym = f'{AX_Y}*100+{k + 1}'
            ws[f'{c}{r}'] = f'=IF(OR({pj}="",{ym}>{AX_YM1}),0,{calc.line("人工费", f"MAX({ym},{AX_OYM})", ym, pj)})'
        B = lambda c: f'{c}{r}'
        for k, c in enumerate(PS_DRV):
            y = f'{c}$2'
            a, b = f'MAX({y}*100+1,{AX_OYM})', f'MIN({y}*100+12,{AX_YM1})'
            base = (f'{calc.line("人工费", a, b, pj)}+{calc.line("分包费", a, b, pj)}+{calc.line("机械运输费", a, b, pj)}')
            extra = f'{calc.line("材料费", a, b, pj)}+{calc.aplog("其他直接费", a, b, pj)}+{calc.jline("其他直接费", a, b, pj)}'
            ws[f'{c}{r}'] = f'=IF(OR({pj}="",{y}>{AX_Y}),0,{base}+IF({PA_DRV}="直接成本",{extra},0))'
        opn = f'({_open_proj(pj, OPJ_REV)}-{_open_proj(pj, OPJ_REC)})'
        ws[f'{PS_OPEN}{r}'] = f'=IF({pj}="",0,{opn})'
        pre = lambda m: B(PS_COL[(m, '前')])
        cur = lambda m: B(PS_COL[(m, '本')])
        ws[f'{PS_AR_B}{r}'] = f'=IF({pj}="",0,ROUND({B(PS_OPEN)}+{pre("确认收入")}-{pre("现金回款")}-{pre("抵账")},2))'
        ws[f'{PS_AR_E}{r}'] = (f'=IF({pj}="",0,ROUND({B(PS_AR_B)}+{cur("确认收入")}-{cur("现金回款")}-{cur("抵账")},2))')
    ws.sheet_state = 'hidden'
    return ws


def ps(meas_or_col, period=None):
    """_项目汇总 的整列（跟项目档案一行对一行）"""
    c = PS_COL[(meas_or_col, period)] if period else meas_or_col
    return f"{SH_PS}!${c}${PS_R0}:${c}${PS_R0 + PJ_R1 - PJ_R0}"


def ps_cell(meas_or_col, period, r):
    c = PS_COL[(meas_or_col, period)] if period else meas_or_col
    return f"{SH_PS}!${c}${r}"


def build_balx(wb, ctx):
    ws = wb.create_sheet(SH_BALX)
    ws['A1'] = '每个单位/人在截止月末和年初的余额、每个账户的余额（取数用，别改）'
    heads = {BX_NAME: '名称', BX_TYPE: '类型', BX_AP_E: '应付(截止)', BX_AP_B: '应付(年初)', BX_WG_E: '欠薪(截止)', BX_WG_B: '欠薪(年初)',
             BX_APAMT: '累计应付', BX_PAID: '累计已付', BX_INVIN: '累计收票', BX_NETPAY: '累计实发应付', BX_WPAID: '累计发放',
             BX_OFFW: '累计总包代发', BX_REFUND: '累计退回',
             BA_NAME: '账户', BA_TYPE: '类型', BA_OWNER: '个人户主人', BA_E: '余额(截止)', BA_B: '余额(年初)'}
    for c, t in heads.items():
        ws[f'{c}2'] = t
    for i in range(UN_R1 - UN_R0 + 1):
        r = BX_R0 + i
        u = UN_R0 + i
        ws[f'{BX_NAME}{r}'] = f'={SH_UNIT}!${UN_NAME}${u}&""'
        ws[f'{BX_TYPE}{r}'] = f'={SH_UNIT}!${UN_TYPE}${u}&""'
        n, t = f'${BX_NAME}{r}', f'${BX_TYPE}{r}'
        isap = f'OR({t}="材料供应商",{t}="分包",{t}="机械运输",{t}="其他")'
        isp = f'OR({t}="管理人员",{t}="工人")'
        apl = f'IF({t}="材料供应商","应付材料款",IF({t}="分包","应付分包款",IF({t}="机械运输","应付机械运输费","应付其他款")))'

        def apb(b):
            amt = (f'SUMIFS({oapr(OAP_AMT)},{oapr(OAP_UNIT)},{n})+SUMIFS({apr(AP_AMT)},{apr(AP_UNIT)},{n},'
                   f'{apr(AP_YM)},">="&{AX_OYM},{apr(AP_YM)},"<="&{b})')
            paid = (f'SUMIFS({oapr(OAP_PAID)},{oapr(OAP_UNIT)},{n})-SUMIFS({jr(J_NET)},{jr(J_UN)},{n},{jr(J_LINE)},{apl},'
                    f'{jr(J_YM)},">="&{AX_OYM},{jr(J_YM)},"<="&{b})+SUMIFS({ofr(OF_AMT)},{ofr(OF_WHO)},{n},{ofr(OF_TYPE)},"总包代付材料分包款",'
                    f'{ofr(OF_YM)},">="&{AX_OYM},{ofr(OF_YM)},"<="&{b})')
            return amt, paid
        amt, paid = apb(AX_YM1)
        ws[f'{BX_APAMT}{r}'] = f'=IF(OR({n}="",NOT({isap})),0,{amt})'
        ws[f'{BX_PAID}{r}'] = f'=IF(OR({n}="",NOT({isap})),0,{paid})'
        ws[f'{BX_AP_E}{r}'] = f'=ROUND({BX_APAMT}{r}-{BX_PAID}{r},2)'
        amt_b, paid_b = apb(AX_PYM)
        ws[f'{BX_AP_B}{r}'] = f'=IF(OR({n}="",NOT({isap})),0,ROUND({amt_b}-({paid_b}),2))'
        ws[f'{BX_INVIN}{r}'] = (f'=IF({n}="",0,SUMIFS({oapr(OAP_INV)},{oapr(OAP_UNIT)},{n})+SUMIFS({ivr(IV_TOTU)},{ivr(IV_DIR)},"进项",'
                                f'{ivr(IV_UNIT)},{n},{ivr(IV_YM)},">="&{AX_OYM},{ivr(IV_YM)},"<="&{AX_YM1}))')

        def wg(b):
            rngj = f'{jr(J_YM)},">="&{AX_OYM},{jr(J_YM)},"<="&{b}'
            return (f'SUMIFS({atr(AT_NETPAY)},{atr(AT_NAME)},{n},{atr(AT_YM)},">="&{AX_OYM},{atr(AT_YM)},"<="&{b})',
                    f'(-SUMIFS({jr(J_NET)},{jr(J_UN)},{n},{jr(J_LINE)},"应付职工薪酬",{jr(J_NET)},"<0",{rngj}))',
                    f'SUMIFS({ofr(OF_AMT)},{ofr(OF_WHO)},{n},{ofr(OF_TYPE)},"总包代发工资",{ofr(OF_YM)},">="&{AX_OYM},{ofr(OF_YM)},"<="&{b})',
                    f'SUMIFS({jr(J_NET)},{jr(J_UN)},{n},{jr(J_LINE)},"应付职工薪酬",{jr(J_NET)},">0",{rngj})')
        owe0 = f'N({SH_UNIT}!${UN_OWE0}${u})'
        np_, wp, ow, rf = wg(AX_YM1)
        ws[f'{BX_NETPAY}{r}'] = f'=IF(OR({n}="",NOT({isp})),0,{np_})'
        ws[f'{BX_WPAID}{r}'] = f'=IF(OR({n}="",NOT({isp})),0,{wp})'
        ws[f'{BX_OFFW}{r}'] = f'=IF(OR({n}="",NOT({isp})),0,{ow})'
        ws[f'{BX_REFUND}{r}'] = f'=IF(OR({n}="",NOT({isp})),0,{rf})'
        ws[f'{BX_WG_E}{r}'] = f'=IF(OR({n}="",NOT({isp})),0,ROUND({owe0}+{BX_NETPAY}{r}-{BX_WPAID}{r}-{BX_OFFW}{r}+{BX_REFUND}{r},2))'
        np_, wp, ow, rf = wg(AX_PYM)
        ws[f'{BX_WG_B}{r}'] = f'=IF(OR({n}="",NOT({isp})),0,ROUND({owe0}+{np_}-{wp}-{ow}+{rf},2))'
    for i in range(AC_R1 - AC_R0 + 1):
        r = BA_R0 + i
        a = AC_R0 + i
        ws[f'{BA_NAME}{r}'] = f'={SH_BASE}!${AC_NAME}${a}&""'
        ws[f'{BA_TYPE}{r}'] = f'={SH_BASE}!${AC_TYPE}${a}&""'
        ws[f'{BA_OWNER}{r}'] = f'={SH_BASE}!${AC_OWNER}${a}&""'
        n = f'${BA_NAME}{r}'

        def bal(b):
            rngj = f'{jr(J_YM)},">="&{AX_OYM},{jr(J_YM)},"<="&{b}'
            return (f'N({SH_BASE}!${AC_OPEN}${a})+SUMIFS({jr(J_NET)},{jr(J_ACC)},{n},{rngj})'
                    f'-SUMIFS({jr(J_NET)},{jr(J_TO)},{n},{jr(J_LINE)},"账户互转",{rngj})')
        ws[f'{BA_E}{r}'] = f'=IF({n}="",0,ROUND({bal(AX_YM1)},2))'
        ws[f'{BA_B}{r}'] = f'=IF({n}="",0,ROUND({bal(AX_PYM)},2))'
    ws.sheet_state = 'hidden'
    return ws


def bx(col):
    return f"{SH_BALX}!${col}${BX_R0}:${col}${BX_R0 + UN_R1 - UN_R0}"


def ba(col):
    return f"{SH_BALX}!${col}${BA_R0}:${col}${BA_R0 + AC_R1 - AC_R0}"
