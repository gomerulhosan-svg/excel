# -*- coding: utf-8 -*-
"""隐藏表 _汇：全书共用、而且必须「先逐个对象算、再加总」的数。
   ① 年度管理费率（待摊池 ÷ 分摊基数）——_收/_付/_考 的「摊费」列按它算；
   ② 日期表（年初前、1～12 月末、利润表本期、资产负债表日）；
   ③ 项目块：每个项目在每个日期的累计开票、收款、产值、确认收入（收入口径要逐项目取大）、应收余额；
   ④ 供应商块：应付余额（正＝应付、负＝预付）；⑤ 账户块：余额；⑥ 人员块：个人户＋报销还款（正＝他手上有公司的钱、负＝公司欠他）；
   ⑦ 结果格（资产负债表、利润表直接用名称取）。"""
from layout import *
from common import F_HELP, esc


def _h(ws, cell, text):
    ws[cell] = text
    ws[cell].font = F_HELP


def build(wb, ctx=None):
    ws = wb[H_SUM]
    Y = SUM_YCOLS
    # ① 年度费率
    _h(ws, 'A1', '① 年度管理费率（待摊池 ÷ 分摊基数；基数按【基本信息】分摊依据：施工费＝人工＋分包＋机械，直接成本＝全部项目成本）')
    for k, c in Y.items():
        _h(ws, f'{c}{SUM_Y0 - 1}', k)
    tans = '+'.join(f'SUMIFS(考_摊{j + 1},考_年月,">="&{{y0}},考_年月,"<="&{{y1}})' for j in range(NPAIR))
    for i in range(NYEARS):
        r = SUM_Y0 + i
        y = f'$A{r}'
        d0, d1 = f'DATE({y},1,1)', f'DATE({y},12,31)'
        ym0, ym1 = f'{y}*100+1', f'{y}*100+12'
        ws[f'{Y["年"]}{r}'] = YEAR0 + i
        ws[f'{Y["待摊池收"]}{r}'] = f'=-SUMIFS(收_净额,收_归属,"待摊费用",收_日期,">="&{d0},收_日期,"<="&{d1})'
        ws[f'{Y["待摊池考"]}{r}'] = f'=SUMIFS(考_未分摊,考_计提,1,考_年月,">="&{ym0},考_年月,"<="&{ym1})'
        ws[f'{Y["待摊池"]}{r}'] = f'=B{r}+C{r}'
        ws[f'{Y["基数"]}{r}'] = (f'=SUMIFS(收_基数,收_日期,">="&{d0},收_日期,"<="&{d1})'
                               f'+SUMIFS(付_基数,付_日期,">="&{d0},付_日期,"<="&{d1})'
                               f'+SUMIFS(考_基数,考_年月,">="&{ym0},考_年月,"<="&{ym1})')
        ws[f'{Y["费率"]}{r}'] = f'=IF(E{r}>0,D{r}/E{r},0)'
        ws[f'{Y["已摊"]}{r}'] = (f'=SUMIFS(收_摊费,收_日期,">="&{d0},收_日期,"<="&{d1})'
                               f'+SUMIFS(付_摊费,付_日期,">="&{d0},付_日期,"<="&{d1})+' + tans.format(y0=ym0, y1=ym1))
        ws[f'{Y["未摊"]}{r}'] = f'=ROUND(D{r}-G{r},2)'

    # ② 日期表
    _h(ws, f'A{SUM_D0 - 1}', '② 日期表（_汇 各块按这些日期算累计数）')
    for k, r in SUM_DROW.items():
        ws[f'A{r}'] = k
        if k == '年初前':
            v = '=P_年初-1'
        elif k.endswith('月') and k[:-1].isdigit():
            m = int(k[:-1])
            v = f'=MAX(P_年初-1,MIN(DATE(P_年度,{m + 1},0),P_截止))'
        elif k == '利润起前':
            v = f'={SUM_EFF["利润起"]}-1'
        elif k == '利润止':
            v = f'={SUM_EFF["利润止"]}'
        else:
            v = f'={SUM_EFF["负债日"]}'
        ws[f'B{r}'] = v
        ws[f'B{r}'].number_format = 'yyyy-mm-dd'
    _h(ws, 'A31', '利润表本期起')
    _h(ws, 'A32', '利润表本期止')
    _h(ws, 'A33', '资产负债表日')
    ws[SUM_EFF['利润起']] = f'=IF(ISNUMBER({PL_IN_START}),INT({PL_IN_START}),P_年初)'
    ws[SUM_EFF['利润止']] = f'=IF(ISNUMBER({PL_IN_END}),INT({PL_IN_END}),P_截止)'
    ws[SUM_EFF['负债日']] = f'=IF(ISNUMBER({BS_IN_DATE}),INT({BS_IN_DATE}),P_截止)'
    for c in SUM_EFF.values():
        ws[c].number_format = 'yyyy-mm-dd'

    # ③ 项目块
    _h(ws, f'A{SUM_PJ_HDR - 1}', '③ 项目：各日期的累计开票、工程款收款、确认产值、确认收入（按收入口径逐项目算）；应收余额＝期初＋收入−收款')
    _h(ws, f'A{SUM_PJ_HDR}', '项目')
    _h(ws, f'B{SUM_PJ_HDR}', '期初应收')
    for k, cols in SUM_PJ_K.items():
        for c, t in zip(cols, ('开票', '收款', '产值', '收入')):
            _h(ws, f'{c}{SUM_PJ_HDR}', f'{k}{t}')
    _h(ws, f'{SUM_PJ_AR0}{SUM_PJ_HDR}', '应收余额年初前')
    _h(ws, f'{SUM_PJ_AR1}{SUM_PJ_HDR}', '应收余额负债日')
    last = SUM_PJ0 + N_PJ - 1
    k0, k1 = SUM_PJ_K['年初前'], SUM_PJ_K['负债日']
    for i in range(N_PJ):
        r = SUM_PJ0 + i
        ws[f'A{r}'] = f'=INDEX(项目_名称,{i + 1})&""'
        a = f'$A{r}'
        ws[f'B{r}'] = f'=IF({a}="",0,SUMIFS(期初_金额,期初_类型,"应收账款",期初_项目,{esc(a)}))'
        for k, (ckp, csk, ccz, crv) in SUM_PJ_K.items():
            d = f'$B${SUM_DROW[k]}'
            ws[f'{ckp}{r}'] = f'=IF({a}="",0,{cum_kp(a, d)})'
            ws[f'{csk}{r}'] = f'=IF({a}="",0,{cum_sk(a, d)})'
            ws[f'{ccz}{r}'] = f'=IF({a}="",0,{cum_cz(a, d)})'
            ws[f'{crv}{r}'] = f'={rev_formula(f"{ckp}{r}", f"{csk}{r}", f"{ccz}{r}")}'
        ws[f'{SUM_PJ_AR0}{r}'] = f'=B{r}+{k0[3]}{r}-{k0[1]}{r}'
        ws[f'{SUM_PJ_AR1}{r}'] = f'=B{r}+{k1[3]}{r}-{k1[1]}{r}'
    rl, ra, rn, rc = SUM_PJ_LIST, SUM_PJ_ALL, SUM_PJ_NONE, SUM_PJ_CO
    ws[f'A{rl}'], ws[f'A{ra}'], ws[f'A{rn}'], ws[f'A{rc}'] = '名单合计', '全部', '未指定项目', '公司合计'
    ws[f'B{rl}'] = f'=SUM(B{SUM_PJ0}:B{last})'
    ws[f'B{ra}'] = '=SUMIFS(期初_金额,期初_类型,"应收账款")'
    ws[f'B{rn}'] = f'=B{ra}-B{rl}'
    for k, (ckp, csk, ccz, crv) in SUM_PJ_K.items():
        d = f'$B${SUM_DROW[k]}'
        for c in (ckp, csk, ccz, crv):
            ws[f'{c}{rl}'] = f'=SUM({c}{SUM_PJ0}:{c}{last})'
        ws[f'{ckp}{ra}'] = f'=SUMIFS(应_开票额,应_日期,"<="&{d})'
        ws[f'{csk}{ra}'] = f'=SUMIFS(收_净额,收_归类,"工程款收款",收_日期,"<="&{d})'
        ws[f'{ccz}{ra}'] = f'=SUMIFS(应_产值额,应_日期,"<="&{d})'
        for c in (ckp, csk, ccz):
            ws[f'{c}{rn}'] = f'=ROUND({c}{ra}-{c}{rl},2)'
        ws[f'{crv}{rn}'] = f'={rev_formula(f"{ckp}{rn}", f"{csk}{rn}", f"{ccz}{rn}")}'
        ws[f'{crv}{rc}'] = f'={crv}{rl}+{crv}{rn}'
    ws[f'{SUM_PJ_AR0}{rn}'] = f'=B{rn}+{k0[3]}{rn}-{k0[1]}{rn}'
    ws[f'{SUM_PJ_AR1}{rn}'] = f'=B{rn}+{k1[3]}{rn}-{k1[1]}{rn}'

    def pos_neg(col, r0, r1, rnone):
        rng = f'{col}{r0}:{col}{r1}'
        return (f'=SUMIF({rng},">0")+MAX(0,{col}{rnone})', f'=-SUMIF({rng},"<0")-MIN(0,{col}{rnone})')

    R = SUM_RROW
    res = {}
    res['应收年初前'], res['预收年初前'] = pos_neg(SUM_PJ_AR0, SUM_PJ0, last, rn)
    res['应收负债日'], res['预收负债日'] = pos_neg(SUM_PJ_AR1, SUM_PJ0, last, rn)

    # ④ 供应商块
    s0, s1 = SUM_SUP0, SUM_SUP0 + N_SUP - 1
    _h(ws, f'A{SUM_SUP_HDR - 1}', '④ 供应商：应付余额＝期初＋应付登记−付款（收支登记里冲应付的）；正＝应付，负＝预付')
    for c, t in zip('ABCD', ('供应商', '期初应付', '余额年初前', '余额负债日')):
        _h(ws, f'{c}{SUM_SUP_HDR}', t)
    for i in range(N_SUP):
        r = s0 + i
        a = f'$A{r}'
        ws[f'A{r}'] = f'=INDEX(供应商_名称,{i + 1})&""'
        ws[f'B{r}'] = f'=IF({a}="",0,SUMIFS(期初_金额,期初_类型,"应付账款",期初_对象,{esc(a)}))'
        for c, k in (('C', '年初前'), ('D', '负债日')):
            d = f'$B${SUM_DROW[k]}'
            ws[f'{c}{r}'] = (f'=IF({a}="",0,B{r}+SUMIFS(付_应付额,付_供应商,{esc(a)},付_日期,"<="&{d})'
                             f'+SUMIFS(收_净额,收_供应商,{esc(a)},收_冲应付,1,收_日期,"<="&{d}))')
    rl, ra, rn = SUM_SUP_LIST, SUM_SUP_ALL, SUM_SUP_NONE
    ws[f'A{rl}'], ws[f'A{ra}'], ws[f'A{rn}'] = '名单合计', '全部', '不在名单里的'
    for c in 'BCD':
        ws[f'{c}{rl}'] = f'=SUM({c}{s0}:{c}{s1})'
        ws[f'{c}{rn}'] = f'=ROUND({c}{ra}-{c}{rl},2)'
    ws[f'B{ra}'] = '=SUMIFS(期初_金额,期初_类型,"应付账款")'
    for c, k in (('C', '年初前'), ('D', '负债日')):
        d = f'$B${SUM_DROW[k]}'
        ws[f'{c}{ra}'] = f'=B{ra}+SUMIFS(付_应付额,付_日期,"<="&{d})+SUMIFS(收_净额,收_冲应付,1,收_日期,"<="&{d})'
    res['应付年初前'], res['预付年初前'] = pos_neg('C', s0, s1, rn)
    res['应付负债日'], res['预付负债日'] = pos_neg('D', s0, s1, rn)

    # ⑤ 账户块
    a0, a1 = SUM_ACC0, SUM_ACC0 + N_ACC - 1
    _h(ws, f'A{SUM_ACC_HDR - 1}', '⑤ 账户余额＝期初＋本账户净额−（内部转账里「对方账户」是它的那些行的净额）')
    for c, t in zip('ABCDE', ('账户', '类型', '所属人', '余额年初前', '余额负债日')):
        _h(ws, f'{c}{SUM_ACC_HDR}', t)
    for i in range(N_ACC):
        r = a0 + i
        a = f'$A{r}'
        ws[f'A{r}'] = f'=INDEX(账户_名称,{i + 1})&""'
        ws[f'B{r}'] = f'=INDEX(账户_类型,{i + 1})&""'
        ws[f'C{r}'] = f'=INDEX(账户_所属人,{i + 1})&""'
        for c, k in (('D', '年初前'), ('E', '负债日')):
            d = f'$B${SUM_DROW[k]}'
            ws[f'{c}{r}'] = (f'=IF({a}="",0,ROUND(INDEX(账户_期初,{i + 1})+SUMIFS(收_净额,收_账户,{esc(a)},收_日期,"<="&{d})'
                             f'-SUMIFS(收_净额,收_对方账户,{esc(a)},收_归类,"内部转账",收_日期,"<="&{d}),2))')
    rng = lambda c: f'{c}{a0}:{c}{a1}'
    res['货币资金年初前'] = f'=SUM({rng("D")})-SUMIFS({rng("D")},{rng("B")},"个人户")'
    res['货币资金负债日'] = f'=SUM({rng("E")})-SUMIFS({rng("E")},{rng("B")},"个人户")'
    res['个人户余额年初前'] = f'=SUMIFS({rng("D")},{rng("B")},"个人户")'
    res['个人户余额负债日'] = f'=SUMIFS({rng("E")},{rng("B")},"个人户")'

    # ⑥ 人员块（垫付/持有）
    p0, p1 = SUM_PER0, SUM_PER0 + N_PER - 1
    _h(ws, f'A{SUM_PER_HDR - 1}', '⑥ 个人：他名下个人户余额合计＋公司还给他的报销款−期初公司欠他的垫付；正＝他手上有公司的钱，负＝公司还欠他')
    for c, t in zip('ABCD', ('姓名', '期初垫付', '年初前', '负债日')):
        _h(ws, f'{c}{SUM_PER_HDR}', t)
    for i in range(N_PER):
        r = p0 + i
        a = f'$A{r}'
        ws[f'A{r}'] = f'=INDEX(人员_姓名,{i + 1})&""'
        ws[f'B{r}'] = f'=IF({a}="",0,SUMIFS(期初_金额,期初_类型,"个人垫付",期初_对象,{esc(a)}))'
        for c, k, bc in (('C', '年初前', 'D'), ('D', '负债日', 'E')):
            d = f'$B${SUM_DROW[k]}'
            ws[f'{c}{r}'] = (f'=IF({a}="",0,SUMIFS({rng(bc)},{rng("B")},"个人户",{rng("C")},{esc(a)})'
                             f'-SUMIFS(收_净额,收_归类,"报销还款",收_人员,{esc(a)},收_日期,"<="&{d})-B{r})')
    rl, ra, rn = SUM_PER_LIST, SUM_PER_ALL, SUM_PER_NONE
    ws[f'A{rl}'], ws[f'A{ra}'], ws[f'A{rn}'] = '名单合计', '全部', '没对上人的'
    for c in 'BCD':
        ws[f'{c}{rl}'] = f'=SUM({c}{p0}:{c}{p1})'
        ws[f'{c}{rn}'] = f'=ROUND({c}{ra}-{c}{rl},2)'
    ws[f'B{ra}'] = '=SUMIFS(期初_金额,期初_类型,"个人垫付")'
    for c, k, bc in (('C', '年初前', 'D'), ('D', '负债日', 'E')):
        d = f'$B${SUM_DROW[k]}'
        ws[f'{c}{ra}'] = (f'=SUMIFS({rng(bc)},{rng("B")},"个人户")'
                          f'-SUMIFS(收_净额,收_归类,"报销还款",收_日期,"<="&{d})-B{ra}')
    res['个人持有年初前'], res['个人垫付年初前'] = pos_neg('C', p0, p1, rn)
    res['个人持有负债日'], res['个人垫付负债日'] = pos_neg('D', p0, p1, rn)

    # ⑦ 结果
    _h(ws, f'K{SUM_RES0 - 1}', '⑦ 结果（报表用名称 汇_xxx 取）')
    for k, r in R.items():
        ws[f'K{r}'] = k
        ws[f'L{r}'] = res[k]
        ws[f'L{r}'].number_format = '#,##0.00'
