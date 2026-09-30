# -*- coding: utf-8 -*-
"""老板看的三张报表：【利润表】（按月）【资产负债表】（截止月末 ＋ 年初）【盈亏平衡表】（保本产值、项目还能花多少、报价计算器）"""
from openpyxl.formatting.rule import FormulaRule
from common import *
from layout import *
import calc
import cats
import s_agg
from s_agg import ms
from s_views import _sheet, _lbl, _val, GREY
import s_proj as sp

POOL = sp.POOL
DIRECT = sp.DIRECT


# ═══════════════════════════ 利润表 ═══════════════════════════
IS_ROWS = []          # (标签, 类型, 公式生成) ；类型：line / sub / calc / pct / head


def build_is(wb, ctx):
    MC = [CL(2 + i) for i in range(12)]       # B..M
    C_YTD, C_PCT, C_PREV = 'N', 'O', 'P'
    ws = _sheet(wb, SH_IS, C_RPT, '利 润 表（管理用 · 按月 · 收入按甲方确认的产值）',
                '💡 全自动，年度和截止日在首页选。收入＝【收入确认】里甲方确认的产值（含税，不是收到的钱）；直接成本按发生（挂账没付的也算）；'
                '税金＝按产值估算的增值税（扣了进项专票）＋附加＋印花；间接费用＝公司日常开支和管理人员在公司那部分工资，在【费用分摊】摊到项目。'
                '这是给老板看赚没赚钱的管理报表，报税用的报表以代账会计的为准。',
                C_PREV, {'A': 30, **{c: 12 for c in MC}, C_YTD: 14, C_PCT: 8, C_PREV: 14})
    put(ws, 'A3', f'={AX_Y}&"年  截止 "&TEXT({AX_E},"m月d日")&"（灰色的月份还没到）"', F_KPI_L, align=AL, border=False)
    H = 5
    header(ws, H, [('A', '项  目'), (C_YTD, '本年累计'), (C_PCT, '占收入'), (C_PREV, '以前年度\n(建账起)')], C_RPT, height=36)
    for i, c in enumerate(MC):
        put(ws, f'{c}{H}', f'{i + 1}月', F_HDR, fill(C_RPT), align=AC)
    cols = MC + [C_YTD, C_PREV]
    src = dict(zip(MC, MS_MCOLS))
    src[C_YTD], src[C_PREV] = MS_YTD, MS_PREV
    rows = [('一、营业收入（确认产值）', 'line', ['营业收入'], 1),
            ('二、直接成本', 'sub', DIRECT, 1),
            *[(f'　　{n}', 'line', [n], 2) for n in DIRECT],
            ('三、税金（估算）', 'line', ['税金估算'], 1),
            ('四、票据贴息', 'line', ['票据贴息'], 1),
            ('项目毛利（一 − 二 − 三 − 四）', 'calc', None, 0),
            ('五、间接费用（管理费用）', 'sub', POOL, 1),
            *[(f'　　{n}', 'line', [n], 2) for n in POOL],
            ('六、财务费用', 'sub', ['利息支出', '手续费', '利息收入'], 1),
            ('　　利息支出', 'line', ['利息支出'], 2), ('　　手续费', 'line', ['手续费'], 2), ('　　利息收入（减）', 'line', ['利息收入'], 2),
            ('七、其他收支（净收入）', 'calc2', None, 1),
            ('　　其他收入', 'line', ['其他收入'], 2), ('　　其他支出', 'line', ['其他支出'], 2), ('　　其他税费', 'line', ['其他税费'], 2),
            ('利润总额', 'calc', None, 0),
            ('减：所得税', 'line', ['所得税'], 1),
            ('净利润', 'calc', None, 0),
            ('净利率', 'pct', None, 0)]
    R = {}
    r = H + 1
    for lab, kind, names, lvl in rows:
        R[lab] = r
        r += 1
    for lab, kind, names, lvl in rows:
        r = R[lab]
        bold = lvl < 2
        fl = FILL_TOT if lvl == 0 else (FILL_SUB if kind in ('sub', 'calc2') else None)
        put(ws, f'A{r}', lab, F_TXTB if bold else F_TXT, fl, align=AL)
        for c in cols:
            if kind in ('line', 'sub'):
                f = '+'.join(ms(n, src[c]) for n in names)
            elif kind == 'calc2':
                f = f'{c}{R["　　其他收入"]}-{c}{R["　　其他支出"]}-{c}{R["　　其他税费"]}'
            elif lab.startswith('项目毛利'):
                f = f'{c}{R["一、营业收入（确认产值）"]}-{c}{R["二、直接成本"]}-{c}{R["三、税金（估算）"]}-{c}{R["四、票据贴息"]}'
            elif lab == '利润总额':
                f = (f'{c}{R["项目毛利（一 − 二 − 三 − 四）"]}-{c}{R["五、间接费用（管理费用）"]}-{c}{R["六、财务费用"]}'
                     f'+{c}{R["七、其他收支（净收入）"]}')
            elif lab == '净利润':
                f = f'{c}{R["利润总额"]}-{c}{R["减：所得税"]}'
            else:
                f = f'IF(N({c}{R["一、营业收入（确认产值）"]})=0,"",{c}{R["净利润"]}/{c}{R["一、营业收入（确认产值）"]})'
            put(ws, f'{c}{r}', f'={f}', F_AUTOB if bold else F_AUTO, fl, PCT if kind == 'pct' else MONEY, AR)
        if kind != 'pct':
            put(ws, f'{C_PCT}{r}', f'=IF(N({C_YTD}${R["一、营业收入（确认产值）"]})=0,"",{C_YTD}{r}/{C_YTD}${R["一、营业收入（确认产值）"]})',
                F_NOTE, fl, PCT, AC)
        else:
            put(ws, f'{C_PCT}{r}', None, fill_=fl)
    last = R['净利率']
    ws.conditional_formatting.add(f'B{H}:M{last}', FormulaRule(formula=[f'{AX_Y}*100+COLUMN(B{H})-1>{AX_YM1}'], font=GREY))
    for lab in ('净利润', '项目毛利（一 − 二 − 三 − 四）', '利润总额'):
        rr = R[lab]
        ws.conditional_formatting.add(f'B{rr}:{C_PREV}{rr}', FormulaRule(formula=[f'N(B{rr})<0'], font=F_RED))
    n0 = last + 2
    notes = ['说明：',
             '· 「项目毛利」＝各项目干活赚的（还没扣公司管理费）；每个项目的明细看【项目利润表】，两边对得上（那张表最下面有对账）。',
             '· 「税金」是按产值估算的，实际交的增值税、附加在流水里冲「应交税费」，不重复算成本；企业所得税单独在最下面。',
             '· 管理人员工资＝【考勤工资】里月薪的人在「公司管理」那部分；在工地的天数已经算进项目人工费了。',
             '· 以前年度的数在最右边一列（建账那年起累计），想看去年的就到首页把年度改成去年。']
    for i, t in enumerate(notes):
        put(ws, f'A{n0 + i}', t, F_NOTE if i else F_TXTB, align=AL, border=False)
        ws.merge_cells(f'A{n0 + i}:{C_PREV}{n0 + i}')
    ws.freeze_panes = f'B{H + 1}'
    global IS_NET_YTD
    IS_NET_YTD = f"{SH_IS}!${C_YTD}${R['净利润']}"
    return ws


IS_NET_YTD = None


# ═══════════════════════════ 资产负债表 ═══════════════════════════
BS_DIFF = None       # 差额格（数据校验用）


def _bs_parts(which):
    """which＝'E' 截止月末 / 'B' 年初。返回各项目的公式（字符串，不带 =）"""
    b = AX_YM1 if which == 'E' else AX_PYM
    bfa = AX_YM1 if which == 'E' else f'IF({AX_PYM}<{AX_OYM},{AX_OYM}-1,{AX_PYM})'   # 年初在建账日之前：固定资产取建账日那天的
    arc, apc, wgc, bac = (PS_AR_E, BX_AP_E, BX_WG_E, BA_E) if which == 'E' else (PS_AR_B, BX_AP_B, BX_WG_B, BA_B)
    yr = f'{jr(J_YM)},">="&{AX_OYM},{jr(J_YM)},"<="&{b}'
    JN = lambda line: f'SUMIFS({jr(J_NET)},{jr(J_LINE)},"{line}",{yr})'
    ofr_ = lambda t: f'SUMIFS({ofr(OF_AMT)},{ofr(OF_TYPE)},"{t}",{ofr(OF_YM)},">="&{AX_OYM},{ofr(OF_YM)},"<="&{b})'
    ar_ps, ap_bx, wg_bx, ba_ = s_agg.ps(arc), s_agg.bx(apc), s_agg.bx(wgc), s_agg.ba(bac)
    p = {}
    ba_t = s_agg.ba(BA_TYPE)
    p['票据'] = f'SUMIFS({ba_},{ba_t},"票据")'
    p['个人正'] = f'SUMIFS({ba_},{ba_t},"个人户",{ba_},">0")'
    p['个人负'] = f'-SUMIFS({ba_},{ba_t},"个人户",{ba_},"<0")'
    p['货币资金'] = f'SUM({ba_})-SUMIFS({ba_},{ba_t},"票据")-SUMIFS({ba_},{ba_t},"个人户")'
    ar_tot = (f'SUM({opr(OPJ_REV)})-SUM({opr(OPJ_REC)})+{calc.rev(AX_OYM, b)}-{JN("应收账款")}'
              f'-({"+".join(ofr_(t) for t in OF_TYPES)})')
    p['应收残'] = f'({ar_tot})-SUM({ar_ps})'
    p['应收账款'] = f'SUMIF({ar_ps},">0")+MAX(0,{p["应收残"]})'
    p['预收账款'] = f'-SUMIF({ar_ps},"<0")+MAX(0,-({p["应收残"]}))'
    ap_lines = ['应付材料款', '应付分包款', '应付机械运输费', '应付其他款']
    ap_tot = (f'SUM({oapr(OAP_AMT)})-SUM({oapr(OAP_PAID)})+' + '+'.join(calc.aplog(t, AX_OYM, b) for t in AP_TYPES)
              + '+' + '+'.join(JN(l) for l in ap_lines) + f'-{ofr_("总包代付材料分包款")}')
    p['应付残'] = f'({ap_tot})-SUM({ap_bx})'
    p['应付账款'] = f'SUMIF({ap_bx},">0")+MAX(0,{p["应付残"]})'
    p['预付账款'] = f'-SUMIF({ap_bx},"<0")+MAX(0,-({p["应付残"]}))'
    wg_tot = (f'SUM({rng(SH_UNIT, UN_OWE0, UN_R0, UN_R1)})+SUMIFS({atr(AT_NETPAY)},{atr(AT_YM)},">="&{AX_OYM},{atr(AT_YM)},"<="&{b})'
              f'+{JN("应付职工薪酬")}-{ofr_("总包代发工资")}')
    p['工资残'] = f'({wg_tot})-SUM({wg_bx})'
    p['应付职工薪酬'] = f'SUMIF({wg_bx},">0")+MAX(0,{p["工资残"]})'
    p['多发工资'] = f'-SUMIF({wg_bx},"<0")+MAX(0,-({p["工资残"]}))'
    tax = f'N({oo(3)})+{calc.tax_est(AX_OYM, b)}+{calc.att_sum(AT_TAX, AX_OYM, b)}+{JN("应交税费")}'
    p['税'] = tax
    p['应交税费'] = f'MAX(0,{tax})'
    p['预缴税费'] = f'MAX(0,-({tax}))'
    p['短期借款'] = f'N({oo(0)})+{JN("短期借款")}'
    p['保证金'] = f'N({oo(1)})-{JN("其他应收款")}'
    p['其他应收别的'] = f'N({oo(2)})'
    p['其他应付别的'] = f'N({oo(4)})'
    idxb = f'({calc._idx(bfa)})'
    FAC, FAS, FAD = (rng(SH_BASE, c, FA_R0, FA_R1) for c in (FA_COST, FA_S, FA_DATE))
    p['固定资产原值'] = f'SUMIFS({FAC},{FAS},">0",{FAS},"<="&({idxb}+1))'
    p['累计折旧'] = calc.fa_dep('190001', bfa)
    p['应付设备款'] = f'N({oo(6)})+SUMIFS({FAC},{FAS},">0",{FAS},"<="&({idxb}+1),{FAD},">="&{OPEN_DATE})+{JN("应付设备款")}'
    p['实收资本'] = f'N({oo(5)})+{JN("实收资本")}'
    allj = f'SUMIFS({jr(J_NET)},{yr})'
    p['待查'] = (f'{JN("待分类")}+({JN("账户互转")}-SUMIFS({jr(J_NET)},{jr(J_LINE)},"账户互转",{jr(J_TOTYPE)},"?*",{yr}))'
                f'-({allj}-SUMIFS({jr(J_NET)},{jr(J_ATYPE)},"?*",{yr}))')
    prof = lambda col: (f'{ms("营业收入", col)}+{ms("其他收入", col)}-('
                        + '+'.join(ms(n, col) for n in cats.IS_NAMES if n not in ('营业收入', '其他收入')) + ')')
    p['利润'] = f'({prof(MS_PREV)})' + (f'+({prof(MS_YTD)})' if which == 'E' else '')
    return p


def build_bs(wb, ctx):
    ws = _sheet(wb, SH_BS, C_RPT, '资 产 负 债 表（管理用 · 截止月末 ＋ 年初）',
                '💡 全自动：钱在哪（银行、甲方欠的、多付的、备用金），欠谁的（材料商、分包、工人工资、税、老板），公司净值多少。'
                '应收、应付、工资都是一个一个项目/单位/人算了再分正负：甲方多付的算预收、多付给材料商的算预付、工人多发的算其他应收、老板个人户负数算欠老板的。'
                '「期初未分配利润」是按建账日的数倒算的（资产−负债−实收资本），下面可以跟老板/会计认的数对一对。最下面「差额」是 0 就说明账是平的。',
                'H', {'A': 34, 'B': 15, 'C': 15, 'D': 2, 'E': 34, 'F': 15, 'G': 15, 'H': 2})
    put(ws, 'A3', f'="截止 "&TEXT({AX_END},"yyyy年m月d日")&"（截止月末）；年初＝"&IF({AX_Y}={AX_Y0},"建账日 "&TEXT({OPEN_DATE},"yyyy年m月d日"),{AX_Y}&"年1月1日")',
        F_KPI_L, align=AL, border=False)
    ws.merge_cells('A3:G3')
    H = 5
    header(ws, H, [('A', '资  产'), ('B', '截止月末'), ('C', '年初'), ('E', '负债和所有者权益'), ('F', '截止月末'), ('G', '年初')], C_RPT, height=30)
    PE, PB = _bs_parts('E'), _bs_parts('B')
    # 期初未分配利润（倒算）
    FAC, FAD, FAS = (rng(SH_BASE, c, FA_R0, FA_R1) for c in (FA_COST, FA_DATE, FA_S))
    plug = (f'SUMIFS({AC_OPENS},{AC_NAMES},"?*")+SUM({opr(OPJ_REV)})-SUM({opr(OPJ_REC)})'
            f'+SUMIFS({FAC},{FAS},">0",{FAD},"<"&{OPEN_DATE})-{calc.fa_dep("190001", f"({AX_OYM}-1)")}'
            f'-(SUM({oapr(OAP_AMT)})-SUM({oapr(OAP_PAID)}))-SUM({rng(SH_UNIT, UN_OWE0, UN_R0, UN_R1)})'
            f'-N({oo(0)})+N({oo(1)})+N({oo(2)})-N({oo(3)})-N({oo(4)})-N({oo(5)})-N({oo(6)})')
    left = [('货币资金（银行＋现金）', 'v', '货币资金'), ('应收票据（在手的承兑汇票）', 'v', '票据'), ('应收账款（甲方还欠的）', 'v', '应收账款'),
            ('预付账款（多付给材料商、分包的）', 'v', '预付账款'), ('其他应收款', 'sum', ['个人正', '保证金', '多发工资', '其他应收别的']),
            ('　其中：老板/负责人手上的备用金', 'v2', '个人正'), ('　　　　保证金、押金', 'v2', '保证金'), ('　　　　工人多发的工资', 'v2', '多发工资'),
            ('　　　　其他', 'v2', '其他应收别的'), ('预缴税费（多交的税）', 'v', '预缴税费'), ('固定资产原值', 'v', '固定资产原值'),
            ('减：累计折旧', 'v', '累计折旧'), ('固定资产净值', 'fa', None), ('', 'blank', None), ('', 'blank', None), ('', 'blank', None),
            ('资产合计', 'tot', None)]
    right = [('短期借款', 'v', '短期借款'), ('应付账款（欠材料商、分包、机械的）', 'v', '应付账款'), ('预收账款（甲方多付的）', 'v', '预收账款'),
             ('应付职工薪酬（欠工人的工资）', 'v', '应付职工薪酬'), ('应交税费（估算还欠的税）', 'v', '应交税费'), ('应付设备款', 'v', '应付设备款'),
             ('其他应付款', 'sum', ['个人负', '其他应付别的']), ('　其中：欠老板/负责人的', 'v2', '个人负'), ('　　　　其他', 'v2', '其他应付别的'),
             ('待查（流水里没分类、没选对账户的）', 'v', '待查'), ('负债合计', 'ltot', None), ('实收资本', 'v', '实收资本'),
             ('未分配利润', 're', None), ('　其中：期初（建账日，倒算）', 'plug', None), ('　　　　建账以来的利润', 'v2', '利润'),
             ('所有者权益合计', 'etot', None), ('负债和所有者权益合计', 'all', None)]
    r0 = H + 1
    rows_l = {lab: r0 + i for i, (lab, _, _) in enumerate(left)}
    rows_r = {lab: r0 + i for i, (lab, _, _) in enumerate(right)}
    for side, spec, rows_, (cl, ce, cb) in ((0, left, rows_l, ('A', 'B', 'C')), (1, right, rows_r, ('E', 'F', 'G'))):
        for i, (lab, kind, key) in enumerate(spec):
            r = r0 + i
            bold = kind in ('tot', 'ltot', 'etot', 'all', 'sum', 're', 'fa')
            fl = FILL_TOT if kind in ('tot', 'all') else (FILL_SUB if kind in ('ltot', 'etot') else None)
            put(ws, f'{cl}{r}', lab or None, F_TXTB if bold else (F_NOTE if kind in ('v2', 'plug') else F_TXT), fl, align=AL)
            for c, P in ((ce, PE), (cb, PB)):
                if kind in ('v', 'v2'):
                    f = f'=ROUND({P[key]},2)'
                elif kind == 'sum':
                    subs = [r + 1 + j for j in range(len(key))]
                    f = '=' + '+'.join(f'{c}{s}' for s in subs)
                elif kind == 'fa':
                    f = f'={c}{rows_["固定资产原值"]}-{c}{rows_["减：累计折旧"]}'
                elif kind == 'tot':
                    items = [rows_[l] for l, k, _ in left if k in ('v', 'sum') and l not in ('固定资产原值', '减：累计折旧')] + [rows_['固定资产净值']]
                    f = '=' + '+'.join(f'{c}{x}' for x in items)
                elif kind == 'ltot':
                    items = [rows_[l] for l, k, _ in right[:10] if k in ('v', 'sum')]
                    f = '=' + '+'.join(f'{c}{x}' for x in items)
                elif kind == 're':
                    f = f'={c}{r + 1}+{c}{r + 2}'
                elif kind == 'plug':
                    f = f'=ROUND({plug},2)'
                elif kind == 'etot':
                    f = f'={c}{rows_["实收资本"]}+{c}{rows_["未分配利润"]}'
                elif kind == 'all':
                    f = f'={c}{rows_["负债合计"]}+{c}{rows_["所有者权益合计"]}'
                else:
                    f = None
                put(ws, f'{c}{r}', f, F_AUTOB if bold else (F_NOTE if kind in ('v2', 'plug') else F_AUTO), fl, MONEY, AR)
    last = r0 + len(left) - 1
    d = last + 2
    put(ws, f'A{d}', '差额（资产合计 − 负债和所有者权益合计，应该是 0）', F_TXTB, FILL_TOT, align=AL)
    for c, (ca, cr) in (('B', ('B', 'F')), ('C', ('C', 'G'))):
        put(ws, f'{c}{d}', f'=ROUND({ca}{rows_l["资产合计"]}-{cr}{rows_r["负债和所有者权益合计"]},2)', F_AUTOB, FILL_TOT, MONEY, AR)
    ws.conditional_formatting.add(f'B{d}:C{d}', FormulaRule(formula=[f'ABS(N(B{d}))>0.01'], fill=FILL_WARN, font=F_RED))
    global BS_DIFF, BS_DIFF_B, BS_PEND, BS_PLUG, BS_ROWS
    BS_DIFF = f"{SH_BS}!$B${d}"
    BS_DIFF_B = f"{SH_BS}!$C${d}"
    BS_ROWS = {**{k: ('B', v) for k, v in rows_l.items()}, **{k: ('F', v) for k, v in rows_r.items()}}
    BS_PEND = f"{SH_BS}!$F${rows_r['待查（流水里没分类、没选对账户的）']}"
    # 期初未分配利润：跟老板/会计认的数对比
    k = d + 2
    section(ws, k, 'A', 'G', '期初未分配利润：倒算的 vs 老板/会计认的', C_RPT)
    _lbl(ws, f'A{k + 1}', '倒算的（建账日资产 − 负债 − 实收资本）')
    _val(ws, f'B{k + 1}', f'=F{rows_r["　其中：期初（建账日，倒算）"]}')
    _lbl(ws, f'A{k + 2}', '老板/会计认的（【期初余额】③ 最后一项）')
    _val(ws, f'B{k + 2}', f'=IF({oo(OO_BOSS)}="","没填",{oo(OO_BOSS)})')
    _lbl(ws, f'A{k + 3}', '差多少')
    _val(ws, f'B{k + 3}', f'=IF(ISNUMBER(B{k + 2}),ROUND(B{k + 1}-B{k + 2},2),"")')
    put(ws, f'C{k + 1}', '差得多：说明建账日的数没填全，常见的是银行期初余额、老板借给公司的钱、欠材料商的、甲方欠的、工人期初欠薪。',
        F_NOTE, align=ALW, border=False)
    ws.merge_cells(f'C{k + 1}:G{k + 3}')
    BS_PLUG = f"{SH_BS}!$B${k + 3}"
    n0 = k + 5
    notes = ['怎么看：',
             '· 货币资金只算银行和现金；老板个人户不算公司的钱：他替公司垫了钱（个人户是负数）算「欠老板/负责人的」，他手上还有公司的备用金算「其他应收」。',
             '· 应收账款＝甲方确认的产值 − 收到的工程款 − 总包代发/代付/扣款；每个项目算完，收多了的放到「预收账款」。',
             '· 应交税费是按产值估算的税减去已经交的，报税以税务局的为准；负数（多交了、异地预缴多了）放在左边「预缴税费」。',
             '· 「待查」不是 0：去【资金流水】看 ✗ 的行（没认出类别、没选账户、账户互转没选对方账户）。']
    for i, t in enumerate(notes):
        put(ws, f'A{n0 + i}', t, F_NOTE if i else F_TXTB, align=AL, border=False)
        ws.merge_cells(f'A{n0 + i}:G{n0 + i}')
    ws.freeze_panes = f'A{H + 1}'
    return ws


BS_PEND = BS_PLUG = BS_DIFF_B = BS_ROWS = None


# ═══════════════════════════ 盈亏平衡表 ═══════════════════════════
BE_ROWS = 60


def build_be(wb, ctx):
    ws = _sheet(wb, SH_BE, C_RPT, '盈 亏 平 衡 表（一年要做多少产值才不亏 · 在建项目还能花多少 · 报价计算器）',
                '💡 ① 公司保本：一年的固定开支（管理费、利息等）÷（1 − 变动成本率）＝保本产值；变动成本率默认用已完工项目的（材料人工分包机械税金占产值的比例），'
                '没有完工项目就用本年的，也可以在黄格子手填。② 在建项目：总价扣掉税金和管理费，还能花多少直接成本不亏。③ 报价计算器：填预算成本，算保本价和建议报价。',
                'J', {'A': 5, 'B': 30, 'C': 15, 'D': 15, 'E': 14, 'F': 14, 'G': 14, 'H': 14, 'I': 9, 'J': 26})
    put(ws, 'B3', f'={AX_Y}&"年  截止 "&TEXT({AX_E},"m月d日")&"（已过 "&{AX_MON}&" 个月）"', F_KPI_L, align=AL, border=False)
    Y = lambda n: ms(n, MS_YTD)
    PPL = lambda c: f'{SH_PPL}!${c}${sp.PPL_R0}:${c}${sp.PPL_R0 + sp.NPJ - 1}'
    done = f'(({PPL(sp.P_STAT)}="完工未结算")+({PPL(sp.P_STAT)}="已结算")+({PPL(sp.P_STAT)}="质保期")+({PPL(sp.P_STAT)}="已完结"))'
    var_done = ('+'.join(f'SUMPRODUCT({done},{PPL(c)})' for c in (sp.C_MAT, sp.C_LAB, sp.C_SUB, sp.C_MACH, sp.C_OTH, sp.C_TAX, sp.C_INT)))
    rev_done = f'SUMPRODUCT({done},{PPL(sp.C_REV)})'
    var_all = '+'.join(f'SUM({PPL(c)})' for c in (sp.C_MAT, sp.C_LAB, sp.C_SUB, sp.C_MACH, sp.C_OTH, sp.C_TAX, sp.C_INT))
    rev_all = f'SUM({PPL(sp.C_REV)})'
    var_y = '+'.join(Y(n) for n in DIRECT + ['税金估算', '票据贴息'])
    section(ws, 5, 'A', 'J', '① 公司保本（全年）', C_RPT)
    fin = f'{Y("利息支出")}+{Y("手续费")}+{Y("利息收入")}+{Y("其他支出")}+{Y("其他税费")}-{Y("其他收入")}'
    items = [('本年已过几个月（建账那年从建账月算）', f'=MAX(1,{AX_MON}-IF({AX_Y}={AX_Y0},MONTH({OPEN_DATE})-1,0))', INT, ''),
             ('本年间接费用（管理费）', '=' + '+'.join(Y(n) for n in POOL), MONEY, '管理人员在公司的工资、社保、办公、招待、车辆、伙食、折旧等'),
             ('本年利息、手续费、其他收支（净）', f'={fin}', MONEY, '公司不分摊的'),
             ('本年固定开支合计', '=C8+C9', MONEY, ''),
             ('全年固定开支（按月平均推一年）', '=IF(N(C7)=0,0,C10/C7*12)', MONEY, '不管做多少活都要花的'),
             ('变动成本率（算出来的）', f'=IF({rev_done}>0,({var_done})/{rev_done},IF({rev_all}>0,({var_all})/{rev_all},'
                                    f'IF(N({Y("营业收入")})>0,({var_y})/{Y("营业收入")},"")))', PCT,
              f'=IF({rev_done}>0,"按已完工项目（开工至今）算的",IF({rev_all}>0,"还没有完工项目，按全部项目开工至今算的（成本没录全会偏低）","按本年的算"))'),
             ('变动成本率（手填，可空着）', None, PCT, '觉得上面的不准就填一个，比如 85%'),
             ('采用的变动成本率', '=IF(ISNUMBER(C13),C13,IF(ISNUMBER(C12),C12,0))', PCT, '每做 100 元产值，材料人工分包机械税金要花掉的'),
             ('边际贡献率（1 − 变动成本率）', '=1-C14', PCT, '每做 100 元产值，能拿来付管理费和赚钱的'),
             ('全年保本产值', '=IF(N(C15)<=0,"做多少亏多少",C11/C15)', MONEY, '一年至少要做这么多产值（含税）才不亏'),
             ('每月保本产值', '=IF(ISNUMBER(C16),C16/12,"")', MONEY, ''),
             ('本年已确认产值', f'={Y("营业收入")}', MONEY, ''),
             ('预计全年产值（按月平均推）', '=IF(N(C7)=0,0,C18/C7*12)', MONEY, ''),
             ('预计全年产值（手填，可空着）', None, MONEY, '手上合同今年能做多少，填这里更准'),
             ('采用的全年产值', '=IF(ISNUMBER(C20),C20,C19)', MONEY, ''),
             ('安全边际', '=IF(OR(N(C21)=0,NOT(ISNUMBER(C16))),"",(C21-C16)/C21)', PCT, '越大越安全；负数＝照这样做下去今年要亏'),
             ('预计全年利润（所得税前）', '=ROUND(C21*C15-C11,2)', MONEY, '')]
    for i, (lab, f, fmt, note) in enumerate(items):
        r = 7 + i
        put(ws, f'A{r}', i + 1, F_AUTO, align=AC)
        put(ws, f'B{r}', lab, F_TXTB if lab.startswith(('全年保本', '采用', '安全')) else F_TXT, align=AL)
        if f is None:
            put(ws, f'C{r}', None, F_IN, FILL_IN, fmt, AR)
        else:
            put(ws, f'C{r}', f, F_KPI_V if lab in ('全年保本产值', '安全边际') else F_AUTOB, FILL_AUTO, fmt, AR)
        put(ws, f'D{r}', note, F_NOTE, align=AL)
        ws.merge_cells(f'D{r}:J{r}')
    r = 7 + len(items)
    put(ws, f'B{r}', '一句话：', F_TXTB, align=AL, border=False)
    put(ws, f'C{r}', '=IF(NOT(ISNUMBER(C16)),"变动成本率≥100%：做得越多亏得越多，先看【项目利润表】哪些项目亏了",'
                     'IF(C21>=C16,"照现在的速度，今年产值约 "&TEXT(C21/10000,"#,##0")&" 万，比保本线多 "&TEXT((C21-C16)/10000,"#,##0")&" 万，预计能赚 "&TEXT(C23/10000,"#,##0.0")&" 万",'
                     '"照现在的速度，今年产值约 "&TEXT(C21/10000,"#,##0")&" 万，还差 "&TEXT((C16-C21)/10000,"#,##0")&" 万产值才保本"))',
        F_RED, align=AL, border=False)
    ws.merge_cells(f'C{r}:J{r}')
    # ② 在建项目
    p0 = r + 4
    section(ws, p0 - 2, 'A', 'J', '② 在建项目：总价扣掉税金和管理费，直接成本最多能花多少', C_RPT)
    header(ws, p0 - 1, [('A', '序号'), ('B', '项目'), ('C', '总价'), ('D', '已确认产值'), ('E', '已花直接成本\n（含贴息）'), ('F', '不亏的\n成本上限'),
                        ('G', '还能花'), ('H', '已花占上限'), ('I', '状态'), ('J', '提示')], C_RPT, height=40)
    tax_rate = (f'IF(AND(SUM({PPL(sp.C_REV)})>0,SUM({PPL(sp.C_TAX)})>SUM({PPL(sp.C_REV)})*0.005),'
                f'SUM({PPL(sp.C_TAX)})/SUM({PPL(sp.C_REV)}),{PA_TAXB})')
    put(ws, f'B{p0 - 3}', f'="税金按产值的 "&TEXT({tax_rate},"0.0%")&" 估，管理费率按今年的 "&TEXT(N(INDEX({SH_ALLOC}!$F${sp.AL_Y0}:$F${sp.AL_Y0 + NYEARS - 1},{AX_Y}-{AX_Y0}+1)),"0.0%")',
        F_NOTE, align=AL, border=False)
    ws.merge_cells(f'B{p0 - 3}:J{p0 - 3}')
    rate_y = f'N(INDEX({SH_ALLOC}!$F${sp.AL_Y0}:$F${sp.AL_Y0 + NYEARS - 1},{AX_Y}-{AX_Y0}+1))'
    HC = 'Z'
    counter(ws, HC, 1, sp.NPJ, lambda i: (f'AND({SH_PPL}!${sp.P_NAME}${sp.PPL_R0 + i}<>"",OR({SH_PPL}!${sp.P_STAT}${sp.PPL_R0 + i}="在建",'
                                           f'{SH_PPL}!${sp.P_STAT}${sp.PPL_R0 + i}="完工未结算"),N({SH_PPL}!${sp.P_TOTAL}${sp.PPL_R0 + i})>0)'))
    for k in range(BE_ROWS):
        r = p0 + k
        ix = f'$Y{r}'
        ws[ix] = f'={kth(k + 1, HC, 1, sp.NPJ)}'
        ws[ix].font = F_HELP
        g = lambda c: f'INDEX({PPL(c)},{ix})'
        direct = '+'.join(g(c) for c in (sp.C_MAT, sp.C_LAB, sp.C_SUB, sp.C_MACH, sp.C_OTH, sp.C_INT))
        drv = '+'.join(g(c) for c in (sp.C_LAB, sp.C_SUB, sp.C_MACH))
        r_eff = f'IF(OR({PA_DRV}="直接成本",({direct})=0),{rate_y},{rate_y}*({drv})/({direct}))'
        ws[f'A{r}'] = f'=IF({ix}=0,"",{k + 1})'
        ws[f'B{r}'] = f'=IF({ix}=0,"",{g(sp.P_NAME)})'
        ws[f'C{r}'] = f'=IF({ix}=0,"",{g(sp.P_TOTAL)})'
        ws[f'D{r}'] = f'=IF({ix}=0,"",{g(sp.C_REV)})'
        ws[f'E{r}'] = f'=IF({ix}=0,"",{direct})'
        # 还能花＝（总价扣税 − 已花直接成本 − 已经摊到的管理费）÷（1＋以后每花 1 元要摊的管理费率）
        ws[f'G{r}'] = f'=IF({ix}=0,"",ROUND((C{r}*(1-{tax_rate})-E{r}-{g(sp.C_ALLOC)})/(1+{r_eff}),2))'
        ws[f'F{r}'] = f'=IF({ix}=0,"",E{r}+G{r})'
    R1 = p0 + BE_ROWS - 1
    style_rows(ws, p0, R1, list('ABCDEFGHIJ'), auto=list('ABCDEFGHIJ'), fmts={**{c: MONEY for c in 'CDEFG'}, 'H': PCT},
               aligns={'B': AL, 'J': AL, **{c: AR for c in 'CDEFG'}})
    for r in range(p0, R1 + 1):
        for c in 'ABCDEFGHIJ':
            ws[f'{c}{r}'].fill = FILL_NONE
    ws.conditional_formatting.add(f'G{p0}:G{R1}', FormulaRule(formula=[f'N($G{p0})<0'], fill=FILL_WARN, font=F_RED))
    ws.conditional_formatting.add(f'J{p0}:J{R1}', FormulaRule(formula=[f'$J{p0}<>""'], font=F_RED))
    hide(ws, 'Y', HC)
    # ③ 报价计算器（右边）
    section(ws, 5, 'L', 'N', '③ 报价计算器', C_CHK)
    widths(ws, {'K': 2, 'L': 26, 'M': 15, 'N': 30})
    calc_rows = [('材料费（预算）', None, MONEY, '黄格子填'), ('人工费（预算）', None, MONEY, ''), ('分包费（预算）', None, MONEY, ''),
                 ('机械运输费（预算）', None, MONEY, ''), ('其他直接费（预算）', None, MONEY, '检测、期间费等'),
                 ('直接成本合计', '=SUM(M7:M11)', MONEY, ''),
                 ('管理费率', f'={rate_y}', PCT, '今年的费率（【费用分摊】）'),
                 ('要摊的管理费', f'=ROUND(IF({PA_DRV}="直接成本",M12,M8+M9+M10)*M13,2)', MONEY, '按分摊依据×费率'),
                 ('税金占产值', f'={tax_rate}', PCT, '按以往项目估'),
                 ('想赚的净利率', 0.08, PCT, '黄格子，比如 8%'),
                 ('保本报价（含税）', '=IF(1-M15<=0,"",ROUND((M12+M14)/(1-M15),2))', MONEY, '低于这个价就亏'),
                 ('建议报价（含税）', '=IF(1-M15-M16<=0,"",ROUND((M12+M14)/(1-M15-M16),2))', MONEY, '按想赚的利润率'),
                 ('按建议价预计能赚', '=IF(ISNUMBER(M18),ROUND(M18*M16,2),"")', MONEY, '')]
    for i, (lab, f, fmt, note) in enumerate(calc_rows):
        r = 7 + i
        put(ws, f'L{r}', lab, F_TXTB if '报价' in lab else F_TXT, align=AL)
        inp = f is None or not str(f).startswith('=')
        put(ws, f'M{r}', f, F_IN if inp else (F_KPI_V if '报价' in lab else F_AUTOB), FILL_IN if inp else FILL_AUTO, fmt, AR)
        put(ws, f'N{r}', note, F_NOTE, align=AL)
    ws.freeze_panes = 'A5'
    return ws


def build_all(wb, ctx):
    build_is(wb, ctx)
    build_bs(wb, ctx)
    build_be(wb, ctx)
