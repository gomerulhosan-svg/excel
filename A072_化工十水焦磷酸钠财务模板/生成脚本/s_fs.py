# -*- coding: utf-8 -*-
"""A072 化工 · 报表：利润表（会小企02＋管理行，1～12 月＋本年）、资产负债表（会小企01，选月份）、现金流量表（会小企03，1～12 月＋本年）。
   数全部来自 s_je 的 _矩（科目×月 借、贷、期初；汇＝一级含下级）；往来逐家拆方向：资产负债表右边私有区按 _位 的来源（期初＋销售＋
   进货货款运费＋销售运费＋资金台帐应收/应付变动＋手工 1122/2202）算到选的月末（条件 日期≤月末）；现金流量表按 _资 每一笔的现流号汇总。

   资产负债表口径：每个科目「本科目」余额（借正贷负，不含下级）按一级科目归行；应收/预收、应付/预付按逐家余额正负拆，
     逐家合计跟科目对不上的差额并进应收（应付）账款/预收（预付）一侧（保证表平），下面核对行报 ✗；其他应收/其他应付按本科目余额方向拆；
     存货＝14xx 存货科目＋全部成本类（4 开头，除 4301 研发→开发支出）；报表行没列到的科目：1 开头进「其他流动资产」、2 开头进「其他流动负债」、
     3 开头（3104 利润分配等）和损益类（5 开头，表结法不结转）进「未分配利润」—— 所以只要凭证平、期初平，表一定平。

   私有（别的模块不要引用；要用就用下面的名称）：
   资产负债表 K..X 隐藏：第 2..151 行＝科目号 1..150：K 科目号、L 首位（编码第一位数字）、M 选的月末本科目余额、N 已归行（一级在报表清单里）；
     第 2..301 行＝往来单位 1..300：P 单位号、Q 应收（截至选的月末）、R 应付；T/U 第 2 行起：报表用到的一级科目代码 → 科目号（没有＝-1）；
     W/X 第 2 行起：标量（输入、认出、选月、月末、逐家拆分合计）。
   利润表 Q/R 隐藏：不在固定代码里的科目 → 科目号；第 FS_PL_DIFF 行隐藏（净利润－损益类合计）。
   现金流量表 第 CF_H1..CF_H4 行隐藏：矩阵月初现金、矩阵月末现金、期末（数）、差。
   名称：利润表_净利润 / 利润表_分后利润 / 利润表_均价 / 利润表_每吨成本 / 利润表_每吨毛利 / 利润表_产量 / 利润表_销量（各 12 格，第 m 个＝m 月）、
     利润表_本年净利润 / 利润表_本年分后利润（1 格）、利润表_核对、资产负债表_平衡、资产负债表_往来核对、现金流量表_核对（文字，√ 开头＝对）、
     资产负债表_选月、资产负债表_货币资金（期末）。"""
import re
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName
from common import *
from layout import *                      # layout 的 norm/esc 盖过 common 的
from s_je import month_raw

QTY3 = '#,##0.000;[Red]-#,##0.000;"-"'
MC = [CL(2 + m) for m in MONTHS]          # C..N：1～12 月
YC = 'O'                                  # 本年
FIX = set(FIXED_CODES)
SIGN = '单位负责人：　　　　　　　　财务主管：　　　　　　　　复核：　　　　　　　　制表：'


def _dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def _lab(ws, a, t):
    ws[a] = t
    ws[a].font = F_HELP


def _grey_months(ws, rng):
    """建账以前、截止以后的月份列灰字"""
    c0 = rng.split(':')[0]
    col = ''.join(ch for ch in c0 if ch.isalpha())
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'OR(COLUMN({col}$1)-2<P_建账月,COLUMN({col}$1)-2>P_截止月)'], font=Font(name=YH, sz=10, color='FFA6A6A6')))


def _head3(ws, last, form):
    put(ws, 'A3', '="编制单位："&P_公司名称', F_TXT, align=AL, border=False)
    ws.row_dimensions[3].height = 22
    put(ws, f'{CL(CI(last) - 1)}3', form, F_NOTE, align=AR, border=False)
    ws.merge_cells(f'{CL(CI(last) - 1)}3:{last}3')


# ═══════════════════════════════ 资产负债表 ═══════════════════════════════
BP = dict(no='K', d1='L', end='M', cov='N', uno='P', ar='Q', ap='R', code='T', kno='U', lab='W', val='X')
KR1 = N_COA + 1                            # 科目块 第 2..151 行
UR1 = N_UNIT + 1                           # 单位块 第 2..301 行
BS_SEL = 'D3'
BS_CODES = ['1001', '1002', '1012', '1101', '1121', '1122', '1123', '1131', '1132', '1221', '1401', '1402', '1403', '1404',
            '1405', '1407', '1408', '1411', '1421', '1501', '1511', '1601', '1602', '1604', '1605', '1606', '1621', '1622',
            '1701', '1702', '1801', '4301', '2001', '2201', '2202', '2203', '2211', '2221', '2231', '2232', '2241', '2401',
            '2501', '2701', '3001', '3002', '3101']
INV_CODES = ['1401', '1402', '1403', '1404', '1405', '1407', '1408', '1411', '1421']
BS_SC = ['输入', '认出', '选月', '月末', '应收正', '应收负', '应收差', '应付正', '应付负', '应付差',
         '应收正0', '应收负0', '应收差0', '应付正0', '应付负0', '应付差0', '合计', '合计0']


def bsc(k):
    return f'${BP["val"]}${2 + BS_SC.index(k)}'


def _rk(c, r1):
    return f'${BP[c]}$2:${BP[c]}${r1}'


def bcol(w):
    return _rk('end', KR1) if w == 'e' else '矩_期初'


def kc(code):
    return f'${BP["kno"]}${2 + BS_CODES.index(code)}'


def A(code, w):
    """一级科目 code（含下级）本科目余额合计（借正贷负）"""
    return f'SUMIF(科目_一级号,{kc(code)},{bcol(w)})'


def DP(code, w, op='>0'):
    """一级科目 code 下各科目（含一级自己）按余额方向：借方余额合计（op='<0' 贷方余额合计，负数）"""
    return f'SUMIFS({bcol(w)},科目_一级号,{kc(code)},{bcol(w)},"{op}")'


def CLS(d, w):
    """编码第一位＝d、一级不在报表清单里的科目余额合计"""
    return f'SUMIFS({bcol(w)},{_rk("d1", KR1)},{d},{_rk("cov", KR1)},0)'


def P5(w):
    return f'SUMIF({_rk("d1", KR1)},5,{bcol(w)})'


def SC(k, w):
    return bsc(k + ('0' if w == 'o' and k not in ('选月', '月末') else ''))


BS_LEFT = [  # (行, 名称, 行次, 公式(w) / 'sec' / 'sum:行,…' / 'expr:' / 0)
    (5, '流动资产：', None, 'sec'),
    (6, '货币资金', 1, lambda w: f'{A("1001", w)}+{A("1002", w)}+{A("1012", w)}'),
    (7, '短期投资', 2, lambda w: A('1101', w)),
    (8, '应收票据', 3, lambda w: A('1121', w)),
    (9, '应收账款', 4, lambda w: f'{SC("应收正", w)}+MAX({SC("应收差", w)},0)+MAX({A("2203", w)},0)'),
    (10, '预付账款', 5, lambda w: f'{A("1123", w)}+{SC("应付负", w)}+MAX(-{SC("应付差", w)},0)'),
    (11, '应收股利', 6, lambda w: A('1131', w)),
    (12, '应收利息', 7, lambda w: A('1132', w)),
    (13, '其他应收款', 8, lambda w: f'{DP("1221", w)}+{DP("2241", w)}'),
    (14, '存货', 9, lambda w: '+'.join(A(c, w) for c in INV_CODES) + '+' + CLS(4, w)),
    (15, '　其中：原材料', 10, lambda w: A('1403', w)),
    (16, '　　　　在产品（生产成本、制造费用）', 11, lambda w: CLS(4, w)),
    (17, '　　　　库存商品', 12, lambda w: A('1405', w)),
    (18, '　　　　周转材料', 13, lambda w: A('1411', w)),
    (19, '其他流动资产', 14, lambda w: CLS(1, w)),
    (20, '流动资产合计', 15, 'sum:6,7,8,9,10,11,12,13,14,19'),
    (21, '非流动资产：', None, 'sec'),
    (22, '长期债券投资', 16, lambda w: A('1501', w)),
    (23, '长期股权投资', 17, lambda w: A('1511', w)),
    (24, '固定资产原价', 18, lambda w: A('1601', w)),
    (25, '减：累计折旧', 19, lambda w: f'-{A("1602", w)}'),
    (26, '固定资产账面价值', 20, 'expr:#24-#25'),
    (27, '在建工程', 21, lambda w: A('1604', w)),
    (28, '工程物资', 22, lambda w: A('1605', w)),
    (29, '固定资产清理', 23, lambda w: A('1606', w)),
    (30, '生产性生物资产', 24, lambda w: f'{A("1621", w)}+{A("1622", w)}'),
    (31, '无形资产', 25, lambda w: f'{A("1701", w)}+{A("1702", w)}'),
    (32, '开发支出', 26, lambda w: A('4301', w)),
    (33, '长期待摊费用', 27, lambda w: A('1801', w)),
    (34, '其他非流动资产', 28, 0),
    (35, '非流动资产合计', 29, 'sum:22,23,26,27,28,29,30,31,32,33,34'),
    (37, '资产总计', 30, 'sum:20,35'),
]
BS_RIGHT = [
    (5, '流动负债：', None, 'sec'),
    (6, '短期借款', 31, lambda w: f'-{A("2001", w)}'),
    (7, '应付票据', 32, lambda w: f'-{A("2201", w)}'),
    (8, '应付账款', 33, lambda w: f'{SC("应付正", w)}+MAX({SC("应付差", w)},0)'),
    (9, '预收账款', 34, lambda w: f'{SC("应收负", w)}+MAX(-{SC("应收差", w)},0)+MAX(-{A("2203", w)},0)'),
    (10, '应付职工薪酬', 35, lambda w: f'-{A("2211", w)}'),
    (11, '应交税费', 36, lambda w: f'-{A("2221", w)}'),
    (12, '应付利息', 37, lambda w: f'-{A("2231", w)}'),
    (13, '应付利润', 38, lambda w: f'-{A("2232", w)}'),
    (14, '其他应付款', 39, lambda w: f'-({DP("1221", w, "<0")}+{DP("2241", w, "<0")})'),
    (15, '其他流动负债', 40, lambda w: f'-{CLS(2, w)}'),
    (16, '流动负债合计', 41, 'sum:6,7,8,9,10,11,12,13,14,15'),
    (17, '非流动负债：', None, 'sec'),
    (18, '长期借款', 42, lambda w: f'-{A("2501", w)}'),
    (19, '长期应付款', 43, lambda w: f'-{A("2701", w)}'),
    (20, '递延收益', 44, lambda w: f'-{A("2401", w)}'),
    (21, '其他非流动负债', 45, 0),
    (22, '非流动负债合计', 46, 'sum:18,19,20,21'),
    (23, '负债合计', 47, 'sum:16,22'),
    (25, '所有者权益（或股东权益）：', None, 'sec'),
    (26, '实收资本（或股本）', 48, lambda w: f'-{A("3001", w)}'),
    (27, '资本公积', 49, lambda w: f'-{A("3002", w)}'),
    (28, '盈余公积', 50, lambda w: f'-{A("3101", w)}'),
    (29, '未分配利润', 51, lambda w: f'-{CLS(3, w)}-{P5(w)}'),
    (30, '所有者权益（或股东权益）合计', 52, 'sum:26,27,28,29'),
    (37, '负债和所有者权益（或股东权益）总计', 53, 'sum:23,30'),
]
BS_TOT = 37                                # 资产总计 / 负债和所有者权益总计 行
BS_CHK, BS_WL, BS_NOTE, BS_SIGN = 39, 40, 41, 43


def build_bs(wb, ws):
    widths(ws, {'A': 30, 'B': 6, 'C': 17, 'D': 17, 'E': 34, 'F': 6, 'G': 17, 'H': 17, 'I': 3})
    title(ws, '资 产 负 债 表', 'H', C_RPT,
          '★ 选月份（D3 黄格，空着＝截止日期那个月）：期末余额＝那个月最后一天，年初余额＝建账期初（建账日前一天晚上）。数全部来自记账凭证。'
          '应收账款/预收账款、应付账款/预付账款按每个客户、供应商（含物流）的余额方向拆开：欠我们的算应收，先收了钱的算预收；我们欠的算应付，先付了钱的算预付。'
          '其他应收款/其他应付款按明细科目余额方向拆。存货＝原材料＋库存商品＋生产成本（在产）＋制造费用（还没分完的）。'
          '损益不结转（表结法）：未分配利润＝期初＋本年到这个月的利润－投资款利息（利润分配）。最下面查资产＝负债＋所有者权益、逐家合计＝应收应付科目。')
    put(ws, 'A3', '="编制单位："&P_公司名称', F_TXT, align=AL, border=False)
    selector(ws, 'C3', '月份', BS_SEL, None, '"1,2,3,4,5,6,7,8,9,10,11,12"',
             prompt='选 1～12（本年度）；空着＝截止日期那个月；填 2026-9、9月 也认')
    S = bsc
    put(ws, 'E3', (f'=P_年度&" 年 "&{S("选月")}&" 月 "&DAY({S("月末")})&" 日"&IF({S("认出")}=0,"　⚠ 月份没认出来，按截止日期那个月",'
                   f'IF({S("选月")}>P_截止月,"　⚠ 晚于截止日期那个月：月末折旧、摊销、计提、成本结转还没做",'
                   f'IF({S("选月")}<P_建账月,"　ℹ 早于建账月：就是建账期初数","")))'), F_TXTB, align=ALW, border=False)
    ws.merge_cells('E3:F3')
    ws.row_dimensions[3].height = 34
    put(ws, 'G3', '会小企01表　单位：元', F_NOTE, align=AR, border=False)
    ws.merge_cells('G3:H3')
    header(ws, 4, [('A', '资    产'), ('B', '行次'), ('C', '期末余额'), ('D', '年初余额'),
                   ('E', '负债和所有者权益（或股东权益）'), ('F', '行次'), ('G', '期末余额'), ('H', '年初余额')], C_RPT, height=26)
    # ── 私有：科目块 ──
    for k, t in (('no', '科目号'), ('d1', '首位'), ('end', '月末本科目'), ('cov', '已归行'), ('uno', '单位号'), ('ar', '应收'),
                 ('ap', '应付'), ('code', '代码'), ('kno', '科目号'), ('lab', '标量'), ('val', '值')):
        _lab(ws, f'{BP[k]}1', t)
    sel = S('选月')
    cov = f'${BP["kno"]}$2:${BP["kno"]}${len(BS_CODES) + 1}'
    for n in range(1, N_COA + 1):
        r = n + 1
        no = f'${BP["no"]}{r}'
        ws[no[1:]] = n
        ws[f'{BP["d1"]}{r}'] = f'=IFERROR(VALUE(LEFT(INDEX(科目_编码,{no}),1)),0)'
        ws[f'{BP["end"]}{r}'] = (f'=ROUND(INDEX(矩_期初,{no})+SUM(INDEX(矩_借,{no},1):INDEX(矩_借,{no},{sel}))'
                                 f'-SUM(INDEX(矩_贷,{no},1):INDEX(矩_贷,{no},{sel})),2)')
        ws[f'{BP["cov"]}{r}'] = f'=IF(ISNA(MATCH(INDEX(科目_一级号,{no}),{cov},0)),0,1)'
    # ── 私有：往来单位逐家（来源同 _位 应收/应付，条件 日期≤选的月末） ──
    me = f'"<="&{S("月末")}'
    for n in range(1, N_UNIT + 1):
        r = n + 1
        u = f'${BP["uno"]}{r}'
        ws[u[1:]] = n
        mj = lambda side, code: f'SUMIFS(手_{side},手_单位号,{u},手_一级号,{kh(code)},手_记账,1,手_日期,{me})'
        ws[f'{BP["ar"]}{r}'] = (f'=IF(INDEX(位_有效,{u})<>1,0,ROUND(INDEX(位_期初应收,{u})'
                                f'+SUMIFS(销_收入含税,销_客户号,{u},销_日期,{me})+SUMIFS(资_应收变动,资_单位号,{u},资_日期,{me})'
                                f'+{mj("借", "1122")}-{mj("贷", "1122")},2))')
        ws[f'{BP["ap"]}{r}'] = (f'=IF(INDEX(位_有效,{u})<>1,0,ROUND(INDEX(位_期初应付,{u})'
                                f'+SUMIFS(进_货款计,进_供应商号,{u},进_日期,{me})+SUMIFS(进_运费计,进_物流号,{u},进_日期,{me})'
                                f'+SUMIFS(销_运费计,销_物流号,{u},销_日期,{me})+SUMIFS(资_应付变动,资_单位号,{u},资_日期,{me})'
                                f'+{mj("贷", "2202")}-{mj("借", "2202")},2))')
    # ── 私有：报表用到的一级科目 ──
    for i, code in enumerate(BS_CODES):
        r = i + 2
        ws[f'{BP["code"]}{r}'] = code
        ws[f'{BP["kno"]}{r}'] = f'=IFERROR(MATCH("{code}",科目_编码,0),-1)'
    # ── 私有：标量 ──
    ar, ap = _rk('ar', UR1), _rk('ap', UR1)
    f = {
        '输入': '=' + month_raw(f'${BS_SEL[0]}${BS_SEL[1:]}'),
        '认出': f'=IF({S("输入")}=-1,1,IF(AND({S("输入")}>=1,{S("输入")}<=12,INT({S("输入")})={S("输入")}),1,0))',
        '选月': f'=IF(OR({S("输入")}=-1,{S("认出")}=0),P_截止月,{S("输入")})',
        '月末': f'=DATE(P_年度,{S("选月")}+1,0)',
        '应收正': f'=SUMIF({ar},">0")', '应收负': f'=-SUMIF({ar},"<0")',
        '应收差': f'=ROUND({A("1122", "e")}-SUM({ar}),2)',
        '应付正': f'=SUMIF({ap},">0")', '应付负': f'=-SUMIF({ap},"<0")',
        '应付差': f'=ROUND(-{A("2202", "e")}-SUM({ap}),2)',
        '应收正0': '=SUMIF(位_期初应收,">0")', '应收负0': '=-SUMIF(位_期初应收,"<0")',
        '应收差0': f'=ROUND({A("1122", "o")}-SUM(位_期初应收),2)',
        '应付正0': '=SUMIF(位_期初应付,">0")', '应付负0': '=-SUMIF(位_期初应付,"<0")',
        '应付差0': f'=ROUND(-{A("2202", "o")}-SUM(位_期初应付),2)',
        '合计': f'=ROUND(SUM({_rk("end", KR1)}),2)', '合计0': '=ROUND(SUM(矩_期初),2)',
    }
    for i, k in enumerate(BS_SC):
        _lab(ws, f'{BP["lab"]}{2 + i}', k)
        ws[f'{BP["val"]}{2 + i}'] = f[k]
    ws[S('月末').replace('$', '')].number_format = DATE
    # ── 报表行 ──
    done = set()
    for side, (cl, cn, cv, cy) in ((BS_LEFT, ('A', 'B', 'C', 'D')), (BS_RIGHT, ('E', 'F', 'G', 'H'))):
        for r, name, no, fx in side:
            done.update(f'{c}{r}' for c in (cl, cn, cv, cy))
            is_sec = fx == 'sec'
            is_tot = isinstance(fx, str) and (fx.startswith('sum:') or fx.startswith('expr:'))
            big = r == BS_TOT
            fl = FILL_TOT if big else (FILL_SUB if is_tot else None)
            put(ws, f'{cl}{r}', name, F_TXTB if (is_sec or is_tot) else F_TXT, fl, align=AL)
            put(ws, f'{cn}{r}', no, F_TXT, fl, align=AC)
            for col, w in ((cv, 'e'), (cy, 'o')):
                if is_sec:
                    val = None
                elif fx == 0:
                    val = 0
                elif isinstance(fx, str) and fx.startswith('sum:'):
                    val = '=' + '+'.join(f'{col}{x}' for x in fx[4:].split(','))
                elif isinstance(fx, str) and fx.startswith('expr:'):
                    val = '=' + fx[5:].replace('#', col)
                else:
                    val = f'=ROUND({fx(w)},2)'
                put(ws, f'{col}{r}', val, F_TXTB if is_tot else F_TXT, fl, MONEY, AR)
    for r in range(5, BS_TOT + 1):
        for c in 'ABCDEFGH':
            if f'{c}{r}' not in done:
                put(ws, f'{c}{r}', None)
    # ── 核对 ──
    put(ws, f'A{BS_CHK}', '平衡检查：', F_TXTB, align=AR, border=False)
    put(ws, f'C{BS_CHK}', (f'=IF(AND(ROUND(C{BS_TOT}-G{BS_TOT},2)=0,ROUND(D{BS_TOT}-H{BS_TOT},2)=0),'
                           f'"√ 资产＝负债＋所有者权益（期末、年初都平）",'
                           f'"✗ 不平：期末差 "&FIXED(C{BS_TOT}-G{BS_TOT},2)&"，年初差 "&FIXED(D{BS_TOT}-H{BS_TOT},2)'
                           f'&"（先看【科目余额表】平不平；科目编码不是 1～5 开头的余额进不了报表）")'),
        F_RED, align=AL, border=False)
    ws.merge_cells(f'C{BS_CHK}:H{BS_CHK}')
    put(ws, f'A{BS_WL}', '往来逐家核对：', F_TXTB, align=AR, border=False)
    put(ws, f'C{BS_WL}', (f'=IF(AND({S("应收差")}=0,{S("应付差")}=0,{S("应收差0")}=0,{S("应付差0")}=0),'
                          f'"√ 客户、供应商逐家余额合计＝应收账款 1122、应付账款 2202（期末、年初）",'
                          f'"✗ 逐家合计跟科目对不上：期末 应收差 "&FIXED({S("应收差")},2)&"、应付差 "&FIXED({S("应付差")},2)'
                          f'&"；年初 应收差 "&FIXED({S("应收差0")},2)&"、应付差 "&FIXED({S("应付差0")},2)'
                          f'&"（差额先并进应收/应付账款；多半是资金台帐、手工凭证的往来单位没填或不在【往来单位】里，看【数据校验】）")'),
        F_RED, align=AL, border=False)
    ws.merge_cells(f'C{BS_WL}:H{BS_WL}')
    put(ws, f'C{BS_NOTE}', (f'=IF(OR(ROUND(C19,2)<>0,ROUND(D19,2)<>0,ROUND(G15,2)<>0,ROUND(H15,2)<>0),'
                            f'"ℹ 「其他流动资产」「其他流动负债」是报表行没列到的科目（自己加的一级科目），明细看【科目余额表】","")'),
        F_NOTE, align=AL, border=False)
    ws.merge_cells(f'C{BS_NOTE}:H{BS_NOTE}')
    put(ws, f'A{BS_SIGN}', SIGN, F_TXT, align=AL, border=False)
    ws.merge_cells(f'A{BS_SIGN}:H{BS_SIGN}')
    for rr in (BS_CHK, BS_WL):
        ws.conditional_formatting.add(f'C{rr}', FormulaRule(formula=[f'LEFT($C${rr},1)="√"'],
                                                          font=Font(name=YH, sz=10, bold=True, color='FF548235')))
    for c in 'KLMNOPQRSTUVWX':
        ws.column_dimensions[c].hidden = True
    for r in range(2, UR1 + 1):
        for c in ('no', 'd1', 'end', 'cov', 'uno', 'ar', 'ap', 'code', 'kno', 'lab', 'val'):
            if ws[f'{BP[c]}{r}'].value is not None:
                ws[f'{BP[c]}{r}'].font = F_HELP
    ws.freeze_panes = 'A5'
    print_setup(ws, '4:4', landscape=True)
    ws.page_setup.fitToHeight = 1
    ws.print_area = f'A1:H{BS_SIGN}'
    _dn(wb, '资产负债表_平衡', f"{SH_BS}!$C${BS_CHK}")
    _dn(wb, '资产负债表_往来核对', f"{SH_BS}!$C${BS_WL}")
    _dn(wb, '资产负债表_选月', f"{SH_BS}!{S('选月')}")
    _dn(wb, '资产负债表_货币资金', f"{SH_BS}!$C$6")
    home_link(ws, 'J1')
    ws.column_dimensions['J'].width = 10


# ═══════════════════════════════ 利润表 ═══════════════════════════════
PL_ROWS = [  # (名称, 行次, 收入类科目(贷－借), 费用类科目(借－贷)) 或 (名称, 行次, 合计式, None)
    ('一、营业收入', 1, ['5001', '5051'], []),
    ('减：营业成本', 2, [], ['5401', '5402']),
    ('　　税金及附加', 3, [], ['5403']),
    ('　　销售费用', 11, [], ['5601']),
    ('　　管理费用', 14, [], ['5602']),
    ('　　其中：开办费（摊销）', 15, [], ['560207']),
    ('　　　　　业务招待费', 16, [], ['560203']),
    ('　　财务费用', 18, [], ['5603']),
    ('　　其中：利息费用（收入以“-”号填列）', 19, [], ['560301', '560302']),
    ('加：投资收益（损失以“-”号填列）', 20, ['5111'], []),
    ('二、营业利润（亏损以“-”号填列）', 21, 'R1-R2-R3-R11-R14-R18+R20', None),
    ('加：营业外收入', 22, ['5301'], []),
    ('减：营业外支出', 24, [], ['5711']),
    ('三、利润总额（亏损总额以“-”号填列）', 30, 'R21+R22-R24', None),
    ('减：所得税费用', 31, [], ['5801']),
    ('四、净利润（净亏损以“-”号填列）', 32, 'R30-R31', None),
]
PL_R0 = 5
PL_ROW = {no: PL_R0 + i for i, (_n, no, *_x) in enumerate(PL_ROWS)}
PL_NET = PL_ROW[32]
PL_MG = PL_NET + 2                          # 管理口径标题行
PL_MGMT = ['产量', '销量', '主营收入', '主营成本', '均价', '每吨成本', '每吨毛利', '毛利率', '投资利息', '分后利润']
PL_MG_NAME = {
    '产量': '产量（吨，生产日报成品）', '销量': '销量（吨，产品）', '主营收入': '主营业务收入（不含税）',
    '主营成本': '="主营业务成本"&IF(P_运费进成本=1,"（含销货运费）","（销货运费记在销售费用）")', '均价': '不含税均价（元/吨）＝主营业务收入÷销量',
    '每吨成本': '每吨主营业务成本（元/吨）', '每吨毛利': '每吨毛利（元/吨）', '毛利率': '毛利率',
    '投资利息': '减：投资款利息（利润分配-应付利润，不进费用）', '分后利润': '分完投资利息后的利润',
}
PL_MROW = {k: PL_MG + 1 + i for i, k in enumerate(PL_MGMT)}
PL_CHK = PL_MROW['分后利润'] + 2            # 核对：损益类合计
PL_CHKT = PL_CHK + 1                        # 核对文字
PL_DIFF = PL_CHK + 2                        # 隐藏：差
PL_SIGN = PL_DIFF + 2
PL_NONFIX = sorted({c for _n, _no, rv, ex in PL_ROWS if isinstance(rv, list) for c in rv + ex if c not in FIX})
PL_LK = dict(code='S', kno='T')


def pl_k(code):
    if code in FIX:
        return kh(code), False
    return f'${PL_LK["kno"]}${2 + PL_NONFIX.index(code)}', True


def pl_amt(code, m, sign):
    """科目 code（含下级）第 m 月发生额：sign=1 借－贷，-1 贷－借"""
    k, guard = pl_k(code)
    a, b = ('矩_汇借', '矩_汇贷') if sign > 0 else ('矩_汇贷', '矩_汇借')
    e = f'INDEX({a},{k},{m})-INDEX({b},{k},{m})'
    return f'IF({k}>0,{e},0)' if guard else f'({e})'


def build_pl(wb, ws):
    widths(ws, {'A': 46, 'B': 5, **{c: 12.5 for c in MC}, YC: 14.5, 'P': 3})
    title(ws, '利 润 表', YC, C_RPT,
          '★ 每个月一列，最后一列是本年合计。数全部来自记账凭证（损益类科目每个月的发生额；表结法，不出结转损益凭证）。'
          '灰字的月份：建账以前没有数；截止日期以后月末的折旧、摊销、工资、利息、成本结转还没做，数不全。'
          '下面「管理口径」几行是老板看的：每吨均价、每吨成本、毛利率，以及「减：投资款利息」（合伙人投资款按约定算的「利息」记在利润分配，不进财务费用）。'
          '最下面核对：净利润＝损益类科目（5 开头）贷方－借方。')
    _head3(ws, YC, '会小企02表　单位：元')
    put(ws, 'C3', '="截止日期 "&TEXT(P_截止,"yyyy-mm-dd")&"（"&P_截止月&" 月）；灰字的月份没有数或月末还没结转"',
        F_NOTE, align=AL, border=False)
    ws.merge_cells('C3:K3')
    header(ws, 4, [('A', '项    目'), ('B', '行次')] + [(c, f'{m}月') for m, c in zip(MONTHS, MC)] + [(YC, '本年累计')],
           C_RPT, height=26)
    for i, code in enumerate(PL_NONFIX):
        _lab(ws, f'{PL_LK["code"]}{2 + i}', code)
        ws[f'{PL_LK["kno"]}{2 + i}'] = f'=IFERROR(MATCH("{code}",科目_编码,0),0)'
        ws[f'{PL_LK["kno"]}{2 + i}'].font = F_HELP
    _lab(ws, f'{PL_LK["code"]}1', '代码')
    _lab(ws, f'{PL_LK["kno"]}1', '科目号')
    for name, no, rev, exp in PL_ROWS:
        r = PL_ROW[no]
        tot = isinstance(rev, str)
        big = no == 32
        fl = FILL_TOT if big else (FILL_SUB if tot else None)
        put(ws, f'A{r}', name, F_TXTB if (tot or not name.startswith('　')) else F_TXT, fl, align=AL)
        put(ws, f'B{r}', no, F_TXT, fl, align=AC)
        for m, col in zip(MONTHS, MC):
            if tot:
                val = '=' + re.sub(r'R(\d+)', lambda x: f'{col}{PL_ROW[int(x.group(1))]}', rev)
            else:
                parts = [pl_amt(c, m, -1) for c in rev] + [pl_amt(c, m, 1) for c in exp]
                val = f'=ROUND({"+".join(parts)},2)'
            put(ws, f'{col}{r}', val, F_TXTB if tot else F_TXT, fl, MONEY, AR)
        put(ws, f'{YC}{r}', f'=SUM(C{r}:N{r})', F_TXTB, fl or FILL_SUB, MONEY, AR)
    # ── 管理口径 ──
    r = PL_MG
    section(ws, r, 'A', YC, '管理口径（老板看的，不是会小企 02 的行）', C_RPT)
    R = PL_MROW
    for k in PL_MGMT:
        r = R[k]
        bold = k in ('均价', '每吨毛利', '分后利润')
        fl = FILL_TOT if k == '分后利润' else None
        put(ws, f'A{r}', PL_MG_NAME[k], F_TXTB if bold else F_TXT, fl, align=AL)
        put(ws, f'B{r}', None, F_TXT, fl)
        fmt = QTY3 if k in ('产量', '销量') else (PCT if k == '毛利率' else MONEY)
        for m, col in list(zip(MONTHS, MC)) + [(0, YC)]:
            c = lambda kk: f'{col}{R[kk]}'
            if k in ('产量', '销量', '主营收入', '主营成本', '投资利息') and m == 0:
                val = f'=SUM(C{r}:N{r})'
            elif k == '产量':
                val = f'=ROUND(SUM(INDEX(本_产量,0,{m})),3)'
            elif k == '销量':
                val = f'=ROUND(SUM(INDEX(品_销量,0,{m})),3)'
            elif k == '主营收入':
                val = f'=ROUND({pl_amt("5001", m, -1)},2)'
            elif k == '主营成本':
                val = f'=ROUND({pl_amt("5401", m, 1)},2)'
            elif k == '投资利息':
                val = f'=ROUND({pl_amt("310401", m, 1)},2)'
            elif k == '均价':
                val = f'=IF({c("销量")}>0,ROUND({c("主营收入")}/{c("销量")},2),"")'
            elif k == '每吨成本':
                val = f'=IF({c("销量")}>0,ROUND({c("主营成本")}/{c("销量")},2),"")'
            elif k == '每吨毛利':
                val = f'=IF({c("销量")}>0,ROUND({c("均价")}-{c("每吨成本")},2),"")'
            elif k == '毛利率':
                val = f'=IF({c("主营收入")}=0,"",ROUND(({c("主营收入")}-{c("主营成本")})/{c("主营收入")},4))'
            else:                                    # 分后利润
                val = f'={col}{PL_NET}-{c("投资利息")}'
            put(ws, f'{col}{r}', val, F_TXTB if bold else F_TXT, fl, fmt, AR)
    # ── 核对 ──
    d1 = f"{SH_BS}!${BP['d1']}$2:${BP['d1']}${KR1}"
    put(ws, f'A{PL_CHK}', '核对：损益类科目（编码 5 开头）贷方－借方', F_NOTE, align=AL, border=False)
    put(ws, f'A{PL_DIFF}', '（隐藏）净利润－损益类合计', F_HELP, align=AL, border=False)
    for m, col in zip(MONTHS, MC):
        put(ws, f'{col}{PL_CHK}', f'=ROUND(SUMIF({d1},5,INDEX(矩_贷,0,{m}))-SUMIF({d1},5,INDEX(矩_借,0,{m})),2)',
            F_NOTE, fmt=MONEY, align=AR, border=False)
    put(ws, f'{YC}{PL_CHK}', f'=SUM(C{PL_CHK}:N{PL_CHK})', F_NOTE, fmt=MONEY, align=AR, border=False)
    for col in MC + [YC]:
        put(ws, f'{col}{PL_DIFF}', f'=ROUND({col}{PL_NET}-{col}{PL_CHK},2)', F_HELP, fmt=MONEY, align=AR, border=False)
    put(ws, f'A{PL_CHKT}', (f'=IF(COUNTIF(C{PL_DIFF}:{YC}{PL_DIFF},"<>0")=0,"√ 净利润＝损益类科目贷方－借方（每个月、本年）",'
                            f'"✗ 净利润跟损益类科目对不上（有损益科目没列进报表）：本年差 "&FIXED({YC}{PL_DIFF},2))'),
        F_RED, align=AL, border=False)
    ws.merge_cells(f'A{PL_CHKT}:{YC}{PL_CHKT}')
    ws.conditional_formatting.add(f'A{PL_CHKT}', FormulaRule(formula=[f'LEFT($A${PL_CHKT},1)="√"'],
                                                           font=Font(name=YH, sz=10, bold=True, color='FF548235')))
    ws.row_dimensions[PL_DIFF].hidden = True
    put(ws, f'A{PL_SIGN}', SIGN, F_TXT, align=AL, border=False)
    ws.merge_cells(f'A{PL_SIGN}:{YC}{PL_SIGN}')
    _grey_months(ws, f'C4:N{PL_MROW["分后利润"]}')
    ws.column_dimensions[PL_LK['code']].hidden = True
    ws.column_dimensions[PL_LK['kno']].hidden = True
    ws.freeze_panes = 'C5'
    print_setup(ws, '4:4', landscape=True)
    ws.page_setup.fitToHeight = 1
    ws.print_area = f'A1:{YC}{PL_SIGN}'
    row = lambda k, rr: _dn(wb, f'利润表_{k}', f"{SH_PL}!$C${rr}:$N${rr}")
    row('净利润', PL_NET)
    for k, nm in (('分后利润', '分后利润'), ('均价', '均价'), ('每吨成本', '每吨成本'), ('每吨毛利', '每吨毛利'),
                  ('产量', '产量'), ('销量', '销量')):
        row(nm, PL_MROW[k])
    _dn(wb, '利润表_本年净利润', f"{SH_PL}!${YC}${PL_NET}")
    _dn(wb, '利润表_本年分后利润', f"{SH_PL}!${YC}${PL_MROW['分后利润']}")
    _dn(wb, '利润表_核对', f"{SH_PL}!$A${PL_CHKT}")
    home_link(ws, 'Q1')
    ws.column_dimensions['Q'].width = 10


# ═══════════════════════════════ 现金流量表 ═══════════════════════════════
CF_LAYOUT = [  # (名称, 行次, 类型)：'sec' 标题；('in'/'out', 现流号)；'sum:±行次,…'；'transit' 在途；'open' 期初；'close' 期末
    ('一、经营活动产生的现金流量：', None, 'sec'),
    ('销售产成品、商品、提供劳务收到的现金', 1, ('in', 1)),
    ('收到其他与经营活动有关的现金', 2, ('in', 2)),
    ('购买原材料、商品、接受劳务支付的现金', 3, ('out', 3)),
    ('支付的职工薪酬', 4, ('out', 4)),
    ('支付的税费', 5, ('out', 5)),
    ('支付其他与经营活动有关的现金', 6, ('out', 6)),
    ('经营活动产生的现金流量净额', 7, 'sum:+1,+2,-3,-4,-5,-6'),
    ('二、投资活动产生的现金流量：', None, 'sec'),
    ('收回短期投资、长期债券投资和长期股权投资收到的现金', 8, ('in', 8)),
    ('取得投资收益收到的现金', 9, ('in', 9)),
    ('处置固定资产、无形资产和其他非流动资产收回的现金净额', 10, ('in', 10)),
    ('短期投资、长期债券投资和长期股权投资支付的现金', 11, ('out', 11)),
    ('购建固定资产、无形资产和其他非流动资产支付的现金', 12, ('out', 12)),
    ('投资活动产生的现金流量净额', 13, 'sum:+8,+9,+10,-11,-12'),
    ('三、筹资活动产生的现金流量：', None, 'sec'),
    ('取得借款收到的现金', 14, ('in', 14)),
    ('吸收投资者投资收到的现金', 15, ('in', 15)),
    ('偿还借款本金支付的现金', 16, ('out', 16)),
    ('偿还借款利息支付的现金', 17, ('out', 17)),
    ('分配利润支付的现金', 18, ('out', 18)),
    ('筹资活动产生的现金流量净额', 19, 'sum:+14,+15,-16,-17,-18'),
    ('四、现金净增加额', 20, 'sum:+7,+13,+19'),
    ('加：资金账户间在途（内部转账两行不在同一个月、没设现金流量项目的）', None, 'transit'),
    ('加：期初现金余额', 21, 'open'),
    ('五、期末现金余额', 22, 'close'),
]
CF_R0 = 5
CF_ROW = {}
for _i, (_nm, _no, _t) in enumerate(CF_LAYOUT):
    if _no:
        CF_ROW[_no] = CF_R0 + _i
    if _t in ('transit', 'open', 'close'):
        CF_ROW[_t] = CF_R0 + _i
CF_LAST = CF_R0 + len(CF_LAYOUT) - 1
CF_CHKT = CF_LAST + 2
CF_SIGN = CF_CHKT + 2
CF_H1, CF_H2, CF_H3, CF_H4 = CF_SIGN + 2, CF_SIGN + 3, CF_SIGN + 4, CF_SIGN + 5
CF_CHKA = f'$C${CF_H4 + 1}'                  # 期末现金－资产负债表货币资金
CASH_CODES = ['1001', '1002', '1012']


def build_cf(wb, ws):
    widths(ws, {'A': 50, 'B': 5, **{c: 12.5 for c in MC}, YC: 14.5})
    title(ws, '现 金 流 量 表', YC, C_RPT,
          '★ 直接法，每个月一列＋本年。现金＝现金、银行、微信支付宝账户（承兑汇票、老板垫付、备用金不是现金）。每一笔按收支类别'
          '（员工报销按报销项目）在【基础资料】里设的「现金流量收/付」归到 1～18 行；内部转账（现金账户之间倒钱）不算流入流出，'
          '两行没在同一个月的、类别没设现金流量项目的显示在「资金账户间在途」。期末现金＝期初＋净增加额＋在途，跟资产负债表的货币资金核对。')
    _head3(ws, YC, '会小企03表　单位：元')
    header(ws, 4, [('A', '项    目'), ('B', '行次')] + [(c, f'{m}月') for m, c in zip(MONTHS, MC)] + [(YC, '本年累计')],
           C_RPT, height=26)
    R = CF_ROW
    for i, (name, no, typ) in enumerate(CF_LAYOUT):
        r = CF_R0 + i
        is_sec = typ == 'sec'
        is_tot = isinstance(typ, str) and not is_sec and typ != 'transit'
        big = typ == 'close'
        fl = FILL_TOT if big else (FILL_SUB if is_tot else None)
        put(ws, f'A{r}', name, F_TXTB if (is_sec or is_tot) else F_TXT, fl, align=AL)
        put(ws, f'B{r}', no, F_TXT, fl, align=AC)
        for m, col in list(zip(MONTHS, MC)) + [(0, YC)]:
            if is_sec:
                val = None
            elif isinstance(typ, tuple):
                sg = '' if typ[0] == 'in' else '-'
                val = (f'={sg}ROUND(SUMIFS(资_现流额,资_现流号,{typ[1]},资_月,{m}),2)' if m else f'=SUM(C{r}:N{r})')
            elif typ.startswith('sum:'):
                val = '=' + ''.join(f'{t[0]}{col}{R[int(t[1:])]}' for t in typ[4:].split(','))
                val = val.replace('=+', '=')
            elif typ == 'transit':
                val = (f'=ROUND(SUMIFS(资_净额,资_现金,1,资_现流号,0,资_月,{m}),2)' if m else f'=SUM(C{r}:N{r})')
            elif typ == 'open':
                val = f'=IF({m}<P_建账月,"",{col}{CF_H1})' if m else f'=C{CF_H1}'
            else:                                                   # close
                val = f'=IF({m}<P_建账月,"",{col}{CF_H3})' if m else f'={YC}{R["open"]}+{YC}{R[20]}+{YC}{R["transit"]}'
            put(ws, f'{col}{r}', val, F_TXTB if is_tot else F_TXT, fl, MONEY, AR)
    # ── 隐藏：矩阵现金 ──
    for rr, t in ((CF_H1, '（隐藏）矩阵月初现金'), (CF_H2, '（隐藏）矩阵月末现金'), (CF_H3, '（隐藏）期末现金（算）'),
                  (CF_H4, '（隐藏）期末算－矩阵月末')):
        put(ws, f'A{rr}', t, F_HELP, align=AL, border=False)
    op = '+'.join(f'INDEX(矩_汇期初,{kh(c)})' for c in CASH_CODES)
    cum = lambda m: '+'.join(f'SUM(INDEX(矩_汇借,{kh(c)},1):INDEX(矩_汇借,{kh(c)},{m}))'
                             f'-SUM(INDEX(矩_汇贷,{kh(c)},1):INDEX(矩_汇贷,{kh(c)},{m}))' for c in CASH_CODES)
    for m, col in zip(MONTHS, MC):
        ws[f'{col}{CF_H1}'] = f'=ROUND({op},2)' if m == 1 else f'=ROUND({op}+{cum(m - 1)},2)'
        ws[f'{col}{CF_H2}'] = f'=ROUND({op}+{cum(m)},2)'
        ws[f'{col}{CF_H3}'] = f'=ROUND({col}{CF_H1}+{col}{R[20]}+{col}{R["transit"]},2)'
        ws[f'{col}{CF_H4}'] = f'=ROUND({col}{CF_H3}-{col}{CF_H2},2)'
        for rr in (CF_H1, CF_H2, CF_H3, CF_H4):
            ws[f'{col}{rr}'].font = F_HELP
            ws[f'{col}{rr}'].number_format = MONEY
    sel = f"{SH_BS}!{bsc('选月')}"
    put(ws, f'A{CF_H4 + 1}', '（隐藏）期末现金（选的月）－资产负债表货币资金', F_HELP, align=AL, border=False)
    ws[CF_CHKA.replace('$', '')] = f'=ROUND(INDEX($C${CF_H3}:$N${CF_H3},{sel})-{SH_BS}!$C$6,2)'
    ws[CF_CHKA.replace('$', '')].font = F_HELP
    for rr in (CF_H1, CF_H2, CF_H3, CF_H4, CF_H4 + 1):
        ws.row_dimensions[rr].hidden = True
    h4 = f'$C${CF_H4}:$N${CF_H4}'
    put(ws, f'A{CF_CHKT}', (f'=IF(AND(COUNTIF({h4},"<>0")=0,{CF_CHKA}=0),'
                            f'"√ 每个月：现金流量合计＝货币资金变动；期末现金＝资产负债表货币资金（"&{sel}&" 月）",'
                            f'IF(COUNTIF({h4},"<>0")>0,"✗ 现金流量合计跟货币资金变动对不上，差 "&FIXED(SUM({h4}),2)'
                            f'&"（收支类别的现金流量项目填的不是 1～18 的行次？）",'
                            f'"✗ 期末现金跟资产负债表货币资金差 "&FIXED({CF_CHKA},2)))'),
        F_RED, align=AL, border=False)
    ws.merge_cells(f'A{CF_CHKT}:{YC}{CF_CHKT}')
    ws.conditional_formatting.add(f'A{CF_CHKT}', FormulaRule(formula=[f'LEFT($A${CF_CHKT},1)="√"'],
                                                           font=Font(name=YH, sz=10, bold=True, color='FF548235')))
    put(ws, f'A{CF_SIGN}', SIGN, F_TXT, align=AL, border=False)
    ws.merge_cells(f'A{CF_SIGN}:{YC}{CF_SIGN}')
    _grey_months(ws, f'C4:N{CF_LAST}')
    home_link(ws, 'Q1')
    ws.column_dimensions['Q'].width = 10
    ws.freeze_panes = 'C5'
    print_setup(ws, '4:4', landscape=True)
    ws.page_setup.fitToHeight = 1
    ws.print_area = f'A1:{YC}{CF_SIGN}'
    _dn(wb, '现金流量表_核对', f"{SH_CF}!$A${CF_CHKT}")


def build(wb, ctx=None):
    build_bs(wb, wb[SH_BS])
    build_pl(wb, wb[SH_PL])
    build_cf(wb, wb[SH_CF])
