# -*- coding: utf-8 -*-
"""报表：【资产负债表】会小企01、【利润表】会小企02、【现金流量表】会小企03（直接法）、【经营报表】（矿山指标）。
资产负债表、利润表从【科目余额表】取（按科目编码精确找行）；现金流量表从【现金流水总表】每一笔的现金流量项目汇总。
期间跟【首页】报表年度/月份走。"""
import re

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from common import *
from s_mend import ORE_R0, ORE_R1, O_AE, O_TE, O_UC, O_TIN, O_TS

TBX = {'期末': '余额期末', '年初': '余额年初', '月初': '余额月初'}


def tbv(code, which='期末'):
    """科目余额表里某个科目（含下级）的净额（借正贷负）"""
    return f'N(IFERROR(INDEX({TBX[which]},MATCH("{code}",余额编码,0)),0))'


def tsum(codes, which='期末'):
    return '(' + '+'.join(tbv(c, which) for c in codes) + ')'


def leaf(codes, which, sign):
    """按明细科目的余额方向拆：sign='>0' 取借方余额的，'<0' 取贷方余额的。
    用「直接记在本科目上的」那一列：末级科目＝它的余额；上级科目直接记的部分也单独算一份，不会漏"""
    col = {'期末': '余额期末直接', '年初': '余额年初直接'}[which]
    return '(' + '+'.join(f'SUMIFS({col},余额编码,"{c}*",{col},"{sign}")' for c in codes) + ')'


def act(code, per, sign=1):
    """本年(Y)/本月(M) 发生额：sign=1 借－贷（费用），-1 贷－借（收入）"""
    d, c = ('余额本年借', '余额本年贷') if per == 'Y' else ('余额本月借', '余额本月贷')
    dv = f'N(IFERROR(INDEX({d},MATCH("{code}",余额编码,0)),0))'
    cv = f'N(IFERROR(INDEX({c},MATCH("{code}",余额编码,0)),0))'
    return f'({dv}-{cv})' if sign > 0 else f'({cv}-{dv})'


def acts(codes, per, sign=1):
    return '(' + '+'.join(act(c, per, sign) for c in codes) + ')' if codes else '0'


def _head(ws, last, title_text, tablename, color, tip):
    title(ws, title_text, last, color, tip)
    ws.row_dimensions[3].height = 22
    put(ws, 'A3', '="编制单位："&公司名称', F_TXT, align=AL, border=False)


# ─────────────────────────── 资产负债表 ───────────────────────────
BS_LEFT = [  # (行号, 名称, 行次, 公式生成器(which) 或 'sec' / 'sum:行次…' / None)
    (5, '流动资产：', None, 'sec'),
    (6, '货币资金', 1, lambda w: tsum(['1001', '1002', '1012'], w)),
    (7, '短期投资', 2, lambda w: tsum(['1101'], w)),
    (8, '应收票据', 3, lambda w: tsum(['1121'], w)),
    (9, '应收账款', 4, lambda w: leaf(['1122', '2203'], w, '>0')),
    (10, '预付账款', 5, lambda w: leaf(['1123', '2202'], w, '>0')),
    (11, '应收股利', 6, lambda w: tsum(['1131'], w)),
    (12, '应收利息', 7, lambda w: tsum(['1132'], w)),
    (13, '其他应收款', 8, lambda w: leaf(['1221', '2241'], w, '>0')),
    (14, '存货', 9, lambda w: tsum(['1401', '1402', '1403', '1404', '1405', '1407', '1408', '1411', '1421',
                                   '4001', '4101', '4401', '4403'], w)),
    (15, '　其中：原材料', 10, lambda w: tsum(['1403'], w)),
    (16, '　　　　在产品（生产成本）', 11, lambda w: tsum(['4001', '4101'], w)),
    (17, '　　　　库存商品（原矿）', 12, lambda w: tsum(['1405'], w)),
    (18, '　　　　周转材料', 13, lambda w: tsum(['1411'], w)),
    (19, '其他流动资产', 14, lambda w: tsum(['1901'], w)),
    (20, '流动资产合计', 15, 'sum:6,7,8,9,10,11,12,13,14,19'),
    (21, '非流动资产：', None, 'sec'),
    (22, '长期债券投资', 16, lambda w: tsum(['1501'], w)),
    (23, '长期股权投资', 17, lambda w: tsum(['1511'], w)),
    (24, '固定资产原价', 18, lambda w: tsum(['1601'], w)),
    (25, '减：累计折旧', 19, lambda w: f'-{tsum(["1602"], w)}'),
    (26, '固定资产账面价值', 20, 'expr:C24-C25'),
    (27, '在建工程', 21, lambda w: tsum(['1604'], w)),
    (28, '工程物资', 22, lambda w: tsum(['1605'], w)),
    (29, '固定资产清理', 23, lambda w: tsum(['1606'], w)),
    (30, '生产性生物资产', 24, lambda w: tsum(['1621', '1622'], w)),
    (31, '无形资产（采矿权等）', 25, lambda w: tsum(['1701', '1702'], w)),
    (32, '开发支出', 26, lambda w: tsum(['4301'], w)),
    (33, '长期待摊费用', 27, lambda w: tsum(['1801'], w)),
    (34, '其他非流动资产（临时设施）', 28, lambda w: tsum(['1631', '1632'], w)),
    (35, '非流动资产合计', 29, 'sum:22,23,26,27,28,29,30,31,32,33,34'),
    (37, '资产总计', 30, 'sum:20,35'),
]
BS_RIGHT = [
    (5, '流动负债：', None, 'sec'),
    (6, '短期借款', 31, lambda w: f'-{tsum(["2001"], w)}'),
    (7, '应付票据', 32, lambda w: f'-{tsum(["2201"], w)}'),
    (8, '应付账款', 33, lambda w: f'-{leaf(["2202", "1123"], w, "<0")}'),
    (9, '预收账款', 34, lambda w: f'-{leaf(["2203", "1122"], w, "<0")}'),
    (10, '应付职工薪酬', 35, lambda w: f'-{tsum(["2211"], w)}'),
    (11, '应交税费', 36, lambda w: f'-{tsum(["2221"], w)}'),
    (12, '应付利息', 37, lambda w: f'-{tsum(["2231"], w)}'),
    (13, '应付利润', 38, lambda w: f'-{tsum(["2232"], w)}'),
    (14, '其他应付款', 39, lambda w: f'-{leaf(["2241", "1221"], w, "<0")}'),
    (15, '其他流动负债', 40, lambda w: '0'),
    (16, '流动负债合计', 41, 'sum:6,7,8,9,10,11,12,13,14,15'),
    (17, '非流动负债：', None, 'sec'),
    (18, '长期借款', 42, lambda w: f'-{tsum(["2501"], w)}'),
    (19, '长期应付款', 43, lambda w: f'-{tsum(["2701"], w)}'),
    (20, '递延收益', 44, lambda w: f'-{tsum(["2401"], w)}'),
    (21, '其他非流动负债', 45, lambda w: '0'),
    (22, '非流动负债合计', 46, 'sum:18,19,20,21'),
    (23, '负债合计', 47, 'sum:16,22'),
    (25, '所有者权益（或股东权益）：', None, 'sec'),
    (26, '实收资本（或股本）', 48, lambda w: f'-{tsum(["3001"], w)}'),
    (27, '资本公积', 49, lambda w: f'-{tsum(["3002"], w)}'),
    (28, '盈余公积', 50, lambda w: f'-{tsum(["3101"], w)}'),
    (29, '未分配利润', 51, lambda w: (f'-({tsum(["3103", "3104", "6000"], w)}'
                                       f'+SUMIFS({TBX[w]},余额编码,"5*",余额级次,1))')),
    (30, '所有者权益（或股东权益）合计', 52, 'sum:26,27,28,29'),
    (37, '负债和所有者权益（或股东权益）总计', 53, 'sum:23,30'),
]


def build_bs(wb, ctx):
    ws = wb.create_sheet(SH_BAL)
    widths(ws, {'A': 30, 'B': 6, 'C': 19, 'D': 19, 'E': 34, 'F': 6, 'G': 19, 'H': 19})
    _head(ws, 'H', '资产负债表', '会小企01表', C_RPT,
          '💡 按【首页】报表年度/月份：期末余额＝报表月末，年初余额＝上年 12 月 31 日。数全部来自【科目余额表】。'
          '应收/预收、应付/预付、其他应收/其他应付按明细科目的余额方向拆开列。损益不结转（表结法），未分配利润＝以前年度累计＋本年利润自动算。')
    put(ws, 'D3', f'={REP_M1}', F_TXTB, None, 'yyyy"年"m"月"d"日"', AC, border=False)
    put(ws, 'G3', '会小企01表', F_NOTE, align=AR_, border=False)
    put(ws, 'H3', '单位：苏姆', F_NOTE, align=AR_, border=False)
    header(ws, 4, [('A', '资    产'), ('B', '行次'), ('C', '期末余额'), ('D', '年初余额'),
                   ('E', '负债和所有者权益（或股东权益）'), ('F', '行次'), ('G', '期末余额'), ('H', '年初余额')], C_RPT, height=26)
    done = set()
    for side, (cl, cn, cv, cy) in ((BS_LEFT, ('A', 'B', 'C', 'D')), (BS_RIGHT, ('E', 'F', 'G', 'H'))):
        for r, name, no, f in side:
            done.update(f'{c}{r}' for c in (cl, cn, cv, cy))
            is_sec = f == 'sec'
            is_tot = isinstance(f, str) and (f.startswith('sum:') or f.startswith('expr:'))
            put(ws, f'{cl}{r}', name, F_TXTB if (is_sec or is_tot) else F_TXT, FILL_SUBH if is_tot else None,
                align=AL)
            put(ws, f'{cn}{r}', no, F_TXT, FILL_SUBH if is_tot else None, align=AC)
            for col, w in ((cv, '期末'), (cy, '年初')):
                if is_sec:
                    val = None
                elif isinstance(f, str) and f.startswith('sum:'):
                    val = '=' + '+'.join(f'{col}{x}' for x in f[4:].split(','))
                elif isinstance(f, str) and f.startswith('expr:'):
                    val = '=' + f[5:].replace('C', col)
                else:
                    val = f'=ROUND({f(w)},2)'
                put(ws, f'{col}{r}', val, F_TXTB if is_tot else F_TXT, FILL_SUBH if is_tot else None, MONEY, AR_)
    for r in range(5, 38):
        for c in 'ABCDEFGH':
            if f'{c}{r}' not in done:
                put(ws, f'{c}{r}', None)
    r = 39
    put(ws, f'A{r}', '平衡检查：', F_TXTB, align=AR_, border=False)
    put(ws, f'C{r}', '=IF(AND(ROUND(C37-G37,2)=0,ROUND(D37-H37,2)=0),"√ 资产＝负债＋所有者权益",'
                     '"✗ 不平：期末差 "&FIXED(C37-G37,2)&"，年初差 "&FIXED(D37-H37,2))', F_RED, align=AL,
        border=False)
    ws.merge_cells(f'C{r}:H{r}')
    put(ws, f'A{r + 1}', '单位负责人：　　　　　　　　财务主管：　　　　　　　　复核：　　　　　　　　制表：', F_TXT, align=AL, border=False)
    ws.merge_cells(f'A{r + 1}:H{r + 1}')
    ws.freeze_panes = 'A5'
    ws.print_area = f'A1:H{r + 1}'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ctx['bs_cash'] = f"{q(SH_BAL)}!$C$6"
    return ws


# ─────────────────────────── 利润表 ───────────────────────────
PL_ROWS = [  # (名称, 行次, 收入类科目(贷-借), 费用类科目(借-贷)) 或 合计公式
    ('一、营业收入', 1, ['5001', '5051'], []),
    ('减：营业成本', 2, [], ['5401', '5402']),
    ('　　税金及附加', 3, [], ['5403']),
    ('　　其中：消费税', 4, [], ['540301']),
    ('　　　　　营业税', 5, [], []),
    ('　　　　　城市维护建设税', 6, [], ['540303']),
    ('　　　　　资源税（含地下资源使用税、水资源使用费）', 7, [], ['540304', '540316', '540317']),
    ('　　　　　土地增值税', 8, [], ['540305']),
    ('　　　　　城镇土地使用税、房产税、车船税、印花税（含财产税、土地税）', 9, [],
     ['540306', '540307', '540308', '540309', '540318', '540319']),
    ('　　　　　教育费附加、矿产资源补偿费、排污费', 10, [], ['540310', '540311', '540312', '540313']),
    ('　　销售费用', 11, [], ['5601']),
    ('　　其中：商品维修费', 12, [], ['560115']),
    ('　　　　　广告费和业务宣传费', 13, [], ['560116']),
    ('　　管理费用', 14, [], ['5602']),
    ('　　其中：开办费', 15, [], []),
    ('　　　　　业务招待费', 16, [], ['560205']),
    ('　　　　　研究费用', 17, [], []),
    ('　　财务费用', 18, [], ['5603']),
    ('　　其中：利息费用（收入以“-”号填列）', 19, [], ['560301']),
    ('加：投资收益（损失以“-”号填列）', 20, ['5111'], []),
    ('二、营业利润（亏损以“-”号填列）', 21, 'R1-R2-R3-R11-R14-R18+R20', None),
    ('加：营业外收入', 22, ['5301'], []),
    ('　　其中：政府补助', 23, ['530102'], []),
    ('减：营业外支出', 24, [], ['5711']),
    ('　　其中：坏账损失', 25, [], ['571103']),
    ('　　　　　无法收回的长期债券投资损失', 26, [], ['571104']),
    ('　　　　　无法收回的长期股权投资损失', 27, [], ['571105']),
    ('　　　　　自然灾害等不可抗力因素造成的损失', 28, [], ['571106']),
    ('　　　　　税收滞纳金', 29, [], ['571107']),
    ('三、利润总额（亏损总额以“-”号填列）', 30, 'R21+R22-R24', None),
    ('减：所得税费用', 31, [], ['5801']),
    ('四、净利润（净亏损以“-”号填列）', 32, 'R30-R31', None),
]
PL_R0 = 5


def pl_row(no):
    return PL_R0 + no - 1


def build_pl(wb, ctx):
    ws = wb.create_sheet(SH_PL)
    widths(ws, {'A': 52, 'B': 6, 'C': 20, 'D': 20})
    _head(ws, 'D', '利润表', '会小企02表', C_RPT,
          '💡 按【首页】报表年度/月份：本年累计＝1 月到报表月，本月＝报表月。数来自【科目余额表】损益类科目的发生额。'
          '业务模块启用日以前的月份跟你原来 8 月报表一个口径（付款即记成本费用）。')
    put(ws, 'C3', f'=报表年度&"年1-"&报表月份&"月"', F_TXTB, align=AC, border=False)
    put(ws, 'D3', '会小企02表　单位：苏姆', F_NOTE, align=AR_, border=False)
    header(ws, 4, [('A', '项    目'), ('B', '行次'), ('C', '本年累计金额'), ('D', '本月金额')], C_RPT, height=26)
    for name, no, rev, exp in PL_ROWS:
        r = pl_row(no)
        tot = isinstance(rev, str)
        put(ws, f'A{r}', name, F_TXTB if (tot or not name.startswith('　')) else F_TXT, FILL_SUBH if tot else None, align=AL)
        put(ws, f'B{r}', no, F_TXT, FILL_SUBH if tot else None, align=AC)
        for col, per in (('C', 'Y'), ('D', 'M')):
            if tot:
                f = re.sub(r'R(\d+)', lambda m: f'{col}{pl_row(int(m.group(1)))}', rev)
                val = '=' + f
            elif not rev and not exp:
                val = 0
            else:
                parts = []
                if rev:
                    parts.append(acts(rev, per, -1))
                if exp:
                    parts.append(acts(exp, per, 1))
                val = f'=ROUND({"+".join(parts)},2)'
            put(ws, f'{col}{r}', val, F_TXTB if tot else F_TXT, FILL_SUBH if tot else None, MONEY, AR_)
    r = pl_row(32) + 2
    put(ws, f'A{r}', '核对：损益类科目发生额合计（贷－借）', F_NOTE, align=AL, border=False)
    for col, (d, c) in (('C', ('余额本年借', '余额本年贷')), ('D', ('余额本月借', '余额本月贷'))):
        put(ws, f'{col}{r}', f'=ROUND(SUMIFS({c},余额编码,"5*",余额级次,1)-SUMIFS({d},余额编码,"5*",余额级次,1),2)',
            F_NOTE, fmt=MONEY, align=AR_, border=False)
    put(ws, f'A{r + 1}', '=IF(AND(ROUND(C{0}-C{1},2)=0,ROUND(D{0}-D{1},2)=0),"√ 跟净利润一致","✗ 跟净利润对不上：有损益科目没列进报表")'
        .format(r, pl_row(32)), F_RED, align=AL, border=False)
    put(ws, f'A{r + 3}', '单位负责人：　　　　　　财务主管：　　　　　　复核：　　　　　　制表：', F_TXT, align=AL, border=False)
    ws.merge_cells(f'A{r + 3}:D{r + 3}')
    ws.freeze_panes = 'A5'
    ws.print_area = f'A1:D{r + 3}'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ctx['pl_net'] = (f"{q(SH_PL)}!$C${pl_row(32)}", f"{q(SH_PL)}!$D${pl_row(32)}")
    ctx['pl_rev'] = (f"{q(SH_PL)}!$C${pl_row(1)}", f"{q(SH_PL)}!$D${pl_row(1)}")
    ctx['pl_cost'] = (f"{q(SH_PL)}!$C${pl_row(2)}", f"{q(SH_PL)}!$D${pl_row(2)}")
    ctx['pl_exp'] = tuple(f"({q(SH_PL)}!${c}${pl_row(11)}+{q(SH_PL)}!${c}${pl_row(14)}+{q(SH_PL)}!${c}${pl_row(18)})" for c in 'CD')
    ctx['pl_tax'] = (f"{q(SH_PL)}!$C${pl_row(3)}", f"{q(SH_PL)}!$D${pl_row(3)}")
    ctx['pl_chk'] = f"{q(SH_PL)}!$A${r + 1}"
    return ws


# ─────────────────────────── 现金流量表 ───────────────────────────
CF_R0 = 5
CF_LAYOUT = [  # (名称, 行次, 类型)  类型：'sec' 标题；('in'/'out', 项目号)；'sum:…'
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
    ('加：期初现金余额', 21, 'open'),
    ('加：汇率变动、账户间划转及换汇差额对现金的影响', None, 'fx'),
    ('五、期末现金余额', 22, 'close'),
]
CASH_CODES = ['1001', '1002', '1012']


def build_cf(wb, ctx):
    ws = wb.create_sheet(SH_CF)
    widths(ws, {'A': 54, 'B': 6, 'C': 20, 'D': 20})
    _head(ws, 'D', '现金流量表', '会小企03表', C_RPT,
          '💡 直接法：【现金流水总表】每一笔按收支类别在【记账规则】里设的现金流量项目归类（收钱按「现金流量(收)」、付钱按「现金流量(付)」），'
          '账户之间转钱、换汇（项目 0）不算流入流出。月末外币调汇、转账/换汇折苏姆的差额单列一行「汇率变动…的影响」，'
          '所以期末现金余额正好等于资产负债表的货币资金。')
    put(ws, 'C3', f'=报表年度&"年1-"&报表月份&"月"', F_TXTB, align=AC, border=False)
    put(ws, 'D3', '会小企03表　单位：苏姆', F_NOTE, align=AR_, border=False)
    header(ws, 4, [('A', '项    目'), ('B', '行次'), ('C', '本年累计金额'), ('D', '本月金额')], C_RPT, height=26)
    rows = {}
    per = {'C': (REP_Y0, REP_M1), 'D': (REP_M0, REP_M1)}
    for i, (name, no, typ) in enumerate(CF_LAYOUT):
        r = CF_R0 + i
        if no:
            rows[no] = r
        if typ in ('open', 'fx', 'close'):
            rows[typ] = r
        is_tot = isinstance(typ, str) and typ not in ('sec',)
        put(ws, f'A{r}', name, F_TXTB if (typ == 'sec' or is_tot) else F_TXT, FILL_SUBH if is_tot else None, align=AL)
        put(ws, f'B{r}', no, F_TXT, FILL_SUBH if is_tot else None, align=AC)
        for col, (d0, d1) in per.items():
            dr = f'{cash(K_DATE)},">="&{d0},{cash(K_DATE)},"<="&{d1}'
            if typ == 'sec':
                val = None
            elif isinstance(typ, tuple):
                sgn = '' if typ[0] == 'in' else '-'
                val = f'={sgn}SUMIFS({cash(K_UZS)},{cash(K_CFI)},{typ[1]},{dr})'
            elif typ.startswith('sum:'):
                val = '=' + ''.join(f'{t[0]}{col}{rows[int(t[1:])]}' for t in typ[4:].split(','))
            elif typ == 'open':
                w = '年初' if col == 'C' else '月初'
                val = f'=ROUND({tsum(CASH_CODES, w)},2)'
            elif typ == 'fx':
                # 调汇凭证 + 手工凭证 里记到现金科目的 + 转账/换汇（项目 0）折苏姆的净额
                parts = []
                for c in CASH_CODES:
                    for src in ('调汇', '手工'):
                        parts.append(f'SUMIFS(分录借方,分录来源,"{src}",分录科目,"{c}*",分录日期,">="&{d0},分录日期,"<="&{d1})'
                                     f'-SUMIFS(分录贷方,分录来源,"{src}",分录科目,"{c}*",分录日期,">="&{d0},分录日期,"<="&{d1})')
                parts.append(f'SUMIFS({cash(K_UZS)},{cash(K_CFI)},0,{dr})')
                val = '=ROUND(' + '+'.join(parts) + ',2)'
            elif typ == 'close':
                val = f'={col}{rows[20]}+{col}{rows[21]}+{col}{r - 1}'
            put(ws, f'{col}{r}', val, F_TXTB if is_tot else F_TXT, FILL_SUBH if is_tot else None, MONEY, AR_)
    r = CF_R0 + len(CF_LAYOUT) + 1
    close = CF_R0 + len(CF_LAYOUT) - 1
    put(ws, f'A{r}', f'=IF(ROUND(C{close}-{ctx["bs_cash"]},2)=0,"√ 期末现金余额＝资产负债表货币资金",'
                     f'"✗ 跟资产负债表货币资金差 "&FIXED(C{close}-{ctx["bs_cash"]},2)&"（有别的凭证直接记了现金科目？）")',
        F_RED, align=AL, border=False)
    put(ws, f'A{r + 2}', '单位负责人：　　　　　　财务主管：　　　　　　复核：　　　　　　制表：', F_TXT, align=AL, border=False)
    ws.merge_cells(f'A{r + 2}:D{r + 2}')
    ws.freeze_panes = 'A5'
    ws.print_area = f'A1:D{r + 2}'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ctx['cf_rows'] = rows
    ctx['cf_close'] = close
    ctx['cf_chk'] = f"{q(SH_CF)}!$A${r}"
    return ws


# ─────────────────────────── 经营报表（矿山） ───────────────────────────
def build_ops(wb, ctx):
    ws = wb.create_sheet(SH_OPS)
    widths(ws, {'A': 30, 'B': 12, 'C': 19, 'D': 19, 'E': 44})
    title(ws, '经营报表（矿山 · 本月 / 本年累计 · 苏姆）', 'E', C_OPS,
          '💡 采矿主要指标：出矿量、掘进量、品位取【产量登记】；销售量取【销售结算】；收入、成本、费用、利润取【利润表】；现金取【现金流量表】【资产负债表】。'
          '还没出矿的月份，吨矿成本这些显示「-」。下面还有一张报表年度 12 个月的走势表。')
    put(ws, 'A3', '="编制单位："&公司名称&"　　期间："&报表年度&"年"&报表月份&"月"', F_TXTB, align=AL, border=False)
    ws.merge_cells('A3:E3')
    header(ws, 4, [('A', '指    标'), ('B', '单位'), ('C', '本月'), ('D', '本年累计'), ('E', '说明 / 取数')], C_OPS, height=26)
    pd = lambda col: mod(SH_PROD, col)
    sl = lambda col: mod(SH_SALE, col)
    per = {'C': (REP_M0, REP_M1), 'D': (REP_Y0, REP_M1)}
    PLc = {'C': 1, 'D': 0}          # ctx 里报表格子：(本年, 本月)
    rowof = {}

    def rep(key, col):
        return ctx[key][PLc[col]]

    def cf(no, col):
        return f"{q(SH_CF)}!${'C' if col == 'D' else 'D'}${ctx['cf_rows'][no]}"

    items = [
        ('sec', '一、生产'),
        ('出矿量（原矿）', '吨', lambda c, d0, d1: f'SUMIFS({pd(PD_TON)},{pd(PD_DATE)},">="&{d0},{pd(PD_DATE)},"<="&{d1})', '【产量登记】出矿量'),
        ('掘进量', '米', lambda c, d0, d1: f'SUMIFS({pd(PD_DEV)},{pd(PD_DATE)},">="&{d0},{pd(PD_DATE)},"<="&{d1})', '【产量登记】掘进量'),
        ('平均出矿品位', '%', lambda c, d0, d1: (f'IF(N({c}{rowof["出矿量（原矿）"]})=0,"-",SUMPRODUCT(({pd(PD_DATE)}>={d0})'
                                                  f'*({pd(PD_DATE)}<={d1})*ISNUMBER({pd(PD_GRADE)})*ISNUMBER({pd(PD_TON)}),'
                                                  f'{pd(PD_GRADE)},{pd(PD_TON)})/{c}{rowof["出矿量（原矿）"]})'), '按出矿量加权'),
        ('生产成本发生额', '苏姆', lambda c, d0, d1: (f'SUMIFS(分录借方,分录科目,"4001*",分录日期,">="&{d0},分录日期,"<="&{d1})'
                                                   f'+SUMIFS(分录借方,分录科目,"4101*",分录日期,">="&{d0},分录日期,"<="&{d1})'
                                                   f'-SUMIFS(分录借方,分录来源,"成本",分录科目,"4*",分录日期,">="&{d0},分录日期,"<="&{d1})'),
         '炸药、油料、配件、人工、折旧等记进生产成本的'),
        ('吨矿生产成本', '苏姆/吨', 'ratio:成本:出矿', '生产成本发生额 ÷ 出矿量'),
        ('sec', '二、销售'),
        ('销售量（原矿）', '吨', lambda c, d0, d1: (f'SUMIFS({sl(S_TON)},{sl(S_VALID)},1,{sl(S_DATE)},">="&{d0},{sl(S_DATE)},"<="&{d1})'),
         '【销售结算】'),
        ('原矿销售收入（不含税）', '苏姆', lambda c, d0, d1: act('500103', 'M' if c == 'C' else 'Y', -1), '主营业务收入-原矿销售'),
        ('平均售价', '苏姆/吨', 'ratio:原矿收入:销量', '原矿销售收入 ÷ 销售量'),
        ('原矿销售成本', '苏姆', lambda c, d0, d1: act('540101', 'M' if c == 'C' else 'Y', 1), '主营业务成本-原矿销售成本'),
        ('吨矿销售成本', '苏姆/吨', 'ratio:原矿成本:销量', '原矿销售成本 ÷ 销售量'),
        ('原矿毛利', '苏姆', 'expr:原矿收入-原矿成本', '原矿销售收入 − 原矿销售成本'),
        ('原矿毛利率', '%', 'pct:毛利:原矿收入', ''),
        ('营业收入（全部）', '苏姆', lambda c, d0, d1: rep('pl_rev', c), '【利润表】营业收入（含启用日以前的采矿服务收入、其他业务收入）'),
        ('营业成本（全部）', '苏姆', lambda c, d0, d1: rep('pl_cost', c), '【利润表】'),
        ('sec', '三、费用与利润'),
        ('税金及附加', '苏姆', lambda c, d0, d1: rep('pl_tax', c), ''),
        ('期间费用（销售＋管理＋财务）', '苏姆', lambda c, d0, d1: rep('pl_exp', c), ''),
        ('吨矿完全成本', '苏姆/吨', 'ratio3:原矿成本+税金+期间费用:销量', '（原矿销售成本＋税金＋期间费用）÷ 销售量；期间费用是全公司的，看本月那列'),
        ('净利润', '苏姆', lambda c, d0, d1: rep('pl_net', c), '【利润表】'),
        ('sec', '四、现金'),
        ('经营活动现金净流量', '苏姆', lambda c, d0, d1: cf(7, c), '【现金流量表】'),
        ('投资活动现金净流量', '苏姆', lambda c, d0, d1: cf(13, c), '买设备、基建、采矿权'),
        ('筹资活动现金净流量', '苏姆', lambda c, d0, d1: cf(19, c), '股东投资、借款'),
        ('期末货币资金', '苏姆', lambda c, d0, d1: ctx['bs_cash'], '【资产负债表】'),
        ('月均现金净流出（经营＋投资）', '苏姆/月', 'burn', '本年经营＋投资净流出 ÷ 月数'),
        ('现有现金可维持', '个月', 'months', '期末货币资金 ÷ 月均净流出'),
        ('sec', '五、资产与往来（报表月末）'),
        ('应收甲方（应收账款-原矿销售）', '苏姆', lambda c, d0, d1: tbv('112203'), '销售结算记了、甲方还没付的'),
        ('应付供应商（应付账款）', '苏姆', lambda c, d0, d1: f'-{tbv("2202")}', ''),
        ('固定资产原值', '苏姆', lambda c, d0, d1: tbv('1601'), '矿车、铲车、四不像…'),
        ('累计折旧', '苏姆', lambda c, d0, d1: f'-{tbv("1602")}', ''),
        ('在建工程（井巷、地面工程）', '苏姆', lambda c, d0, d1: tbv('1604'), ''),
        ('采矿权及无形资产', '苏姆', lambda c, d0, d1: tbv('1701'), '67号矿井溢价款'),
        ('原矿结存', '吨', lambda c, d0, d1: f'IFERROR(INDEX({rng(SH_ORE, O_TE, ORE_R0, ORE_R1)},{REP_N}),0)', '【原矿产销存】'),
        ('材料库存（原材料＋周转材料）', '苏姆', lambda c, d0, d1: f'{tbv("1403")}+{tbv("1411")}', '【存货收发存】'),
        ('sec', '六、人工'),
        ('计提工资', '苏姆', lambda c, d0, d1: (f'SUMIFS({mod(SH_PAY, W_AMT)},{mod(SH_PAY, W_VALID)},1,{mod(SH_PAY, W_DATE)},">="&{d0},'
                                              f'{mod(SH_PAY, W_DATE)},"<="&{d1})'), '【工资计提】（启用日以前没计提）'),
        ('支付工资（流水）', '苏姆', lambda c, d0, d1: (f'-SUMIFS({cash(K_UZS)},{cash(K_CAT)},"应付职工薪酬",{cash(K_DATE)},">="&{d0},'
                                                  f'{cash(K_DATE)},"<="&{d1})'), '【现金流水】「应付职工薪酬」'),
        ('在册人数（报表月）', '人', lambda c, d0, d1: (f'SUMIFS({mod(SH_PAY, W_HEAD)},{mod(SH_PAY, W_VALID)},1,{mod(SH_PAY, W_DATE)},">="&{REP_M0},'
                                                  f'{mod(SH_PAY, W_DATE)},"<="&{REP_M1})'), '【工资计提】人数'),
    ]
    r = 5
    for it in items:
        if it[0] == 'sec':
            put(ws, f'A{r}', it[1], F_SEC, FILL_SUBH, align=AL)
            for c in 'BCDE':
                put(ws, f'{c}{r}', None, fill_=FILL_SUBH)
            r += 1
            continue
        name = it[0]
        rowof[name] = r
        r += 1
    alias = {'成本': '生产成本发生额', '出矿': '出矿量（原矿）', '原矿收入': '原矿销售收入（不含税）', '销量': '销售量（原矿）',
             '原矿成本': '原矿销售成本', '毛利': '原矿毛利', '税金': '税金及附加', '期间费用': '期间费用（销售＋管理＋财务）'}
    for it in items:
        if it[0] == 'sec':
            continue
        name, unit, f, note = it
        r = rowof[name]
        put(ws, f'A{r}', name, F_TXT, align=AL)
        put(ws, f'B{r}', unit, F_TXT, align=AC)
        put(ws, f'E{r}', note or None, F_NOTE, align=AL)
        for col, (d0, d1) in per.items():
            R = lambda k: f'{col}{rowof[alias[k]]}'
            if callable(f):
                val = '=' + f(col, d0, d1)
            elif f.startswith('ratio:'):
                a, b = f[6:].split(':')
                val = f'=IF(N({R(b)})=0,"-",ROUND({R(a)}/{R(b)},2))'
            elif f.startswith('ratio3:'):
                a, b = f[7:].split(':')
                num = '+'.join(R(x) for x in a.split('+'))
                val = f'=IF(N({R(b)})=0,"-",ROUND(({num})/{R(b)},2))'
            elif f.startswith('pct:'):
                a, b = f[4:].split(':')
                val = f'=IF(N({R(b)})=0,"-",{R(a)}/{R(b)})'
            elif f.startswith('expr:'):
                a, b = f[5:].split('-')
                val = f'={R(a)}-{R(b)}'
            elif f == 'burn':
                e, v = rowof['经营活动现金净流量'], rowof['投资活动现金净流量']
                val = f'=IF(-(D{e}+D{v})<=0,"-",-(D{e}+D{v})/报表月份)' if col == 'D' else '="-"'
            elif f == 'months':
                b_, c_ = rowof['月均现金净流出（经营＋投资）'], rowof['期末货币资金']
                val = f'=IF(ISNUMBER(D{b_}),ROUND(D{c_}/D{b_},1),"-")' if col == 'D' else '="-"'
            fmt = {'吨': QTY2, '米': '#,##0.0;-#,##0.0;"-"', '%': '0.00', '人': '#,##0;-#,##0;"-"', '个月': '0.0'}.get(unit, MONEY)
            if name == '原矿毛利率':
                fmt = PCT
            put(ws, f'{col}{r}', val, F_TXTB if name in ('净利润', '出矿量（原矿）', '期末货币资金') else F_TXT, None, fmt, AR_)
    # 余额类指标本月/本年相同：本年列写「同左」
    last = r
    # ── 12 个月走势 ──
    t0 = max(rowof.values()) + 3
    put(ws, f'A{t0 - 1}', '七、报表年度 12 个月走势', F_SEC, border=False)
    cols = ['月份', '出矿(吨)', '掘进(米)', '销售(吨)', '营业收入', '营业成本', '期间费用', '净利润', '经营现金净流量',
            '资本性支出', '月末货币资金']
    letters = [CL(i + 1) for i in range(len(cols))]
    for L, t in zip(letters, cols):
        put(ws, f'{L}{t0}', t, F_HDR, fill(C_OPS), align=ACW)
        if L not in ('A', 'B', 'C', 'D', 'E'):
            ws.column_dimensions[L].width = 17
    ws.row_dimensions[t0].height = 30

    def jsum(side, code, d0, d1, src=None):
        col = '分录借方' if side == 'D' else '分录贷方'
        s = f',分录来源,"{src}"' if src else ''
        return f'SUMIFS({col},分录科目,"{code}*",分录日期,">="&{d0},分录日期,"<="&{d1}{s})'

    for m in range(1, 13):
        r = t0 + m
        d0, d1 = f'DATE(报表年度,{m},1)', f'DATE(报表年度,{m}+1,0)'
        put(ws, f'A{r}', f'={d0}', F_TXT, None, MONTH, AC)
        dr = lambda rg: f'{rg},">="&{d0},{rg},"<="&{d1}'
        vals = [
            f'=SUMIFS({pd(PD_TON)},{dr(pd(PD_DATE))})',
            f'=SUMIFS({pd(PD_DEV)},{dr(pd(PD_DATE))})',
            f'=SUMIFS({sl(S_TON)},{sl(S_VALID)},1,{dr(sl(S_DATE))})',
            '=' + '+'.join(f'({jsum("C", c, d0, d1)}-{jsum("D", c, d0, d1)})' for c in ('5001', '5051')),
            '=' + '+'.join(f'({jsum("D", c, d0, d1)}-{jsum("C", c, d0, d1)})' for c in ('5401', '5402')),
            '=' + '+'.join(f'({jsum("D", c, d0, d1)}-{jsum("C", c, d0, d1)})' for c in ('5601', '5602', '5603')),
            f'={jsum("C", "5", d0, d1)}-{jsum("D", "5", d0, d1)}',
            f'=SUMIFS({cash(K_UZS)},{cash(K_CFI)},">=1",{cash(K_CFI)},"<=6",{dr(cash(K_DATE))})',
            f'=-SUMIFS({cash(K_UZS)},{cash(K_CFI)},12,{dr(cash(K_DATE))})',
            '=ROUND(' + '+'.join(f'SUMIF({acc(AC_CODE)},"{c}*",{acc(AC_OPEN)})+SUMIFS(分录借方,分录科目,"{c}*",分录日期,"<="&{d1})'
                                 f'-SUMIFS(分录贷方,分录科目,"{c}*",分录日期,"<="&{d1})' for c in CASH_CODES) + ',2)',
        ]
        for L, v in zip(letters[1:], vals):
            fmt = QTY2 if L in ('B', 'D') else ('#,##0.0;-#,##0.0;"-"' if L == 'C' else MONEY)
            put(ws, f'{L}{r}', v, F_TXT, None, fmt, AR_)
    r = t0 + 13
    put(ws, f'A{r}', '合计', F_TXTB, FILL_TOT, align=AC)
    for L in letters[1:]:
        f = f'=SUM({L}{t0 + 1}:{L}{t0 + 12})' if L != letters[-1] else None
        put(ws, f'{L}{r}', f, F_TXTB, FILL_TOT, QTY2 if L in ('B', 'D') else ('#,##0.0' if L == 'C' else MONEY), AR_)
    ws.conditional_formatting.add(f'A{t0 + 1}:{letters[-1]}{t0 + 12}',
                                  FormulaRule(formula=[f'MONTH($A{t0 + 1})=报表月份'], fill=fill('FFFFF2CC')))
    ws.freeze_panes = 'A5'
    ctx['ops_rows'] = rowof
    return ws
