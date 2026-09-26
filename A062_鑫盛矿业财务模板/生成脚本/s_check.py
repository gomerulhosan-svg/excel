# -*- coding: utf-8 -*-
"""【数据校验】【首页】【使用说明】"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from common import *
from s_mend import ORE_R0, ORE_R1, O_AE, O_TE
from s_reports import tbv
from s_je import VL, VL_R1

CODES = acc(AC_CODE)


def _cnt(sh, col, pre='✗'):
    return f'COUNTIF({mod(sh, col)},"{pre}*")'


def _ok0(n, bad, good='√ 没有'):
    """n=0 → good；否则 bad & n"""
    return f'=IF({n}=0,"{good}","{bad}"&{n}&" 行")'


def build_check(wb, ctx):
    ws = wb.create_sheet(SH_CHK)
    widths(ws, {'A': 5, 'B': 34, 'C': 46, 'D': 60})
    title(ws, '数据校验（打开先看这张：✗ 要改，⚠ 要看一眼）', 'D', C_CHK,
          '💡 全部自动检查，期间跟【首页】报表年度/月份走。「结果」列是 ✗ 的，按右边「怎么处理」去改；⚠ 是提醒，不一定是错。')
    ws.row_dimensions[3].height = 26
    for c in 'ABCD':
        put(ws, f'{c}3', None, fill_=FILL_KPI, border=False)
    header(ws, 4, [('A', '序'), ('B', '检查项目'), ('C', '结果'), ('D', '怎么处理')], C_CHK, height=24)
    R0 = 5
    ar, ae = ctx['acct_rows']
    ir0, ir1 = ctx['inv_rows']
    cash_m = cash(K_UZS)
    rate_miss = '+'.join(f'IF(COUNTIF(账户币种,"{c}")>0,IF(N(INDEX(记账汇率区,{REP_N},{i + 1}))=0,1,0),0)'
                         for i, c in enumerate(RATE_CURS) if c != 'UZS')
    fa_cost = f'SUMIFS({mod(SH_FA, F_COST)},{mod(SH_FA, F_DATE)},"<="&{REP_M1})'
    checks = [
        ('现金流水：有 ✗ 的行', _ok0(f'COUNTIF({cash(K_CHK)},"✗*")', '✗ '),
         '到【现金流水总表】最右边「核对」列筛选 ✗，按提示改（✗ 的行不生成凭证）'),
        ('现金流水：有 ⚠ 的行', _ok0(f'COUNTIF({cash(K_CHK)},"⚠*")', '⚠ '),
         '多数是原表本来就缺金额/摘要的行（备注写着【待补】），补上就好'),
        ('现金流水：有金额但没生成凭证', _ok0(f'(COUNTIF({cash_m},">0")+COUNTIF({cash_m},"<0")-SUM({cash(K_VALID)}))', '✗ '),
         '一般是类别没选、类别不在【记账规则】里、或对方科目不在科目表里——看那几行的「核对」列'),
        ('采购入库：有 ✗ 的行', _ok0(_cnt(SH_BUY, B_CHK), '✗ '), '到【采购入库】「校验」列筛 ✗'),
        ('领用出库：有 ✗ 的行', _ok0(_cnt(SH_ISS, I_CHK), '✗ '), '到【领用出库】「校验」列筛 ✗'),
        ('产量登记：有 ✗ 的行', _ok0(_cnt(SH_PROD, PD_CHK), '✗ '), '到【产量登记】「校验」列筛 ✗'),
        ('销售结算：有 ✗ 的行', _ok0(_cnt(SH_SALE, S_CHK), '✗ '), '到【销售结算】「校验」列筛 ✗'),
        ('工资计提：有 ✗ 的行', _ok0(_cnt(SH_PAY, W_CHK), '✗ '), '到【工资计提】「校验」列筛 ✗'),
        ('固定资产：有 ✗ 的卡片', _ok0(_cnt(SH_FA, F_CHK), '✗ '), '到【固定资产】「校验」列筛 ✗（✗ 的卡片不提折旧）'),
        ('手工凭证：有 ✗ 的行', _ok0(_cnt(SH_MAN, MN_CHK), '✗ '), '到【手工凭证】「校验」列筛 ✗（一张凭证有一行 ✗ 整张不记账）'),
        ('记账分录：全部借方＝贷方', '=IF(ROUND(SUM(分录借方)-SUM(分录贷方),2)=0,"√ 平衡","✗ 差 "&FIXED(SUM(分录借方)-SUM(分录贷方),2))',
         '不平说明模板公式被改动过，找我'),
        ('本月每张凭证借贷平衡', _ok0(f'COUNTIF({rng(SH_VLIST, VL["ok"], VL_R0, VL_R1)},"✗*")', '✗ ', '√ 都平'),
         '到【凭证汇总】看「借贷平衡」列'),
        ('科目余额表借贷平衡', f'={q(SH_TB)}!F3', '看【科目余额表】第 3 行提示'),
        ('资产负债表平衡', f'={q(SH_BAL)}!C39', '不平一般是科目表里加了新的一级科目，报表没归进去——找我'),
        ('利润表＝损益科目合计', f'={ctx["pl_chk"]}', ''),
        ('现金流量表期末＝货币资金', f'={ctx["cf_chk"]}', ''),
        ('资金账户：原币余额×月末汇率＝账上苏姆',
         f'=IF(ROUND(SUMIF({q(SH_ACCT)}!$K${ar}:$K${ae},">0")-SUMIF({q(SH_ACCT)}!$K${ar}:$K${ae},"<0"),2)=0,"√ 一致",'
         f'"✗ 有账户对不上，看【资金账户余额表】K 列")',
         '多半是流水有 ✗ 的行没记账，或这个月的汇率没填（月末调汇没做）'),
        ('报表月的外币汇率', f'=IF(({rate_miss})=0,"√ 都填了","✗ 报表月有外币的汇率没填")', '到【月度汇率】把这个月的人民币、美元汇率填上'),
        ('资金划转待对冲（122104）余额',
         f'=IF(ROUND({tbv("122104")},2)=0,"√ 0（账户之间的转账都对上了）","⚠ 余额 "&FIXED({tbv("122104")},2))',
         '账户之间转钱要转出、转入两笔都记（类别都选「内部转账」）；不是 0 说明有一笔只记了一边或金额不一样——到【明细账】选 122104 查'),
        ('外币兑换待对冲（122105）余额',
         f'=IF(ROUND({tbv("122105")},2)=0,"√ 0","⚠ 余额 "&FIXED({tbv("122105")},2))',
         '买卖外汇两边折苏姆的差额（＋一边没记的）。核对清楚后把【基础资料】「外币兑换差额月末转汇兑损益」改成「是」，每月自动转进财务费用-汇兑损益'),
        ('应付账款（220202）有借方余额',
         f'=IF({tbv("220202")}<=0,"√ 没有","⚠ 借方余额 "&FIXED({tbv("220202")},2))',
         '启用日以后「采购付款」冲的是应付账款：付了钱但【采购入库】没登记进货，就会变成借方（像预付）'),
        ('应付职工薪酬（2211 各明细）有借方余额',
         f'=IF(SUMIFS(余额期末直接,余额编码,"2211*",余额期末直接,">0")=0,"√ 没有",'
         f'"⚠ 借方余额 "&FIXED(SUMIFS(余额期末直接,余额编码,"2211*",余额期末直接,">0"),2))',
         '启用日以后发工资冲的是应付职工薪酬：先在【工资计提】把当月工资计提了；外籍员工的工资发放选「发外籍工资」（冲 221102）'),
        ('出矿的月份生产成本转完了',
         f'=IF(N(IFERROR(INDEX({rng(SH_ORE, "D", ORE_R0, ORE_R1)},{REP_N}),0))=0,"√ 本月没出矿（生产成本留在在产品）",'
         f'IF(ROUND({tbv("4001")}+{tbv("4101")},2)=0,"√ 转完了","✗ 还剩 "&FIXED({tbv("4001")}+{tbv("4101")},2)&" 没转进原矿"))',
         '生产成本有记在上级科目（4001、400101）的，或者 4001/4101 下的末级科目超过 20 个——改记到末级明细，或找我加'),
        ('生产成本末级科目个数', f'=IF(COUNT({rng(SH_ACC, "P", AC_R0, AC_R1)})>20,"✗ 超过 20 个，多出来的不会转原矿，找我加",'
                               f'"√ "&COUNT({rng(SH_ACC, "P", AC_R0, AC_R1)})&" 个")', '月末结转③最多转 20 个 4001/4101 下的末级科目'),
        ('上级科目直接记了账', _ok0(f'COUNTIFS(余额末级,"否",余额期末直接,"<>0")', '⚠ ', '√ 没有'),
         '有凭证直接记在上级科目（比如 2241、4001）而不是明细上：报表不会错，但明细账看不清。到【科目余额表】找「末级＝否」且有余额的科目，改记到明细'),
        ('外币应付账款的汇率差', f'=IF(ABS(N({ctx["arap_fx"]}))<1,"√ 没有","⚠ 合计 "&FIXED({ctx["arap_fx"]},2))',
         '人民币/美元进货按进货月汇率记应付，付款按付款时的汇率冲，差额留在应付账款里。到【往来余额表】P 列看是哪个供应商，做张手工凭证转汇兑损益'),
        ('原矿产销存＝库存商品-原矿',
         f'=IF(ROUND(N(IFERROR(INDEX({rng(SH_ORE, O_AE, ORE_R0, ORE_R1)},{REP_N}),0))-{tbv("140501")},2)=0,"√ 一致",'
         f'"✗ 差 "&FIXED(N(IFERROR(INDEX({rng(SH_ORE, O_AE, ORE_R0, ORE_R1)},{REP_N}),0))-{tbv("140501")},2))',
         '有手工凭证或流水直接记了 140501？原矿的进出只走月末结转'),
        ('原矿结存吨数为负', _ok0(f'COUNTIF({rng(SH_ORE, O_TE, ORE_R0, ORE_R1)},"<0")', '✗ ', '√ 没有'),
         '卖的比出的多：检查【产量登记】是不是漏登'),
        ('材料：存货收发存＝原材料＋周转材料',
         f'=IF(ABS(N({q(SH_INV)}!L3))<1,"√ 一致","⚠ 差 "&FIXED({q(SH_INV)}!L3,2))',
         '领用单价四舍五入会有几分钱的差；差得多是有凭证直接记了 1403/1411，或领用行有 ✗'),
        ('材料月末数量为负', _ok0(f'COUNTIF({q(SH_INV)}!$L${ir0}:$L${ir1},"<0")', '✗ ', '√ 没有'),
         '领的比进的多：先把采购入库补上'),
        ('固定资产：科目原值＝卡片原值',
         f'=IF(ROUND({tbv("1601")}-{fa_cost},2)=0,"√ 一致","⚠ 科目 "&FIXED({tbv("1601")},0)&"，卡片 "&FIXED({fa_cost},0))',
         '买设备的钱记进了 1601，但【固定资产】没建卡片——没卡片就不提折旧。把矿车、铲车、四不像…一台一张卡片补上'),
        ('基础资料：模块科目都在科目表里', _ok0(f'COUNTIF({q(SH_BASE)}!$C${P_ROW[ACC_PARAMS[0]]}:$C${P_ROW[ACC_PARAMS[-1]]},"✗*")', '✗ ', '√ 都在'),
         '【基础资料】参数里「…科目」右边显示 ✗ 的，改成科目表里有的末级科目'),
        ('资金账户：对应科目都在科目表里',
         _ok0(f'SUMPRODUCT((账户名<>"")*ISNA(MATCH(账户科目,{CODES},0)))', '✗ ', '√ 都在'), '【基础资料·资金账户】「对应科目」'),
        ('科目表：编码重复', _ok0(f'SUMPRODUCT(({CODES}<>"")*(COUNTIF({CODES},{CODES})>1))', '✗ ', '√ 没有'),
         '【科目表】编码列标红的是重复的'),
        ('科目表：编码格式不对', _ok0(f'SUMPRODUCT(({CODES}<>"")*((({acc(AC_KEY)}="")+ISNA(MATCH(LEN({CODES}),{{4,6,8,10}},0)))>0))',
                                 '✗ ', '√ 没有'),
         '编码只能是 4、6、8、10 位数字（前后不能有空格、字母）；格式不对的科目不进科目余额表'),
        ('录入表里有公式报错（＃VALUE！之类）的格子',
         _ok0('+'.join(f'SUMPRODUCT(--ISERROR({x}))' for x in (cash(K_CHK), mod(SH_BUY, B_CHK), mod(SH_ISS, I_CHK),
                                                                 mod(SH_SALE, S_CHK), mod(SH_PAY, W_CHK), mod(SH_FA, F_CHK),
                                                                 mod(SH_MAN, MN_CHK), mod(SH_PROD, PD_CHK))),
              '✗ ', '√ 没有'),
         '多半是数字格子里填了文字（比如「5年」「3000吨」）——到那张表的「核对/校验」列找报错的格子'),
        ('容量：本月凭证张数 / 打印页数',
         f'=IF(OR(COUNTIFS(凭证月份,报表年度*100+报表月份,凭证有,1)>{VL_N},本月总页数>{VP_PAGES}),"✗ 超出容量，找我加",'
         f'"√ "&COUNTIFS(凭证月份,报表年度*100+报表月份,凭证有,1)&" 张 / "&本月总页数&" 页")', f'每月最多 {VL_N} 张凭证、{VP_PAGES} 页'),
        ('容量：现金流水剩余空行',
         f'=IF({CASH_R1 - CASH_R0 + 1}-COUNTA({cash(K_DATE)})<200,"⚠ 只剩 "&({CASH_R1 - CASH_R0 + 1}-COUNTA({cash(K_DATE)}))&" 行",'
         f'"√ 还剩 "&({CASH_R1 - CASH_R0 + 1}-COUNTA({cash(K_DATE)}))&" 行")', f'共 {CASH_R1 - CASH_R0 + 1} 行，快满了找我加'),
    ]
    for i, (name, f, how) in enumerate(checks):
        r = R0 + i
        put(ws, f'A{r}', i + 1, F_TXT, align=AC)
        put(ws, f'B{r}', name, F_TXT, align=AL)
        put(ws, f'C{r}', f, F_TXTB, align=AL)
        put(ws, f'D{r}', how or None, F_NOTE, align=ALW)
        ws.row_dimensions[r].height = 30
    R1 = R0 + len(checks) - 1
    ws.conditional_formatting.add(f'C{R0}:C{R1}', FormulaRule(formula=[f'LEFT($C{R0},1)="✗"'], fill=FILL_WARN,
                                                               font=Font(name=YH, sz=10, bold=True, color='FFC00000')))
    ws.conditional_formatting.add(f'C{R0}:C{R1}', FormulaRule(formula=[f'LEFT($C{R0},1)="⚠"'], fill=fill('FFFFEB9C')))
    ws.conditional_formatting.add(f'C{R0}:C{R1}', FormulaRule(formula=[f'LEFT($C{R0},1)="√"'], fill=FILL_OK))
    put(ws, 'A3', '汇总：', F_KPI_L, FILL_KPI, align=AR_, border=False)
    ws.merge_cells('A3:B3')
    put(ws, 'C3', f'=IF(COUNTIF(C{R0}:C{R1},"✗*")=0,"√ 没有 ✗","✗ "&COUNTIF(C{R0}:C{R1},"✗*")&" 项要改")'
                  f'&"　　⚠ "&COUNTIF(C{R0}:C{R1},"⚠*")&" 项提醒"', Font(name=YH, sz=12, bold=True, color='FFC00000'),
        FILL_KPI, align=AL, border=False)
    ws.freeze_panes = 'A5'
    ctx['chk_range'] = (R0, R1)
    return ws


# ─────────────────────────── 首页 ───────────────────────────
NAV = [
    ('设置', [(SH_BASE, '参数（公司名、本位币、启用日、税率、凭证号打不打印、模块科目）、资金账户、项目、往来单位、物料档案'),
              (SH_RULE, '每类业务记到哪个科目：收支类别、物料类别、领用用途、部门、固定资产类别、现金流量项目'),
              (SH_RATE, '每月记账汇率（人民币、美元…折苏姆）'),
              (SH_ACC, '会计科目表（原来 287 个＋采矿补的明细）')]),
    ('日常录入', [(SH_CASH, '★ 每一笔收付款（五个账户合在一张），自动出凭证'),
                (SH_BUY, '进货：进口设备、炸药、油料、配件、五金、劳保、食材…'),
                (SH_ISS, '领料：从库里领出来用，自动按月加权平均算成本'),
                (SH_PROD, '出矿量、掘进量、品位（每天或每班）'),
                (SH_SALE, '卖原矿给甲方（吨 × 苏姆单价），自动算销项税、结转成本'),
                (SH_PAY, '每月计提工资（按部门）'),
                (SH_FA, '设备卡片，自动提折旧'),
                (SH_MAN, '调账、计提、分摊等多借多贷凭证')]),
    ('月末（全自动）', [(SH_MEND, '折旧、外币调汇、生产成本转原矿、结转销售成本'),
                     (SH_ORE, '原矿每月产、销、存和每吨成本')]),
    ('凭证', [(SH_JE, '所有凭证的分录总表'),
            (SH_VLIST, '报表月的凭证清单（凭证号、借贷合计、页数）'),
            (SH_VPRT, '打印记账凭证（A5 横向）')]),
    ('账簿', [(SH_TB, '科目余额表：月初、本月、本年累计、月末'),
            (SH_GL, '明细账：选科目看每一笔'),
            (SH_ACCT, '五个资金账户的原币、折苏姆余额，跟账核对'),
            (SH_ARAP, '按往来单位：应收、应付、其他应收应付、实收资本'),
            (SH_INV, '材料收发存')]),
    ('报表', [(SH_BAL, '资产负债表（会小企01）'),
            (SH_PL, '利润表（会小企02）'),
            (SH_CF, '现金流量表（会小企03，直接法）'),
            (SH_OPS, '经营报表：出矿、销售、吨矿成本、现金可维持几个月')]),
    ('检查', [(SH_CHK, '★ 打开先看：✗ 要改、⚠ 要看'),
            (SH_HELP, '怎么用（每天、每月、年底）')]),
]


def build_home(wb, ctx):
    ws = wb.create_sheet(SH_HOME)
    widths(ws, {'A': 2, 'B': 16, 'C': 18, 'D': 16, 'E': 18, 'F': 16, 'G': 18, 'H': 16, 'I': 18})
    title(ws, '=公司名称&" · 财务账套（本位币：苏姆 UZS）"', 'I', C_HOME,
          '💡 黄格子改报表年度、月份，所有账簿、报表、凭证打印都跟着变。平时只录【现金流水总表】和各业务表，凭证、账、报表全自动。'
          '点下面的表名直接跳过去。打开先看【数据校验】有没有 ✗。')
    ws.row_dimensions[3].height = 10
    ws.row_dimensions[4].height = 30
    put(ws, 'B4', '报表年度', F_KPI_L, FILL_KPI, align=AC)
    put(ws, 'C4', ctx['home_year'], F_SEL, FILL_SEL, '0', AC)
    put(ws, 'D4', '报表月份', F_KPI_L, FILL_KPI, align=AC)
    # 注意：HOME_MONTH 固定在 F4（common 里约定），E4 放说明
    put(ws, 'E4', None, fill_=FILL_KPI)
    put(ws, 'F4', ctx['home_month'], F_SEL, FILL_SEL, '0', AC)
    ws.merge_cells('D4:E4')
    put(ws, 'G4', '← 改这两格，全册跟着变', F_TIP, align=AL, border=False)
    ws.merge_cells('G4:I4')
    add_list_dv(ws, 'C4', '"2025,2026,2027"')
    add_list_dv(ws, 'F4', '"1,2,3,4,5,6,7,8,9,10,11,12"')
    ws.row_dimensions[5].height = 22
    put(ws, 'B5', '="期间："&报表年度&"年"&报表月份&"月　　建账起始："&YEAR(起始月份)&"年"&MONTH(起始月份)&"月　　业务模块启用日："&YEAR(启用日)&"年"&MONTH(启用日)&"月"&DAY(启用日)&"日"',
        F_NOTE, align=AL, border=False)
    ws.merge_cells('B5:I5')
    # KPI 卡片
    R = ctx['ops_rows']
    ops = lambda name, col='C': f"{q(SH_OPS)}!${col}${R[name]}"
    c0, c1 = ctx['chk_range']
    kpis = [('期末货币资金', f'={ctx["bs_cash"]}', MONEY), ('本月营业收入', f'={ctx["pl_rev"][1]}', MONEY),
            ('本月净利润', f'={ctx["pl_net"][1]}', MONEY), ('本年净利润', f'={ctx["pl_net"][0]}', MONEY),
            ('本月出矿（吨）', f'={ops("出矿量（原矿）")}', QTY2), ('本月销售（吨）', f'={ops("销售量（原矿）")}', QTY2),
            ('应收甲方', f'={ops("应收甲方（应收账款-原矿销售）")}', MONEY),
            ('数据校验', f'=IF(COUNTIF({q(SH_CHK)}!$C${c0}:$C${c1},"✗*")=0,"√ 没有 ✗",COUNTIF({q(SH_CHK)}!$C${c0}:$C${c1},"✗*")&" 项 ✗")'
                        f'&" / "&COUNTIF({q(SH_CHK)}!$C${c0}:$C${c1},"⚠*")&" 项 ⚠"', None)]
    for i, (lab, f, fmt) in enumerate(kpis):
        row = 7 + (i // 4) * 3
        col = CL(2 + (i % 4) * 2)
        col2 = CL(3 + (i % 4) * 2)
        ws.merge_cells(f'{col}{row}:{col2}{row}')
        ws.merge_cells(f'{col}{row + 1}:{col2}{row + 1}')
        put(ws, f'{col}{row}', lab, F_KPI_L, FILL_KPI, align=AC)
        put(ws, f'{col2}{row}', None, fill_=FILL_KPI)
        put(ws, f'{col}{row + 1}', f, F_KPI_V, FILL_KPI, fmt, AC)
        put(ws, f'{col2}{row + 1}', None, fill_=FILL_KPI)
        ws.row_dimensions[row + 1].height = 28
    link(ws['H11'], SH_CHK)
    # 导航
    r = 14
    put(ws, f'B{r}', '表 格 导 航（点表名跳过去）', F_SEC, border=False)
    r += 1
    header(ws, r, [('B', '分类'), ('C', '表名'), ('E', '做什么用')], C_HOME, height=24)
    ws.merge_cells(f'C{r}:D{r}')
    ws.merge_cells(f'E{r}:I{r}')
    put(ws, f'D{r}', None, fill_=fill(C_HOME))
    for c in 'FGHI':
        put(ws, f'{c}{r}', None, fill_=fill(C_HOME))
    r += 1
    for grp, items in NAV:
        r_start = r
        for sh, desc in items:
            ws.merge_cells(f'C{r}:D{r}')
            ws.merge_cells(f'E{r}:I{r}')
            put(ws, f'B{r}', grp if r == r_start else None, F_TXTB, FILL_SUBH, align=AC)
            c = put(ws, f'C{r}', sh, Font(name=YH, sz=11, bold=True, color='FF0563C1', underline='single'), align=AL)
            link(c, sh)
            put(ws, f'D{r}', None)
            put(ws, f'E{r}', desc, F_TXT, align=AL)
            for cc in 'FGHI':
                put(ws, f'{cc}{r}', None)
            ws.row_dimensions[r].height = 20
            r += 1
        if r - r_start > 1:
            ws.merge_cells(f'B{r_start}:B{r - 1}')
    # 每月流程
    r += 1
    put(ws, f'B{r}', '每 月 怎 么 做', F_SEC, border=False)
    steps = ['① 平时：每一笔收付款记【现金流水总表】（选账户、类别，填金额）；进货记【采购入库】，领料记【领用出库】，出矿记【产量登记】，卖矿记【销售结算】。',
             '② 月底：【月度汇率】填当月汇率；【工资计提】按部门计提当月工资；新买的设备在【固定资产】建卡片。',
             '③ 月末结转不用做：折旧、调汇、生产成本转原矿、结转销售成本全自动，凭证日期是月末最后一天。',
             '④ 【首页】把报表月份改成这个月 → 看【数据校验】有没有 ✗ → 看【凭证汇总】【科目余额表】→ 打印【记账凭证】和三张报表。']
    for i, t in enumerate(steps):
        rr = r + 1 + i
        ws.merge_cells(f'B{rr}:I{rr}')
        put(ws, f'B{rr}', t, F_TXT, align=ALW, border=False)
        ws.row_dimensions[rr].height = 30
    ws.freeze_panes = 'A6'
    ws.sheet_view.showGridLines = False
    return ws


HELP = [
    ('这本账怎么搭的', None),
    ('1', '本位币苏姆（UZS）。五个资金账户（人民币现金、美元现金、苏姆现金、NBU 银行-6688、公务卡-9200）的流水全部记在【现金流水总表】，外币按【月度汇率】折苏姆。'),
    ('2', '每一笔流水按「收/支类别」自动找对方科目（【记账规则】），一笔一张凭证；本账户那一边按【基础资料·资金账户】的「对应科目」。'),
    ('3', '业务表（采购入库、领用出库、销售结算、工资计提、手工凭证）每一行自动出凭证；月末结转（折旧、调汇、成本结转）全自动。'),
    ('4', '所有凭证的分录在【记账分录】，凭证号按月、按日期自动编；科目余额表、明细账、三张报表都从这里取，期间跟【首页】的报表年度/月份走。'),
    ('5', '损益类科目不做结转（表结法）：利润表直接取发生额，资产负债表的未分配利润自动算＝以前累计＋本年利润。年底要不要出结转损益凭证，找我。'),
    ('历史数据（2025-05 ～ 2026-08）', None),
    ('6', '原表【1现金流水总表】2,402 行原样搬过来（行号不变，第 6 行起），公式照原表，只是原来指向另一个文件的名字（账户名、类别表、记账汇率区…）现在指向本册。'),
    ('7', '业务模块启用日（【基础资料】参数，现在是 2026-09-01）以前：那时候没有采购入库、工资计提、销售结算，「采购付款」「生产成本」记营业成本-采矿服务成本（投产前）、'
          '「应付职工薪酬」记管理费用-基本工资、「销售回款」记主营业务收入-矿石开采服务——跟你 8 月报表一个口径。启用日以后按应付、应收、计提走。'),
    ('8', '有几笔按证据改了记账科目（原来的类别不动，写在「指定对方科目」列，格子上有批注说明为什么）：67 号矿井溢价款记采矿权；银行「1900 注册资本」是把现金存进公户，'
          '记资金划转（不再重复记实收资本）；记在「内部转账」里的银行手续费记财务费用。详见 说明.md。'),
    ('平时怎么录', None),
    ('9', '现金流水：日期、账户、类别、收入或支出（只填一个）必填；外币有实际成交汇率的填「实际成交汇率」。往来单位（客户、供应商、股东、借支的员工）尽量写，往来余额表靠它。'),
    ('10', '要改某一笔的科目：「指定对方科目」填编码或下拉选。几笔想并成一张凭证：「凭证合并号」写同一个号（同一个月）。'),
    ('11', '账户之间转钱、换汇：转出、转入两笔都要记，类别都选「内部转账」（换汇选「购汇/结汇」），中间科目余额回到 0 就说明对上了。'),
    ('12', '进货：【采购入库】一行一个物料，同一张单子「单据号」写一样的就并成一张凭证。付供应商在流水选「采购付款」（进口设备选「付设备款（进口）」），往来单位写同一个供应商；'
          '付关税、清关费、运费选「付进口税费及运杂」；海关交的进口增值税选「进口增值税（海关）」。人民币/美元进货的汇率差看【往来余额表】P 列。'),
    ('13', '领料：【领用出库】选物料、用途，单价自动（月加权平均）。用途决定记生产成本的哪个明细还是在建工程、管理费用。'),
    ('14', '出矿、卖矿：【产量登记】记出矿量；【销售结算】记吨数、单价，自动出「借 应收账款、贷 收入、贷 销项税」；月底按每吨成本自动结转销售成本。甲方付款在流水选「销售回款」。'),
    ('15', '工资：每月底【工资计提】按部门记应发工资；发工资在流水选「应付职工薪酬」（外籍员工选「发外籍工资」；9 月发的 8 月工资选「发放启用日前工资」）。'
          '个税选「代扣个人所得税」，单位社保选「社保（单位负担）」，增值税选「缴纳增值税」（别选「税费」）。'),
    ('16', '设备：【采购入库】（或流水「固定资产」）记进 1601，同时在【固定资产】建卡片（原值、入账日期、类别、部门），下个月起自动提折旧。'),
    ('17', '清单里没有的（账户、类别、物料、往来单位、科目）：先去【基础资料】【记账规则】【科目表】往下面空行加一行，再录。不要插行、删行。'),
    ('18', '手工凭证：同一张凭证几行挨着写、「凭证组号」一样，日期写在第一行；组号每个月可以从 1 重新编。一张凭证有一行 ✗，整张先不记账。'),
    ('打印凭证', None),
    ('19', '【首页】选报表月 → 【凭证汇总】看总页数 → 【记账凭证】打印，页码范围 1 到总页数（A5 横向，一页 8 条分录，超过的接下一页）。'),
    ('20', '凭证号默认不印（留空手写，照右边红字写）；想印上，把【基础资料】「凭证号是否打印」改「是」。制单、出纳、复核、会计主管的名字在【基础资料】参数里填。'),
    ('其他', None),
    ('21', '打开时 Excel/WPS 会自己全部重算，第一次打开慢一点正常（公式多）。改完数等右下角算完再看。'),
    ('22', '数字不对先看【数据校验】；科目余额表不平、报表不平都会在那里亮 ✗。'),
]


def build_help(wb, ctx):
    ws = wb.create_sheet(SH_HELP)
    widths(ws, {'A': 6, 'B': 120})
    title(ws, '使用说明', 'B', C_HOME, '💡 送人或以后自己忘了，看这张就行；更细的改动记录、需要你确认的事项在 说明.md。')
    r = 4
    for no, text in HELP:
        if text is None:
            ws.merge_cells(f'A{r}:B{r}')
            put(ws, f'A{r}', no, F_SEC, FILL_SUBH, align=AL)
            ws.row_dimensions[r].height = 24
        else:
            put(ws, f'A{r}', int(no), F_TXT, align=AC)
            put(ws, f'B{r}', text, F_TXT, align=ALW)
            ws.row_dimensions[r].height = 18 if len(text) < 70 else 34
        r += 1
    ws.sheet_view.showGridLines = False
    return ws
