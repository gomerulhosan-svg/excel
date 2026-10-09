# -*- coding: utf-8 -*-
"""【首页】：报表年度、截止日期（全书默认期间）、关键数、提醒、各表入口、常见情况怎么记。"""
import datetime as dt
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from layout import *
from common import *
import s_check

NAV = [
    ('平时录（蓝）', C_IN, [(SH_CASH, '所有账户每一笔钱（就是你的收支登记表）'), (SH_AR_IN, '补充协议、签证、扣款、开票'),
                         (SH_AP_IN, '材料送货、分包结算、机械台班'), (SH_ATT, '一人一月一行，8 个工地'),
                         (SH_INV, '电子税务局导出的发票（核对用）')]),
    ('基本信息（灰）', C_MASTER, [(SH_PJ, '项目、合同、保证金、质保金、开工'), (SH_CUS, '甲方/总包'), (SH_SUP, '材料商、分包、机械'),
                             (SH_PER, '工人、管理人员、老板、过账人员'), (SH_RATE, '日工资/月薪，涨工资加一行'),
                             (SH_BASE, '参数、账户、收支项目'), (SH_OPEN, '建账日以前的余额')]),
    ('查看（绿）', C_VIEW, [(SH_FUND, '周报/月报/年报：各账户收支余额'), (SH_ACCT, '选账户看每一笔和余额'),
                         (SH_AR, '按项目：应收、已收、未收、开票'), (SH_ARD, '选项目或甲方对账'),
                         (SH_AP, '按供应商：应付、已付、未付、开票'), (SH_APD, '选供应商对账'),
                         (SH_PERSON, '借款、垫付报销、备用金、过账'), (SH_PAY, '一个月的工资表'),
                         (SH_PAYSUM, '选时间段看工资和欠薪'), (SH_LABOR, '选时间段、选工地看人工'),
                         (SH_ARREAR, '未发清工资、每月补发'), (SH_INVSUM, '销项、进项、按月、对税局')]),
    ('项目（橙）', C_PROJ, [(SH_PJBOOK, '选项目：收入、成本、利润和每一笔'), (SH_PJPL, '所有项目赚亏'),
                         (SH_ALLOC, '管理费怎么摊、摊了多少')]),
    ('报表（深红）', C_RPT, [(SH_PL, '利润表（按月）'), (SH_BS, '资产负债表'), (SH_BE, '盈亏平衡'), (SH_CHK, '哪里录错了')]),
]

HOWTO = [
    ('老板/负责人自己的钱替公司付了', '账户选他的个人账户（聂辉微信、王伟微信……），收支项目按实际用途；有票填「已开票金额」；要摊的「费用归属」选待摊费用'),
    ('公司把垫付的钱还给他', '公司账户付，收支项目「报销还款」，人员＝他'),
    ('账户之间倒钱', '转出的账户记一行，收支项目「内部转账」，对方账户选转入的账户（两边自动都记）'),
    ('甲方从农民工专户代发工资', '账户「农民工专户」：收入记工程款（项目、客户）；支出记劳务费（人员＝工人）或分包（供应商＝班组长）；余额应为 0'),
    ('工人把多拿的工资退回', '收到钱的账户记收入，收支项目「工资退回」，人员＝那个工人；张总再转给聂总记内部转账'),
    ('过账的人（姚俊强）', '人员信息「过账＝是」；公司替他交社保/发工资、他退回，都记「其他应收款-代扣社保」，人员＝他；不算成本'),
    ('借老板的钱 / 还老板', '收支项目「个人借款-姓名」：收入＝借入，支出＝归还'),
    ('补发以前欠的工资', '照常记劳务费，人员＝他；【欠薪补发】自动按月显示'),
    ('红包、公关', '收支项目记业务招待费，费用归属选待摊费用（为某个项目花的就填项目、选项目成本）；不要记成材料费'),
    ('跟材料商/分包对了账、收到送货单或结算单', '【应付登记】记一行（可同时填已开票金额、开票日期、税率）；付款在收支登记，供应商名称填这家，自动冲应付'),
    ('甲方签了补充协议、签证、扣款，或者开了发票', '【应收登记】记一行；收款在收支登记记工程款，项目名称选对'),
]

# 关键数：(标签, 公式, 跳到哪张表)。各报表模块写好后可以在 KPI_EXTRA 里加
KPI_EXTRA = []


def build(wb, ctx=None):
    ws = wb[SH_HOME]
    widths(ws, {'A': 2, 'B': 17, 'C': 15, 'D': 3, 'E': 17, 'F': 15, 'G': 3, 'H': 17, 'I': 15, 'J': 3, 'K': 17, 'L': 15})
    title(ws, '建 筑 安 装 财 务 模 板（消防改造 · 水电安装 · 强弱电）', 'L', C_HOME,
          '💡 平时只录蓝色那几张（收支登记、应收登记、应付登记、考勤工资，发票登记核对用），其余全是自动的；所有表都可以插行、删行。'
          '下面黄格子选报表年度和截止日期（空着＝最后一笔流水那天），各报表没单独选日期时就按这个算。')
    ws['A1'].value = '=IF(P_公司="","",P_公司&" · ")&"建筑安装财务模板"'
    put(ws, 'B4', '报表年度', F_KPI_L, fill('FFD9E1F2'), align=AC)
    put(ws, HOME_YEAR, None, F_SEL, FILL_SEL, '0', AC)
    dv = DataValidation(type='whole', operator='between', formula1='2000', formula2='2100', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(HOME_YEAR)
    put(ws, 'E4', '截止日期', F_KPI_L, fill('FFD9E1F2'), align=AC)
    put(ws, HOME_END, None, F_SEL, FILL_SEL, DATE, AC)
    dv_date(ws, HOME_END)
    put(ws, 'H4', '="现在按："&P_年度&" 年，"&TEXT(P_年初,"yyyy-mm-dd")&" ～ "&TEXT(P_截止,"yyyy-mm-dd")'
                  '&"（最后一笔流水 "&TEXT(P_最后流水日,"yyyy-mm-dd")&"）"', F_NOTE, align=AL, border=False)
    ws.merge_cells('H4:L4')
    put(ws, 'B5', '空着＝最后一笔流水那年 / 那天。改了年度、截止日期，所有报表跟着变。', F_NOTE, align=AL, border=False)
    ws.merge_cells('B5:L5')

    # 关键数
    section(ws, 6, 'B', 'L', '关键数（截至截止日期；点名字跳到那张表）', C_HOME)
    kpis = [('本年确认收入', '=汇_收入12月-汇_收入年初前', SH_PL),
            ('本年净利润', f'={SH_PL}!$C$23', SH_PL),
            ('欠工人工资（应付工资）', f'={SH_BS}!$F$8', SH_PAYSUM),
            ('本年保本收入', f'=IF(ISNUMBER({SH_BE}!$B$14),{SH_BE}!$B$14,0)', SH_BE),
            ('银行＋现金＋专户', '=汇_货币资金负债日', SH_FUND),
            ('甲方还欠（应收）', '=汇_应收负债日-汇_预收负债日', SH_AR),
            ('欠供应商（应付）', '=汇_应付负债日-汇_预付负债日', SH_AP),
            ('公司欠老板借款', '=SUMIFS(期初_金额,期初_类型,"个人借款")+SUMIFS(收_净额,收_归类,"个人借款",收_日期,"<="&P_截止)', SH_PERSON),
            ('公司欠老板垫付（未报销）', '=汇_个人垫付负债日-汇_个人持有负债日', SH_PERSON),
            ('本年待摊管理费', f'=IFERROR(INDEX(汇_待摊池,P_年度-{YEAR0}+1),0)', SH_ALLOC),
            ('本年管理费率', f'=IFERROR(INDEX(汇_费率,P_年度-{YEAR0}+1),0)', SH_ALLOC)] + KPI_EXTRA
    for i, (lab, f, sh) in enumerate(kpis[:12]):
        r = 7 + (i // 4) * 2
        c1 = ['B', 'E', 'H', 'K'][i % 4]
        c2 = CL(CI(c1) + 1)
        c = put(ws, f'{c1}{r}', lab, Font(name=YH, sz=10, bold=True, color='FF0563C1', underline='single'), fill('FFD9E1F2'), align=AC)
        link(c, sh)
        ws.merge_cells(f'{c1}{r}:{c2}{r}')
        put(ws, f'{c1}{r + 1}', f, F_KPI_V, fill('FFFFFFFF'), PCT if '率' in lab else MONEY, AC)
        ws.merge_cells(f'{c1}{r + 1}:{c2}{r + 1}')
        ws.row_dimensions[r + 1].height = 26
    nr = 7 + 2 * ((min(len(kpis), 12) + 3) // 4)
    put(ws, f'B{nr}', '资产负债表日期没单独选时＝截止日期。「公司欠老板垫付」＝老板个人账户替公司付了、公司还没还的（负数＝他手上有公司的钱）。',
        F_NOTE, align=AL, border=False)
    ws.merge_cells(f'B{nr}:L{nr}')

    # 提醒
    r = nr + 2
    section(ws, r, 'B', 'L', '提醒', C_RPT)
    CR = s_check.CHECK_ROWS
    chk = lambda lab: f'{SH_CHK}!$C${CR[lab]}'
    rem = [f'="数据校验：要改 "&{SH_CHK}!{s_check.CHK_X}&" 项、要看 "&{SH_CHK}!{s_check.CHK_W}&" 项"&'
           f'IF({SH_CHK}!{s_check.CHK_X}+{SH_CHK}!{s_check.CHK_W}=0,"，都没问题 √","  → 点这里去看")',
           f'=IF(N({chk(f"{SH_CASH}逐行")})>0,"收支登记有 "&{chk(f"{SH_CASH}逐行")}&" 行要改（账户、收支项目、项目名不在基本信息里，或者没填日期）","收支登记都认出来了 √")',
           f'=IF(ABS(N({chk("专户（农民工专户、九安账户这类）期末余额合计")}))>=0.01,"专户还有余额 "&TEXT({chk("专户（农民工专户、九安账户这类）期末余额合计")},"#,##0.00")&"：收了没付出去，或者漏记了一边","专户余额都是 0 √")',
           f'=IF(N({chk("供应商多付了（应付余额为负）的家数")})>0,{chk("供应商多付了（应付余额为负）的家数")}&" 家供应商付款比应付登记多：送货单/结算没录全，或者付款记错了家 → 看【应付账款】","")',
           f'=IF(ABS(N({chk("待摊费用没摊出去的（各年合计）")}))>=0.01,"有 "&TEXT({chk("待摊费用没摊出去的（各年合计）")},"#,##0.00")&" 待摊费用那一年没有项目可摊，留在公司 → 看【费用分摊】","")']
    for i, f in enumerate(rem):
        rr = r + 1 + i
        put(ws, f'B{rr}', f, F_TXTB if i == 0 else F_TXT, align=AL, border=False)
        ws.merge_cells(f'B{rr}:L{rr}')
    link(ws[f'B{r + 1}'], SH_CHK)
    ws.conditional_formatting.add(f'B{r + 1}:B{r + len(rem)}',
                                  FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("要改",B{r + 1})),ISNUMBER(SEARCH("余额 ",B{r + 1})),ISNUMBER(SEARCH("多",B{r + 1})))'],
                                              font=Font(name=YH, sz=10, bold=True, color='FFC00000')))

    # 导航
    r = r + len(rem) + 2
    section(ws, r, 'B', 'L', '各张表（点名字跳过去；每张表右上角有「← 回首页」）', C_HOME)
    r += 1
    col_sets = [('B', 'C', 'F'), ('H', 'I', 'L')]
    for grp, color, sheets in NAV:
        put(ws, f'B{r}', grp, F_SEC, fill(color), align=AL)
        ws.merge_cells(f'B{r}:L{r}')
        r += 1
        for j, (sh, desc) in enumerate(sheets):
            c1, c2, c3 = col_sets[j % 2]
            rr = r + j // 2
            cell = put(ws, f'{c1}{rr}', sh, Font(name=YH, sz=10, bold=True, color='FF0563C1', underline='single'), align=AL)
            link(cell, sh)
            put(ws, f'{c2}{rr}', desc, F_NOTE, align=AL)
            ws.merge_cells(f'{c2}{rr}:{c3}{rr}')
        r += (len(sheets) + 1) // 2

    # 常见情况怎么记
    r += 1
    section(ws, r, 'B', 'L', '常见情况怎么记（不知道记哪就看这里）', C_HOME)
    for i, (k, v) in enumerate(HOWTO):
        rr = r + 1 + i
        put(ws, f'B{rr}', k, F_TXTB, fill('FFF2F2F2'), align=ALW)
        ws.merge_cells(f'B{rr}:C{rr}')
        put(ws, f'E{rr}', v, F_TXT, align=ALW)
        ws.merge_cells(f'E{rr}:L{rr}')
        ws.row_dimensions[rr].height = 32
    r = r + len(HOWTO) + 2
    section(ws, r, 'B', 'L', '注意', C_HOME)
    notes = ['① 名称（项目、甲方、供应商、人员、账户、收支项目）都从下拉选；新的先到基本信息那几张表加一行。',
             '② 录入表哪里都可以插行、删行；灰色、绿色、橙色、深红的表是自动的，有保护（没有密码），只有黄格子能选。',
             '③ 记错了直接改或删整行都行；改完看一眼【数据校验】。',
             '④ 金额都含税；利润表、资产负债表是管理用的（看赚没赚钱、钱在哪、欠谁的），报税以代账会计为准。']
    for i, t in enumerate(notes):
        rr = r + 1 + i
        put(ws, f'B{rr}', t, F_TXT, align=ALW, border=False)
        ws.merge_cells(f'B{rr}:L{rr}')
        ws.row_dimensions[rr].height = 20
    ws.sheet_view.showGridLines = False
    print_setup(ws, None, landscape=False)
