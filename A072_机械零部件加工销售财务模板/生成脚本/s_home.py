# -*- coding: utf-8 -*-
"""【首页】：截止日期（全书默认）、钱（可用资金、各账户）、往来、付款和计划、本月/本年经营（利润表口径）、提醒、各表入口、常见情况怎么记。
   首页的数都直接按定义名称算（不引用别的查看表的格子，查看表改了选择格也不影响首页）。"""
from openpyxl.formatting.rule import FormulaRule
from layout import *
from common import *
import s_check

LINKF = Font(name=YH, sz=10, bold=True, color='FF0563C1', underline='single')
LBL = fill('FFD9E1F2')
WHITE = fill('FFFFFFFF')

NAV = [
    ('平时录（蓝）', C_IN, [(SH_CASH, '7 个账户的每一笔钱（混在一张表，有即时余额）'), (SH_SALE, '发货一行（还没定价的先空着单价）'),
                       (SH_PUR, '到货一行（钢材、外协、外购成品、辅料……）'), (SH_INV, '开出、收到的发票（手工登记）'),
                       (SH_APV, '要付的款申请一行，老板审批'), (SH_PP, '每天要买的东西提交一行')]),
    ('查看（绿，自动）', C_VIEW, [(SH_DAY, '选一天：各账户收支、余额，当天每一笔'), (SH_MON, '选一个月：各账户、各类别、全年各月'),
                           (SH_FEE, '费用按月统计，选一个类别看明细'), (SH_PL, '1～12 月利润表')]),
    ('计划（紫，自动）', C_PLAN, [(SH_FPLAN, '这个月的钱够不够：固定支出、批了没付、等审批的'), (SH_APRINT, '打印给老板签字'),
                           (SH_PPSUM, '每天的采购计划单、按月汇总')]),
    ('往来（橙，自动）', C_WL, [(SH_AR, '每个客户欠多少、多久没回款、开票'), (SH_AP, '欠每个供应商多少、欠票、待审批'),
                          (SH_STMT, '选一家出对账单（A4 打印）'), (SH_INVS, '每月开票收票、没开票的客户、欠票的供应商')]),
    ('基础（灰）', C_MASTER, [(SH_BASE, '公司名称、建账日、账户、收支类别、人员、库存估值'), (SH_UNIT, '客户、供应商（期初应收应付）'),
                          (SH_FIX, '每月固定要付的')]),
    ('校验', C_CHK, [(SH_CHK, '哪里录错了、该跟进什么')]),
]

HOWTO = [
    ('客户打来货款', '【资金台帐】一行：账户选收到钱的账户，类别「收客户货款」，往来单位选客户，填收入。退给客户的钱同样类别，记在支出栏'),
    ('收到承兑汇票', '【资金台帐】账户选「承兑汇票」，类别「收客户货款」，往来单位选客户，填收入（票号、到期日写在摘要）。承兑不算可用资金'),
    ('承兑背书付货款', '【资金台帐】账户选「承兑汇票」，类别「付供应商货款」，往来单位选供应商，填支出'),
    ('承兑贴现、到期托收', '记两行「内部转账」：承兑汇票账户支出票面金额、银行账户收入票面金额；贴息另记一行（银行账户支出，类别「承兑贴息」）'),
    ('发货（接单加工、外购成品转卖）', '【销售登记】一个产品一行：客户、产品、数量、单价；还没定价的单价金额先空着，定了价回到原行补（日期不改）。买成品回来卖的业务类型选「外购成品」'),
    ('钢材、外协、外购成品、辅料到货', '【采购登记】一行：供应商、物料、数量、单价或金额（按送货单、磅单、结算单）；月底才定价的先空着'),
    ('付供应商货款', '先在【付款审批】申请一行 → 老板批（【待审批付款单】打印签字）→ 付款时【资金台帐】类别「付供应商货款」、往来单位选供应商、关联单号选那张审批单'),
    ('零星现买（当场付清、不用对账）', '【资金台帐】类别「零星料钱（现买）」，往来单位不用填；登过采购登记的货付款一定选「付供应商货款」，不然成本算两遍'),
    ('工资、电费、房租、社保', '【资金台帐】选对应类别；工资、电费是上个月的，「所属月份」填几月（利润算到那个月）。每月固定的在【固定支出】登记一次，【资金计划】自动看哪些还没付'),
    ('计划要付的（电费、税、利息……）', '在【付款审批】申请一行，计划付款日期填哪天；老板批了，【资金计划】那个月就算进要付的'),
    ('退货、扣款、折让、质量罚款', '【销售登记】或【采购登记】记一行负数金额，产品名称写原因'),
    ('开发票、收发票', '【发票登记】一张一行：开出（我们开给客户）或收到（供应商开来的）、往来单位、价税合计；税额空着自动算'),
    ('账户之间倒钱（转账、取现、存现）', '【资金台帐】记两行「内部转账」：转出账户记支出、转进账户记收入，金额一样；手续费另记一行'),
    ('老板投钱拿钱、借款还款、保证金', '【资金台帐】类别选「老板存取」「借款」「还借款」「保证金」：进出资金，不算收入支出'),
    ('买机床等设备', '【资金台帐】类别「买设备」：不算当月费用；每月折旧在【基础资料】估一个数'),
    ('每天采购计划', '【采购计划】提交一行 → 老板在「审批」列选 → 买了选「已下单」「已到货」；【采购计划汇总】打印当天的单'),
    ('月底', '在【基础资料】⑤ 估一下仓库里料、在制品、成品值多少（可不填）；看【利润表】、【资金月报】、【客户往来】'),
]

PRODUCE = ('生产这块怎么处理：财务不跟生产过程，也不用盘进销存。料、外协、外购成品买进来（采购登记）就算成本，工人工资、电费、房租、刀具辅料也算生产成本；'
           '月底大概估一下仓库里还有多少料和货（填在基础资料⑤，不填也行），利润表用「上次估的－这次估的」把没用掉的扣回来。'
           '以后能盘点了，把估值填准，利润就准了。销售收入按发货登记（定了价才有金额），回款按资金台帐。')


def _pl_month(j):
    """第 j 月（P_年度）的利润表各项（跟【利润表】同一口径）；返回 dict 公式片段"""
    m = f'(P_年度*100+{j})'
    upto = '资_日期,"<="&P_截止'
    cash = lambda kind: f'SUMIFS(资_支出额,资_归类,"{kind}",资_所属年月,{m},{upto})'
    has = f'COUNTIFS(库存_年月,{m},库存_有值,1)>0'
    val = f'SUMIFS(库存_估值,库存_年月,{m},库存_有值,1)'
    prevm = f'IFERROR(LARGE(库存_键,COUNTIF(库存_键,">="&{m})+1),0)'
    inv = (f'IF({has},IF({prevm}>0,SUMIFS(库存_估值,库存_年月,{prevm},库存_有值,1)-{val},'
           f'IF(ISNUMBER(P_期初库存),P_期初库存-{val},0)),0)')
    return dict(
        收入=f'SUMIFS(销_金额,销_计入往来,1,销_年月,{m},销_日期,"<="&P_截止)',
        采购=f'SUMIFS(采_金额,采_计入往来,1,采_年月,{m},采_日期,"<="&P_截止)',
        生产=cash('生产成本'),
        折旧=f'IF(AND({m}>={ym("P_建账日")},{m}<=P_截止年月),P_月折旧,0)',
        库存=f'IF({m}>P_截止年月,0,{inv})',
        费用=f'{cash("销售费用")}+{cash("管理费用")}+{cash("财务费用")}',
        税金=cash('税金'),
        其他=f'-SUMIFS(资_支出额,资_归类,"其他收入",资_所属年月,{m},{upto})-{cash("其他支出")}',
    )


def build(wb, ctx=None):
    ws = wb[SH_HOME]
    widths(ws, {'A': 2, 'B': 24, 'C': 15, 'D': 15, 'E': 3, 'F': 24, 'G': 15, 'H': 15, 'I': 3, 'J': 22, 'K': 15, 'L': 15})
    title(ws, '机 械 零 部 件 加 工 销 售 财 务', 'L', C_HOME,
          '💡 平时只录蓝色的几张：【资金台帐】（每一笔钱）、【销售登记】（发货）、【采购登记】（到货）、【发票登记】、【付款审批】、【采购计划】；'
          '日报、月报、费用、利润表、往来、对账单、资金计划都是自动的。登记表都可以插行、删行、排序。灰底行是示例，正式用之前整行删掉。')
    ws['A1'].value = '=IF(P_公司名称="","",P_公司名称&" · ")&"机械零部件加工销售财务"'
    put(ws, 'B4', '截止日期', F_KPI_L, LBL, align=AC)
    put(ws, HOME_END, None, F_SEL, FILL_SEL, DATE, AC)
    dv_date(ws, HOME_END)
    put(ws, 'D4', '="现在按 "&TEXT(P_截止,"yyyy-mm-dd")&" 算（空着＝最后一笔 "&TEXT(P_最后日期,"yyyy-mm-dd")&"）；建账日 "&TEXT(P_建账日,"yyyy-mm-dd")',
        Font(name=YH, sz=11, bold=True, color='FF1F3864'), align=AL, border=False)
    ws.merge_cells('D4:L4')
    put(ws, 'B5', (f'=IF(AND(ISNUMBER({HOME_END}),P_截止<P_最后日期),"⚠ 截止日期比最新录入（"&TEXT(P_最后日期,"yyyy-mm-dd")&"）早，'
                   f'后面录的都没算进来——是不是忘了清空？",IF(AND(NOT(ISNUMBER({HOME_END})),P_截止被挡=1),"⚠ 有 "&P_近月笔数&" 笔日期比别的晚一个多月，'
                   f'多半年份填错了，截止日期没跟过去（看【数据校验】）","截止日期决定：余额、欠款算到哪天，「今天」「本月」是哪天哪个月。平时空着就行（＝最后一笔的日期）；'
                   f'想按今天算，在黄格里输入 =TODAY()。"))'), F_NOTE, align=AL, border=False)
    ws.merge_cells('B5:L5')
    ws.conditional_formatting.add('B5', FormulaRule(formula=['LEFT(B5,1)="⚠"'], font=F_RED))

    # ① 钱
    r = 7
    section(ws, r, 'B', 'L', '① 钱（截至截止日期；按日期算。点名字看资金日报）', C_HOME)
    link(ws[f'B{r}'], SH_DAY)
    bal = lambda a: f'INDEX(账户_期初余额,{a})+SUMIFS(资_净额,资_账户号,{a},资_日期,"<="&P_截止)'
    avail = acc_expr(1)
    bills = acc_expr(0)
    nt = '资_归类,"<>内部转账",资_资金有效,1'
    kpis = [('可用资金（不含承兑）', f'=ROUND({avail},2)'), ('手上承兑汇票', f'=ROUND({bills},2)'),
            ('今天收入（不含内部转账）', f'=SUMIFS(资_收入,{nt},资_日期,P_截止)'), ('今天支出（不含内部转账）', f'=SUMIFS(资_支出,{nt},资_日期,P_截止)'),
            ('本月收入（不含内部转账）', f'=SUMIFS(资_收入,{nt},资_年月,P_截止年月,资_日期,"<="&P_截止)'),
            ('本月支出（不含内部转账）', f'=SUMIFS(资_支出,{nt},资_年月,P_截止年月,资_日期,"<="&P_截止)')]
    for i, (lab, f) in enumerate(kpis):
        c0, cv = (('B', 'C'), ('F', 'G'), ('J', 'K'))[i % 3]
        rr = r + 1 + i // 3
        put(ws, f'{c0}{rr}', lab, F_KPI_L, LBL, align=AL)
        put(ws, f'{cv}{rr}', f, F_KPI_V if i == 0 else F_AUTOB, WHITE, MONEY, AR)
        ws.merge_cells(f'{cv}{rr}:{CL(CI(cv) + 1)}{rr}')
        ws.row_dimensions[rr].height = 22
    r += 3
    NA = N_ACC
    for i in range(NA):
        c0, cv = (('B', 'C'), ('F', 'G'), ('J', 'K'))[i % 3]
        rr = r + 1 + i // 3
        a = f'INDEX(账户_名称,{i + 1})'
        put(ws, f'{c0}{rr}', f'=IF({a}="","",{a}&IF(INDEX(账户_可用,{i + 1})=0,"（不算可用）",""))', F_TXTB, align=AL, border=False)
        put(ws, f'{cv}{rr}', f'=IF({a}="","",ROUND({bal(i + 1)},2))', F_AUTOB, None, MONEY, AR, border=False)
        ws.merge_cells(f'{cv}{rr}:{CL(CI(cv) + 1)}{rr}')
    last_acc = r + (NA + 2) // 3
    ws.conditional_formatting.add(f'B{r + 1}:L{last_acc}', FormulaRule(formula=[f'AND(ISNUMBER(B{r + 1}),B{r + 1}<-0.005)'], font=F_RED))
    for c0, cv in (('B', 'D'), ('F', 'H'), ('J', 'L')):        # 有账户的格子才填色、画框
        ws.conditional_formatting.add(f'{c0}{r + 1}:{cv}{last_acc}', FormulaRule(formula=[f'${c0}{r + 1}<>""'], border=BD))
        ws.conditional_formatting.add(f'{c0}{r + 1}:{c0}{last_acc}', FormulaRule(formula=[f'${c0}{r + 1}<>""'], fill=LBL))
    r = last_acc + 2

    # ② 往来
    section(ws, r, 'B', 'L', '② 往来（截至截止日期；点名字看明细）', C_WL)
    od = '位_是客户,1,位_没回款天数,">"&P_回款天数'
    wl = [('客户欠我们（合计）', '=SUMIFS(位_应收,位_是客户,1,位_应收,">0")', MONEY, SH_AR),
          ('客户预收（先收了钱）', '=-SUMIFS(位_应收,位_是客户,1,位_应收,"<0")', MONEY, SH_AR),
          ('="超 "&P_回款天数&" 天没回款"', f'=COUNTIFS({od})&" 家 / "&TEXT(SUMIFS(位_应收,{od}),"#,##0")', '@', SH_AR),
          ('我们欠供应商（合计）', '=SUMIFS(位_应付,位_是供应商,1,位_应付,">0")', MONEY, SH_AP),
          ('预付供应商', '=-SUMIFS(位_应付,位_是供应商,1,位_应付,"<0")', MONEY, SH_AP),
          ('还没定价（发货 / 到货）', '=COUNTIFS(销_待定价,1,销_计入往来,1)&" 笔 / "&COUNTIFS(采_待定价,1,采_计入往来,1)&" 笔"', '@', SH_SALE),
          ('客户还没开票', '=SUMIFS(位_未开票,位_是客户,1,位_未开票,">0")', MONEY, SH_INVS),
          ('供应商欠我们发票', '=SUMIFS(位_欠票,位_是供应商,1,位_欠票,">0")', MONEY, SH_INVS),
          ('对账单', '选一家、选起止日期，A4 打印', '@', SH_STMT)]
    for i, (lab, f, fmt, sh) in enumerate(wl):
        c0, cv = (('B', 'C'), ('F', 'G'), ('J', 'K'))[i % 3]
        rr = r + 1 + i // 3
        c = put(ws, f'{c0}{rr}', lab, LINKF, LBL, align=AL)
        if not str(lab).startswith('='):
            link(c, sh)
        put(ws, f'{cv}{rr}', f, F_AUTOB if lab != '对账单' else F_NOTE, WHITE, fmt, AR if fmt == MONEY else AC)
        ws.merge_cells(f'{cv}{rr}:{CL(CI(cv) + 1)}{rr}')
        ws.row_dimensions[rr].height = 22
    r += 5

    # ③ 付款和计划
    section(ws, r, 'B', 'L', '③ 付款和计划（本月＝截止日期那个月；点名字看资金计划）', C_PLAN)
    link(ws[f'B{r}'], SH_FPLAN)
    fx_due = 'SUMIFS(固_每月金额,固_本月生效,1)'
    fx_paid = 'SUMIFS(固_本月已付,固_本月生效,1)'
    fx_left = f'SUMPRODUCT((固_本月生效=1)*(固_每月金额>固_本月已付)*(固_每月金额-固_本月已付))'
    ap_left = 'SUMIFS(批_未付,批_计划年月,"<="&P_截止年月)'
    pay = [('等老板审批', '=COUNTIFS(批_状态码,0)&" 笔 / "&TEXT(SUMIFS(批_申请金额,批_状态码,0),"#,##0")', '@', SH_APRINT),
           ('批了还没付（本月及以前计划的）', f'=COUNTIFS(批_未付,">0",批_计划年月,"<="&P_截止年月)&" 笔 / "&TEXT({ap_left},"#,##0")', '@', SH_FPLAN),
           ('本月固定支出 应付 / 已付', f'=TEXT({fx_due},"#,##0")&" / "&TEXT({fx_paid},"#,##0")', '@', SH_FIX),
           ('本月固定支出还没付', f'=ROUND({fx_left},2)', MONEY, SH_FPLAN),
           ('预计还剩（可用资金－固定没付－批了没付）', f'=ROUND({avail}-{fx_left}-{ap_left},2)', MONEY, SH_FPLAN),
           ('采购计划 待审批 / 同意没买', '=COUNTIFS(购_状态码,0)&" 条 / "&COUNTIFS(购_状态码,1)&" 条"', '@', SH_PPSUM)]
    for i, (lab, f, fmt, sh) in enumerate(pay):
        c0, cv = (('B', 'C'), ('F', 'G'), ('J', 'K'))[i % 3]
        rr = r + 1 + i // 3
        link(put(ws, f'{c0}{rr}', lab, LINKF, LBL, align=ALW), sh)
        put(ws, f'{cv}{rr}', f, F_AUTOB, WHITE, fmt, AR if fmt == MONEY else AC)
        ws.merge_cells(f'{cv}{rr}:{CL(CI(cv) + 1)}{rr}')
        ws.row_dimensions[rr].height = 30
    ws.conditional_formatting.add(f'C{r + 2}:L{r + 2}', FormulaRule(formula=[f'AND(ISNUMBER(C{r + 2}),C{r + 2}<-0.005)'], font=F_RED))
    r += 4

    # ④ 本月、本年经营（利润表口径）
    section(ws, r, 'B', 'L', '④ 本月、本年经营（利润表口径；点名字看利润表）', C_VIEW)
    link(ws[f'B{r}'], SH_PL)
    hc = {k: CL(CI('AB') + j) for j, k in enumerate(['收入', '采购', '生产', '折旧', '库存', '费用', '税金', '其他'])}
    ws['AA1'] = '月'
    for k, c in hc.items():
        ws[f'{c}1'] = k
    for j in range(1, 13):
        ws[f'AA{j + 1}'] = j
        for k, f in _pl_month(j).items():
            ws[f'{hc[k]}{j + 1}'] = f'=ROUND({f},2)'
    for c in ['AA'] + list(hc.values()):
        ws[f'{c}1'].font = F_HELP
        hide(ws, c)
    mon = lambda k: f'INDEX(${hc[k]}$2:${hc[k]}$13,MONTH(P_截止))'
    yr = lambda k: f'SUM(${hc[k]}$2:${hc[k]}$13)'
    lines = [('销售收入（发货）', lambda g: g('收入')),
             ('销售成本', lambda g: f'{g("采购")}+{g("生产")}+{g("折旧")}+{g("库存")}'),
             ('毛利', lambda g: f'{g("收入")}-({g("采购")}+{g("生产")}+{g("折旧")}+{g("库存")})'),
             ('期间费用（销售管理财务）', lambda g: g('费用')),
             ('税金', lambda g: g('税金')),
             ('其他收入－其他支出', lambda g: g('其他')),
             ('利润', lambda g: f'{g("收入")}-({g("采购")}+{g("生产")}+{g("折旧")}+{g("库存")})-{g("费用")}-{g("税金")}+{g("其他")}')]
    header(ws, r + 1, [('B', '项目'), ('C', '本月'), ('D', '本年'), ('F', '钱（资金台帐）'), ('G', '本月'), ('H', '本年')], C_VIEW, height=22)
    for i, (lab, fn) in enumerate(lines):
        rr = r + 2 + i
        bold = lab in ('毛利', '利润')
        link(put(ws, f'B{rr}', lab, LINKF, FILL_TOT if bold else LBL, align=ALW), SH_PL)
        put(ws, f'C{rr}', f'=ROUND({fn(mon)},2)', F_TXTB if bold else F_AUTOB, FILL_TOT if bold else WHITE, MONEY, AR)
        put(ws, f'D{rr}', f'=ROUND({fn(yr)},2)', F_TXTB if bold else F_AUTOB, FILL_TOT if bold else WHITE, MONEY, AR)
    ycond = f'资_年月,">="&(P_年度*100+1),资_年月,"<="&P_截止年月,资_日期,"<="&P_截止'
    mcond = '资_年月,P_截止年月,资_日期,"<="&P_截止'
    money = [('回款（收客户货款）', 'SUMIFS(资_净额,资_归类,"收客户货款",{c})'),
             ('付供应商货款', 'SUMIFS(资_支出额,资_归类,"付供应商货款",{c})'),
             ('付的费用（不含买设备）', '+'.join(f'SUMIFS(资_支出额,资_归类,"{k}",{{c}})' for k in EXPENSE_KINDS)),
             ('买设备', 'SUMIFS(资_支出额,资_归类,"买设备",{c})'),
             ('老板存取、借款等（净进）', 'SUMIFS(资_净额,资_归类,"不算收支",{c})')]
    for i, (lab, f) in enumerate(money):
        rr = r + 2 + i
        put(ws, f'F{rr}', lab, F_KPI_L, LBL, align=AL)
        put(ws, f'G{rr}', '=ROUND(' + f.format(c=mcond) + ',2)', F_AUTOB, WHITE, MONEY, AR)
        put(ws, f'H{rr}', '=ROUND(' + f.format(c=ycond) + ',2)', F_AUTOB, WHITE, MONEY, AR)
    put(ws, f'J{r + 2}', '说明', F_KPI_L, LBL, align=AC)
    put(ws, f'K{r + 2}', ('利润＝发货的收入－采购、生产工资电费、折旧、库存变动－费用、税金（金额按实际成交价，带票的是含税价）。'
                         '工资电费按「所属月份」算；没估库存的月份只供参考。'
                         '「本年」＝截止日那年 1 月到截止日。'), F_NOTE, align=ALW, border=False)
    ws.merge_cells(f'K{r + 2}:L{r + 8}')
    ws.merge_cells(f'J{r + 2}:J{r + 8}')
    ws.conditional_formatting.add(f'C{r + 2}:D{r + 8}', FormulaRule(formula=[f'C{r + 2}<-0.005'], font=F_RED))
    put(ws, f'B{r + 9}', PRODUCE, F_NOTE, align=ALW, border=False)
    ws.merge_cells(f'B{r + 9}:L{r + 9}')
    ws.row_dimensions[r + 9].height = 46
    r += 11

    # ⑤ 提醒
    section(ws, r, 'B', 'L', '⑤ 提醒（点第一行看数据校验）', C_CHK)
    CR = s_check.CHECK_ROWS
    cc = lambda lab: f'N({SH_CHK}!$C${CR[lab]})'
    big = f'{cc(f"资金台帐超过 {N_CASH} 笔（多出来的不算）")}+{cc(f"销售登记超过 {N_SALE} 行")}+{cc(f"采购登记超过 {N_PUR} 行")}' \
          f'+{cc("基础资料某一块超了（多出来的不算）")}'
    rows_x = '+'.join(cc(sh) for sh, *_r in s_check.SHEETS)
    rem = [f'="数据校验：要改 "&{SH_CHK}!{s_check.CHK_X}&" 项、要看 "&{SH_CHK}!{s_check.CHK_W}&" 项"&'
           f'IF({SH_CHK}!{s_check.CHK_X}+{SH_CHK}!{s_check.CHK_W}=0,"，都没问题 √","  → 点这里去看")',
           f'=IF({rows_x}>0,"登记表有 "&({rows_x})&" 行要改（看那一行最右边的「这一行的问题」）","登记表都没有要改的 √")',
           f'=IF({cc("示例行还没删（灰底、备注写着「示例」）")}>0,"还有 "&{cc("示例行还没删（灰底、备注写着「示例」）")}&" 行示例没删（灰底、备注写着「示例」）：正式用之前整行删掉","")',
           f'=IF({cc("内部转账没配对（转出、转进合计不为 0）")}<>0,"内部转账有一笔只记了一边（差 "&TEXT({cc("内部转账没配对（转出、转进合计不为 0）")},"#,##0.00")&"）","")',
           f'=IF({cc("账户余额是负数的账户个数（按日期算到截止日）")}>0,"有 "&{cc("账户余额是负数的账户个数（按日期算到截止日）")}&" 个账户余额是负数：多半漏记了收款，或者期初余额没填","")',
           f'=IF({cc("还没定价的发货（笔）")}+{cc("还没定价的到货（笔）")}>0,"还有 "&{cc("还没定价的发货（笔）")}&" 笔发货、"&{cc("还没定价的到货（笔）")}&" 笔到货没定价：钱还没算进往来和利润","")',
           f'=IF({cc("本月固定支出过了日子还没付清（条）")}>0,"本月有 "&{cc("本月固定支出过了日子还没付清（条）")}&" 条固定支出过了日子还没付清（看资金计划）","")',
           '=IF(P_公司名称="","【基础资料】还没填公司名称（首页、对账单抬头要用）","")',
           f'=IF({big}>0,"有表超过容量了，多出来的没算进去 → 看【数据校验】","")']
    for i, f in enumerate(rem):
        rr = r + 1 + i
        put(ws, f'B{rr}', f, F_TXTB if i == 0 else F_TXT, align=AL, border=False)
        ws.merge_cells(f'B{rr}:L{rr}')
    link(ws[f'B{r + 1}'], SH_CHK)
    ws.conditional_formatting.add(f'B{r + 1}:B{r + len(rem)}', FormulaRule(
        formula=[f'OR(AND(ISNUMBER(SEARCH("要改 ",B{r + 1})),ISERROR(SEARCH("要改 0 项",B{r + 1}))),ISNUMBER(SEARCH("只记了一边",B{r + 1})),'
                 f'ISNUMBER(SEARCH("超过容量",B{r + 1})),ISNUMBER(SEARCH("负数",B{r + 1})),ISNUMBER(SEARCH("没删",B{r + 1})))'],
        font=Font(name=YH, sz=10, bold=True, color='FFC00000')))

    # 导航
    r = r + len(rem) + 2
    section(ws, r, 'B', 'L', '各张表（点名字跳过去；每张表右上角有「← 回首页」）', C_HOME)
    r += 1
    col_sets = [('B', 'C', 'E'), ('F', 'G', 'H'), ('J', 'K', 'L')]
    for grp, color, sheets in NAV:
        put(ws, f'B{r}', grp, F_SEC, fill(color), align=AL)
        ws.merge_cells(f'B{r}:L{r}')
        r += 1
        for j, (sh, desc) in enumerate(sheets):
            c1, c2, c3 = col_sets[j % 3]
            rr = r + j // 3
            link(put(ws, f'{c1}{rr}', sh, LINKF, align=AL), sh)
            put(ws, f'{c2}{rr}', desc, F_NOTE, align=ALW)
            if c3 != c2:
                ws.merge_cells(f'{c2}{rr}:{c3}{rr}')
            ws.row_dimensions[rr].height = 28
        r += (len(sheets) + 2) // 3

    # 怎么记
    r += 1
    section(ws, r, 'B', 'L', '常见情况怎么记', C_HOME)
    for i, (kk, v) in enumerate(HOWTO):
        rr = r + 1 + i
        put(ws, f'B{rr}', kk, F_TXTB, fill('FFF2F2F2'), align=ALW)
        ws.merge_cells(f'B{rr}:C{rr}')
        put(ws, f'D{rr}', v, F_TXT, align=ALW)
        ws.merge_cells(f'D{rr}:L{rr}')
        ws.row_dimensions[rr].height = 32
    r = r + len(HOWTO) + 2
    section(ws, r, 'B', 'L', '注意', C_HOME)
    notes = ['① 启用三步：【基础资料】填公司名称、建账日（期初是前一天晚上的数）、核对 8 个账户的期初余额（承兑汇票有票要填）；'
             '【往来单位】核对期初应收、期初应付；各表灰底的示例行整行删掉。然后从建账日起每天录。',
             '② 客户、供应商、账户、收支类别都从下拉选；新的先到【往来单位】【基础资料】加一行。往来单位录全称、其他叫法也认。',
             '③ 登记表、基础资料都可以插行、删行、排序（整行一起动）；绿色、紫色、橙色的表是自动的，有保护（没有密码），只有黄格子能改。',
             '④ 这本表一个人（出纳）改：采购员、业务员把要买的、要付的微信报过来录。老板看首页、资金计划、待审批付款单、利润表。',
             '⑤ 发票登记只登本公司真实开出、收到的发票。',
             f'⑥ 资金台帐最多 {N_CASH} 笔、销售和采购各 {N_SALE} 行、发票和付款审批各 {N_INV} 行；用满了另起一本（期初余额、期初欠款填上一本的余额）。']
    for i, t in enumerate(notes):
        rr = r + 1 + i
        put(ws, f'B{rr}', t, F_TXT, align=ALW, border=False)
        ws.merge_cells(f'B{rr}:L{rr}')
        ws.row_dimensions[rr].height = 30
    print_setup(ws, None, landscape=False)
    ws.oddFooter.center.text = '第 &P 页 共 &N 页'


def acc_expr(flag):
    """可用资金（flag=1）或手上承兑（flag=0，账户_可用=0 的账户）到截止日的余额"""
    return (f'(SUMIFS(账户_期初余额,账户_可用,{flag},账户_名称,"?*")'
            + ''.join(f'+IF(AND(INDEX(账户_名称,{a})<>"",INDEX(账户_可用,{a})={flag}),SUMIFS(资_净额,资_账户号,{a},资_日期,"<="&P_截止),0)'
                      for a in range(1, N_ACC + 1)) + ')')
