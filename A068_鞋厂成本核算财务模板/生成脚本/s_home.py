# -*- coding: utf-8 -*-
"""【首页】选月份、关键数（本月 / 本年累计）、资金、往来、提醒、每天录什么、各表入口；
   【数据校验】哪里录错了、漏了、账平不平、快满了（第二阶段）。

   取数规矩：别的表的格子只用 layout.py 里定死的（各录入表的校验列 / ★接口列、BS_CHECK、BS_M、会计科目表 C3、
   成本分摊表 alr()/al()、_款式月 sm_col()……）；关键数自己用 je_sum / SUMIFS 算，不依赖查看表的格子。
   数据校验的汇总格：C3＝✗ 几项、E3＝⚠ 几项（首页引用）。"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Protection
from common import *
from layout import *          # 注意：layout 的 C_IN / C_CHK / C_HOME 等颜色覆盖 common 的同名常量

YR, OPEN, CO, LAGD = P['YEAR'], P['OPEN'], P['CO'], P['LAGD']
KPI_FILL = fill('FFD9E1F2')
F_LINK = Font(name=YH, sz=10, bold=True, color='FF0563C1', underline='single')
F_LINKS = Font(name=YH, sz=10, color='FF0563C1', underline='single')
F_GOOD = Font(name=YH, sz=10, bold=True, color='FF00B050')
F_GREY = Font(name=YH, sz=10, bold=True, color='FF7F7F7F')
F_KV = Font(name=YH, sz=11, bold=True, color='FF1F3864')
F_KV10 = Font(name=YH, sz=10, bold=True, color='FF1F3864')
F_BOX = Font(name=YH, sz=11, bold=True, color='FFFFFFFF')
F_ARROW = Font(name=YH, sz=14, bold=True, color='FF7F7F7F')
F_DEMO = Font(name=YH, sz=10, bold=True, color='FFC00000')
FILL_YEL = fill('FFFFEB9C')
FILL_GREY = fill('FFF2F2F2')
QTY = '#,##0;[Red]-#,##0;"-"'

CHK_X, CHK_W = 'C3', 'E3'            # 数据校验：✗ 几项、⚠ 几项（首页引用）
CHK_HDR, CHK_R0 = 4, 5
OS_N = 600                           # 【订单汇总】最多列几张订单（开发说明 s_cost 一节）
# 首页、数据校验共用的「看哪个月」：首页 C4 不是 1～12 的数就按 12 月算
HM = f'IF(AND(ISNUMBER({HOME_M}),{HOME_M}>=1,{HOME_M}<=12),INT({HOME_M}),12)'


# ─────────────────────────── 小工具 ───────────────────────────
def _net(code, mc=None, credit=False):
    """某科目（编码前缀，含下级）发生额：credit=True 贷−借，否则 借−贷；mc＝月份条件（None＝全部）"""
    d, c = je_sum(f'"{code}*"', 'D', mc), je_sum(f'"{code}*"', 'C', mc)
    return f'({c}-{d})' if credit else f'({d}-{c})'


def _open(code):
    """【会计科目表】年初余额（正数＝正常方向）：编码前缀下的末级科目之和"""
    return f'SUMIFS({COA_OPENS},{COA_CODES},"{code}*",{COA_LEAFS},1)'


def _bal(code, mc=None, credit=False):
    """科目余额（正常方向为正）＝年初＋发生"""
    return f'ROUND({_open(code)}+{_net(code, mc, credit)},2)'


def _dup(r):
    """范围里重名的格子数（空的不算）"""
    return f'SUMPRODUCT((TRIM({r}&"")<>"")*(COUNTIF({r},{r})>1))'


def _ins(sheet, cols, chk, r0, r1):
    """插进来的行：有内容、校验格却是空的（没有公式）"""
    has = '+'.join(f'({rng(sheet, c, r0, r1)}<>"")' for c in cols)
    return f'SUMPRODUCT((({has})>0)*({rng(sheet, chk, r0, r1)}=""))'


def _lnk(ws, coord, text, sheet, ref='A1', font=F_LINK, align=AC, fill_=None, border=True):
    c = put(ws, coord, text, font, fill_, align=align, border=border)
    link(c, sheet, ref)
    return c


# ═══════════════════════════ 数据校验 ═══════════════════════════
INPUTS = [  # (表, 校验列范围, 校验列表头格, ✗ 怎么改, ⚠ 是什么)
    (SH_CASH, jr(J_CHK), f'{J_CHK}{J_HDR}',
     '在「校验」列筛选 ✗，照提示改（多是没选账户、往来单位没登记、费用支出没选费用项目）；✗ 的行不进报表',
     '提醒照样记账：用在哪的款式没登记、单位类型不对、账户余额成负数……看一眼是不是真有问题'),
    (SH_DN, dnr(DN_CHK), f'{DN_CHK}{DN_HDR}',
     '多是供应商没在【往来单位】登记、日期没填成日期、没金额；✗ 的行不算材料成本、不算应付',
     '晚到的单算到哪个月、没填收单日期、金额跟 数量×单价 差得多、用在哪的款式没登记'),
    (SH_OUT, otr(OT_CHK), f'{OT_CHK}{OT_HDR}',
     '加工厂先在【往来单位】登记（类型选外发加工厂）；日期要填成日期；✗ 的行不算加工费',
     '没填收单日期、加工厂类型不对、用在哪的款式没登记'),
    (SH_ORD, odr(OD_CHK), f'{OD_CHK}{OD_HDR}',
     '没填订单号、款式编码、数量；交货日期要填成日期；✗ 的行不算收入',
     '款式没登记（没单价、分不到成本）、没有结算单价、实交比订单多、客户没登记（「⏳ 还没交货」不算）'),
    (SH_WAGE, wgr(WG_CHK), f'{WG_CHK}{WG_HDR}',
     '月份填 1～12、选部门；✗ 的行不算工资',
     '没填姓名、用在哪的款式没登记、管理销售部门填了用在哪'),
    (SH_MJ, mjr(MJ_CHK), f'{MJ_CHK}{MJ_HDR}',
     '科目编码要在【会计科目表】、用末级科目；400104、540101～540104 不能手工记',
     '往来科目没填单位、费用科目没填费用项目、动了资金科目'),
]
CAPS = [  # (表, 序号列范围, 容量, 序号列表头)
    (SH_CASH, jr(J_SEQ), J_N), (SH_DN, dnr(DN_SEQ), DN_N), (SH_OUT, otr(OT_SEQ), OT_N), (SH_ORD, odr(OD_SEQ), OD_N),
    (SH_WAGE, wgr(WG_SEQ), WG_N), (SH_MJ, mjr(MJ_SEQ), MJ_N),
    (SH_UNIT, unr(UN_SEQ), UN_R1 - UN_R0 + 1), (SH_STY, str_(ST_SEQ), ST_N),
    (SH_MAT, rng(SH_MAT, MT_SEQ, MT_R0, MT_R1), MT_R1 - MT_R0 + 1),
    (SH_BASE, rng(SH_BASE, FA_SEQ, FA_R0, FA_R1), FA_R1 - FA_R0 + 1),
]
CHK_ROW = {}          # 检查项键 → 数据校验里的行号（首页提醒用）


def _items():
    """[(键, 检查项, 结果公式, 状态种类, 数字格式, 怎么改, 去哪改(表, 格)) 或 ('#', 分组标题)]"""
    it = [('#', '一、录入表（各表最右边「校验」列：✗ 的行进不了账，⚠ 照样记账）')]
    for sh, chk, hdr, how_x, how_w in INPUTS:
        it.append((f'x:{sh}', f'【{sh}】✗ 要改的行', f'COUNTIF({chk},"✗*")', 'x', '0" 行"', how_x, (sh, hdr)))
        it.append((f'w:{sh}', f'【{sh}】⚠ 提醒的行', f'COUNTIF({chk},"⚠*")', 'w', '0" 行"', how_w, (sh, hdr)))
    it.append(('nojz', '【资金日记账】有金额、没进账的笔数', f'COUNTIFS({jr(J_NET)},"<>0",{jr(J_OK)},0)', 'x', '0" 笔"',
               '这些钱算进了账户余额，却没进报表（校验是 ✗ 的行）：照「校验」列的提示改好就进账了', (SH_CASH, f'{J_CHK}{J_HDR}')))
    ins = '+'.join([
        _ins(SH_CASH, (J_DATE, J_ACC, J_IN, J_OUT), J_CHK, J_R0, J_R1),
        _ins(SH_DN, (DN_DATE, DN_SUP, DN_NAME, DN_QTY, DN_AMTIN), DN_CHK, DN_R0, DN_R1),
        _ins(SH_OUT, (OT_DATE, OT_SUP, OT_QTY, OT_AMTIN), OT_CHK, OT_R0, OT_R1),
        _ins(SH_ORD, (OD_NO, OD_STY, OD_QTY, OD_DDATE), OD_CHK, OD_R0, OD_R1),
        _ins(SH_WAGE, (WG_MON, WG_NAME, WG_DEPT, WG_BASE), WG_CHK, WG_R0, WG_R1),
        _ins(SH_MJ, (MJ_DATE, MJ_DRIN, MJ_CRIN, MJ_AMTIN), MJ_CHK, MJ_R0, MJ_R1)])
    it.append(('ins', '录入表里插进来的行（没有公式，不会算）', ins, 'x', '0" 行"',
               '录入表里插了行：校验列是空的就是插的行。把那几行内容剪下来粘到最下面的空行，再把插的行删掉', (SH_CASH, 'A1')))

    it.append(('#', '二、账平不平'))
    coa_d = (f'ROUND(SUMIFS({COA_OPENS},{COA_LEAFS},1,{COA_DIRS},"借")'
             f'-SUMIFS({COA_OPENS},{COA_LEAFS},1,{COA_DIRS},"贷"),2)')
    coa_c3 = cell(SH_COA, 'C3')
    it.append(('coa', '期初余额平不平（【会计科目表】C3：借方减贷方）', coa_d,
               f'=IF(OR(LEFT({coa_c3}&"",1)="✗",ABS(N({{v}}))>=0.01),"✗","√")', MONEY,
               '【会计科目表】年初余额：资产要等于负债＋权益。多半是「年初未分配利润」（3104）没填对，可以倒挤', (SH_COA, 'C3')))
    it.append(('bs', f'="资产负债表平不平（"&{BS_M}&" 月末，资产−负债和权益）"', f'N({BS_CHECK})', 'x0', MONEY,
               '【资产负债表】H3 不是 0 说明有公式被改坏了（或者自己加的科目没归类），请联系做表的人', (SH_BS, 'H3')))
    now_r = rng(SH_BASE, AC_NOW, AC_R0, AC_R1)
    terms = []
    for k in range(1, AC_R1 - AC_R0 + 2):
        code = f'INDEX({AC_CODES},{k})&""'
        jb = f'N(INDEX({AC_OPENS},{k}))+{je_sum(code, "D")}-{je_sum(code, "C")}'
        terms.append(f'IF(OR(TRIM(INDEX({AC_NAMES},{k})&"")="",{code}=""),0,'
                     f'IF(ABS(ROUND(N(INDEX({now_r},{k}))-({jb}),2))>=0.01,1,0))')
    it.append(('acc', '资金账户：日记账算的余额跟资金科目余额对不上的账户', '+'.join(terms), 'w', '0" 个"',
               '多半是【资金日记账】有 ✗ 的行（有金额没进账），或者【手工分录】动了资金科目；改好就对上了。逐个账户看【账户余额表】',
               (SH_ACCB, 'A1')))
    diff = rng(SH_BASE, AC_DIFF, AC_R0, AC_R1)
    it.append(('real', '资金账户：跟银行 App / 支付宝实际余额有差额的账户', f'COUNTIF({diff},">=0.01")+COUNTIF({diff},"<=-0.01")',
               'w', '0" 个"', '【基础资料】② 填的实际余额跟表里算的不一样：漏记、记错金额或者记错账户', (SH_BASE, f'{AC_REAL}{AC_R0}')))
    it.append(('rev', '主营业务收入 跟【订单明细】交货金额 差多少',
               f'ROUND({_net("5001", None, True)}-SUMIFS({odr(OD_AMT)},{odr(OD_OK)},1),2)', 'w0', MONEY,
               '收入按【订单明细】交货金额自动记；不一样多半是【手工分录】记了 5001（调账），确认一下是不是故意的', (SH_IS, 'A1')))
    moh = '+'.join(f'IF(ABS({_bal(MOH, f"{chr(34)}<={chr(34)}&{m}")})>=0.01,1,0)' for m in range(1, 13))
    it.append(('moh', '制造费用（4101）月末余额不是 0 的月份', moh, 'x', '0" 个月"',
               '制造费用每月末自动转到生产成本，余额应该是 0；不是 0 多半是【会计科目表】4101 填了年初余额，或者公式被改坏了',
               (SH_TB, 'A1')))

    it.append(('#', '三、成本和订单'))
    al_bad = '+'.join(f'IF(ABS(ROUND(SUM({smr("成本", m)})-N({alr("转出合计", m)}),2))>=0.01,1,0)' for m in range(1, 13))
    it.append(('alloc', '【成本分摊表】「核对」不平的月份（各款式成本合计 − 本月转出合计）', al_bad, 'x', '0" 个月"',
               '各款式分到的成本加起来应该正好等于每月转出合计；不平说明公式被改坏了，请联系做表的人', (SH_ALLOC, f'A{AL_CHECK_ROW}')))
    sm_cost = f"{q(SH_SM)}!${sm_col('成本', 1)}${SM_R0}:${sm_col('成本', 12)}${SM_R1}"
    it.append(('ocost', '订单成本合计 跟 各款式成本合计 差多少', f'ROUND(SUM({odr(OD_COSTV)})-SUM({sm_cost}),2)', 'x1', MONEY,
               '订单每行成本＝实交双数×那个款那个月的单双成本，差几分钱是四舍五入（1 元以内不用管）；差得多请联系做表的人',
               (SH_OSUM, 'A1')))
    it.append(('wip', f'="生产成本余额（在制，到 "&{HM}&" 月末；月份跟首页走）"', _bal('4001', f'"<="&{HM}'), 'i+', MONEY,
               '＝还挂在车间、没交货的款的直接成本（等交货自动转出）＋没交货月份结转下来的公共成本，有数是正常的；负数要查',
               (SH_ALLOC, f'A{AL_ROWS["月末在制"]}')))
    it.append(('snew', '订单里有、【款式档案】没登记的款式', f'COUNT({odr(OD_SNEW)})', 'w', '0" 个"',
               '【款式档案】右边列出来了，抄过去登记（填结算单价）；不登记没有单价、分不到成本', (SH_STY, f'{ST_NCODE}{ST_HDR}')))
    it.append(('p0', '交了货、结算单价是 0 的行', f'COUNTIFS({odr(OD_DQ)},"<>0",{odr(OD_PRICEU)},0)', 'w', '0" 行"',
               '在【款式档案】填结算单价，或者在【订单明细】这一行填「结算单价」；单价 0 的交货没有收入', (SH_ORD, f'{OD_PRICE}{OD_HDR}')))
    it.append(('nnew', '送货单里有、【品名档案】没登记的品名', f'COUNT({dnr(DN_NNEW)})', 'w', '0" 个"',
               '【品名档案】右边列出来了，抄过去选个材料类别；不登记算「未分类」，不影响成本', (SH_MAT, f'{MT_NNAME}{MT_HDR}')))
    it.append(('late', f'="晚到的单（收单比送货晚 "&{LAGD}&" 天以上）：送货单＋外发"',
               f'COUNT({dnr(DN_LATEK)})+COUNT({otr(OT_LATEK)})', 'i', '0" 行"',
               '供应商单子拿来得晚，录进去以后那个月的成本会跟着变（除非结了账）；看【跨月补单】，月底催供应商把单子送来',
               (SH_LATE, 'A1')))

    it.append(('#', '四、往来和工资'))
    apk = (f'SUMPRODUCT(({unr(UN_APK)}<>"")*(({unr(UN_AP0)}+SUMIFS({jer(JE_AMT)},{jer(JE_CR)},"2202",{jer(JE_UNIT)},{UN_NAMES})'
           f'-SUMIFS({jer(JE_AMT)},{jer(JE_DR)},"2202",{jer(JE_UNIT)},{UN_NAMES}))<-0.005))')
    it.append(('apneg', '付的钱比登记的送货单 / 加工单多的供应商、加工厂（全部已记的）', apk, 'w', '0" 家"',
               '多半还有送货单没拿到（或者是预付款）：问供应商要单子；看【应付账款汇总】最右边的提示', (SH_APS, 'A1')))
    nocust = (f'ROUND(SUMIFS({odr(OD_AMT)},{odr(OD_OK)},1)'
              f'-SUMPRODUCT(({AX_CUST}<>"")*SUMIFS({odr(OD_AMT)},{odr(OD_OK)},1,{odr(OD_CUSTK)},{AX_CUST})),2)')
    it.append(('nocust', '交货收入里客户没登记成「客户」的金额（记在「未登记客户」名下）', nocust, 'w0', MONEY,
               '【订单明细】的客户（空＝默认客户）要在【往来单位】登记、类型选「客户」，不然应收按客户看不准', (SH_UNIT, f'{UN_TYPE}{UN_HDR}')))
    it.append(('ncust', f'客户家数（记账分录按客户分收入，最多 {JE_NCUST} 家）', f'COUNT({unr(UN_ARK)})',
              f'=IF(N({{v}})>{JE_NCUST},"⚠","√")', '0" 家"',
               f'超过 {JE_NCUST} 家时，第 {JE_NCUST + 1} 家以后的收入记在「未登记客户」名下；不再往来的客户把类型改成「其他」',
               (SH_UNIT, f'{UN_TYPE}{UN_HDR}')))
    it.append(('wage', '未发工资（应付职工薪酬 2211 余额，全部已记的）', _bal('2211', None, True), 'i+', MONEY,
               '＝【工资登记】算了、还没在【资金日记账】记「付工资」的；一般是最近一个月的。负数＝发的比应发多', (SH_WSUM, 'A1')))

    it.append(('#', '五、档案和设置'))
    it.append(('open', '建账日期（【基础资料】①）', f'IF(ISNUMBER({OPEN}),{OPEN},"")',
               f'=IF(NOT(ISNUMBER({OPEN})),"✗",IF(YEAR({OPEN})>{YR},"✗",IF(OR(YEAR({OPEN})<{YR},DAY({OPEN})<>1),"⚠","√")))', DATE,
               '要填成日期，不能晚于会计年度；最好是某个月的 1 号（报表按月算）。在上一年＝新的一年接着用：期初余额要改成 1 月 1 日的数',
               (SH_BASE, f'{PA_VAL}{PA_ROW["OPEN"]}')))
    it.append(('accode', '资金账户算不出科目编码（没选类型 / 银行超过 6 个 / 支付宝微信超过 4 个）',
               f'SUMPRODUCT((TRIM({AC_NAMES}&"")<>"")*({AC_CODES}=""))', 'x', '0" 个"',
               '【基础资料】② 账户要选类型（银行 / 现金 / 支付宝微信）', (SH_BASE, f'{AC_TYPE}{AC_R0}')))
    it.append(('fecls', '费用项目没选归类', f'SUMPRODUCT((TRIM({FE_NAMES}&"")<>"")*({FE_CODES}=""))', 'x', '0" 个"',
               '【基础资料】④ 每个费用项目要选归类（制造费用 / 管理费用 / 销售费用……），不然算不出科目', (SH_BASE, f'{FE_CLS}{FE_R0}')))
    it.append(('untype', '往来单位没选类型', f'SUMPRODUCT((TRIM({UN_NAMES}&"")<>"")*(TRIM({UN_TYPES_R}&"")=""))', 'w', '0" 家"',
               '【往来单位】类型决定资金日记账怎么自动认（材料供应商＝付材料款、客户＝收货款……）', (SH_UNIT, f'{UN_TYPE}{UN_HDR}')))
    it.append(('dupun', '往来单位重名', _dup(UN_NAMES), 'x', '0" 个"',
               '同名的加个字区分（比如 张伟(电工)）；重名会把两家的账算到一起', (SH_UNIT, f'{UN_NAME}{UN_HDR}')))
    it.append(('dupst', '款式编码重名', _dup(ST_KEYS), 'x', '0" 个"',
               '【款式档案】同一个款式编码只能登记一行（重的那行标红了）', (SH_STY, f'{ST_CODE}{ST_HDR}')))
    it.append(('dupmt', '品名档案重名', _dup(rng(SH_MAT, MT_KEY, MT_R0, MT_R1)), 'w', '0" 个"',
               '【品名档案】同一个品名只登记一行（空格、全角半角括号不算区别）', (SH_MAT, f'{MT_NAME}{MT_HDR}')))
    it.append(('dupbase', '【基础资料】里重名（资金账户、收支类别、费用项目、部门）',
               '+'.join(_dup(r) for r in (AC_NAMES, CT_NAMES, FE_NAMES, DP_NAMES)), 'x', '0" 个"',
               '同一块里名字不能重复（下拉、自动认都靠名字）', (SH_BASE, 'A4')))
    fa_d, fa_c = rng(SH_BASE, FA_DATE, FA_R0, FA_R1), rng(SH_BASE, FA_COST, FA_R0, FA_R1)
    fa = (f'ROUND(SUMIFS({fa_c},{fa_d},">="&MAX({OPEN},DATE({YR},1,1)),{fa_d},"<="&DATE({YR},12,31))'
          f'-{je_sum(chr(34) + "1601*" + chr(34), "D")},2)')
    it.append(('fa', '建账后买的固定资产原值 跟 固定资产（1601）借方 差多少', fa, 'w0', MONEY,
               '买设备在【资金日记账】记「购置设备」，还要在【基础资料】⑦ 登记一行（才会提折旧）；正数＝登记了、钱没记，负数＝钱记了、没登记',
               (SH_BASE, f'{FA_NAME}{FA_R0}')))
    fund = '+'.join(f'(LEFT({mjr(x)},4)="{p}")' for x in (MJ_DR, MJ_CR) for p in ('1001', '1002', '1012'))
    it.append(('mjfund', '【手工分录】动了资金科目的行', f'SUMPRODUCT(({mjr(MJ_OK)}=1)*(({fund})>0))', 'w', '0" 行"',
               '资金日记账的余额里没有这一笔，账户余额会跟科目对不上；收付款尽量在【资金日记账】记', (SH_MJ, f'{MJ_CHK}{MJ_HDR}')))

    it.append(('#', '六、容量（用到第几行 / 一共几行，超过 90% 提醒）'))
    for sh, seq, cap in CAPS:
        lab = f'【{sh}】用到第几行' if sh != SH_BASE else '【基础资料】⑦ 固定资产用到第几行'
        it.append((f'cap:{sh}', lab, f'MAX({seq})', f'cap{cap}', f'0"/{cap}"',
                   '快满了请联系做表的人加行（或者新的一年另存一本）', (sh, 'A1')))
    it.append(('cap:os', f'订单张数（【订单汇总】最多列 {OS_N} 张）', f'COUNT({odr(OD_OFIRST)})', f'cap{OS_N}', f'0"/{OS_N}"',
               '超过了【订单汇总】只列前面的，订单明细照样算', (SH_OSUM, 'A1')))
    return it


def _status(kind, v):
    if kind.startswith('='):
        return kind.replace('{v}', v)
    if kind == 'x':
        return f'=IF(N({v})>0,"✗","√")'
    if kind == 'w':
        return f'=IF(N({v})>0,"⚠","√")'
    if kind == 'x0':
        return f'=IF(ABS(N({v}))>=0.01,"✗","√")'
    if kind == 'w0':
        return f'=IF(ABS(N({v}))>=0.01,"⚠","√")'
    if kind == 'x1':
        return f'=IF(ABS(N({v}))>=1,"✗","√")'
    if kind == 'i':
        return f'=IF(N({v})=0,"√","○")'
    if kind == 'i+':
        return f'=IF(N({v})<=-0.01,"⚠",IF(ABS(N({v}))<0.01,"√","○"))'
    if kind.startswith('cap'):
        return f'=IF(N({v})>{int(kind[3:])}*0.9,"⚠","√")'
    raise ValueError(kind)


def build_chk(wb, ctx):
    ws = wb[SH_CHK]
    widths(ws, {'A': 5, 'B': 46, 'C': 15, 'D': 9, 'E': 66, 'F': 18})
    title(ws, '数 据 校 验（录错的、漏录的、账平不平、快满了的，都在这里）', 'F', C_CHK,
          '💡 全自动，打开就是最新的。状态：✗＝一定要改（不改报表就不准）；⚠＝看一下是不是真有问题；√＝没问题；○＝只是告诉你一个数（正常情况也会有）。'
          '点最右边「去哪改」跳到那张表，在那张表的「校验」列筛选 ✗ / ⚠ 就能找到是哪几行。首页上的 ✗ / ⚠ 项数就是 C3、E3 这两格。')
    home_link(ws, 'G1')
    ws.column_dimensions['G'].width = 10
    items = _items()
    r = CHK_R0
    seq = 0
    for item in items:
        if item[0] == '#':
            section(ws, r, 'A', 'F', item[1], 'FF7F7F7F')
            ws.row_dimensions[r].height = 20
            r += 1
            continue
        key, lab, f, kind, fmt, how, (sh, ref) = item
        seq += 1
        CHK_ROW[key] = r
        put(ws, f'A{r}', seq, F_AUTO, align=AC)
        put(ws, f'B{r}', lab, F_TXT, align=ALW)
        put(ws, f'C{r}', f'={f}', F_AUTOB, FILL_AUTO, fmt, AR)
        put(ws, f'D{r}', _status(kind, f'C{r}'), F_TXTB, align=AC)
        put(ws, f'E{r}', how, F_NOTE, align=ALW)
        _lnk(ws, f'F{r}', f'去【{sh}】', sh, ref, F_LINKS)
        ws.row_dimensions[r].height = 30 if len(how) > 40 or len(str(lab)) > 26 else 20
        r += 1
    r1 = r - 1
    header(ws, CHK_HDR, [('A', '序号'), ('B', '检查什么'), ('C', '结果'), ('D', '状态'), ('E', '怎么改'), ('F', '去哪改')], C_CHK)
    # 第 3 行：✗ 几项（C3）、⚠ 几项（E3）——首页引用
    put(ws, 'A3', None, None, KPI_FILL)
    put(ws, 'B3', '✗ 一定要改的（项）', F_KPI_L, KPI_FILL, align=AC)
    put(ws, CHK_X, f'=COUNTIF($D${CHK_R0}:$D${r1},"✗")', F_KPI_V, KPI_FILL, '0" 项"', AC)
    put(ws, 'D3', '⚠ 要看', F_KPI_L, KPI_FILL, align=ACW)
    put(ws, CHK_W, f'=COUNTIF($D${CHK_R0}:$D${r1},"⚠")', F_KPI_V, KPI_FILL, '0" 项（看一下是不是真有问题）"', AL)
    put(ws, 'F3', f'=IF(C3+E3=0,"√ 都没问题","")', F_GOOD, KPI_FILL, align=AC)
    ws.row_dimensions[3].height = 28
    ws.conditional_formatting.add(CHK_X, FormulaRule(formula=[f'N({CHK_X})>0'], fill=FILL_WARN, font=F_RED))
    ws.conditional_formatting.add(CHK_W, FormulaRule(formula=[f'N({CHK_W})>0'], fill=FILL_YEL))
    rg = f'D{CHK_R0}:D{r1}'
    ws.conditional_formatting.add(rg, FormulaRule(formula=[f'$D{CHK_R0}="✗"'], fill=FILL_WARN, font=F_RED))
    ws.conditional_formatting.add(rg, FormulaRule(formula=[f'$D{CHK_R0}="⚠"'], fill=FILL_YEL))
    ws.conditional_formatting.add(rg, FormulaRule(formula=[f'$D{CHK_R0}="√"'], font=F_GOOD))
    ws.conditional_formatting.add(rg, FormulaRule(formula=[f'$D{CHK_R0}="○"'], font=F_GREY))
    ws.conditional_formatting.add(f'C{CHK_R0}:C{r1}', FormulaRule(
        formula=[f'$D{CHK_R0}="✗"'], font=F_RED))
    put(ws, f'B{r1 + 2}', '说明：「结果」是个数的，0 就是没问题；是金额的，0 就是对得上。资产负债表的月份跟【资产负债表】C3 走；'
                          '生产成本余额的月份跟【首页】C4 走；其余都是全部已记的数。', F_NOTE, align=ALW, border=False)
    ws.merge_cells(f'B{r1 + 2}:F{r1 + 3}')
    ws.freeze_panes = f'A{CHK_R0}'
    print_setup(ws, f'{CHK_HDR}:{CHK_HDR}', landscape=True)
    return ws


# ═══════════════════════════ 首页 ═══════════════════════════
NAV = [
    ('录入（每天 / 每笔录 · 蓝色）', C_IN, [
        (SH_CASH, '银行1、银行2、现金、支付宝混在一张表，一笔一行，余额逐行算'),
        (SH_DN, '供应商送货单，一个品名一行；单子晚到也照送货日期记'),
        (SH_OUT, '外发加工厂的加工单（加工费）；能分到款式就填款式'),
        (SH_ORD, '电商部采购单整张粘进来；交货了填交货日期'),
        (SH_WAGE, '每人每月一行；计件的可以按款式分几行'),
        (SH_MJ, '调账、暂估、计提用，平时一般不用'),
    ]),
    ('成本、订单（自动 · 绿色）', C_VIEW, [
        (SH_SPL, '每个款式的双数、收入、成本（材料/外发/人工/制造）、毛利'),
        (SH_OSUM, '每张订单交了多少、还差多少、赚了多少'),
        (SH_OQ, '选一张订单号，看逐行交货、成本、毛利'),
        (SH_ALLOC, '每月成本怎么分到款式；只有「在制估计」一行可以填'),
        (SH_MSUM, '材料按类别、按品名看花了多少、单价涨没涨'),
        (SH_LATE, '供应商单子晚到多久、晚到的是哪几张'),
    ]),
    ('往来、资金（自动 · 绿色）', C_VIEW, [
        (SH_APS, '欠每家供应商 / 加工厂多少、最后什么时候付的'),
        (SH_SST, '选供应商和日期，打印对账单'),
        (SH_CST, '跟电商部对账：交货、收款、还欠多少'),
        (SH_ACCB, '每个账户每月余额、收支类别每月汇总'),
        (SH_FEE, '水电房租运费……每个费用项目每月花多少'),
        (SH_WSUM, '各部门每月工资、发了多少、还欠多少'),
    ]),
    ('报表（自动 · 深红）', C_RPT, [
        (SH_IS, '1～12 月利润表：收入、营业成本、费用、净利润'),
        (SH_BS, '选月份看资产负债表（家底）'),
        (SH_CF, '1～12 月现金流量表：钱从哪来、花到哪去'),
        (SH_TB, '选月份看每个科目的发生额和余额'),
        (SH_GL, '选科目（和往来单位）看逐笔明细'),
        (SH_JE, '全部记账分录（自动生成，只能看）'),
    ]),
    ('校验、档案（偶尔改 · 灰色）', C_ARC, [
        (SH_CHK, '哪里录错了、漏了、账平不平、快满了'),
        (SH_BASE, '参数（年度、建账日期）、资金账户、收支类别、费用项目、部门、固定资产'),
        (SH_UNIT, '供应商、加工厂、客户、老板；期初欠款'),
        (SH_STY, '款式编码、结算单价、分摊系数'),
        (SH_MAT, '品名属于哪类材料（可以不填）'),
        (SH_COA, '会计科目、年初余额'),
    ]),
]
FLOW_IN = [(SH_CASH, '每一笔收付款（4 个账户混着记，只填收入或支出一边）'),
           (SH_DN, '供应商拿来送货单就记（送货日期照单子写，收单日期写拿到那天）'),
           (SH_OUT, '外发加工厂的回货单 / 结算单'),
           (SH_ORD, '采购单粘进来；交货了填交货日期（分批交就复制一行）'),
           (SH_WAGE, '月底每人一行（计件的可按款式分）')]
FLOW_OUT = [('成本', SH_SPL, C_VIEW, '款式成本利润 · 成本分摊表'),
            ('订单', SH_OSUM, C_VIEW, '订单汇总 · 订单查询'),
            ('利润', SH_IS, C_RPT, '利润表（每月）'),
            ('对账单', SH_SST, C_VIEW, '供应商对账单 · 客户对账单 · 应付汇总'),
            ('报表', SH_BS, C_RPT, '资产负债表 · 现金流量表 · 科目余额表')]
PAIRS = [('B', 'C'), ('D', 'E'), ('F', 'G'), ('H', 'I'), ('J', 'K')]
HROW = 4                       # HOME_M 所在行（C4）
HX = '$Z$4'                    # 首页隐藏帮手：实际用的月份（C4 不是 1～12 就按 12 月）


def _box(ws, r, c1, c2, text, color, sheet=None, font=F_BOX, height=None):
    if c1 != c2:
        ws.merge_cells(f'{c1}{r}:{c2}{r}')
    c = put(ws, f'{c1}{r}', text, font, fill(color), align=ACW)
    for i in range(CI(c1) + 1, CI(c2) + 1):
        ws.cell(row=r, column=i).border = BD
    if sheet:
        link(c, sheet)
    if height:
        ws.row_dimensions[r].height = height
    return c


def _note(ws, r, text, c1='B', c2='K', height=None, font=F_NOTE):
    ws.merge_cells(f'{c1}{r}:{c2}{r}')
    c = put(ws, f'{c1}{r}', text, font, align=ALW, border=False)
    if height:
        ws.row_dimensions[r].height = height
    return c


def build_home(wb, ctx):
    ws = wb[SH_HOME]
    assert HOME_M == cell(SH_HOME, 'C4')
    widths(ws, {'A': 2, **{CL(i): 15 for i in range(2, 12)}, 'L': 2})
    title(ws, '鞋 厂 成 本 核 算 财 务 账', 'K', C_HOME,
          '💡 平时只录蓝色那 5 张（资金日记账、送货单登记、外发加工登记、订单明细、工资登记），成本、利润、对账单、报表全是自动的。'
          'C4 黄格选月份，下面「本月」「本年累计」跟着变。点表名就能跳过去，每张表右上角有「← 回首页」。灰色、白色格子是公式，不要往里打字。')
    ws['A1'].value = f'={CO}&"　"&{YR}&" 年 · 成本核算财务账（按款式 × 月算成本利润）"'

    # ── 第 4 行：看哪个月（HOME_M）
    selector(ws, f'B{HROW}', '看哪个月', f'C{HROW}', 9, f'={AX_M}', '0"月"', '选 1～12：下面「本月」「本年累计」跟着变')
    ws[HX.replace('$', '')] = f'={HM.replace(HOME_M, "$C$4")}'
    ws[HX.replace('$', '')].font = F_HELP
    ws.merge_cells(f'D{HROW}:K{HROW}')
    put(ws, f'D{HROW}', f'=IF({HX}<>$C$4,"⚠ C4 要选 1～12 的月份（现在按 12 月算）",'
                        f'"本月＝"&{HX}&" 月，本年累计＝1～"&{HX}&" 月；资金是现在的余额（全部已记的），应付、应收是 "&{HX}&" 月末的")',
        F_NOTE, align=AL, border=False)
    ws.row_dimensions[HROW].height = 24
    ws.conditional_formatting.add(f'D{HROW}', FormulaRule(formula=[f'LEFT($D${HROW},1)="⚠"'], font=F_RED))
    r = HROW + 1
    if ctx.get(SH_CASH) or ctx.get(SH_ORD):
        c = _note(ws, r, '现在表里是演示数据（备注写「示例」的都是编的；中茂 9/5 送货单、电商部采购单 353651 是照片上的真单子）。'
                         '开始记自己的账：把 5 张录入表、手工分录的演示行清空（只清内容，不要删行），再改【基础资料】【往来单位】【款式档案】的期初和档案。'
                         '这一行看完可以选中按 Delete 删掉。', height=34, font=F_DEMO)
        c.fill = FILL_TIP
        c.protection = Protection(locked=False)
        r += 1
    r += 1

    # ── 经营情况：本月 / 本年累计
    section(ws, r, 'B', 'K', f'="经营情况（"&{HX}&" 月 ／ 本年 1～"&{HX}&" 月累计 · 单位：元）"', C_HOME)
    r += 1
    heads = ['交货双数', '交货收入', '营业成本', '毛利', '毛利率', '期间费用', '净利润', '净利率', '平均每双成本']
    hdr = r
    put(ws, f'B{r}', '', F_HDR, fill('FF2F5597'), align=ACW)
    for i, t in enumerate(heads):
        put(ws, f'{CL(3 + i)}{r}', t, F_HDR, fill('FF2F5597'), align=ACW)
    ws.row_dimensions[r].height = 22
    dq_blk = f"{q(SH_SM)}!${sm_col('双数', 1)}${SM_R0}:${sm_col('双数', 12)}${SM_R1}"
    dq_mon = f"{q(SH_SM)}!${sm_col('双数', 1)}${SM_MROW}:${sm_col('双数', 12)}${SM_MROW}"
    for k, (lab, op, mc) in enumerate([(f'="本月（"&{HX}&" 月）"', '=', HX), (f'="本年累计（1～"&{HX}&" 月）"', '<=', f'"<="&{HX}')]):
        r += 1
        put(ws, f'B{r}', lab, F_TXTB, KPI_FILL, align=AC)
        C, D, E, F, G, H, I, J, K = (f'{CL(3 + i)}{r}' for i in range(9))
        oth = f'Z{r}'          # 隐藏帮手：税金及附加＋营业外支出＋所得税－营业外收入
        ws[oth] = f'=ROUND({_net("5403", mc)}+{_net("5711", mc)}+{_net("5801", mc)}-{_net("5301", mc, True)},2)'
        ws[oth].font = F_HELP
        vals = {
            C: (f'=SUMPRODUCT(({dq_mon}{op}{HX})*{dq_blk})', QTY),
            D: (f'=ROUND({_net("5001", mc, True)}+{_net("5051", mc, True)},2)', MONEY0),
            E: (f'=ROUND({_net("5401", mc)}+{_net("5402", mc)},2)', MONEY0),
            F: (f'=ROUND({D}-{E},2)', MONEY0),
            G: (f'=IF(N({D})=0,"",ROUND({F}/{D},4))', PCT),
            H: (f'=ROUND({_net("5601", mc)}+{_net("5602", mc)}+{_net("5603", mc)},2)', MONEY0),
            I: (f'=ROUND({F}-{H}-{oth},2)', MONEY0),
            J: (f'=IF(N({D})=0,"",ROUND({I}/{D},4))', PCT),
            K: (f'=IF(N({C})=0,"",ROUND({E}/{C},2))', MONEY),
        }
        for coord, (f, fmt) in vals.items():
            put(ws, coord, f, F_KV, fill('FFFFFFFF'), fmt, AC)
        ws.row_dimensions[r].height = 26
    kr0, kr1 = hdr + 1, r
    ws.conditional_formatting.add(f'C{kr0}:K{kr1}', FormulaRule(formula=[f'AND(ISNUMBER(C{kr0}),C{kr0}<0)'], font=F_RED))
    r += 1
    c = _note(ws, r, '交货双数、交货收入按【订单明细】交货日期算到月（收入含其他业务收入）；营业成本＝【成本分摊表】每月结转的（只算交了货的鞋，'
                     '没交货的挂在生产成本里）；期间费用＝销售＋管理＋财务费用；净利润跟【利润表】一样（再减税金及附加、营业外收支、所得税）。'
                     '点这里看【利润表】 →', height=32, font=F_LINKS)
    link(c, SH_IS)
    r += 2

    # ── 资金：每个账户当前余额 ＋ 总余额
    section(ws, r, 'B', 'K', '资金（现在的余额 ＝ 期初 ＋ 资金日记账全部已记的）', C_HOME)
    r += 1
    now_r = rng(SH_BASE, AC_NOW, AC_R0, AC_R1)
    n_ac = AC_R1 - AC_R0 + 1
    _lnk(ws, f'B{r}', '总余额', SH_CASH, 'A1', Font(name=YH, sz=10, bold=True, color='FF1F3864', underline='single'), fill_=KPI_FILL)
    put(ws, f'B{r + 1}', f'=SUM({now_r})', F_KPI_V, fill('FFFFFFFF'), MONEY, AC)
    for k in range(1, n_ac + 1):
        col = CL(2 + k)
        put(ws, f'{col}{r}', f'=IF(TRIM(INDEX({AC_NAMES},{k})&"")="","",INDEX({AC_NAMES},{k})&"")', F_KPI_L, KPI_FILL, align=ACW)
        put(ws, f'{col}{r + 1}', f'=IF({col}{r}="","",N(INDEX({now_r},{k})))', F_KV10, fill('FFFFFFFF'), MONEY, AC)
    _lnk(ws, f'K{r}', '跟银行 App\n对不上的', SH_BASE, f'{AC_REAL}{AC_R0}',
         Font(name=YH, sz=10, bold=True, color='FF1F3864', underline='single'), fill_=KPI_FILL)
    ws[f'K{r}'].alignment = ACW
    put(ws, f'K{r + 1}', f"={q(SH_CHK)}!$C${CHK_ROW['real']}", F_KV10, fill('FFFFFFFF'), '0" 个"', AC)
    ws.row_dimensions[r].height = 30
    ws.row_dimensions[r + 1].height = 26
    ws.conditional_formatting.add(f'B{r + 1}:J{r + 1}', FormulaRule(formula=[f'AND(ISNUMBER(B{r + 1}),B{r + 1}<0)'], font=F_RED))
    ws.conditional_formatting.add(f'K{r + 1}', FormulaRule(formula=[f'N(K{r + 1})>0'], fill=FILL_YEL))
    r += 2
    _note(ws, r, '账户余额跟【基础资料】② 的「当前余额」一样（含内部转账）。每月照银行 App / 支付宝在【基础资料】② 填一次实际余额，差额是 0 就对了。',
          height=20)
    r += 2

    # ── 往来和要留意的
    section(ws, r, 'B', 'K', f'="往来和要留意的（应付、应收到 "&{HX}&" 月末）"', C_HOME)
    r += 1
    mc = f'"<="&{HX}'
    tiles = [
        (('B', 'C'), f'="应付供应商 / 加工厂（"&{HX}&" 月末）"', f'={_bal("2202", mc, True)}', MONEY, SH_APS),
        (('D', 'E'), f'="电商部 / 客户欠我们（"&{HX}&" 月末）"', f'={_bal("1122", mc)}', MONEY, SH_CST),
        (('F', 'G'), '还没交货的订单双数',
         f'=MAX(0,SUM({odr(OD_OQ)})-SUMIFS({odr(OD_DQ)},{odr(OD_DQ)},">0"))', '#,##0" 双"', SH_OSUM),
        (('H', 'I'), f'="晚到的送货单（晚 "&{LAGD}&" 天以上）"', f'=COUNT({dnr(DN_LATEK)})', '0" 行"', SH_LATE),
        (('J', 'J'), '数据校验\n✗ 要改', f"={q(SH_CHK)}!${CHK_X[0]}${CHK_X[1:]}", '0" 项"', SH_CHK),
        (('K', 'K'), '数据校验\n⚠ 要看', f"={q(SH_CHK)}!${CHK_W[0]}${CHK_W[1:]}", '0" 项"', SH_CHK),
    ]
    for (c1, c2), lab, f, fmt, sh in tiles:
        if c1 != c2:
            ws.merge_cells(f'{c1}{r}:{c2}{r}')
            ws.merge_cells(f'{c1}{r + 1}:{c2}{r + 1}')
        _lnk(ws, f'{c1}{r}', lab, sh, 'A1', Font(name=YH, sz=10, bold=True, color='FF1F3864', underline='single'),
             fill_=KPI_FILL, align=ACW)
        put(ws, f'{c1}{r + 1}', f, F_KPI_V, fill('FFFFFFFF'), fmt, AC)
        for rr in (r, r + 1):
            ws[f'{c2}{rr}'].border = BD
    ws.row_dimensions[r].height = 30
    ws.row_dimensions[r + 1].height = 28
    ws.conditional_formatting.add(f'J{r + 1}', FormulaRule(formula=[f'N(J{r + 1})>0'], fill=FILL_WARN, font=F_RED))
    ws.conditional_formatting.add(f'K{r + 1}', FormulaRule(formula=[f'N(K{r + 1})>0'], fill=FILL_YEL))
    r += 3

    # ── 提醒（引用数据校验的结果格）
    section(ws, r, 'B', 'K', '提醒', 'FFC00000')
    r += 1
    cr = lambda key: f"{q(SH_CHK)}!$C${CHK_ROW[key]}"
    X, W = f"{q(SH_CHK)}!${CHK_X[0]}${CHK_X[1:]}", f"{q(SH_CHK)}!${CHK_W[0]}${CHK_W[1:]}"
    rem = [
        (f'=IF(N({X})>0,"✗ ","")&"数据校验：要改 "&N({X})&" 项、要看 "&N({W})&" 项"&IF(N({X})+N({W})=0,"，都没问题 √","　→ 点这里去看")', SH_CHK),
        (f'=IF(N({cr("nojz")})>0,"✗ 资金日记账有 "&{cr("nojz")}&" 笔有金额、没进账（校验是 ✗）：账户余额里有、报表里没有，照「校验」列的提示改",'
         f'"资金日记账每一笔都进账了 √")', SH_CASH),
        (f'=IF(N({cr("late")})>0,"有 "&{cr("late")}&" 行单子收单比送货晚 "&{LAGD}&" 天以上：录进去以后那几个月的成本已经跟着变了，看【跨月补单】","")',
         SH_LATE),
        (f'=IF(ABS(N({cr("wip")}))>=0.01,"车间里还挂着 "&TEXT({cr("wip")},"#,##0.00")&" 元成本（"&{HX}&" 月末，没交货的款的直接成本＋没分完的公共成本），'
         f'交了货自动转到营业成本 → 看【成本分摊表】","")', SH_ALLOC),
        (f'=IF(N({cr("apneg")})>0,"⚠ "&{cr("apneg")}&" 家供应商 / 加工厂付的钱比登记的单子多：可能还有送货单没拿到 → 看【应付账款汇总】","")', SH_APS),
        (f'=IF(N({cr("snew")})>0,"⚠ 订单里有 "&{cr("snew")}&" 个款式没在【款式档案】登记：没有单价、分不到成本 → 去登记","")', SH_STY),
    ]
    rem0 = r
    for f, sh in rem:
        # 文字是公式、可能是空的：链接不能放在同一格（LibreOffice 会把空结果存成链接地址），单独放在 K 列
        c = _note(ws, r, f, 'B', 'J', font=F_TXT)
        c.alignment = AL
        _lnk(ws, f'K{r}', f'去看 →', sh, 'A1', F_LINKS, border=False)
        ws.row_dimensions[r].height = 20
        r += 1
    rg = f'B{rem0}:J{r - 1}'
    ws.conditional_formatting.add(rg, FormulaRule(formula=[f'LEFT($B{rem0},1)="✗"'], font=F_RED))
    ws.conditional_formatting.add(rg, FormulaRule(formula=[f'LEFT($B{rem0},1)="⚠"'], fill=FILL_YEL))
    ws.conditional_formatting.add(rg, FormulaRule(formula=[f'RIGHT($B{rem0},1)="√"'], font=F_GOOD))
    r += 1

    # ── 每天录什么（流程图）
    section(ws, r, 'B', 'K', '每天录什么：蓝色 5 张录进去，下面的全是自动出来的（点格子跳过去）', C_HOME)
    r += 1
    for (c1, c2), (sh, _d) in zip(PAIRS, FLOW_IN):
        _box(ws, r, c1, c2, sh, C_IN, sh, height=28)
    r += 1
    for (c1, c2), (_sh, d) in zip(PAIRS, FLOW_IN):
        _box(ws, r, c1, c2, d, 'FFDDEBF7', None, Font(name=YH, sz=9, color='FF1F3864'), height=40)
    r += 1
    for c1, c2 in PAIRS:
        _box(ws, r, c1, c2, '↓', 'FFFFFFFF', None, F_ARROW, height=20)
        for cc in (c1, c2):
            ws[f'{cc}{r}'].border = NOBD
    r += 1
    _box(ws, r, 'B', 'K', '自动：记账分录（每一笔一借一贷）→ 成本分摊表（每月把成本分到款式：直接记到款式的等交货转出，公共的按交货双数 × 系数分）'
                         '→ 科目余额表', 'FF7F7F7F', SH_ALLOC, Font(name=YH, sz=10, bold=True, color='FFFFFFFF'), height=34)
    r += 1
    for c1, c2 in PAIRS:
        _box(ws, r, c1, c2, '↓', 'FFFFFFFF', None, F_ARROW, height=20)
        for cc in (c1, c2):
            ws[f'{cc}{r}'].border = NOBD
    r += 1
    for (c1, c2), (t, sh, color, _d) in zip(PAIRS, FLOW_OUT):
        _box(ws, r, c1, c2, t, color, sh, height=28)
    r += 1
    for (c1, c2), (_t, _sh, _c, d) in zip(PAIRS, FLOW_OUT):
        _box(ws, r, c1, c2, d, 'FFE2EFDA', None, Font(name=YH, sz=9, color='FF375623'), height=32)
    r += 2

    # ── 各张表（导航）
    section(ws, r, 'B', 'K', '各张表（点表名跳过去；每张表右上角有「← 回首页」）', C_HOME)
    r += 1
    for grp, color, sheets in NAV:
        ws.merge_cells(f'B{r}:K{r}')
        put(ws, f'B{r}', grp, F_SEC, fill(color), align=AL)
        r += 1
        for j, (sh, desc) in enumerate(sheets):
            cn, cd1, cd2 = (('B', 'C', 'F'), ('G', 'H', 'K'))[j % 2]
            rr = r + j // 2
            _lnk(ws, f'{cn}{rr}', sh, sh, 'A1', F_LINK, align=AL)
            ws.merge_cells(f'{cd1}{rr}:{cd2}{rr}')
            put(ws, f'{cd1}{rr}', desc, F_NOTE, align=AL)
            for i in range(CI(cd1) + 1, CI(cd2) + 1):
                ws.cell(row=rr, column=i).border = BD
            ws.row_dimensions[rr].height = 20
        r += (len(sheets) + 1) // 2
    r += 1

    # ── 怎么用
    section(ws, r, 'B', 'K', '怎么用（第一次看这里）', C_HOME)
    steps = ['① 第一次用：【基础资料】改公司名称、会计年度、建账日期，填资金账户的期初余额；【往来单位】登记供应商、加工厂、客户（和期初欠款）；'
             '【款式档案】登记款式和结算单价；【会计科目表】填其余年初余额，第 3 行要显示「√ 平」。',
             '② 每天：收付款记【资金日记账】（类别可以不选，选了供应商 / 客户 / 费用项目会自动认）；供应商拿来送货单就记【送货单登记】'
             '（送货日期照单子写，收单日期写拿到那天）；外发的加工单记【外发加工登记】。',
             '③ 电商部的采购单整张粘进【订单明细】，每行补上订单号；交货了填交货日期（分批交就复制一行）。每月月底录【工资登记】。',
             '④ 某个款专用的材料、外发、计件工资、模具，「用在哪」填款式编码，成本直接记到这个款；不填的按交货双数 × 系数分到各款（系数在【款式档案】调）。',
             '⑤ 看结果：【款式成本利润】每个款赚多少；【订单汇总】【订单查询】每张单赚多少；【利润表】【资产负债表】给老板看；'
             '跟供应商、电商部对账用【供应商对账单】【客户对账单】。',
             '⑥ 每月对一次：【基础资料】② 填银行 App 的实际余额；看【数据校验】有没有 ✗。供应商单子晚到会改动以前月份的成本——'
             '不想再改了，就在【基础资料】填「已结账到几月」。',
             '⑦ 注意：录错了就改或清空那几格，不要插行、删行（右边有隐藏公式）；灰色格子是公式，不要往里打字。新的一年另存一本，改会计年度和期初。']
    for t in steps:
        r += 1
        _note(ws, r, t, height=34, font=F_TXT)

    hide(ws, 'Z')
    ws.sheet_view.showGridLines = False
    print_setup(ws, None, landscape=False)
    return ws


def build(wb, ctx):
    build_chk(wb, ctx)          # 先建数据校验（首页提醒要用它的行号）
    build_home(wb, ctx)
