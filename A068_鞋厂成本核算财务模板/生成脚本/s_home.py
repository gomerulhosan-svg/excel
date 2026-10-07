# -*- coding: utf-8 -*-
"""【首页】选月份（空着＝最新月份）、关键数（本月 / 本年累计）、资金、往来、提醒、每天录什么、各表入口、
   每月月底做什么、新的一年怎么开、开始记自己的账（清演示数据）；
   【数据校验】哪里录错了、漏了、账平不平、快满了（第二阶段）。

   取数规矩：别的表的格子只用 layout.py 里定死的（各录入表的校验列 / ★接口列、BS_CHECK、BS_MX、会计科目表 C3、
   成本分摊表 alr()/al()、_款式月 sm_col()、科目×月矩阵 mat()……）；关键数自己算，不依赖查看表的格子。
   科目发生额一律用 mat()（【科目余额表】隐藏的科目×月矩阵），不再用 je_sum 扫记账分录。
   数据校验的汇总格：C3＝✗ 几项、E3＝⚠ 几项（首页引用）。"""
import re
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Protection
from common import *
from layout import *          # 注意：layout 的 C_IN / C_CHK / C_HOME 等颜色覆盖 common 的同名常量

try:                          # 【会计科目表】「检查」列（s_base 定的；自己加的科目挂得对不对）
    from s_base import COA_CHK as _COA_CHK
except Exception:             # noqa
    _COA_CHK = 'K'

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
F_TILE = Font(name=YH, sz=10, bold=True, color='FF1F3864', underline='single')
FILL_YEL = fill('FFFFEB9C')
FILL_GREY = fill('FFF2F2F2')
QTY = '#,##0;[Red]-#,##0;"-"'

CHK_X, CHK_W = 'C3', 'E3'            # 数据校验：✗ 几项、⚠ 几项（首页引用）
CHK_HDR, CHK_R0 = 4, 5
OS_N = 600                           # 【订单汇总】最多列几张订单（开发说明 s_cost 一节）
HROW = 4                             # HOME_M 所在行（C4）
HX = '$Z$4'                          # 首页隐藏帮手：实际用的月份（C4 空着＝最新月份）
HXR = f"{q(SH_HOME)}!{HX}"           # 别的表（数据校验）引用首页的实际月份
UN_NOWC = 'M'                        # 【往来单位】M 列「现在余额」（s_base：(应付＋其他应付)−(应收＋其他应收)，客户取反；全部已记的）
SUPS = ('材料供应商', '外发加工厂')
# 首页、数据校验共用的「看哪个月」：C4 是 1～12 就用它；空着（或乱填）＝最新月份（【基础资料】LASTM；还没业务＝建账月）
_LAST = (f'IF(N({P["LASTM"]})>=1,MIN(12,INT(N({P["LASTM"]}))),'
         f'IF(AND(ISNUMBER({OPEN}),YEAR({OPEN})={YR}),MONTH({OPEN}),1))')
HM = f'IF(AND(ISNUMBER({HOME_M}),{HOME_M}>=1,{HOME_M}<=12),INT({HOME_M}),{_LAST})'
NW0 = 'H'                            # 数据校验「有交货、没录工资」那一行：H～S＝1～12 月是不是（1/0），T＝月份清单（隐藏）
NW_LIST = 'T'


# ─────────────────────────── 小工具 ───────────────────────────
_SUM_ROWS = [(c, COA_R0 + i, d) for i, (c, _n, _cl, d, _u, _l, src) in enumerate(COA) if src == 'sum']


def _opencell(r):
    return f"{q(SH_COA)}!${COA_OPEN}${r}"


def _open(code):
    """【会计科目表】年初余额（正数＝正常方向）：编码前缀下每个科目「自己的」年初之和（预置的下级合计行 1002/1012/4001/5401 不重复算）"""
    sub = ''.join(f'-{_opencell(r)}' for c, r, _d in _SUM_ROWS if c.startswith(code))
    return f'(SUMIFS({COA_OPENS},{COA_CODES},"{code}*"){sub})'


def _net(code, m=None, cum=False, credit=False):
    """某科目（编码前缀，含下级）发生额，从【科目余额表】科目×月矩阵取：credit=True 贷−借，否则 借−贷；
       m＝月份（数字或格子）、cum＝1～m 月累计；m=None＝全部已记的（1～12 月累计）"""
    if m is None:
        m, cum = 12, True
    p = f'"{code}*"'
    d, c = mat(p, 'D', m, cum), mat(p, 'C', m, cum)
    return f'({c}-{d})' if credit else f'({d}-{c})'


def _bal(code, m=None, credit=False):
    """科目余额（正常方向为正）＝年初＋到 m 月末的累计发生（m=None＝全部已记的）"""
    return f'ROUND({_open(code)}+{_net(code, m, True, credit)},2)'


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


def _unsum(sign):
    """【往来单位】供应商、加工厂「现在余额」正数（sign='>'）或负数（'<'）部分的合计"""
    m = unr(UN_NOWC)
    return '+'.join(f'SUMIFS({m},{UN_TYPES_R},"{t}",{m},"{sign}0")' for t in SUPS)


# ═══════════════════════════ 数据校验 ═══════════════════════════
INPUTS = [  # (表, 校验列范围, 校验列表头格, ✗ 怎么改, ⚠ 是什么)
    (SH_CASH, jr(J_CHK), f'{J_CHK}{J_HDR}',
     '在「校验」列筛选 ✗，照提示改（多是没选账户、往来单位没登记（名字要跟【往来单位】一字不差）、费用支出没选费用项目）；✗ 的行不进报表',
     '提醒照样记账：用在哪的款式没登记或已停产、单位类型不对、账户余额成负数……看一眼是不是真有问题'),
    (SH_DN, dnr(DN_CHK), f'{DN_CHK}{DN_HDR}',
     '多是供应商没在【往来单位】登记、日期没填成日期、没金额；✗ 的行不算材料成本、不算应付',
     '晚到的单算到哪个月、已结账的月份没填收单日期、金额跟 数量×单价 差得多、用在哪的款式没登记或已停产'),
    (SH_OUT, otr(OT_CHK), f'{OT_CHK}{OT_HDR}',
     '加工厂先在【往来单位】登记（类型选外发加工厂）；日期要填成日期；✗ 的行不算加工费',
     '没填收单日期、加工厂类型不对、用在哪的款式没登记或已停产'),
    (SH_ORD, odr(OD_CHK), f'{OD_CHK}{OD_HDR}',
     '没填订单号、款式编码、数量；交货日期要填成日期；✗ 的行不算收入',
     '款式没登记（没单价、分不到成本）、没有结算单价、实交比订单多、客户没登记（「⏳ 还没交货」「√ 退货」「√ 返修重交」不算）'),
    (SH_WAGE, wgr(WG_CHK), f'{WG_CHK}{WG_HDR}',
     '月份填 1～12、选部门；✗ 的行不算工资',
     '没填姓名、用在哪的款式没登记或已停产、管理销售部门填了用在哪'),
    (SH_MJ, mjr(MJ_CHK), f'{MJ_CHK}{MJ_HDR}',
     '科目编码要在【会计科目表】、用末级科目；400104、540101～540104 不能手工记',
     '往来科目没填单位、费用科目没填费用项目、动了资金科目、冲减了材料 / 外发成本（冲回要跟真送货单同一个成本月）'),
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
    """[(键, 检查项, 结果公式, 状态种类, 数字格式, 怎么改, 去哪改(表, 格)) 或 ('#', 分组标题)]
       公式、说明里的 {row} 换成这一项在数据校验里的行号"""
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
    # 期初平不平：每个科目只算「自己的」年初（预置的下级合计行 1002/1012/4001/5401 不重复算；跟【会计科目表】C3 同一口径）
    side = {k: f'SUMIFS({COA_OPENS},{COA_DIRS},"{k}")' + ''.join(f'-{_opencell(r)}' for _c, r, d in _SUM_ROWS if d == k)
            for k in ('借', '贷')}
    coa_d = f'ROUND({side["借"]}-({side["贷"]}),2)'
    coa_c3 = cell(SH_COA, 'C3')
    it.append(('coa', '期初余额平不平（【会计科目表】C3：借方减贷方）', coa_d,
               f'=IF(OR(LEFT({coa_c3}&"",1)="✗",ABS(N({{v}}))>=0.01),"✗","√")', MONEY,
               '【会计科目表】年初余额：资产要等于负债＋权益。先核对：资金账户期初（【基础资料】②）、⑦ 固定资产（演示设备删了没有）、'
               '【往来单位】期初、会计科目表淡黄格（示例数换了没有）；都对了，最后用「年初未分配利润」3104 倒挤（C3 写借方多就加、贷方多就减）',
               (SH_COA, 'C3')))
    kchk = rng(SH_COA, _COA_CHK, COA_R0, COA_R1)
    it.append(('coasub', '【会计科目表】自己加的科目挂错了（加在系统记账的科目下面、编码重复、上级编码不对）', f'COUNTIF({kchk},"✗*")', 'x',
               '0" 个"', '系统自动记账的科目（2202、5602……）下面不能加下级（系统分录照样记在原科目上）：上级编码空着，换个新编码；'
                         '费用要分细用【基础资料】④ 费用项目。照【会计科目表】K 列的提示改', (SH_COA, f'{_COA_CHK}{COA_HDR}')))
    how_bs = (f'=IF(AND(ABS(N(C{{row_coa}}))>=0.01,ABS(N({BS_CHECK})-N(C{{row_coa}}))<0.01),'
              f'"期初余额就不平（上一项），资产负债表跟着差同一个数：先把【会计科目表】年初余额改平，这一项会自己变 √",'
              f'"【资产负债表】H3 不是 0：先看上一项期初平不平；期初平了还不平，说明有公式被改坏了（或者自己加的科目没归类），请联系做表的人")')
    it.append(('bs', f'="资产负债表平不平（"&{BS_MX}&" 月末，资产−负债和权益）"', f'N({BS_CHECK})', 'x0', MONEY, how_bs, (SH_BS, 'H3')))
    now_r = rng(SH_BASE, AC_NOW, AC_R0, AC_R1)
    terms = []
    for k in range(1, AC_R1 - AC_R0 + 2):
        code = f'INDEX({AC_CODES},{k})&""'
        jb = f'N(INDEX({AC_OPENS},{k}))+{mat(code, "D", 12, True)}-{mat(code, "C", 12, True)}'
        terms.append(f'IF(OR(TRIM(INDEX({AC_NAMES},{k})&"")="",{code}=""),0,'
                     f'IF(ABS(ROUND(N(INDEX({now_r},{k}))-({jb}),2))>=0.01,1,0))')
    it.append(('acc', '资金账户：日记账算的余额跟资金科目余额对不上的账户', '+'.join(terms), 'w', '0" 个"',
               '多半是【资金日记账】有 ✗ 的行（有金额没进账），或者【手工分录】动了资金科目，或者有两个「现金」账户（都记在 1001）；'
               '改好就对上了。逐个账户看【账户余额表】', (SH_ACCB, 'A1')))
    diff = rng(SH_BASE, AC_DIFF, AC_R0, AC_R1)
    it.append(('real', '资金账户：跟银行 App / 支付宝实际余额有差额的账户', f'COUNTIF({diff},">=0.01")+COUNTIF({diff},"<=-0.01")',
               'w', '0" 个"', '【基础资料】② 填的实际余额跟表里算的不一样：漏记、记错金额或者记错账户', (SH_BASE, f'{AC_REAL}{AC_R0}')))
    it.append(('rev', '主营业务收入 跟【订单明细】交货金额 差多少',
               f'ROUND({_net("5001", None, credit=True)}-SUMIFS({odr(OD_AMT)},{odr(OD_OK)},1),2)', 'w0', MONEY,
               '收入按【订单明细】交货金额自动记；不一样多半是【手工分录】记了 5001（调账），确认一下是不是故意的', (SH_IS, 'A1')))
    moh = '+'.join(f'IF(ABS({_bal(MOH, m)})>=0.01,1,0)' for m in range(1, 13))
    it.append(('moh', '制造费用（4101）月末余额不是 0 的月份', moh, 'x', '0" 个月"',
               '制造费用每月末自动转到生产成本，余额应该是 0；不是 0 多半是【会计科目表】4101 填了年初余额，或者公式被改坏了',
               (SH_TB, 'A1')))

    it.append(('#', '三、成本和订单'))
    al_bad = '+'.join(f'IF(ABS(ROUND(SUM({smr("成本", m)})-N({alr("转出合计", m)}),2))>=0.01,1,0)' for m in range(1, 13))
    it.append(('alloc', '【成本分摊表】「核对」不平的月份（各款式成本合计 − 本月转出合计）', al_bad, 'x', '0" 个月"',
               '各款式分到的成本加起来应该正好等于每月转出合计；不平说明公式被改坏了，请联系做表的人', (SH_ALLOC, f'A{AL_CHECK_ROW}')))
    # #6 负的成本分到款式：成本分摊表任一月任一组件「公共本月分摊」<0，或 _款式月 任一「本X」<0
    b0, b1 = SM_BLOCKS.index('本材料'), SM_BLOCKS.index('本制造')
    assert b1 - b0 == 3 and SM_BLOCKS[b0:b1 + 1] == [f'本{c}' for c in COMPS]
    neg_m = '+'.join('IF(OR(' + ','.join(f'N({al(c, "公共本月分摊", m)})<=-0.01' for c in COMPS) + '),1,0)' for m in range(1, 13))
    own = f"{q(SH_SM)}!${sm_col('本材料', 1)}${SM_R0}:${sm_col('本制造', 12)}${SM_R1}"
    it.append(('negal', '分到款式的成本是负数（公共分摊为负的月份＋款式·月·组件）', f'{neg_m}+COUNTIF({own},"<=-0.01")', 'w', '0" 处"',
               '多半是【手工分录】冲回暂估记到了下个月（没结账时晚到的送货单按送货日期算，不用暂估），或者那个月退货比进货多；'
               '把冲回改到真送货单算进去的那个月，或把暂估、冲回两行都清空', (SH_ALLOC, f'A{AL_CROWS["材料"]["公共本月分摊"]}')))
    # #39 订单成本合计 vs 款式成本合计：订单每行成本四舍五入到分（每行最多差半分钱）＋单双 4 位小数（每双最多差 0.00005）
    sm_cost = f"{q(SH_SM)}!${sm_col('成本', 1)}${SM_R0}:${sm_col('成本', 12)}${SM_R1}"
    sm_q = f"{q(SH_SM)}!${sm_col('双数', 1)}${SM_R0}:${sm_col('双数', 12)}${SM_R1}"
    tol = f'MAX(1,0.005*COUNT({odr(OD_UCV)})+0.00005*SUM({sm_q}))'
    it.append(('ocost', f'="订单成本合计 跟 各款式成本合计 差多少（四舍五入最多差 "&TEXT({tol},"0.00")&" 元）"',
               f'ROUND(SUM({odr(OD_COSTV)})-SUM({sm_cost}),2)',
               f'=IF(ABS(N({{v}}))<0.01,"√",IF(ABS(N({{v}}))<{tol},"○","✗"))', MONEY,
               '订单每行成本＝实交双数×那个款那个月的单双成本，每行四舍五入到分，行多了加起来差几块钱是正常的（○ 不用管）；'
               '✗＝差得超过四舍五入可能的范围，说明【订单明细】成本列或公式被改坏了，请联系做表的人', (SH_OSUM, 'A1')))
    it.append(('wip', f'="生产成本余额（在制，到 "&{HXR}&" 月末；月份跟首页走）"', _bal('4001', HXR), 'i+', MONEY,
               '＝还挂在车间、没交货的款的直接成本（等交货按进度转出）＋没交货月份结转下来的公共成本，有数是正常的；负数要查',
               (SH_ALLOC, f'A{AL_ROWS["月末在制"]}')))
    S, J = str_(ST_WIP), str_(ST_STAT)
    it.append(('stwip', '款式还挂着没转出的直接成本（【款式档案】S 列，状态不是「停产」的款）', f'ROUND(SUMIFS({S},{J},"<>停产"),2)', 'i+',
               MONEY, '直接记到款式的成本按交货进度转出，还有订单没交完的挂着是正常的；订单取消了的把【订单明细】订单数量改成实交数；'
                      '不会再做的款把【款式档案】状态改「停产」（以后它的专用成本按公共分）', (SH_STY, f'{ST_WIP}{ST_HDR}')))
    it.append(('stwip2', '停产的款还挂着没转出的直接成本（以后不交货就转不出去）', f'ROUND(SUMIFS({S},{J},"停产"),2)', 'w0', MONEY,
               '停产的款最后一次交货时还有没交完的订单数量，成本按比例留下了：把【订单明细】这个款没交完的订单数量改成实交数（订单取消了），'
               '就在最后交货那个月全部转出', (SH_STY, f'{ST_WIP}{ST_HDR}')))
    # #19 有交货、工资登记那个月一笔都没有：H～S 是每个月 1/0（隐藏），T 是月份清单（首页提醒用）
    it.append(('nowage', '有交货、【工资登记】那个月一笔工资都没有的月份', f'SUM({NW0}{{row}}:{CL(CI(NW0) + 11)}{{row}})', 'w', '0" 个月"',
               f'=IF(${NW_LIST}${{row}}="","有交货的月份都录了工资",'
               f'"是 "&${NW_LIST}${{row}}&" 月。工资月底才录的话，这几个月交货的款、订单成本里没有工资（房租水电没记也一样）：'
               f'单双成本偏低、毛利虚高，录完工资再看")', (SH_WAGE, 'A1')))
    it.append(('snew', '订单里有、【款式档案】没登记的款式', f'COUNT({odr(OD_SNEW)})', 'w', '0" 个"',
               '【款式档案】右边列出来了，抄过去登记（填结算单价）；不登记没有单价、分不到成本', (SH_STY, f'{ST_NCODE}{ST_HDR}')))
    it.append(('p0', '交了货、结算单价是 0 的行', f'COUNTIFS({odr(OD_DQ)},"<>0",{odr(OD_PRICEU)},0)', 'w', '0" 行"',
               '在【款式档案】填结算单价，或者在【订单明细】这一行填「结算单价」；单价 0 的交货没有收入', (SH_ORD, f'{OD_PRICE}{OD_HDR}')))
    it.append(('nnew', '送货单里有、【品名档案】没登记的品名', f'COUNT({dnr(DN_NNEW)})', 'w', '0" 个"',
               '【品名档案】右边列出来了，抄过去选个材料类别；不登记算「未分类」，不影响成本', (SH_MAT, f'{MT_NNAME}{MT_HDR}')))
    it.append(('late', f'="晚到的单（收单比送货晚 "&{LAGD}&" 天以上）：送货单＋外发"',
               f'COUNT({dnr(DN_LATEK)})+COUNT({otr(OT_LATEK)})', 'i', '0" 行"',
               f'供应商单子拿来得晚，录进去以后那个月的成本会跟着变（那个月在【基础资料】⑩ 填了结账日期的除外：算到收单月）；'
               f'看【{SH_LATE}】，月底催供应商把单子送来', (SH_LATE, 'A1')))
    cb = []
    for src, sm_, rd, ok, dt in ((dnr, DN_SM, DN_RDATE, DN_OK, DN_DATE), (otr, OT_SM, OT_RDATE, OT_OK, OT_DATE)):
        cb += [f'IF(ISNUMBER(INDEX({CLOSE_DATES},{m})),COUNTIFS({src(sm_)},{m},{src(rd)},"",{src(ok)},1,{src(dt)},">="&{OPEN}),0)'
               for m in range(1, 13)]
    it.append(('clblank', '已结账的月份（【基础资料】⑩ 填了结账日期）里、收单日期空着的送货单 / 外发单', '+'.join(cb), 'w', '0" 行"',
               '结账以后才录的单一定要填收单日期（填了就算到收单月，已结账的数不变；空着按送货日期算回已结账的月份）；'
               '结账以前就录好的，把收单日期补上也不会改数', (SH_DN, f'{DN_RDATE}{DN_HDR}')))

    it.append(('#', '四、往来和工资（全部已记的）'))
    m_ = unr(UN_NOWC)
    apk = '+'.join(f'COUNTIFS({m_},"<-0.005",{UN_TYPES_R},"{t}")' for t in SUPS)
    it.append(('apneg', '付多了的供应商、加工厂（【往来单位】现在余额是负数的）', apk, 'w', '0" 家"',
               '多半还有送货单没拿到（或者是预付款）：问供应商要单子；首页「预付给供应商」是金额，【应付账款汇总】看是哪几家', (SH_APS, 'A1')))
    nocust = (f'ROUND(SUMIFS({odr(OD_AMT)},{odr(OD_OK)},1)'
              f'-SUMPRODUCT(({AX_CUST}<>"")*SUMIFS({odr(OD_AMT)},{odr(OD_OK)},1,{odr(OD_CUSTK)},{AX_CUST})),2)')
    it.append(('nocust', '交货收入里客户没登记成「客户」的金额（记在「未登记客户」名下）', nocust, 'w0', MONEY,
               '【订单明细】的客户（空＝默认客户）要在【往来单位】登记、类型选「客户」，不然应收按客户看不准（首页往来也看不到）',
               (SH_UNIT, f'{UN_TYPE}{UN_HDR}')))
    it.append(('ncust', f'客户家数（记账分录按客户分收入，最多 {JE_NCUST} 家）', f'COUNT({unr(UN_ARK)})',
              f'=IF(N({{v}})>{JE_NCUST},"⚠","√")', '0" 家"',
               f'超过 {JE_NCUST} 家时，第 {JE_NCUST + 1} 家以后的收入记在「未登记客户」名下；不再往来的客户把类型改成「其他」',
               (SH_UNIT, f'{UN_TYPE}{UN_HDR}')))
    it.append(('wage', '未发工资（应付职工薪酬 2211 余额，全部已记的）', _bal('2211', None, True), 'i+', MONEY,
               '＝【工资登记】算了、还没在【资金日记账】记「付工资」的；一般是最近一个月的。负数＝发的比应发多', (SH_WSUM, 'A1')))

    it.append(('#', '五、档案和设置'))
    it.append(('open', '建账日期（【基础资料】①）', f'IF(ISNUMBER({OPEN}),{OPEN},"")',
               f'=IF(NOT(ISNUMBER({OPEN})),"✗",IF(YEAR({OPEN})<>{YR},"✗",IF(DAY({OPEN})<>1,"⚠","√")))', DATE,
               '要填成日期、在会计年度里，最好是某个月的 1 号（报表按月算）。新的一年另存一本时，建账日期改成新一年的 1 月 1 日'
               '（不改，固定资产期初、折旧会算错），期初换成上年 12 月末的数（看首页「新的一年怎么开」）',
               (SH_BASE, f'{PA_VAL}{PA_ROW["OPEN"]}')))
    miss = '+'.join(f'IF(COUNTIF({CT_NAMES},"{n}")=0,1,0)' for n, *_r in CATS)
    it.append(('catname', f'【基础资料】③ 预置的收支类别名字被改了（或删了）（预置 {len(CATS)} 个）', miss, 'x', '0" 个"',
               '预置的类别（收货款、付材料款……内部转账）公式是按名字认的，一改账户余额、自动认类别就错：改回原来的名字；'
               '要别的叫法，在下面空行另加一行', (SH_BASE, f'{CT_NAME}{CT_R0}')))
    it.append(('cash2', '「现金」类型的资金账户个数', f'COUNTIFS({AC_TYPES_R},"现金",{AC_NAMES},"<>")',
               '=IF(N({v})>1,"⚠","√")', '0" 个"',
               '现金账户都记在 1001 库存现金上，有两个就分不开，账户余额表、上面「资金账户对不上」会一直报差：两个钱箱合成一个账户记；'
               '一定要分开，就把其中一个类型选「支付宝微信」（记到其他货币资金，余额能分开对）', (SH_BASE, f'{AC_TYPE}{AC_R0}')))
    demo_cnt = '+'.join(f'COUNTIF({r},"*示例*")' for r in (
        rng(SH_BASE, FA_NOTE, FA_R0, FA_R1), rng(SH_BASE, AC_NO, AC_R0, AC_R1), unr(UN_NOTE), str_(ST_NOTE),
        rng(SH_MAT, MT_NOTE, MT_R0, MT_R1), rng(SH_COA, COA_NOTE, COA_R0, COA_R1)))
    it.append(('demo', '档案里还留着演示数据（备注、说明里写着「示例」的行）', demo_cnt, 'w', '0" 行"',
               '开始记自己的账要清掉（看首页「开始记自己的账」）：尤其【基础资料】⑦ 的演示设备，不删会每月提折旧进款式成本；'
               '自己的数据就把备注里的「示例」两个字删掉', (SH_BASE, f'{FA_NAME}{FA_R0}')))
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
          f'-{mat(chr(34) + "1601*" + chr(34), "D", 12, True)},2)')
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


def _tlen(t):
    """显示出来大约多少字：公式就按里面最长的那段文字算"""
    t = str(t or '')
    if t.startswith('='):
        parts = re.findall(r'"([^"]*)"', t)
        return (max(len(x) for x in parts) + 6) if parts else 10
    return len(t)


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
    if kind == 'i':
        return f'=IF(N({v})=0,"√","○")'
    if kind == 'i+':
        return f'=IF(N({v})<=-0.01,"⚠",IF(ABS(N({v}))<0.01,"√","○"))'
    if kind.startswith('cap'):
        return f'=IF(N({v})>{int(kind[3:])}*0.9,"⚠","√")'
    raise ValueError(kind)


def _nowage_helpers(ws, r):
    """「有交货、工资还没录」：H～S＝1～12 月（这个月有成本双数、工资登记里这个月一行能记账的都没有 → 1），T＝月份清单「8、10」"""
    cells = []
    for m in range(1, 13):
        c = f'{CL(CI(NW0) + m - 1)}{r}'
        ws[c] = f'=IF(AND(SUM({smr("双数", m)})>0,COUNTIFS({wgr(WG_CM)},{m},{wgr(WG_OK)},1)=0),1,0)'
        ws[c].font = F_HELP
        cells.append(c)
    ws[f'{NW_LIST}{r}'] = '=MID(' + '&'.join(f'IF({c}=1,"、{m}","")' for m, c in enumerate(cells, 1)) + ',2,99)'
    ws[f'{NW_LIST}{r}'].font = F_HELP


def build_chk(wb, ctx):
    ws = wb[SH_CHK]
    widths(ws, {'A': 6, 'B': 46, 'C': 15, 'D': 9, 'E': 66, 'F': 18})
    title(ws, '数 据 校 验（录错的、漏录的、账平不平、快满了的，都在这里）', 'F', C_CHK,
          '💡 全自动，打开就是最新的。状态：✗＝一定要改（不改报表就不准）；⚠＝看一下是不是真有问题；√＝没问题；○＝只是告诉你一个数（正常情况也会有）。'
          '点最右边「去哪改」跳到那张表，在那张表的「校验」列筛选 ✗ / ⚠ 就能找到是哪几行。首页上的 ✗ / ⚠ 项数就是 C3、E3 这两格。')
    items = _items()
    # 先排行号（说明里会引用别的项的行号，比如 bs 引用 coa）
    r, rows = CHK_R0, {}
    for item in items:
        if item[0] != '#':
            rows[item[0]] = r
        r += 1
    r = CHK_R0
    seq = 0
    for item in items:
        if item[0] == '#':
            section(ws, r, 'A', 'F', item[1], 'FF7F7F7F')
            ws.row_dimensions[r].height = 20
            r += 1
            continue
        key, lab, f, kind, fmt, how, (sh, ref) = item
        sub = lambda s: s.replace('{row}', str(r)).replace('{row_coa}', str(rows.get('coa', 0))) if isinstance(s, str) else s
        f, how, lab = sub(f), sub(how), sub(lab)
        seq += 1
        CHK_ROW[key] = r
        put(ws, f'A{r}', seq, F_AUTO, align=AC)
        put(ws, f'B{r}', lab, F_TXT, align=ALW)
        put(ws, f'C{r}', f'={f}', F_AUTOB, FILL_AUTO, fmt, AR)
        put(ws, f'D{r}', _status(kind, f'C{r}'), F_TXTB, align=AC)
        put(ws, f'E{r}', how, F_NOTE, align=ALW)
        _lnk(ws, f'F{r}', f'去【{sh}】', sh, ref, F_LINKS)
        lines = max(-(-_tlen(how) // 37), -(-_tlen(lab) // 23), 1)     # E 列 66 宽 9 号字一行约 37 字；B 列 46 宽 10 号字约 23 字
        ws.row_dimensions[r].height = max(20, 12.5 * lines + 6)
        if key == 'nowage':
            _nowage_helpers(ws, r)
        r += 1
    r1 = r - 1
    hide(ws, *[CL(CI(NW0) + i) for i in range(13)])
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
    put(ws, f'B{r1 + 2}', '说明：「结果」是个数的，0 就是没问题；是金额的，0 就是对得上。资产负债表的月份跟【资产负债表】C3 走（空着＝最新月份）；'
                          '生产成本余额的月份跟【首页】C4 走（空着＝最新月份）；其余都是全部已记的数。', F_NOTE, align=ALW, border=False)
    ws.merge_cells(f'B{r1 + 2}:F{r1 + 3}')
    ws.freeze_panes = f'B{CHK_R0}'
    print_setup(ws, f'{CHK_HDR}:{CHK_HDR}', landscape=True)
    ws.print_area = f'A1:F{r1 + 3}'
    return ws


# ═══════════════════════════ 首页 ═══════════════════════════
NAV = [
    ('录入（每天 / 每笔录 · 蓝色）', C_IN, [
        (SH_CASH, '银行1、银行2、现金、支付宝混在一张表，一笔一行，余额逐行算'),
        (SH_DN, '供应商送货单，一个品名一行；单子晚到也照送货日期记'),
        (SH_OUT, '外发加工厂的加工单（加工费）；能分到款式就填款式'),
        (SH_ORD, '电商部采购单整张粘进来；交货了填交货日期'),
        (SH_WAGE, '每人每月一行；计件的可以按款式分几行'),
        (SH_MJ, '调账、计提、冲销用，平时一般不用'),
    ]),
    ('成本、订单（自动 · 绿色）', C_VIEW, [
        (SH_SPL, '每个款式的双数、收入、成本（材料/外发/人工/制造）、毛利'),
        (SH_OSUM, '每张订单交了多少、还差多少、赚了多少'),
        (SH_OQ, '选一张订单号，看逐行交货、成本、毛利'),
        (SH_ALLOC, '每月成本怎么分到款式；只有「在制估计」一行可以填'),
        (SH_MSUM, '材料按类别、按品名看花了多少、单价涨没涨'),
        (SH_LATE, '供应商送货单、加工单晚到多久、晚到的是哪几张'),
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
        (SH_BASE, '参数（年度、建账日期）、每月结账日期、资金账户、收支类别、费用项目、部门、固定资产'),
        (SH_UNIT, '供应商、加工厂、客户、老板；期初欠款、现在余额'),
        (SH_STY, '款式编码、结算单价、分摊系数、停产；挂着没转的成本'),
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


def _h(text, per_line=64):
    """B:J 合并格（9 列 × 15 宽）放 10 号字：按字数估行高"""
    return max(20, 15 * (-(-len(text) // per_line)) + 6)


def _steps(ws, r, head, items, color=C_HOME):
    """一段带编号的说明：每条 B:J 文字，K 列「去 →」链接（没有链接的空着）。返回最后一行"""
    section(ws, r, 'B', 'K', head, color)
    for t, sh, ref in items:
        r += 1
        _note(ws, r, t, 'B', 'J', height=_h(t), font=F_TXT)
        if sh:
            _lnk(ws, f'K{r}', '去 →', sh, ref, F_LINKS, align=AC, border=False)
    return r


def build_home(wb, ctx):
    ws = wb[SH_HOME]
    assert HOME_M == cell(SH_HOME, 'C4')
    widths(ws, {'A': 2, **{CL(i): 15 for i in range(2, 12)}, 'L': 2})
    title(ws, '鞋 厂 成 本 核 算 财 务 账', 'K', C_HOME,
          '💡 平时只录蓝色那 5 张（资金日记账、送货单登记、外发加工登记、订单明细、工资登记），成本、利润、对账单、报表全是自动的。'
          'C4 黄格选月份（空着＝最新的月份），下面「本月」「本年累计」跟着变。点表名就能跳过去，每张表左上角有「←首页」回到这里。'
          '灰色、白色格子是公式，不要往里打字。')
    ws['A1'].value = f'={CO}&"　"&{YR}&" 年 · 成本核算财务账（按款式 × 月算成本利润）"'
    demo = bool(ctx.get(SH_CASH) or ctx.get(SH_ORD))

    # ── 第 4 行：看哪个月（HOME_M）：出厂空着＝最新月份（有业务的最后一个月）
    selector(ws, f'B{HROW}', '看哪个月', f'C{HROW}', None, f'={AX_M}', '0"月"',
             '选 1～12；空着＝最新的月份（有业务的最后一个月）')
    ws[HX.replace('$', '')] = f'={HM.replace(HOME_M, "$C$4")}'
    ws[HX.replace('$', '')].font = F_HELP
    ws.merge_cells(f'D{HROW}:K{HROW}')
    tail = f'"本月＝"&{HX}&" 月，本年累计＝1～"&{HX}&" 月；资金、往来是现在的余额（全部已记的）"'
    put(ws, f'D{HROW}', f'=IF($C$4="","空着＝最新月份 "&{HX}&" 月"&IF(N({P["LASTM"]})>=1,"（有业务的最后一个月）","（还没录业务，按建账月）")&"："&{tail},'
                        f'IF({HX}<>$C$4,"⚠ C4 要选 1～12 的月份（现在按 "&{HX}&" 月算；清空就看最新月份）",{tail}))',
        F_NOTE, align=AL, border=False)
    ws.row_dimensions[HROW].height = 24
    ws.conditional_formatting.add(f'D{HROW}', FormulaRule(formula=[f'LEFT($D${HROW},1)="⚠"'], font=F_RED))
    r = HROW + 1
    if demo:
        c = _note(ws, r, '现在表里是演示数据（备注写「示例」的都是编的；中茂 9/5 送货单、电商部采购单 353651 是照片上的真单子）。'
                         '开始记自己的账要先清掉，清哪些看下面「开始记自己的账」一段（在「怎么用」后面）。这一行看完可以选中按 Delete 删掉。',
                  height=34, font=F_DEMO)
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
    for lab, op, cum in [(f'="本月（"&{HX}&" 月）"', '=', False), (f'="本年累计（1～"&{HX}&" 月）"', '<=', True)]:
        r += 1
        put(ws, f'B{r}', lab, F_TXTB, KPI_FILL, align=AC)
        C, D, E, F, G, H, I, J, K = (f'{CL(3 + i)}{r}' for i in range(9))
        oth = f'Z{r}'          # 隐藏帮手：税金及附加＋营业外支出＋所得税－营业外收入
        n = lambda code, credit=False: _net(code, HX, cum, credit)
        ws[oth] = f'=ROUND({n("5403")}+{n("5711")}+{n("5801")}-{n("5301", True)},2)'
        ws[oth].font = F_HELP
        vals = {
            C: (f'=SUMPRODUCT(({dq_mon}{op}{HX})*{dq_blk})', QTY),
            D: (f'=ROUND({n("5001", True)}+{n("5051", True)},2)', MONEY0),
            E: (f'=ROUND({n("5401")}+{n("5402")},2)', MONEY0),
            F: (f'=ROUND({D}-{E},2)', MONEY0),
            G: (f'=IF(N({D})=0,"",ROUND({F}/{D},4))', PCT),
            H: (f'=ROUND({n("5601")}+{n("5602")}+{n("5603")},2)', MONEY0),
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
    c = _note(ws, r, '交货双数按【订单明细】交货日期算到月（不含退货、返修重交）；交货收入含其他业务收入；营业成本＝【成本分摊表】每月结转的'
                     '（只算交了货的鞋，没交货的挂在生产成本里）；期间费用＝销售＋管理＋财务费用；净利润跟【利润表】一样'
                     '（再减税金及附加、营业外收支、所得税）。月底工资、房租水电没录全之前，本月成本偏低。点这里看【利润表】 →',
              height=44, font=F_LINKS)
    link(c, SH_IS)
    r += 2

    # ── 资金：每个账户当前余额 ＋ 总余额
    section(ws, r, 'B', 'K', '资金（现在的余额 ＝ 期初 ＋ 资金日记账全部已记的）', C_HOME)
    r += 1
    now_r = rng(SH_BASE, AC_NOW, AC_R0, AC_R1)
    n_ac = AC_R1 - AC_R0 + 1
    _lnk(ws, f'B{r}', '总余额', SH_CASH, 'A1', F_TILE, fill_=KPI_FILL)
    put(ws, f'B{r + 1}', f'=SUM({now_r})', F_KPI_V, fill('FFFFFFFF'), MONEY, AC)
    for k in range(1, n_ac + 1):
        col = CL(2 + k)
        put(ws, f'{col}{r}', f'=IF(TRIM(INDEX({AC_NAMES},{k})&"")="","",INDEX({AC_NAMES},{k})&"")', F_KPI_L, KPI_FILL, align=ACW)
        put(ws, f'{col}{r + 1}', f'=IF({col}{r}="","",N(INDEX({now_r},{k})))', F_KV10, fill('FFFFFFFF'), MONEY, AC)
    _lnk(ws, f'K{r}', '跟银行 App\n对不上的', SH_BASE, f'{AC_REAL}{AC_R0}', F_TILE, fill_=KPI_FILL)
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

    # ── 往来（现在＝全部已记的，按【往来单位】M 列「现在余额」分正负）和要留意的
    section(ws, r, 'B', 'K', '往来（现在的余额＝全部已记的，不跟上面的月份走；按【往来单位】每家「现在余额」分开算）和要留意的', C_HOME)
    r += 1
    mc_ = unr(UN_NOWC)
    tiles = [
        (('B', 'B'), '欠供应商/加工厂\n（应付）', f'=ROUND({_unsum(">")},2)', MONEY, SH_APS),
        (('C', 'C'), '预付给供应商\n（付多了的）', f'=ROUND(-({_unsum("<")}),2)', MONEY, SH_APS),
        (('D', 'D'), '电商部/客户\n欠我们（应收）', f'=ROUND(SUMIFS({mc_},{UN_TYPES_R},"客户",{mc_},">0"),2)', MONEY, SH_CST),
        (('E', 'E'), '客户预付\n（预收）', f'=ROUND(-SUMIFS({mc_},{UN_TYPES_R},"客户",{mc_},"<0"),2)', MONEY, SH_CST),
        (('F', 'G'), '还没交货的订单双数',
         f'=MAX(0,SUM({odr(OD_OQ)})-SUM({odr(OD_CQ)}))', '#,##0" 双"', SH_OSUM),
        (('H', 'I'), f'="晚到的单据（晚 "&{LAGD}&" 天以上）"', f'=COUNT({dnr(DN_LATEK)})+COUNT({otr(OT_LATEK)})', '0" 行"', SH_LATE),
        (('J', 'J'), '数据校验\n✗ 要改', f"={q(SH_CHK)}!${CHK_X[0]}${CHK_X[1:]}", '0" 项"', SH_CHK),
        (('K', 'K'), '数据校验\n⚠ 要看', f"={q(SH_CHK)}!${CHK_W[0]}${CHK_W[1:]}", '0" 项"', SH_CHK),
    ]
    tile_r = r
    for (c1, c2), lab, f, fmt, sh in tiles:
        if c1 != c2:
            ws.merge_cells(f'{c1}{r}:{c2}{r}')
            ws.merge_cells(f'{c1}{r + 1}:{c2}{r + 1}')
        _lnk(ws, f'{c1}{r}', lab, sh, 'A1', F_TILE, fill_=KPI_FILL, align=ACW)
        put(ws, f'{c1}{r + 1}', f, F_KPI_V, fill('FFFFFFFF'), fmt, AC)
        for rr in (r, r + 1):
            ws[f'{c2}{rr}'].border = BD
    ws.row_dimensions[r].height = 32
    ws.row_dimensions[r + 1].height = 28
    for cc in ('C', 'E'):                     # 预付、预收：有数就淡黄（多半要找供应商要单子 / 跟客户对账）
        ws.conditional_formatting.add(f'{cc}{r + 1}', FormulaRule(formula=[f'N({cc}{r + 1})>=0.01'], fill=FILL_YEL))
    ws.conditional_formatting.add(f'J{r + 1}', FormulaRule(formula=[f'N(J{r + 1})>0'], fill=FILL_WARN, font=F_RED))
    ws.conditional_formatting.add(f'K{r + 1}', FormulaRule(formula=[f'N(K{r + 1})>0'], fill=FILL_YEL))
    r += 2
    _note(ws, r, '应付 / 预付：【往来单位】类型是材料供应商、外发加工厂的，每家「现在余额」正数加起来是欠他们的，负数加起来是付多了的（预付，或者还有送货单没拿到）；'
                 '应收 / 预收：类型是客户的，正数是他欠我们的，负数是他多付的。资产负债表按所选月末拆，数可能不一样。',
          height=30)
    r += 2

    # ── 提醒（引用数据校验的结果格）
    section(ws, r, 'B', 'K', '提醒', 'FFC00000')
    r += 1
    cr = lambda key: f"{q(SH_CHK)}!$C${CHK_ROW[key]}"
    X, W = f"{q(SH_CHK)}!${CHK_X[0]}${CHK_X[1:]}", f"{q(SH_CHK)}!${CHK_W[0]}${CHK_W[1:]}"
    nw_list = f"{q(SH_CHK)}!${NW_LIST}${CHK_ROW['nowage']}"
    st_n = f'COUNTIF({str_(ST_WIP)},">=0.005")'
    st_all = f'(N({cr("stwip")})+N({cr("stwip2")}))'
    rem = [
        (f'=IF(N({X})>0,"✗ ","")&"数据校验：要改 "&N({X})&" 项、要看 "&N({W})&" 项"&IF(N({X})+N({W})=0,"，都没问题 √","　→ 点这里去看")',
         SH_CHK, 'A1', 20),
        (f'=IF(N({cr("nojz")})>0,"✗ 资金日记账有 "&{cr("nojz")}&" 笔有金额、没进账（校验是 ✗）：账户余额里有、报表里没有，照「校验」列的提示改",'
         f'"资金日记账每一笔都进账了 √")', SH_CASH, 'A1', 20),
        (f'=IF(N({cr("nowage")})>0,"⚠ "&{nw_list}&" 月有交货、【工资登记】还没录这个月的工资：这个月交货的款、订单单双成本偏低、毛利虚高，录完再看","")',
         SH_WAGE, 'A1', 30),
        (f'=IF(N({cr("late")})>0,"有 "&{cr("late")}&" 行单子收单比送货晚 "&{LAGD}&" 天以上：录进去以后那几个月的成本跟着变了'
         f'（结了账的月份除外），看【{SH_LATE}】","")', SH_LATE, 'A1', 30),
        (f'=IF(ABS(N({cr("wip")}))>=0.01,"车间里还挂着 "&TEXT({cr("wip")},"#,##0.00")&" 元成本（"&{HX}&" 月末，没交货的款的直接成本＋没分完的公共成本），'
         f'交了货自动转到营业成本 → 看【成本分摊表】","")', SH_ALLOC, 'A1', 30),
        (f'=IF(ABS({st_all})<0.01,"",IF(ABS(N({cr("stwip2")}))>=0.01,"⚠ ","")&"款式在制：直接记到款式、还挂着没转出的 "&TEXT({st_all},"#,##0.00")'
         f'&" 元（"&{st_n}&" 个款；订单没交完的正常，交货时按进度转）"&IF(ABS(N({cr("stwip2")}))>=0.01,"；其中停产的款 "'
         f'&TEXT({cr("stwip2")},"#,##0.00")&" 元转不出去：订单取消了的把订单数量改成实交数","")&" → 看【款式档案】S 列")',
         SH_STY, f'{ST_WIP}{ST_HDR}', 30),
        (f'=IF(N({cr("apneg")})>0,"⚠ "&{cr("apneg")}&" 家供应商 / 加工厂付多了，共 "&TEXT(C{tile_r + 1},"#,##0.00")&" 元：'
         f'预付款，或者还有送货单没拿到 → 看【应付账款汇总】","")', SH_APS, 'A1', 20),
        (f'=IF(N({cr("clblank")})>0,"⚠ 已结账的月份里有 "&{cr("clblank")}&" 行送货单 / 外发单没填收单日期：结账以后才录的要填收单日期，'
         f'不然会改动已结账的数","")', SH_DN, f'{DN_RDATE}{DN_HDR}', 30),
        (f'=IF(N({cr("snew")})>0,"⚠ 订单里有 "&{cr("snew")}&" 个款式没在【款式档案】登记：没有单价、分不到成本 → 去登记","")',
         SH_STY, f'{ST_NCODE}{ST_HDR}', 20),
    ]
    rem0 = r
    for f, sh, ref, h in rem:
        # 文字是公式、可能是空的：链接不能放在同一格（LibreOffice 会把空结果存成链接地址），单独放在 K 列
        c = _note(ws, r, f, 'B', 'J', font=F_TXT)
        c.alignment = ALW
        _lnk(ws, f'K{r}', '去看 →', sh, ref, F_LINKS, border=False)
        ws.row_dimensions[r].height = h
        r += 1
    rg = f'B{rem0}:J{r - 1}'
    ws.conditional_formatting.add(rg, FormulaRule(formula=[f'LEFT($B{rem0},1)="✗"'], font=F_RED))
    ws.conditional_formatting.add(f'K{rem0}:K{r - 1}', FormulaRule(formula=[f'$B{rem0}=""'], font=Font(name=YH, sz=10, color='FFFFFFFF')))
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
    _box(ws, r, 'B', 'K', '自动：记账分录（每一笔一借一贷）→ 成本分摊表（每月把成本分到款式：直接记到款式的按交货进度转出，公共的按交货双数 × 系数分）'
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
    section(ws, r, 'B', 'K', '各张表（点表名跳过去；每张表左上角有「←首页」回到这里）', C_HOME)
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
            put(ws, f'{cd1}{rr}', desc, F_NOTE, align=ALW)
            for i in range(CI(cd1) + 1, CI(cd2) + 1):
                ws.cell(row=rr, column=i).border = BD
            ws.row_dimensions[rr].height = max(ws.row_dimensions[rr].height or 0, 20 if len(desc) <= 30 else 30)
        r += (len(sheets) + 1) // 2
    r += 1

    # ── 怎么用
    section(ws, r, 'B', 'K', '怎么用（第一次看这里）', C_HOME)
    steps = ['① 第一次用：【基础资料】改公司名称、会计年度、建账日期（最好是某个月的 1 号），填资金账户的期初余额；'
             '【往来单位】登记供应商、加工厂、客户（和期初欠款）；【款式档案】登记款式和结算单价；【会计科目表】填其余年初余额，第 3 行要显示「√ 平」。'
             + ('表里现在的演示数据怎么清，看下面「开始记自己的账」。' if demo else ''),
             '② 每天：收付款记【资金日记账】（类别可以不选，选了供应商 / 客户 / 费用项目会自动认）；供应商拿来送货单就记【送货单登记】'
             '（送货日期照单子写，收单日期写拿到那天）；外发的加工单记【外发加工登记】。',
             '③ 电商部的采购单整张粘进【订单明细】（选择性粘贴→数值），每行补上订单号；交货了填交货日期（分批交就复制一行）；'
             '退货填负数实交；返修好再交回去的另起一行、备注写「返修」。每月月底录【工资登记】。',
             '④ 某个款专用的材料、外发、计件工资、模具，「用在哪」填款式编码，成本直接记到这个款，按交货进度转出；不填的按交货双数 × 系数分到各款'
             '（系数在【款式档案】调）。款不再做了，把【款式档案】状态改「停产」，以后它的专用成本按公共分。',
             '⑤ 看结果：【款式成本利润】每个款赚多少；【订单汇总】【订单查询】每张单赚多少；【利润表】【资产负债表】给老板看；'
             '跟供应商、电商部对账用【供应商对账单】【客户对账单】。',
             '⑥ 每月月底照下面「每月月底做什么」做一遍。供应商单子晚到会改动以前月份的成本——某个月的数不想再变了，'
             '就在【基础资料】⑩ 填那个月的结账日期。',
             '⑦ 注意：录错了就改或清空那几格，不要插行、删行（右边有隐藏公式）；灰色格子是公式，不要往里打字。每张表左上角「←首页」回到这里。'
             '新的一年看最下面「新的一年怎么开」。']
    for t in steps:
        r += 1
        _note(ws, r, t, height=max(34, _h(t, 72)), font=F_TXT)

    # ── 开始记自己的账：清演示数据（#18）
    if demo:
        r += 2
        co = ctx.get('coa_open', {})
        alias = {'3104': '年初未分配利润'}
        co_txt = '、'.join(f'{alias.get(code, nm)} {v:,.2f}'.rstrip('0').rstrip('.') for code, v in co.items()
                          for c2, nm, *_x in COA if c2 == code)
        coa_in_row = next((COA_R0 + i for i, x in enumerate(COA) if x[0] in co), COA_R0)
        r = _steps(ws, r, '开始记自己的账：先把演示数据清掉（备注写「示例」的都是编的）', [
            ('1. 5 张录入表（资金日记账、送货单登记、外发加工登记、订单明细、工资登记）和【手工分录】：选中演示行的内容按 Delete'
             '（只清内容，不要删行、插行）。', SH_CASH, 'A1'),
            ('2. 【往来单位】【款式档案】【品名档案】：演示的单位、款式、品名换成自己的（期初应付、期初应收填自己建账日那天的数）。',
             SH_UNIT, f'{UN_NAME}{UN_R0}'),
            ('3. 【基础资料】② 资金账户期初换成建账日那天的实际余额（「示例余额」字样删掉）；⑦ 固定资产的演示设备删掉、换成自己的'
             '（不删会每月提折旧，算进款式成本和管理费用）。', SH_BASE, f'{FA_NAME}{FA_R0}'),
            (f'4. 【会计科目表】G 列淡黄格的年初余额是示例（{co_txt}），换成自己的。' if co_txt else
             '4. 【会计科目表】G 列淡黄格的年初余额换成自己的。', SH_COA, f'{COA_OPEN_IN}{coa_in_row}'),
            ('5. 【基础资料】改公司名称、会计年度、建账日期；首页 C4 空着就看最新的月份。', SH_BASE, f'{PA_VAL}{PA_ROW["CO"]}'),
            ('6. 清完先看【会计科目表】C3「期初平不平」：要显示「√ 平」。不平先核对上面 2～4 几处，都对了最后再用 3104 年初未分配利润倒挤；'
             '再看【数据校验】有没有 ✗。', SH_COA, 'C3'),
        ], color='FFC00000')
    r += 2

    # ── 每月月底做什么（#33）
    r = _steps(ws, r, '每月月底做什么（按顺序做；不做这几步，当月成本、利润不准）', [
        ('1. 本月工资全部录进【工资登记】（月份填本月）：不录，本月交货的款、订单单双成本偏低（首页「提醒」会说哪个月没录）。', SH_WAGE, 'A1'),
        ('2. 本月已付的房租、水电、维修都记进【资金日记账】（按付款日期算月份）。', SH_CASH, 'A1'),
        (f'3. 催供应商、加工厂把本月的送货单 / 加工单拿来录上；【{SH_LATE}】看哪些单子晚到。', SH_LATE, 'A1'),
        ('4. 车间还有一大批做了一半的鞋，可以在【成本分摊表】「在制估计」本月那格填个估计数（可不填；不填＝本月公共成本全分给本月交货的款）。',
         SH_ALLOC, f'A{AL_ROWS["在制估计"]}'),
        ('5. 照银行 App、支付宝在【基础资料】② 填实际余额，差额要是 0。', SH_BASE, f'{AC_REAL}{AC_R0}'),
        ('6. 看【数据校验】：✗ 改到 0，⚠ 看一遍是不是真有问题。', SH_CHK, 'A1'),
        ('7. 这个月的数报给老板、不想再变了：在【基础资料】⑩ 这个月那一行填结账日期（比如 8 月填 9/5）。'
         '以后才收到的这个月的送货单 / 外发单算到收单那个月，这个月的数不再变；填了就不要改、不要清。'
         '只管供应商单据：工资、交货、收付款、手工分录补录到已结账的月份，那个月的数照样会变。', SH_BASE, f'{CL_DATE}{CL_R0}'),
    ])
    r += 2

    # ── 新的一年怎么开（#34）
    ny = [
        ('1. 旧本先做完：【基础资料】「折旧提到几月」填 12（12 月没业务也把全年折旧提足）；【数据校验】没有 ✗。', SH_BASE, f'{PA_VAL}{PA_ROW["DEPTO"]}'),
        ('2. 把旧本另存一份当新本（旧本留着查）。', None, None),
        ('3. 新本【基础资料】：会计年度改成新的一年；建账日期改成新一年的 1 月 1 日；⑩ 每月结账日期全部清空。', SH_BASE, f'{PA_VAL}{PA_ROW["YEAR"]}'),
        ('4. 新本清空 6 张录入表（资金日记账、送货单登记、外发加工登记、订单明细、工资登记、手工分录）的内容，不要删行；'
         '【订单明细】里还没交货的订单行留着（交了一部分的，订单数量改成还差的双数）。', SH_CASH, 'A1'),
        ('5. 期初抄旧本 12 月末的数：资金账户期初＝旧本【科目余额表】选 12 月时 1001、1002xx、1012xx 的月末余额'
         '（数据校验没有 ✗、也没录下一年的流水时，就是【基础资料】② 的「当前余额」）；'
         f'往来单位期初＝旧本每家 12 月末余额（旧本【资产负债表】C3 选 12，【往来单位】隐藏的 {UN_BAP}～{UN_BOP} 列：应付、应收、其他应收、其他应付，'
         f'取消隐藏抄到新本 {UN_AP0}～{UN_OP0} 列）。', SH_UNIT, f'{UN_AP0}{UN_HDR}'),
        (f'6. 固定资产：⑦ 原样抄过去，「建账前已提折旧」＝旧本这一台的建账前已提＋旧本本年提的（⑦ 右边隐藏的 {fa_mcol(1)}～{fa_mcol(12)} 列之和）。',
         SH_BASE, f'{FA_NAME}{FA_R0}'),
        ('7. 【会计科目表】淡黄格的年初＝旧本【科目余额表】选 12 月的月末余额；「年初未分配利润」3104＝旧本【资产负债表】12 月的「未分配利润」期末数；'
         '在制：旧本 12 月末 400101～400104 的余额，填到新本同一科目的年初（新本 1 月按公共成本分到交货的款）。', SH_COA,
         f'{COA_OPEN_IN}{next(COA_R0 + i for i, x in enumerate(COA) if x[6] == "in")}'),
        ('8. 填完看新本【会计科目表】C3 要显示「√ 平」，新本【资产负债表】的年初数要等于旧本 12 月末的数。'
         '上年送货、今年才拿到的单子记在新本（送货日期照写、收单日期填今年的，自动算到收单月）。', SH_COA, 'C3'),
    ]
    r = _steps(ws, r, '新的一年怎么开（一年一本）', ny)

    hide(ws, 'Z')
    ws.sheet_view.showGridLines = False
    print_setup(ws, None, landscape=False)
    return ws


def build(wb, ctx):
    build_chk(wb, ctx)          # 先建数据校验（首页提醒要用它的行号）
    build_home(wb, ctx)
