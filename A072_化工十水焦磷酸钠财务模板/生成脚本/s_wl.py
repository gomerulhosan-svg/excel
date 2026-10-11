# -*- coding: utf-8 -*-
"""A072 化工 · 往来组（橙）：客户往来、供应商往来、对账单、发票跟进，以及隐藏表 _序（对账单、发票清单的合并排序键）。
   照 A069 s_wl.py（build_party / build_stmt / build_seq）和 s_pl.build_inv 移植，口径改成本模板（设计 §6.5）：
   只用 layout 的定义名称、本模块自己的表格子；对账单、发票清单另外按「第 n 条」取录入表的车牌号、合同号、摘要、备注
   （INDEX(录入表!$X$4:$X$1048576, n+1)，插行删行排序不错位）。

【逐家余额口径】跟 _位 W 应收、X 应付一样（期初、运费腿、手工凭证都含）：
   客户（应收 1122）：期初应收 ＋ 销售收入含税（客户号；待定价＝0）＋ 资金台帐应收变动（对方一级 1122：收钱为负）
                     ＋ 手工凭证（一级 1122、明细＝这家、记账）借－贷。
   供应商（应付 2202）：期初应付 ＋ 进货货款计（供应商号；待定价按参考价暂估）＋ 进货运费计（物流号）＋ 销售运费计（物流号）
                     ＋ 资金台帐应付变动（对方一级 2202：付钱为负）＋ 手工凭证（一级 2202）贷－借。
   起日前欠＝期初＋到「前界」的发生；本期＝[起, 止]；止日欠款＝起日前欠＋本期货－本期钱＋其他（手工凭证）。止＝截止日时＝_位 的 应收/应付。

【客户往来】黄格 C3 起（空＝P_月初；早于建账日按建账日）、E3 止（空＝P_截止；晚于截止日按截止日，早于建账日按建账日前一天）。
   前界＝MIN(起－1, 止)。第 4 行固定格：C4 客户欠款合计（止日欠款只算正数）、E4 预收合计（正数显示）、G4 超期没回款家数
   （到截止日 位_没回款天数 > P_回款天数）、J4 超期的欠款（这些家的 位_应收）、M4 待定价发货笔数（列出的各家，到截止日）。
   第 7 行起每家一行：列入＝往来单位类型是客户/客户和供应商，或者（别的类型）这段时间有应收余额、发生额（状态里写「类型是…」）；
   按止日欠款从大到小（预收的排最后，一样多按往来单位表顺序）；后面合计行、「共 N 家」。
   列：A 序号 B 客户 C 期初应收 D 起日前欠 E 本期发货 F 本期回款 G 其他（手工凭证等）H 止日欠款 I 预收 J 已开票 K 未开票
   L 待定价发货(笔) M 最后发货 N 最后回款 O 没回款天数 P 状态（J～O 是到截止日的数）。
【供应商往来】同样（类型＝供应商、物流、客户和供应商；别的类型有应付发生的也列出来）。C4 欠供应商合计、E4 预付合计、
   G4 其中欠物流单位、J4 供应商欠票合计、M4 待定价到货笔数。本期到货＝进货货款（作为供应商）＋进货运费、销售运费（作为物流单位）。
   列：A 序号 B 供应商/物流 C 类型 D 期初应付 E 起日前欠 F 本期到货 G 本期付款 H 其他 I 止日欠款 J 预付 K 已收票 L 欠票
   M 待定价到货(笔) N 最后到货（含拉货：进货、销售运费）O 最后付款 P 状态。
   隐藏（两张表一样）：AA/AB 标量（第 3 行起）；AD..AV 每家一行（第 u 家在第 6+u 行）：家、有效、列入、期初、前欠、货、钱、其他、止欠、
   分键、名次、（AO＝清单第 k 行取第几家：-1 合计、-2 共几家）、简称重复、类型、最后到货（供应商表）、已开/已收票、未开/欠票、待定价。

【对账单】黄格 B3 往来单位（简称、全称、别名都认；空＝第一家客户）、F3 对账方向（客户/供应商；空＝按单位类型；
   这家没有那个方向时按类型，D4 ⚠；既不是客户也不是供应商：选了就按选的，没选按余额在哪边）、H3 起（空＝建账日）、J3 止（空＝截止日）。
   期初＝到前界 MIN(起－1, 止) 的欠款（不算截止日以后的行）。第 6 行抬头、第 7 行期间和联系人、第 8 行期末大字、
   第 9 行 上期结欠＋本期货－本期钱（±调账）＝期末、第 10 行本期发票；第 11 行表头（每页重复）、第 12 行期初、第 13 行起明细（最多 300 行），
   紧跟本期合计、共几笔、签字。明细：客户＝销售（每车）、资金台帐（对方一级 1122）、手工凭证（一级 1122、明细＝这家）；
   供应商＝进货货款（供应商）、进货运费（物流单位）、销售运费（物流单位）、资金台帐（对方一级 2202）、手工凭证（一级 2202）。
   同一天先货后钱、手工最后。每行：A 日期 B 车牌/编号 C 物料或产品/摘要 D 吨 E 单价/运费 F 税点 G 发货(到货)金额 H 回款(付款)金额
   I 余额（＝期初＋SUM(G)－SUM(H)）J 备注。运费行：日期｜车牌｜物料或产品｜吨｜运费｜税点｜金额。
   隐藏：AA/AB 标量；AD..AH 每行：k、类型（1 明细 2 本期合计 3 共几笔 4 签字 5 日期）、键、源（1 销售 2 进货货款 3 进货运费
   4 销售运费 5 资金 6 手工）、第几条。
【_序】A 列（第 2 行起）：对账单合并排序键＝日期×SORT_M＋偏移；偏移按段：销售 1..N_SALE、进货货款、进货运费、销售运费、资金、手工
   （SEG_OFF），同一天先货后钱。C 列：发票清单键（发票登记第 n 条＝日期×M＋n，只算有效、且不是「开出＋号码已在销售台账」的；
   销售台账填了发票号的第 n 车＝日期×M＋N_INV＋n）。E 列：拉货键（进货第 n 条运费计>0：物流号×M＋日期；销售第 n 条：同）。

【发票跟进】黄格 B3 ④ 的月份（1～12；也认 2026-9、202609、9月、那月的日期；只认本年度；空＝截止月）。
   第 4、5 行 KPI（截至截止日）：B5 客户没开票、D5 多开给客户的票、E5 供应商欠票、G5 供应商多给的票（位_未开票、位_欠票）。
   ① 1～12 月（建账月～截止月，日期 ≤ 截止）：开出＝销售台账填了发票号的车（收入含税、销项）＋发票登记开出且号码不在销售台账里；
   收到＝发票登记收到（含补票）；其中专票税额＝票种可抵税率>0 的税额；开出税额－收到专票税额。
   ②③④ 从第 FLOW_R0 行起一块接一块往下排：② 客户开票情况（是客户、未开票≠0，未开票多的在前）、③ 供应商欠票（是供应商、欠票≠0）、
   ④ 选的月的发票清单（_序 C 列的键，[月初, MIN(月末, 截止)]）。②③ 合计行的未开票、欠票只加正数（「不含多开 / 不含多给」）。
   隐藏：AA/AB 标量；AD..AG 每家（第 3 行起）：客户未开票、客户键、供应商欠票、供应商键；AH..AL 往下排的段、段内行、行类、第几条、来源；
   AN..AZ 票种 × 月的收到税额（AN＝可抵税率>0）。"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.workbook.defined_name import DefinedName
from common import *
from layout import *          # layout 的 norm / esc 在后面，盖过 common 的
from s_mirror import unit_no
from s_mirror3 import ym_any, ym_ok

M = SORT_M
NL = 300                                   # 对账单明细最多显示 300 行
SEGS = [('销售', N_SALE), ('进货货款', N_PUR), ('进货运费', N_PUR), ('销售运费', N_SALE), ('资金', N_CASH), ('手工', N_MJ)]
SEG_OFF = []                               # 第 s 段（1 起）偏移从 SEG_OFF[s-1]+1 到 SEG_OFF[s-1]+段长
_o = 0
for _nm, _n in SEGS:
    SEG_OFF.append(_o)
    _o += _n
NSEQ = _o
SEG_END = [SEG_OFF[i] + SEGS[i][1] for i in range(len(SEGS))]
assert NSEQ < SORT_M, '合并排序键的条数要小于 SORT_M'
SEQ = f'{H_SORT}!$A$2:$A${NSEQ + 1}'
NIV = N_INV + N_SALE                       # 发票清单键：发票登记 1..N_INV，销售 N_INV+1..
assert NIV < SORT_M
IVK = f'{H_SORT}!$C$2:$C${NIV + 1}'
NFR = N_PUR + N_SALE                       # 拉货键：进货 1..N_PUR，销售 N_PUR+1..
FRK = f'{H_SORT}!$E$2:$E${NFR + 1}'
K1122, K2202 = kh('1122'), kh('2202')
NOTKT = 'IFERROR(MATCH("无票",票种_名称,0),-1)'

_thin = Side(style='thin', color='FFBFBFBF')
CF_BD = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
LBL = fill('FFD9E1F2')
F_KV = Font(name=YH, sz=10, bold=True, color='FFC00000')
F_BIG = Font(name=YH, sz=16, bold=True, color='FF000000')
F_BIG2 = Font(name=YH, sz=14, bold=True, color='FF1F3864')
F_HEADB = Font(name=YH, sz=11, bold=True, color='FF000000')
F_KPI_B = Font(name=YH, sz=11, bold=True, color='FF1F3864')
DATE_SEL = 'IF(AND(ISNUMBER({x}),{x}>=36526,{x}<73051),INT({x}),{d})'
OK_SEL = 'IF(TRIM({x}&"")="",1,IF(AND(ISNUMBER({x}),{x}>=36526,{x}<73051),1,0))'
QTY = 'General'
PRICE = '#,##0.00##'
DAYS = '0;-0;0'                            # 0 天也显示 0（没欠款时 _位 给 ""，自然是空）


def cf_fill(rgb):
    return PatternFill('solid', fgColor=rgb, bgColor=rgb)


def R2(x):
    return f'ROUND({x},2)'


class Sc:
    """隐藏标量：AA 列名字、AB 列值（第 r0 行起）"""

    def __init__(self, ws, keys, r0=3, cl='AA', cv='AB'):
        self.ws, self.cv = ws, cv
        self.r = {k: r0 + i for i, k in enumerate(keys)}
        for k, r in self.r.items():
            ws[f'{cl}{r}'] = k
            ws[f'{cl}{r}'].font = F_HELP

    def __call__(self, k):
        return f'${self.cv}${self.r[k]}'

    def ext(self, k):
        """别的表引用这个标量"""
        return f"'{self.ws.title}'!${self.cv}${self.r[k]}"

    def set(self, k, f):
        c = self.ws[f'{self.cv}{self.r[k]}']
        c.value = '=' + f
        c.font = F_HELP


def hset(ws, coord, f):
    ws[coord] = f if isinstance(f, (int, float)) else '=' + f
    ws[coord].font = F_HELP


def cell(ws, coord, f, font=F_TXT, fmt=None, align=AC):
    c = ws[coord]
    c.value = '=' + f
    c.font = font
    if fmt:
        c.number_format = fmt
    c.alignment = align
    return c


def kpi(ws, lbl_cell, lbl, val_cell, f, fmt=MONEY):
    put(ws, lbl_cell, lbl, F_KPI_L, LBL, align=ACW)
    put(ws, val_cell, '=' + f, F_KV, FILL_AUTO, fmt, AR)


def finish(ws, last, hdr_rows, start, end_cell, cap_row, freeze, landscape):
    ws.freeze_panes = freeze
    print_setup(ws, hdr_rows, landscape=landscape)
    ws.oddFooter.center.text = '第 &P 页 共 &N 页'
    q = f"'{ws.title}'"
    ws.defined_names['Print_Area'] = DefinedName(
        'Print_Area', attr_text=f'{q}!$A${start}:INDEX({q}!${last}$1:${last}${cap_row},{q}!{end_cell})')


def chain(pairs, default='""'):
    out = default
    for c, v in reversed(pairs):
        out = f'IF({c},{v},{out})'
    return out


# ═══════════════════════════════ 口径（客户 1122 / 供应商 2202） ═══════════════════════════════
def goods(cust, u, cd):
    """货：客户＝销售收入含税；供应商＝进货货款计＋进货运费计（物流号）＋销售运费计（物流号）。cd(日期名) → 日期条件"""
    if cust:
        return f'SUMIFS(销_收入含税,销_客户号,{u},{cd("销_日期")})'
    return (f'SUMIFS(进_货款计,进_供应商号,{u},{cd("进_日期")})+SUMIFS(进_运费计,进_物流号,{u},{cd("进_日期")})'
            f'+SUMIFS(销_运费计,销_物流号,{u},{cd("销_日期")})')


def chg(cust, u, cd):
    """资金台帐对余额的影响（带符号）：应收变动（收钱为负）/ 应付变动（付钱为负）"""
    col = '资_应收变动' if cust else '资_应付变动'
    return f'SUMIFS({col},资_单位号,{u},{cd("资_日期")})'


def mj(cust, side, u, cd):
    """手工凭证一级 1122/2202、明细＝这家、记账的 借/贷"""
    return f'SUMIFS(手_{side},手_单位号,{u},手_一级号,{K1122 if cust else K2202},手_记账,1,{cd("手_日期")})'


def other(cust, u, cd):
    """手工凭证对余额的影响：应收＝借－贷；应付＝贷－借"""
    a, b = ('借', '贷') if cust else ('贷', '借')
    return f'{mj(cust, a, u, cd)}-{mj(cust, b, u, cd)}'


# ═══════════════════════════════ 客户往来 / 供应商往来 ═══════════════════════════════
P_H, P_L0 = 6, 7                       # 表头行、清单第一行
P_N = N_UNIT + 2                       # 每家一行＋合计＋共几家
U0 = P_L0                              # 隐藏每家表：第 u 家在第 U0+u-1 行
HU = dict(家='AD', 有='AE', 是='AF', 期初='AG', 前欠='AH', 货='AI', 钱='AJ', 他='AK', 止欠='AL', 分='AM', 名次='AN', 取='AO',
          重='AP', 类='AQ', 末货='AR', 票='AS', 欠票='AT', 待='AU')
COLS_AR = dict(序='A', 名='B', 期初='C', 前欠='D', 货='E', 钱='F', 他='G', 止欠='H', 预='I', 票='J', 欠票='K', 待='L', 末货='M',
               末钱='N', 天='O', 状态='P')
COLS_AP = dict(序='A', 名='B', 类='C', 期初='D', 前欠='E', 货='F', 钱='G', 他='H', 止欠='I', 预='J', 票='K', 欠票='L', 待='M',
               末货='N', 末钱='O', 状态='P')


def _rng(col):
    return f'${col}${U0}:${col}${U0 + N_UNIT - 1}'


def ap_last_arrive_col():
    """供应商往来隐藏「最后到货（含拉货）」那一列（第 u 家第 u 个），发票跟进 ③ 用"""
    return f"'{SH_AP}'!{_rng(HU['末货'])}"


TIP_AR = ('★ 每家客户欠多少、这段时间发了多少货、回了多少款。黄格选起止日期：起空着＝截止日那个月 1 号，止空着＝截止日'
          '（截止日在首页改；晚于截止日的按截止日算，早于建账日的按建账日），右边灰字是实际用的日期。列的是【往来单位】里类型是客户'
          '（或客户和供应商）的，别的类型有应收发生的也列出来（状态里写类型）。按止日欠款从大到小排，预收（负数）的排最后。'
          '起日前欠＝期初应收＋起日以前的发货－回款±手工凭证；止日欠款＝起日前欠＋本期发货－本期回款＋其他（手工凭证调账）。'
          '发货按提货日期、含税、待定价的不算；回款＝资金台帐里记到应收账款的（收客户货款，退款是负数）。'
          '已开票、未开票、待定价、最后发货、最后回款、没回款天数都是到截止日的数。要给哪家对账，去【对账单】选它。')
TIP_AP = ('★ 欠每家供应商、物流单位多少，这段时间到了多少货、付了多少款。黄格选起止日期（同客户往来）。列的是【往来单位】里类型是'
          '供应商、物流、客户和供应商的，别的类型有应付发生的也列出来。本期到货＝进货货款（这家是供应商）＋进货运费、销售运费'
          '（这家是物流单位；进货没填物流单位、有运费的算供应商的）；待定价的货款按物料档案参考单价暂估。本期付款＝资金台帐里记到'
          '应付账款的（原材料材料款、原材料运输费、销货运费……，退回是负数）；其他＝手工凭证调账。'
          '已收票、欠票、待定价、最后到货、最后付款都是到截止日的数；欠票是负数＝对方发票多给了。')


def build_party(ws, cust):
    who = '客户' if cust else '供应商'
    C = COLS_AR if cust else COLS_AP
    LAST = 'P'
    if cust:
        W = dict(A=5, B=16, C=13.5, D=13.5, E=13.5, F=13.5, G=12, H=13.5, I=6, J=13.5, K=13.5, L=8, M=12.5, N=12.5, O=8,
                 P=30)
    else:
        W = dict(A=5, B=16, C=9, D=13.5, E=13.5, F=13.5, G=13.5, H=12, I=13.5, J=6, K=13.5, L=13.5, M=8, N=12.5, O=12.5,
                 P=30)
    widths(ws, W)
    title(ws, f'=IF(P_公司名称="","",P_公司名称&"　")&"{who}往来"', LAST, C_WL, TIP_AR if cust else TIP_AP)
    selector(ws, 'B3', '起日期', 'C3', None, fmt=DATE)
    selector(ws, 'D3', '止日期', 'E3', None, fmt=DATE)
    dv_date(ws, 'C3')
    dv_date(ws, 'E3')
    home_link(ws, f'{LAST}3')
    ws.row_dimensions[3].height = 30

    keys = ['起选', '起', '起认出', '止输入', '止0', '止', '止认出', '前界', '家数', '末行', '回款天数',
            't_期初', 't_前欠', 't_货', 't_钱', 't_他', 't_止欠', 't_票', 't_欠票', 't_待', 'k_欠', 'k_预', 'k_3', 'k_4']
    sc = Sc(ws, keys)
    S0, S1 = sc('起'), sc('止')
    sc.set('起选', DATE_SEL.format(x='C3', d='P_月初'))
    sc.set('起', f'MAX({sc("起选")},P_建账日)')                      # 早于建账日按建账日（建账前的数都在期初里）
    sc.set('起认出', OK_SEL.format(x='C3'))
    sc.set('止输入', DATE_SEL.format(x='E3', d='P_截止'))
    sc.set('止0', f'MIN({sc("止输入")},P_截止)')
    sc.set('止', f'MAX({sc("止0")},P_建账日-1)')                     # 早于建账日按建账日前一天（只有期初）
    sc.set('止认出', OK_SEL.format(x='E3'))
    sc.set('前界', f'MIN({S0}-1,{S1})')                              # 起日前欠算到这天（≤ 止 ≤ 截止）
    sc.set('回款天数', 'P_回款天数')

    flag = '位_是客户' if cust else '位_是供应商'
    open_ = '位_期初应收' if cust else '位_期初应付'
    bef = lambda nm: f'{nm},"<="&{sc("前界")}'
    win = lambda nm: dr(nm, S0, S1)

    # ── 隐藏：每家一行 ──
    H = HU
    for i in range(N_UNIT):
        u, r = i + 1, U0 + i
        c = lambda k: f'${H[k]}{r}'
        ws[f'{H["家"]}{r}'] = u
        ws[f'{H["家"]}{r}'].font = F_HELP
        hset(ws, f'{H["有"]}{r}', f'IF(INDEX(位_有效,{u})=1,1,0)')
        ok = f'{c("有")}=1'
        hset(ws, f'{H["期初"]}{r}', f'IF({ok},INDEX({open_},{u}),0)')
        hset(ws, f'{H["前欠"]}{r}',
             f'IF({ok},{R2(c("期初") + "+" + goods(cust, u, bef) + "+" + chg(cust, u, bef) + "+" + other(cust, u, bef))},0)')
        hset(ws, f'{H["货"]}{r}', f'IF({ok},{R2(goods(cust, u, win))},0)')
        hset(ws, f'{H["钱"]}{r}', f'IF({ok},{R2("-" + chg(cust, u, win))},0)')
        hset(ws, f'{H["他"]}{r}', f'IF({ok},{R2(other(cust, u, win))},0)')
        hset(ws, f'{H["止欠"]}{r}', R2(f'{c("前欠")}+{c("货")}-{c("钱")}+{c("他")}'))
        hset(ws, f'{H["是"]}{r}', (f'IF({ok},IF(OR(INDEX({flag},{u})=1,ABS({c("前欠")})>0.005,ABS({c("止欠")})>0.005,'
                                  f'{c("货")}<>0,{c("钱")}<>0,{c("他")}<>0),1,0),0)'))
        isin = f'{c("是")}=1'
        hset(ws, f'{H["分"]}{r}', f'IF({isin},ROUND({c("止欠")}*100,0),"")')
        kk = c('分')
        above = (f'+COUNTIF(${H["分"]}${U0}:{H["分"]}{r - 1},{kk})' if i else '')
        hset(ws, f'{H["名次"]}{r}', f'IF({isin},1+COUNTIF({_rng(H["分"])},">"&{kk}){above},"")')
        # 简称重复：前面已有同名的（发生额都记在第一家）
        hset(ws, f'{H["重"]}{r}', (f'IF({ok},IF(IFERROR(MATCH({esc(f"INDEX(位_简称,{u})")},位_简称,0),{u})<{u},1,0),0)'
                                  if u > 1 else '0'))
        hset(ws, f'{H["类"]}{r}', f'IF({ok},INDEX(位_类型,{u})&"","")')
        if cust:
            hset(ws, f'{H["票"]}{r}', f'IF({ok},INDEX(位_已开票,{u}),0)')
            hset(ws, f'{H["欠票"]}{r}', f'IF({ok},INDEX(位_未开票,{u}),0)')
            hset(ws, f'{H["待"]}{r}', f'IF({ok},INDEX(位_待定价发货,{u}),0)')
        else:
            hset(ws, f'{H["票"]}{r}', f'IF({ok},INDEX(位_已收票,{u}),0)')
            hset(ws, f'{H["欠票"]}{r}', f'IF({ok},INDEX(位_欠票,{u}),0)')
            hset(ws, f'{H["待"]}{r}', f'IF({ok},INDEX(位_待定价到货,{u}),0)')
            # 最后到货（到截止日）：进货（作为供应商，_位）和拉货（作为物流单位，_序 E 列）里最晚的
            v = f'IFERROR(LARGE({FRK},COUNTIF({FRK},">"&({u}*{M}+P_截止))+1),0)'
            hset(ws, f'{H["末货"]}{r}', f'IF({ok},MAX(INDEX(位_最后到货,{u}),IF({v}>{u}*{M},{v}-{u}*{M},0)),0)')
            ws[f'{H["末货"]}{r}'].number_format = DATE
    rg = lambda k: _rng(H[k])
    sc.set('家数', f'COUNTIF({rg("是")},1)')
    sc.set('末行', f'{P_L0 - 1}+{sc("家数")}+2')
    for k, hk in (('t_期初', '期初'), ('t_前欠', '前欠'), ('t_货', '货'), ('t_钱', '钱'), ('t_他', '他'), ('t_止欠', '止欠'),
                  ('t_票', '票'), ('t_欠票', '欠票')):
        sc.set(k, R2(f'SUMIFS({rg(hk)},{rg("是")},1)'))
    sc.set('t_待', f'SUMIFS({rg("待")},{rg("是")},1)')
    sc.set('k_欠', R2(f'SUMIFS({rg("止欠")},{rg("是")},1,{rg("止欠")},">0")'))
    sc.set('k_预', R2(f'-SUMIFS({rg("止欠")},{rg("是")},1,{rg("止欠")},"<0")'))
    if cust:
        sc.set('k_3', 'COUNTIFS(位_有效,1,位_没回款天数,">"&P_回款天数)')
        sc.set('k_4', R2('SUMIFS(位_应收,位_有效,1,位_没回款天数,">"&P_回款天数)'))
    else:
        sc.set('k_3', R2(f'SUMIFS({rg("止欠")},{rg("是")},1,{rg("止欠")},">0",{rg("类")},"物流")'))
        sc.set('k_4', R2(f'SUMIFS({rg("欠票")},{rg("是")},1,{rg("欠票")},">0")'))

    # ── 第 3、4、5 行 ──
    ws.merge_cells('F3:O3')
    put(ws, 'F3', (f'="实际用的："&TEXT({S0},"yyyy-mm-dd")&" ～ "&TEXT({S1},"yyyy-mm-dd")'
                   f'&IF(TRIM(C3&"")="","（起空着＝截止日那个月 1 号"&IF(P_月初<P_建账日,"，早于建账日按建账日","")&"）","")'
                   f'&IF(TRIM(E3&"")="","（止空着＝截止日）","")'
                   f'&IF({sc("起认出")}=0,"　⚠ 起日期没认出，按截止日那个月 1 号","")'
                   f'&IF({sc("止认出")}=0,"　⚠ 止日期没认出，按截止日","")'
                   f'&IF(AND(TRIM(C3&"")<>"",{sc("起认出")}=1,{sc("起选")}<P_建账日),"　⚠ 起早于建账日 "&TEXT(P_建账日,"yyyy-mm-dd")'
                   f'&"，按建账日（建账前的数都在期初里）","")'
                   f'&IF({sc("止输入")}<P_建账日,"　⚠ 止早于建账日 "&TEXT(P_建账日,"yyyy-mm-dd")&"，按建账日前一天（只有期初）","")'
                   f'&IF({sc("止输入")}>P_截止,"　⚠ 止晚于截止日 "&TEXT(P_截止,"yyyy-mm-dd")&"，按截止日算","")'
                   f'&IF({sc("起选")}>{sc("止0")},"　✗ 起晚于止：本期都是 0，请改日期","")'), F_NOTE, align=ALW, border=False)
    ws.conditional_formatting.add('F3', FormulaRule(formula=['OR(ISNUMBER(FIND("⚠",$F$3)),ISNUMBER(FIND("✗",$F$3)))'],
                                                    font=Font(bold=True, color='FFC00000')))
    if cust:
        kpi(ws, 'B4', '客户欠款合计\n（只算欠的）', 'C4', sc('k_欠'))
        kpi(ws, 'D4', '预收合计', 'E4', sc('k_预'))
        put(ws, 'F4', '="超 "&P_回款天数&" 天没回款（到截止日）"', F_KPI_L, LBL, align=ACW)
        put(ws, 'G4', '=' + sc('k_3'), F_KV, FILL_AUTO, '0" 家"', AR)
        kpi(ws, 'I4', '超期的欠款\n（到截止日）', 'J4', sc('k_4'))
        ws.merge_cells('K4:L4')
        put(ws, 'K4', '待定价发货', F_KPI_L, LBL, align=ACW)
    else:
        kpi(ws, 'B4', '欠供应商合计\n（只算欠的）', 'C4', sc('k_欠'))
        kpi(ws, 'D4', '预付合计', 'E4', sc('k_预'))
        kpi(ws, 'F4', '其中欠物流单位', 'G4', sc('k_3'))
        kpi(ws, 'I4', '供应商欠票\n（到截止日）', 'J4', sc('k_4'))
        ws.merge_cells('K4:L4')
        put(ws, 'K4', '待定价到货', F_KPI_L, LBL, align=ACW)
    ws['L4'].border = BD
    put(ws, 'M4', '=' + sc('t_待'), F_KV, FILL_AUTO, '0" 笔"', AR)
    ws.row_dimensions[4].height = 32
    ws.merge_cells(f'A5:{LAST}5')
    if cust:
        t5 = ('="按止日欠款从大到小排，预收（负数）的排最后。已开票、未开票、待定价、最后发货、最后回款、没回款天数、超期都是到截止日 "'
              '&TEXT(P_截止,"yyyy-mm-dd")&" 的数；没回款天数超过 "&P_回款天数&" 天标红（天数在【基础资料】改）。"')
    else:
        t5 = ('="按止日欠款从大到小排，预付（负数）的排最后。已收票、欠票、待定价、最后到货、最后付款都是到截止日 "'
              '&TEXT(P_截止,"yyyy-mm-dd")&" 的数；欠票是负数＝对方发票多给了；最后到货含给我们拉货的日子。"')
    put(ws, 'A5', t5, F_NOTE, align=ALW, border=False)
    ws.row_dimensions[5].height = 20

    # ── 表头 ──
    if cust:
        heads = [('序', '序号'), ('名', '客户'), ('期初', '期初应收\n（建账日）'), ('前欠', '起日前欠'), ('货', '本期发货'),
                 ('钱', '本期回款'), ('他', '其他\n（手工凭证等）'), ('止欠', '止日欠款\n（负数＝预收）'), ('预', '预收'),
                 ('票', '已开票'), ('欠票', '未开票'), ('待', '待定价\n发货(笔)'), ('末货', '最后发货'), ('末钱', '最后回款'),
                 ('天', '没回款\n天数'), ('状态', '状态')]
    else:
        heads = [('序', '序号'), ('名', '供应商、物流'), ('类', '类型'), ('期初', '期初应付\n（建账日）'), ('前欠', '起日前欠'),
                 ('货', '本期到货\n（含运费）'), ('钱', '本期付款'), ('他', '其他\n（手工凭证等）'), ('止欠', '止日欠款\n（负数＝预付）'),
                 ('预', '预付'), ('票', '已收票'), ('欠票', '欠票'), ('待', '待定价\n到货(笔)'), ('末货', '最后到货'),
                 ('末钱', '最后付款'), ('状态', '状态')]
    header(ws, P_H, [(C[k], t) for k, t in heads], C_WL, height=36)

    # ── 清单 ──
    n = sc('家数')
    for k in range(1, P_N + 1):
        r = P_L0 + k - 1
        ix = f'${H["取"]}{r}'
        hset(ws, f'{H["取"]}{r}', f'IF({k}<={n},IFERROR(MATCH({k},{rg("名次")},0),0),IF({k}={n}+1,-1,IF({k}={n}+2,-2,0)))')
        pk = lambda rng_, tot: f'IF({ix}=0,"",IF({ix}>0,INDEX({rng_},{ix}),IF({ix}=-1,{tot},"")))'
        pos = lambda rng_: f'IF({ix}>0,IF(INDEX({rng_},{ix})>0,INDEX({rng_},{ix}),""),"")'
        cell(ws, f'{C["序"]}{r}', f'IF({ix}>0,{k},"")', F_NOTE)
        cell(ws, f'{C["名"]}{r}', f'IF({ix}>0,INDEX(位_简称,{ix}),IF({ix}=-1,"合计",IF({ix}=-2,"共 "&{n}&" 家","")))',
             align=AL)
        for hk, tk in (('期初', 't_期初'), ('前欠', 't_前欠'), ('货', 't_货'), ('钱', 't_钱'), ('他', 't_他'),
                       ('止欠', 't_止欠'), ('票', 't_票'), ('欠票', 't_欠票')):
            cell(ws, f'{C[hk]}{r}', pk(rg(hk), sc(tk)), fmt=MONEY, align=AR)
        cell(ws, f'{C["待"]}{r}', pk(rg('待'), sc('t_待')), fmt=INT)
        zq = f'INDEX({rg("止欠")},{ix})'
        cell(ws, f'{C["预"]}{r}', f'IF({ix}>0,IF({zq}<-0.005,"{"预收" if cust else "预付"}",""),"")', F_TXTB)
        tp = f'INDEX({rg("类")},{ix})'
        if cust:
            cell(ws, f'{C["末货"]}{r}', pos('位_最后发货'), fmt=DATE)
            cell(ws, f'{C["末钱"]}{r}', pos('位_最后回款'), fmt=DATE)
            cell(ws, f'{C["天"]}{r}', f'IF({ix}>0,INDEX(位_没回款天数,{ix}),"")', fmt=DAYS)
            st = (f'IF({zq}<-0.005,"、预收","")&IF(N(INDEX(位_没回款天数,{ix}))>P_回款天数,"、超 "&P_回款天数&" 天没回款","")'
                  f'&IF(INDEX({rg("待")},{ix})>0,"、有待定价","")&IF(INDEX({rg("欠票")},{ix})<-0.005,"、发票多开","")'
                  f'&IF(INDEX({rg("重")},{ix})=1,"、简称重复（数记在同名第一家）","")'
                  f'&IF(INDEX(位_是客户,{ix})=0,IF({tp}="","、往来单位没填类型","、类型是「"&{tp}&"」"),"")')
        else:
            cell(ws, f'{C["类"]}{r}', f'IF({ix}>0,{tp},"")')
            cell(ws, f'{C["末货"]}{r}', pos(rg('末货')), fmt=DATE)
            cell(ws, f'{C["末钱"]}{r}', pos('位_最后付款'), fmt=DATE)
            st = (f'IF({zq}<-0.005,"、预付","")&IF(INDEX({rg("待")},{ix})>0,"、有待定价","")'
                  f'&IF(INDEX({rg("欠票")},{ix})<-0.005,"、对方发票多给","")'
                  f'&IF(INDEX({rg("重")},{ix})=1,"、简称重复（数记在同名第一家）","")'
                  f'&IF(INDEX(位_是供应商,{ix})=0,IF({tp}="","、往来单位没填类型","、类型是「"&{tp}&"」"),"")')
        cell(ws, f'{C["状态"]}{r}', f'IF({ix}>0,MID({st},2,100),"")', align=AL)
    a0, a1 = P_L0, P_L0 + P_N - 1
    z0 = f'${H["取"]}{a0}'
    cf = ws.conditional_formatting
    rng = f'A{a0}:{LAST}{a1}'
    cf.add(rng, FormulaRule(formula=[f'{z0}=-1'], fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{z0}=-2'], font=Font(color='FF808080'), stopIfTrue=True))
    if cust:
        tc = C['天']
        cf.add(f'{tc}{a0}:{tc}{a1}', FormulaRule(formula=[f'AND({z0}>0,N(${tc}{a0})>{sc("回款天数")})'],
                                               font=Font(bold=True, color='FFC00000'), border=CF_BD, stopIfTrue=True))
    s_ = C['状态']
    cf.add(f'{s_}{a0}:{s_}{a1}', FormulaRule(formula=[f'AND({z0}>0,OR(ISNUMBER(FIND("超",${s_}{a0})),'
                                                      f'ISNUMBER(FIND("类型",${s_}{a0}))))'],
                                           font=Font(bold=True, color='FFC00000'), border=CF_BD, stopIfTrue=True))
    cf.add(f'{s_}{a0}:{s_}{a1}', FormulaRule(formula=[f'AND({z0}>0,${s_}{a0}<>"")'],
                                           font=Font(bold=True, color='FFC65911'), border=CF_BD, stopIfTrue=True))
    p_ = C['预']
    cf.add(f'{p_}{a0}:{p_}{a1}', FormulaRule(formula=[f'AND({z0}>0,${p_}{a0}<>"")'], font=Font(bold=True, color='FF2F75B5'),
                                           fill=cf_fill('FFDDEBF7'), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{z0}>0'], border=CF_BD))
    hide(ws, *[CL(i) for i in range(CI('AA'), CI(H['待']) + 1)])
    finish(ws, LAST, f'{P_H}:{P_H}', 1, sc('末行'), a1, 'C7', True)
    return sc


# ═══════════════════════════════ 对账单 ═══════════════════════════════
ST_LAST = 'J'
ST_H, ST_O, ST_D0 = 11, 12, 13          # 表头、期初行、明细第一行
ST_N = NL + 7                           # 明细 300 行＋本期合计、共几笔、空、签字、空、空、日期
ST_D1 = ST_D0 + ST_N - 1
HS = dict(k='AD', 类型='AE', 键='AF', 源='AG', 条='AH')

TIP_ST = ('★ 给客户、供应商、物流单位对账，A4 竖着打（只印第 6 行往下），请对方核对盖章。黄格：往来单位（下拉选，也可打简称、全称、'
          '其他叫法；空＝第一家客户）、对账方向（空＝按单位类型）、起（空＝建账日）、止（空＝截止日）。明细按日期排，同一天先货后钱，'
          '手工凭证调账排在当天最后。客户：货＝销售台账每一车（按提货日期，待定价的按 0 算），钱＝资金台帐记到应收账款的（退款是负数）；'
          '供应商：货＝进货货款（待定价按参考单价暂估）＋给我们拉货的运费（进货运费、销售运费，每车一行：车牌、物料或产品、吨、运费、税点、'
          '金额），钱＝资金台帐记到应付账款的；两边都含手工凭证里明细是这家的调账。')


def build_stmt(ws):
    LAST = ST_LAST
    W = dict(A=11, B=12.5, C=18, D=8, E=10.5, F=9, G=13.5, H=13.5, I=14, J=26, K=10)
    widths(ws, W)
    title(ws, '对账单', LAST, C_WL, TIP_ST)
    home_link(ws, 'K1')
    selector(ws, 'A3', '往来单位', 'B3', None, '=往来单位列表',
             prompt='从下拉选；也可以打简称、全称、其他叫法；空着＝第一家客户')
    ws.merge_cells('B3:C3')
    selector(ws, 'D3', '对账方向', 'F3', None, '"客户,供应商"',
             prompt='空着＝按单位类型（是客户就按客户：应收；不然按供应商：应付，物流单位也选供应商）')
    ws.merge_cells('D3:E3')
    ws['E3'].border = BD
    selector(ws, 'G3', '起', 'H3', None, fmt=DATE)
    selector(ws, 'I3', '止', 'J3', None, fmt=DATE)
    dv_date(ws, 'H3')
    dv_date(ws, 'J3')
    ws.row_dimensions[3].height = 26

    keys = ['名', '名规', 'u找', '首客户', 'u', '是客户', '是供应商', '位应收', '位应付', '方向输入', 'dir',
            '起选', '起', '起认出', '止输入', '止0', '止', '止认出', '前界', 'lo', 'hi', '无票号',
            '期初', '货', '钱', '调G', '调H', '合G', '合H', '期末', '待定价', '票', '欠票', '笔数', 'n', '末行',
            '简称', '抬头名', '联系人', '电话']
    sc = Sc(ws, keys)
    U, D, S0, S1 = sc('u'), sc('dir'), sc('起'), sc('止')
    sc.set('名', 'TRIM(B3&"")')
    sc.set('名规', norm('B3'))
    sc.set('u找', unit_no(sc('名规')))
    sc.set('首客户', 'IFERROR(MATCH(1,位_是客户,0),0)')
    sc.set('u', f'IF({sc("名")}="",{sc("首客户")},{sc("u找")})')
    sc.set('是客户', f'IF({U}=0,0,INDEX(位_是客户,{U}))')
    sc.set('是供应商', f'IF({U}=0,0,INDEX(位_是供应商,{U}))')
    sc.set('位应收', f'IF({U}=0,0,INDEX(位_应收,{U}))')
    sc.set('位应付', f'IF({U}=0,0,INDEX(位_应付,{U}))')
    sc.set('方向输入', 'TRIM(F3&"")')
    di, ic, isp = sc('方向输入'), sc('是客户'), sc('是供应商')
    # 选的方向这家没有（供应商选了「客户」）就按单位类型（不然会打出「双方两清」），D4 用 ⚠ 说出来；
    # 既不是客户也不是供应商（其他、出借人…）：选了就按选的，没选按余额在哪边（只有应收有数＝客户）
    sc.set('dir', (f'IF(AND({di}="客户",OR({ic}=1,{isp}=0)),1,IF(AND({di}="供应商",OR({isp}=1,{ic}=0)),2,'
                   f'IF({ic}=1,1,IF({isp}=1,2,IF(AND(ABS({sc("位应收")})>0.005,ABS({sc("位应付")})<=0.005),1,2)))))'))
    sc.set('起选', DATE_SEL.format(x='H3', d='P_建账日'))
    sc.set('起', f'MAX({sc("起选")},P_建账日)')
    sc.set('起认出', OK_SEL.format(x='H3'))
    sc.set('止输入', DATE_SEL.format(x='J3', d='P_截止'))
    sc.set('止0', f'MIN({sc("止输入")},P_截止)')
    sc.set('止', f'MAX({sc("止0")},P_建账日-1)')
    sc.set('止认出', OK_SEL.format(x='J3'))
    sc.set('前界', f'MIN({S0}-1,{S1})')
    sc.set('lo', f'{S0}*{M}')
    sc.set('hi', f'({S1}+1)*{M}')
    sc.set('无票号', NOTKT)
    nt = sc('无票号')
    bef = lambda nm: f'{nm},"<="&{sc("前界")}'
    per = lambda nm: dr(nm, S0, S1)
    upto = lambda nm: f'{nm},"<="&{S1}'
    issued = lambda cd: (f'SUMIFS(销_收入含税,销_客户号,{U},销_发票号键,"#?*",{cd("销_日期")})'
                         f'+SUMIFS(票_价税合计,票_单位号,{U},票_方向,1,票_销售重复,0,票_有效,1,{cd("票_日期")})')
    got = lambda cd: f'SUMIFS(票_价税合计,票_单位号,{U},票_方向,2,票_补票行,0,票_有效,1,{cd("票_日期")})'
    cs = dict(
        期初=(f'INDEX(位_期初应收,{U})+{goods(True, U, bef)}+{chg(True, U, bef)}+{other(True, U, bef)}',
              f'INDEX(位_期初应付,{U})+{goods(False, U, bef)}+{chg(False, U, bef)}+{other(False, U, bef)}'),
        货=(goods(True, U, per), goods(False, U, per)),
        钱=('-' + chg(True, U, per), '-' + chg(False, U, per)),
        调G=(mj(True, '借', U, per), mj(False, '贷', U, per)),
        调H=(mj(True, '贷', U, per), mj(False, '借', U, per)),
        票=(issued(per), got(per)),
        欠票=(f'INDEX(位_期初未开票,{U})+SUMIFS(销_收入含税,销_客户号,{U},{upto("销_日期")})'
              f'-SUMIFS(销_收入含税,销_客户号,{U},销_票种号,{nt},{upto("销_日期")})-({issued(upto)})',
              f'INDEX(位_期初欠票,{U})+SUMIFS(进_货款计,进_供应商号,{U},进_货款票种号,"<>"&{nt},{upto("进_日期")})'
              f'+SUMIFS(进_运费计,进_物流号,{U},进_运费票种号,"<>"&{nt},{upto("进_日期")})'
              f'+SUMIFS(销_运费计,销_物流号,{U},销_运费票种号,"<>"&{nt},{upto("销_日期")})-{got(upto)}'),
    )
    for k, (a, b) in cs.items():
        sc.set(k, f'IF({U}=0,0,IF({D}=1,{R2(a)},{R2(b)}))')
    sc.set('合G', R2(f'{sc("货")}+{sc("调G")}'))
    sc.set('合H', R2(f'{sc("钱")}+{sc("调H")}'))
    sc.set('期末', R2(f'{sc("期初")}+{sc("合G")}-{sc("合H")}'))
    sc.set('待定价', (f'IF({U}=0,0,IF({D}=1,COUNTIFS(销_客户号,{U},销_待定价,1,销_有效,1,{per("销_日期")}),'
                    f'COUNTIFS(进_供应商号,{U},进_待定价,1,进_有效,1,{per("进_日期")})))'))
    sc.set('笔数', f'COUNT({SEQ})')
    sc.set('n', f'MIN({sc("笔数")},{NL})')
    sc.set('末行', f'{ST_D0 + 6}+{sc("n")}')
    sc.set('简称', f'IF({U}=0,{sc("名")},INDEX(位_简称,{U}))')
    sc.set('抬头名', f'IF({U}=0,"",IF(INDEX(位_全称规,{U})="",INDEX(位_简称,{U}),INDEX(位_全称规,{U})))')
    sc.set('联系人', f'IF({U}=0,"",INDEX(位_联系人,{U})&"")')
    sc.set('电话', f'IF({U}=0,"",INDEX(位_电话,{U})&"")')

    # ── 第 4、5 行：实际用的、提醒 ──
    put(ws, 'A4', '实际用的', F_NOTE, align=AC, border=False)
    ws.merge_cells('B4:C4')
    put(ws, 'B4', (f'=IF({U}=0,IF({sc("名")}="","⚠ 往来单位表里还没有客户","⚠ 不在【往来单位】表里"),'
                   f'"✓ "&{sc("简称")}&"（"&IF(INDEX(位_类型,{U})="","没填类型",INDEX(位_类型,{U}))&"）"'
                   f'&IF({sc("名")}="","　空着＝第一家客户",""))'),
        F_AUTOB, FILL_AUTO, align=ACW)
    ws.merge_cells('D4:F4')
    w1 = (f'IF(AND({di}<>"",{di}<>"客户",{di}<>"供应商"),"⚠ 方向只认客户、供应商，",'
          f'IF(AND({ic}=0,{isp}=0),"⚠ 既不是客户也不是供应商，",'
          f'IF(OR(AND({di}="客户",{D}<>1),AND({di}="供应商",{D}<>2)),"⚠ 不是"&{di}&"，","")))')
    put(ws, 'D4', (f'=IF({U}=0,"—",{w1}&IF({D}=1,"按客户对（应收）","按供应商对（应付）")'
                   f'&IF({di}="","（按单位类型）",""))'), F_AUTOB, FILL_AUTO, align=ACW)
    ws.conditional_formatting.add('D4', FormulaRule(formula=['LEFT($D$4,1)="⚠"'], font=Font(bold=True, color='FFC00000')))
    ws['E4'].border = ws['F4'].border = BD
    put(ws, 'H4', f'={S0}', F_AUTOB, FILL_AUTO, DATE, AC)
    put(ws, 'J4', f'={S1}', F_AUTOB, FILL_AUTO, DATE, AC)
    ws.row_dimensions[4].height = 30
    ws.merge_cells(f'A5:{LAST}5')
    warn = (f'IF({sc("起认出")}=0,"⚠ 起日期没认出，按建账日。","")&IF({sc("止认出")}=0,"⚠ 止日期没认出，按截止日。","")'
            f'&IF({sc("起选")}<P_建账日,"⚠ 起早于建账日 "&TEXT(P_建账日,"yyyy-mm-dd")&"，按建账日（建账前的数都在期初里）。","")'
            f'&IF({sc("止输入")}<P_建账日,"⚠ 止早于建账日 "&TEXT(P_建账日,"yyyy-mm-dd")&"，按建账日前一天（只有期初）。","")'
            f'&IF({sc("止输入")}>P_截止,"⚠ 止晚于截止日 "&TEXT(P_截止,"yyyy-mm-dd")&"，按截止日算。","")'
            f'&IF({sc("起选")}>{sc("止0")},"✗ 起晚于止，请改日期。","")'
            f'&IF({sc("笔数")}>{NL},"⚠ 共 "&{sc("笔数")}&" 笔，只列了前 {NL} 笔（合计是全部的），请缩短日期分几张对。","")'
            f'&IF(AND({U}>0,{sc("欠票")}<-0.005),"ℹ 截至止日发票比应开的多 "&TEXT(-{sc("欠票")},"#,##0.00")&" 元，对账单上没印'
            f'（多半是给建账前的货开的票：在【往来单位】填期初未开票/期初欠票）。","")')
    put(ws, 'A5', (f'=IF({warn}="","起空着＝建账日 "&TEXT(P_建账日,"yyyy-mm-dd")&"，止空着＝截止日 "&TEXT(P_截止,"yyyy-mm-dd")'
                   f'&"；往来单位空着＝第一家客户。下面第 6 行起是打印出来的对账单。",{warn})'), F_NOTE, align=ALW, border=False)
    ws.conditional_formatting.add('A5', FormulaRule(formula=['OR(LEFT($A$5,1)="⚠",LEFT($A$5,1)="✗")'],
                                                    font=Font(bold=True, color='FFC00000')))
    ws.row_dimensions[5].height = 30

    # ── 抬头 ──
    ws.merge_cells(f'A6:{LAST}6')
    put(ws, 'A6', (f'=IF({U}=0,IF({sc("名")}="","⚠【往来单位】表里还没有客户，先去加一家",'
                   f'"⚠ 「"&{sc("名")}&"」不在【往来单位】表里，先去【往来单位】加一行（简称、全称、其他叫法都认）"),'
                   f'IF(P_公司名称="","",P_公司名称&" 与 ")&{sc("抬头名")}&" 对账单")'), F_BIG, align=ACW, border=False)
    ws.conditional_formatting.add('A6', FormulaRule(formula=['LEFT($A$6,1)="⚠"'], font=Font(bold=True, color='FFC00000')))
    ws.row_dimensions[6].height = 40
    ws.merge_cells('A7:E7')
    put(ws, 'A7', f'=IF({U}=0,"","对账期间："&TEXT({S0},"yyyy-mm-dd")&" 至 "&TEXT({S1},"yyyy-mm-dd"))', F_HEADB, align=AL,
        border=False)
    ws.merge_cells(f'F7:{LAST}7')
    lx, dh = sc('联系人'), sc('电话')        # 没有的那半句不印，都没有就空着
    put(ws, 'F7', (f'=IF({U}=0,"",IF({lx}="","","对方联系人："&{lx})&IF(AND({lx}<>"",{dh}<>""),"　　","")'
                   f'&IF({dh}="","","电话："&{dh}))'), F_TXT, align=AR, border=False)
    ws.row_dimensions[7].height = 22
    E = sc('期末')
    amt = lambda x: f'TEXT({x},"#,##0.00")'
    big = (f'IF({U}=0,"","截至 "&TEXT({S1},"yyyy-mm-dd")&"，"&IF({D}=1,'
           f'IF({E}>0.005,"贵公司尚欠我公司 "&{amt(E)}&" 元",IF({E}<-0.005,"我公司预收贵公司 "&{amt("-" + E)}&" 元","双方两清")),'
           f'IF({E}>0.005,"我公司尚欠贵公司 "&{amt(E)}&" 元",IF({E}<-0.005,"我公司预付贵公司 "&{amt("-" + E)}&" 元","双方两清"))))')
    ws.merge_cells(f'A8:{LAST}8')
    put(ws, 'A8', '=' + big, F_BIG2, fill('FFFFF2CC'), align=ACW, border=False)
    ws.row_dimensions[8].height = 32
    ws.merge_cells(f'A9:{LAST}9')
    adj = f'ROUND({sc("调G")}-{sc("调H")},2)'
    put(ws, 'A9', (f'=IF({U}=0,"","上期结欠 "&{amt(sc("期初"))}&" ＋ 本期"&IF({D}=1,"发货","到货")&" "&{amt(sc("货"))}'
                   f'&" － 本期"&IF({D}=1,"回款","付款")&" "&{amt(sc("钱"))}'
                   f'&IF({adj}>0.005," ＋ 调账 "&{amt(adj)},IF({adj}<-0.005," － 调账 "&{amt("-" + adj)},""))'
                   f'&" ＝ 期末余额 "&{amt(E)}&" 元"'
                   f'&IF({sc("待定价")}>0,"；另有 "&{sc("待定价")}&" 笔"&IF({D}=1,"发货单价待定（金额没算）",'
                   f'"到货单价待定（按参考单价暂估，已算在里面）"),""))'),
        F_TXT, align=ALW, border=False)
    ws.row_dimensions[9].height = 20
    ws.merge_cells(f'A10:{LAST}10')
    q = sc('欠票')
    # 到止日发票比应开的多（q<0）不印给对方：多半是给建账前的货开的票（期初未开票/期初欠票没填），第 5 行打印区外用 ℹ 说
    put(ws, 'A10', (f'=IF({U}=0,"",IF({D}=1,"本期我公司开出发票 "&{amt(sc("票"))}&" 元"'
                    f'&IF({q}<-0.005,"","；截至 "&TEXT({S1},"yyyy-mm-dd")&"，还有 "&{amt(q)}&" 元没开票"),'
                    f'"本期收到贵公司发票 "&{amt(sc("票"))}&" 元"'
                    f'&IF({q}<-0.005,"","；截至 "&TEXT({S1},"yyyy-mm-dd")&"，贵公司还欠我公司发票 "&{amt(q)}&" 元")))'),
        F_TXT, align=ALW, border=False)
    ws.row_dimensions[10].height = 20

    # ── 表头、期初行 ──
    header(ws, ST_H, [('A', '日期'), ('B', '车牌／单号'), ('C', '物料或产品／摘要'), ('D', '吨'),
                      ('E', f'=IF({D}=1,"单价","单价／运费")'), ('F', '税点'),
                      ('G', f'=IF({D}=1,"发货金额","到货金额")'), ('H', f'=IF({D}=1,"回款金额","付款金额")'), ('I', '余额'),
                      ('J', '备注')], C_WL, height=30)
    r = ST_O
    for col in 'ABCDEFGHIJ':
        put(ws, f'{col}{r}', None, F_TXTB, fill('FFF2F2F2'), align=AC)
    put(ws, f'A{r}', f'=IF({U}=0,"","上期结欠")', F_TXTB, fill('FFF2F2F2'), align=AL)
    put(ws, f'C{r}', f'=IF({U}=0,"","截至 "&TEXT({sc("前界")},"yyyy-mm-dd"))', F_TXTB, fill('FFF2F2F2'), align=AL)
    put(ws, f'I{r}', f'=IF({U}=0,"",{sc("期初")})', F_TXTB, fill('FFF2F2F2'), MONEY, AR)

    # ── 明细（活动区：明细 → 本期合计 → 共几笔 → 签字） ──
    n, N_ = sc('n'), sc('笔数')
    ends = SEG_END
    for kk in range(ST_N):
        k, r = kk + 1, ST_D0 + kk
        t, key, s, i = (f'${HS[x]}{r}' for x in ('类型', '键', '源', '条'))
        ws[f'{HS["k"]}{r}'] = k
        ws[f'{HS["k"]}{r}'].font = F_HELP
        hset(ws, f'{HS["类型"]}{r}',
             f'IF({U}=0,0,IF({k}<={n},1,IF({k}={n}+1,2,IF({k}={n}+2,3,IF({k}={n}+4,4,IF({k}={n}+7,5,0))))))')
        hset(ws, f'{HS["键"]}{r}', f'IF({t}=1,SMALL({SEQ},{k}),0)')
        off = f'MOD({key},{M})'
        hset(ws, f'{HS["源"]}{r}', (f'IF({t}<>1,0,IF({off}<={ends[0]},1,IF({off}<={ends[1]},2,IF({off}<={ends[2]},3,'
                                   f'IF({off}<={ends[3]},4,IF({off}<={ends[4]},5,6))))))'))
        hset(ws, f'{HS["条"]}{r}', f'IF({s}=0,0,{off}-CHOOSE({s},{",".join(str(x) for x in SEG_OFF)}))')
        z = lambda nm: f'INDEX({nm},{i})'
        inp = lambda sh, cols, f: f'INDEX({inref(sh, cols[f])},{i}+1)&""'

        def by(d, dflt='""'):
            out = dflt
            for s_, e in sorted(d.items(), reverse=True):
                out = f'IF({s}={s_},{e},{out})'
            return out
        nm_sale = (f'IF({z("销_品号")}>0,INDEX(品档_名称,{z("销_品号")}),'
                   f'IF({z("销_料号")}>0,INDEX(料档_名称,{z("销_料号")}),""))')
        nm_pur = (f'IF({z("进_料号")}>0,INDEX(料档_名称,{z("进_料号")}),'
                  f'IF({z("进_品号")}>0,INDEX(品档_名称,{z("进_品号")}),""))')
        nz = lambda x: f'IF({x}=0,"",{x})'
        cell(ws, f'A{r}', f'IF({t}=1,INT({key}/{M}),IF({t}=2,"本期合计",IF({t}=3,"共 "&{N_}&" 笔","")))', fmt=DATE)
        cell(ws, f'B{r}', by({1: inp(SH_SALE, SALE_COLS, '车牌号'), 2: inp(SH_PUR, PUR_COLS, '车牌号'),
                              3: inp(SH_PUR, PUR_COLS, '车牌号'), 4: inp(SH_SALE, SALE_COLS, '车牌号'),
                              5: f'IF({z("资_编号键")}="","",MID({z("资_编号键")},2,100))',
                              6: f'"凭证组 "&{z("手_组号")}'}), align=AL)
        cash_lbl = f'IF({D}=1,IF({z("资_净额")}<0,"退款","回款"),IF({z("资_净额")}>0,"退回","付款"))'
        tail = (f'IF({t}=3,IF({N_}>{NL},"⚠ 只列了前 {NL} 笔",""),IF({t}=4,"我方（盖章）：",IF({t}=5,"日期：","")))')
        cell(ws, f'C{r}', by({1: nm_sale, 2: nm_pur, 3: nm_pur, 4: nm_sale,
                              5: f'{cash_lbl}&"（"&INDEX(账户_名称,{z("资_账户号")})&"）"', 6: '"调账（手工凭证）"'}, tail),
             align=AL)
        cell(ws, f'D{r}', by({1: nz(z('销_吨位')), 2: nz(z('进_数量')), 3: nz(z('进_数量')), 4: nz(z('销_吨位'))}),
             fmt=QTY, align=AR)
        cell(ws, f'E{r}', by({1: nz(z('销_单价')), 2: nz(z('进_单价')), 3: nz(z('进_运费')), 4: nz(z('销_运费'))}),
             fmt=PRICE, align=AR)
        cell(ws, f'F{r}', by({3: nz(z('进_税点')), 4: nz(z('销_税点'))}), fmt=MONEY, align=AR)
        mjg = f'IF({D}=1,{nz(z("手_借"))},{nz(z("手_贷"))})'
        mjh = f'IF({D}=1,{nz(z("手_贷"))},{nz(z("手_借"))})'
        cell(ws, f'G{r}', by({1: f'IF({z("销_待定价")}=1,"待定价",{z("销_收入含税")})', 2: z('进_货款计'), 3: z('进_运费计'),
                              4: z('销_运费计'), 6: mjg}, f'IF({t}=2,{sc("合G")},"")'), fmt=MONEY, align=AR)
        cell(ws, f'H{r}', by({5: f'IF({D}=1,{z("资_净额")},-{z("资_净额")})', 6: mjh}, f'IF({t}=2,{sc("合H")},"")'),
             fmt=MONEY, align=AR)
        cell(ws, f'I{r}', (f'IF({t}=1,ROUND({sc("期初")}+SUM(G${ST_D0}:G{r})-SUM(H${ST_D0}:H{r}),2),'
                           f'IF({t}=2,{E},""))'), fmt=MONEY, align=AR)
        bz_s, bz_p = inp(SH_SALE, SALE_COLS, '备注'), inp(SH_PUR, PUR_COLS, '备注')
        bz_c, bz_m, zy_m = inp(SH_CASH, CASH_COLS, '备注'), inp(SH_MJ, MJ_COLS, '备注'), inp(SH_MJ, MJ_COLS, '摘要')
        ht = inp(SH_SALE, SALE_COLS, '合同号')
        fp = z('销_发票号键')
        add = lambda x: f'IF({x}="","","；"&{x})'
        n_sale = (f'MID(IF({z("销_待定价")}=1,"；单价待定","")&IF({ht}="","","；合同 "&{ht})'
                  f'&IF({fp}="","","；发票 "&MID({fp},2,100))&{add(bz_s)},2,300)')
        n_pur = f'MID(IF({z("进_待定价")}=1,"；单价待定，按参考单价暂估","")&{add(bz_p)},2,300)'
        n_fr = f'"进货运费"&IF({z("进_供应商号")}<>{U},"·"&{z("进_供应商")},"")&{add(bz_p)}'
        n_sfr = f'"送货运费·"&{z("销_客户")}&{add(bz_s)}'
        n_cash = f'{z("资_显示摘要")}&{add(bz_c)}'
        n_mj = f'MID({add(zy_m)}&{add(bz_m)},2,300)'
        jt = (f'IF({t}=2,IF({sc("待定价")}>0,"另有 "&{sc("待定价")}&" 笔单价待定",""),'
              f'IF({t}=4,"对方确认（盖章）：",IF({t}=5,"日期：","")))')
        cell(ws, f'J{r}', by({1: n_sale, 2: n_pur, 3: n_fr, 4: n_sfr, 5: n_cash, 6: n_mj}, jt), align=AL)
    reg = f'A{ST_D0}:{LAST}{ST_D1}'
    t0 = f'${HS["类型"]}{ST_D0}'
    cf = ws.conditional_formatting
    cf.add(reg, FormulaRule(formula=[f'{t0}=2'], fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD, stopIfTrue=True))
    cf.add(reg, FormulaRule(formula=[f'{t0}=3'], font=Font(color='FF808080'), stopIfTrue=True))
    cf.add(reg, FormulaRule(formula=[f'OR({t0}=4,{t0}=5)'], font=Font(bold=True), stopIfTrue=True))
    cf.add(f'J{ST_D0}:J{ST_D1}', FormulaRule(formula=[f'AND({t0}=1,ISNUMBER(FIND("单价待定",$J{ST_D0})))'],
                                             font=Font(bold=True, color='FFC00000'), border=CF_BD, stopIfTrue=True))
    cf.add(f'G{ST_D0}:G{ST_D1}', FormulaRule(formula=[f'AND({t0}=1,$G{ST_D0}="待定价")'],
                                             font=Font(bold=True, color='FFC00000'), border=CF_BD, stopIfTrue=True))
    cf.add(reg, FormulaRule(formula=[f'{t0}=1'], border=CF_BD))
    for rr in range(ST_D0, ST_D1 + 1):
        ws.row_dimensions[rr].height = 18
    hide(ws, *[CL(i) for i in range(CI('AA'), CI(HS['条']) + 1)])
    finish(ws, LAST, f'{ST_H}:{ST_H}', 6, sc('末行'), ST_D1, 'A5', False)
    ws.print_options.horizontalCentered = True
    return sc


# ═══════════════════════════════ _序 ═══════════════════════════════
def build_seq(wb, sc):
    ws = wb[H_SORT]
    ws['A1'] = f'对账单排序键：日期×{M}＋偏移（销售、进货货款、进货运费、销售运费、资金、手工各一段）'
    ws['C1'] = f'发票清单键：日期×{M}＋（发票登记 n；销售 {N_INV}＋n）'
    ws['E1'] = f'拉货键：物流号×{M}＋日期（进货 1～{N_PUR}，销售往后）'
    for a in ('A1', 'C1', 'E1'):
        ws[a].font = F_HELP
    u, d, lo, hi = sc.ext('u'), sc.ext('dir'), sc.ext('lo'), sc.ext('hi')
    ix = lambda nm, n: f'INDEX({nm},{n})'
    for s, (nm, cap) in enumerate(SEGS, 1):
        off = SEG_OFF[s - 1]
        for n in range(1, cap + 1):
            if s in (1, 4):
                dt_, cond = ix('销_日期', n), (f'{ix("销_客户号", n)}={u},{ix("销_有效", n)}=1' if s == 1 else
                                              f'{ix("销_物流号", n)}={u},{ix("销_有效", n)}=1,{ix("销_运费计", n)}<>0')
                need = f'{d}=1' if s == 1 else f'{d}=2'
            elif s in (2, 3):
                dt_ = ix('进_日期', n)
                cond = (f'{ix("进_供应商号", n)}={u},{ix("进_有效", n)}=1' if s == 2 else
                        f'{ix("进_物流号", n)}={u},{ix("进_有效", n)}=1,{ix("进_运费计", n)}<>0')
                need = f'{d}=2'
            elif s == 5:
                dt_ = ix('资_日期', n)
                cond = f'{ix("资_单位号", n)}={u},{ix("资_有效", n)}=1,{ix("资_原对方一级", n)}=IF({d}=1,{K1122},{K2202})'
                need = 'TRUE'
            else:
                dt_ = ix('手_日期', n)
                cond = f'{ix("手_单位号", n)}={u},{ix("手_记账", n)}=1,{ix("手_一级号", n)}=IF({d}=1,{K1122},{K2202})'
                need = 'TRUE'
            k = f'{dt_}*{M}+{off + n}'
            pre = f'IF(OR({u}=0,NOT({need})),"",' if need != 'TRUE' else f'IF({u}=0,"",'
            ws[f'A{off + n + 1}'] = f'={pre}IF(AND({cond},{k}>={lo},{k}<{hi}),{k},""))'
    for n in range(1, N_INV + 1):
        ws[f'C{n + 1}'] = (f'=IF(AND({ix("票_有效", n)}=1,OR({ix("票_方向", n)}=2,{ix("票_销售重复", n)}=0)),'
                           f'{ix("票_日期", n)}*{M}+{n},"")')
    for n in range(1, N_SALE + 1):
        ws[f'C{N_INV + n + 1}'] = (f'=IF(AND({ix("销_有效", n)}=1,{ix("销_发票号键", n)}<>""),'
                                   f'{ix("销_日期", n)}*{M}+{N_INV + n},"")')
    for n in range(1, N_PUR + 1):
        ws[f'E{n + 1}'] = (f'=IF(AND({ix("进_有效", n)}=1,{ix("进_运费计", n)}>0,{ix("进_物流号", n)}>0),'
                           f'{ix("进_物流号", n)}*{M}+{ix("进_日期", n)},"")')
    for n in range(1, N_SALE + 1):
        ws[f'E{N_PUR + n + 1}'] = (f'=IF(AND({ix("销_有效", n)}=1,{ix("销_运费计", n)}>0,{ix("销_物流号", n)}>0),'
                                   f'{ix("销_物流号", n)}*{M}+{ix("销_日期", n)},"")')


# ═══════════════════════════════ 发票跟进 ═══════════════════════════════
def cf_blocks(ws, r0, r1, rules):
    """条件格式：rules＝[(起列, 止列, 公式(用本块第一列、第 r0 行写), 样式)]，按优先顺序。
       列切成「规则集合一样」的连续块，每格只落在一个块里（LibreOffice / 有的 WPS 对重叠的块只认一个）。"""
    lo = min(CI(a) for a, _b, _c, _k in rules)
    hi = max(CI(b) for _a, b, _c, _k in rules)
    runs = []
    for i in range(lo, hi + 1):
        sig = tuple(j for j, (a, b, _c, _k) in enumerate(rules) if CI(a) <= i <= CI(b))
        if not sig:
            continue
        if runs and runs[-1][2] == sig and runs[-1][1] == i - 1:
            runs[-1][1] = i
        else:
            runs.append([i, i, sig])
    for i0, i1, sig in runs:
        for j in sig:
            _a, _b, cond, kw = rules[j]
            ws.conditional_formatting.add(f'{CL(i0)}{r0}:{CL(i1)}{r1}',
                                          FormulaRule(formula=[cond(CL(i0))], stopIfTrue=True, **kw))


IV_LAST = 'J'
IV_R1, IV_H1, IV_H2, IV_M0 = 6, 7, 8, 9              # ① 标题、两行表头、1 月所在行
IV_TOT = IV_M0 + 12                                  # ① 合计行
FLOW_R0 = IV_TOT + 2                                 # ②③④ 从这一行起往下排
N2, N3, N4 = 150, 150, 120                           # 最多列几家 / 几张
FLOW_N = (N2 + 4) + 1 + (N3 + 4) + 1 + (N4 + 5)
IU0, IU1 = 3, 3 + N_UNIT - 1                         # 每家一行的排序辅助：第 3 行起
KC, KS = 'AE', 'AG'                                  # 客户键、供应商键
KC_RNG, KS_RNG = f'${KC}${IU0}:${KC}${IU1}', f'${KS}${IU0}:${KS}${IU1}'
FH = dict(段='AH', 行='AI', 类='AJ', 第几='AK', 源='AL')
TK_FLAG, TK_M0, TK_R0 = 'AN', 'AO', 3                # 票种 k（第 TK_R0+k-1 行）：可抵（税率>0）、1～12 月收到税额

TIP_IV = ('★ 全自动。开出的票：销售台账填了发票号的车（一车算一张，金额＝这一车含税额）＋发票登记里开出、号码不在销售台账里的；'
          '收到的票：发票登记（含补票）。B3 月份管 ④（填 1～12，也认 2026-9、202609；空着＝截止日期那个月）。'
          '② ③ 是截至截止日期的数：未开票＝期初未开票＋要开票的发货（票种不是无票）－已开票；欠票＝期初欠票＋要票的到货'
          '（货款票种、运费票种不是无票；物流单位的进货运费、销售运费也算）－收到的票；合计行的未开票、欠票只加正数（多开、多给的看第 5 行）。')


def build_inv(wb):
    ws = wb[SH_INVS]
    W = {'A': 27, 'B': 13.5, 'C': 14, 'D': 28, 'E': 17, 'F': 13.5, 'G': 13.5, 'H': 14, 'I': 16, 'J': 22}
    widths(ws, W)
    title(ws, '发 票 跟 进', IV_LAST, C_WL, TIP_IV)

    # ── 第 3 行：④ 月份 ──
    selector(ws, 'A3', '④ 月份', 'B3', None, '"1,2,3,4,5,6,7,8,9,10,11,12"',
             prompt='选 1～12（本年度）；空着＝截止日期那个月；填 2026-9 或 202609 也认')
    SC = ['月份输入', '清单月', '月份认出', '月初', '月末', '清单止', '清单前', '清单数', '清单显示', '开出张数', '收到张数',
          '客户数', '客户显示', '供应商数', '供应商显示', 'L2', 's3', 'L3', 's4', 'L4', '尾', '末行',
          '客应开', '客已开', '客未开', '客欠款', '供应收', '供已收', '供欠票', '供欠他',
          '开价税', '开税额', '开不含税', '收价税', '收税额', '收不含税']
    rows = {k: 3 + i for i, k in enumerate(SC)}
    S = lambda k: f'$AB${rows[k]}'
    mv = S('月份输入')
    win_p = lambda: f'票_日期,">="&{S("月初")},票_日期,"<="&{S("清单止")}'
    win_s = lambda: f'销_日期,">="&{S("月初")},销_日期,"<="&{S("清单止")}'
    out_p = '票_方向,1,票_销售重复,0,票_有效,1'
    in_p = '票_方向,2,票_有效,1'
    out_s = '销_发票号键,"#?*",销_有效,1'

    def unit_sum(col, flag, key, pos=False):      # pos：只加正数（多开、多给的不抵）
        extra = f',{col},">0"' if pos else ''
        return f'=ROUND(SUMIFS({col},位_有效,1,{flag},1,{key},">0"{extra}),2)'
    f = {
        '月份输入': '=' + ym_any('B3'),
        '月份认出': f'=IF({mv}=-1,1,IF(AND({ym_ok(mv)},INT({mv}/100)=P_年度),1,0))',
        '清单月': f'=IF(AND({mv}<>-1,{S("月份认出")}=1),MOD({mv},100),P_截止月)',
        '月初': f'=DATE(P_年度,{S("清单月")},1)',
        '月末': f'=DATE(P_年度,{S("清单月")}+1,0)',
        '清单止': f'=MIN({S("月末")},P_截止)',
        '清单前': f'=COUNTIF({IVK},"<"&{S("月初")}*{M})',
        '清单数': (f'=IF({S("清单止")}<{S("月初")},0,COUNTIFS({IVK},">="&{S("月初")}*{M},'
                 f'{IVK},"<"&({S("清单止")}+1)*{M}))'),
        '清单显示': f'=MIN({S("清单数")},{N4})',
        '开出张数': f'=COUNTIFS({out_p},{win_p()})+COUNTIFS({out_s},{win_s()})',
        '收到张数': f'=COUNTIFS({in_p},{win_p()})',
        '客户数': f'=COUNT({KC_RNG})',
        '客户显示': f'=MIN({S("客户数")},{N2})',
        '供应商数': f'=COUNT({KS_RNG})',
        '供应商显示': f'=MIN({S("供应商数")},{N3})',
        'L2': f'={S("客户显示")}+4',
        's3': f'={S("L2")}+1',
        'L3': f'={S("供应商显示")}+4',
        's4': f'={S("s3")}+{S("L3")}+1',
        'L4': f'={S("清单显示")}+5',
        '尾': f'={S("s4")}+{S("L4")}',
        '末行': f'={FLOW_R0}+{S("尾")}-1',
        '客应开': unit_sum('位_应开票', '位_是客户', KC_RNG),
        '客已开': unit_sum('位_已开票', '位_是客户', KC_RNG),
        '客未开': unit_sum('位_未开票', '位_是客户', KC_RNG, True),      # 只加正数（多开的不抵）＝ B5
        '客欠款': unit_sum('位_应收', '位_是客户', KC_RNG),
        '供应收': unit_sum('位_应收票', '位_是供应商', KS_RNG),
        '供已收': unit_sum('位_已收票', '位_是供应商', KS_RNG),
        '供欠票': unit_sum('位_欠票', '位_是供应商', KS_RNG, True),      # 只加正数（多给的不抵）＝ E5
        '供欠他': unit_sum('位_应付', '位_是供应商', KS_RNG),
        '开价税': f'=ROUND(SUMIFS(票_价税合计,{out_p},{win_p()})+SUMIFS(销_收入含税,{out_s},{win_s()}),2)',
        '开税额': f'=ROUND(SUMIFS(票_税额,{out_p},{win_p()})+SUMIFS(销_销项,{out_s},{win_s()}),2)',
        '开不含税': f'=ROUND(SUMIFS(票_不含税,{out_p},{win_p()})+SUMIFS(销_收入,{out_s},{win_s()}),2)',
        '收价税': f'=ROUND(SUMIFS(票_价税合计,{in_p},{win_p()}),2)',
        '收税额': f'=ROUND(SUMIFS(票_税额,{in_p},{win_p()}),2)',
        '收不含税': f'=ROUND(SUMIFS(票_不含税,{in_p},{win_p()}),2)',
    }
    for k in SC:
        ws[f'AA{rows[k]}'] = k
        ws[f'AB{rows[k]}'] = f[k]
        ws[f'AA{rows[k]}'].font = ws[f'AB{rows[k]}'].font = F_HELP
    ws.merge_cells('C3:I3')
    put(ws, 'C3', (f'="实际用：④ "&P_年度&" 年 "&{S("清单月")}&" 月"&IF(TRIM(B3&"")="","（空着＝截止日期那个月）","")'
                   f'&IF({S("月份认出")}=0,"　⚠ 月份没认出来或不是本年度（先按截止日期那个月）","")'), F_NOTE, align=ALW,
        border=False)
    ws.conditional_formatting.add('C3', FormulaRule(formula=[f'{S("月份认出")}=0'], font=F_RED))
    home_link(ws, f'{IV_LAST}3')
    ws.row_dimensions[3].height = 26

    # ── 第 4、5 行：KPI（截至截止日） ──
    cust = '位_有效,1,位_是客户,1,位_未开票'
    supp = '位_有效,1,位_是供应商,1,位_欠票'
    kpis = [('A', 'A', '截至日期', '=P_截止', DATE),
            ('B', 'C', f'="客户没开票（"&COUNTIFS({cust},">0.005")&" 家）"', f'=ROUND(SUMIFS(位_未开票,{cust},">0.005"),2)', MONEY),
            ('D', 'D', f'="多开给客户的票（"&COUNTIFS({cust},"<-0.005")&" 家）"', f'=ROUND(-SUMIFS(位_未开票,{cust},"<-0.005"),2)', MONEY),
            ('E', 'F', f'="供应商欠票（"&COUNTIFS({supp},">0.005")&" 家）"', f'=ROUND(SUMIFS(位_欠票,{supp},">0.005"),2)', MONEY),
            ('G', 'H', f'="供应商多给的票（"&COUNTIFS({supp},"<-0.005")&" 家）"', f'=ROUND(-SUMIFS(位_欠票,{supp},"<-0.005"),2)', MONEY)]
    for c1, c2, lbl, val, fmt in kpis:
        if c1 != c2:
            ws.merge_cells(f'{c1}4:{c2}4')
            ws.merge_cells(f'{c1}5:{c2}5')
        put(ws, f'{c1}4', lbl, F_KPI_L, LBL, align=ACW)
        put(ws, f'{c1}5', val, F_KPI_B, FILL_TOT, fmt, AC)
        if c1 != c2:
            put(ws, f'{c2}4', None, F_KPI_L, LBL)
            put(ws, f'{c2}5', None, F_KPI_B, FILL_TOT)
    ws.merge_cells('I4:J5')
    put(ws, 'I4', '多开、多给：票比货多，一般是给建账以前的货开的票（在【往来单位】填期初未开票、期初欠票）', F_NOTE, align=ALW,
        border=False)
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 26

    # ── 票种 × 月：收到的税额（可抵＝票种可抵税率>0） ──
    ws[f'{TK_FLAG}{TK_R0 - 1}'] = '可抵'
    ws[f'{TK_FLAG}{TK_R0 - 1}'].font = F_HELP
    for kt in range(1, N_TK + 1):
        r = TK_R0 + kt - 1
        hset(ws, f'{TK_FLAG}{r}', f'IF(N(INDEX(票种_税率,{kt}))>0,1,0)')
        for m in range(1, 13):
            hset(ws, f'{CL(CI(TK_M0) + m - 1)}{r}',
                 f'SUMIFS(票_税额,{in_p},票_月,{m},票_票种号,{kt},票_日期,"<="&P_截止)')
    tk_flag_rng = f'${TK_FLAG}${TK_R0}:${TK_FLAG}${TK_R0 + N_TK - 1}'

    # ── ① 各月 ──
    section(ws, IV_R1, 'A', 'I', '="① "&P_年度&" 年各月开出、收到的发票（开出按提货日期 / 开票日期；只算建账日～截止日期）"', C_WL)
    ws.row_dimensions[IV_R1].height = 20
    ws.merge_cells(f'A{IV_H1}:A{IV_H2}')
    ws.merge_cells(f'B{IV_H1}:D{IV_H1}')
    ws.merge_cells(f'E{IV_H1}:H{IV_H1}')
    ws.merge_cells(f'I{IV_H1}:I{IV_H2}')
    header(ws, IV_H1, [('A', '月份'), ('B', '开出（开给客户的）'), ('C', None), ('D', None), ('E', '收到（供应商、物流开来的）'),
                       ('F', None), ('G', None), ('H', None), ('I', '开出税额－收到专票税额\n（只供参考，不是报税数）')], C_WL, height=22)
    header(ws, IV_H2, [('A', None), ('B', '张数\n（销售一车算一张）'), ('C', '价税合计'), ('D', '税额'), ('E', '张数'),
                       ('F', '价税合计'), ('G', '税额'), ('H', '其中专票税额\n（票种可抵税率>0）'), ('I', None)], C_WL, height=42)
    cut_p, cut_s = '票_日期,"<="&P_截止', '销_日期,"<="&P_截止'
    cols = {
        'B': lambda m: f'COUNTIFS({out_p},票_月,{m},{cut_p})+COUNTIFS({out_s},销_月,{m},{cut_s})',
        'C': lambda m: f'ROUND(SUMIFS(票_价税合计,{out_p},票_月,{m},{cut_p})+SUMIFS(销_收入含税,{out_s},销_月,{m},{cut_s}),2)',
        'D': lambda m: f'ROUND(SUMIFS(票_税额,{out_p},票_月,{m},{cut_p})+SUMIFS(销_销项,{out_s},销_月,{m},{cut_s}),2)',
        'E': lambda m: f'COUNTIFS({in_p},票_月,{m},{cut_p})',
        'F': lambda m: f'ROUND(SUMIFS(票_价税合计,{in_p},票_月,{m},{cut_p}),2)',
        'G': lambda m: f'ROUND(SUMIFS(票_税额,{in_p},票_月,{m},{cut_p}),2)',
        'H': lambda m: (f'ROUND(SUMIF({tk_flag_rng},1,${CL(CI(TK_M0) + m - 1)}${TK_R0}:${CL(CI(TK_M0) + m - 1)}'
                        f'${TK_R0 + N_TK - 1}),2)'),
    }
    for m in range(1, 13):
        r = IV_M0 + m - 1
        hide_m = f'OR({m}<P_建账月,{m}>P_截止月)'
        put(ws, f'A{r}', f'{m}月', F_TXT, align=AC)
        for c, fn in cols.items():
            put(ws, f'{c}{r}', f'=IF({hide_m},"",{fn(m)})', F_AUTO, None, INT if c in 'BE' else MONEY, AR)
        put(ws, f'I{r}', f'=IF({hide_m},"",ROUND(D{r}-H{r},2))', F_AUTO, FILL_AUTO, MONEY, AR)
        ws.conditional_formatting.add(f'A{r}', FormulaRule(formula=[hide_m], font=Font(color='FFA6A6A6')))
    put(ws, f'A{IV_TOT}', '合计', F_TXTB, FILL_TOT, align=AC)
    for c in 'BCDEFGHI':
        put(ws, f'{c}{IV_TOT}', f'=ROUND(SUM({c}{IV_M0}:{c}{IV_M0 + 11}),2)', F_AUTOB, FILL_TOT, INT if c in 'BE' else MONEY, AR)

    # ── 每家一行的排序辅助（隐藏）：客户按未开票、供应商按欠票，从大到小 ──
    for c, t in (('AD', '客户未开票'), ('AE', '客户键'), ('AF', '供应商欠票'), ('AG', '供应商键')):
        ws[f'{c}{IU0 - 1}'] = t
        ws[f'{c}{IU0 - 1}'].font = F_HELP
    for i in range(N_UNIT):
        r, n = IU0 + i, i + 1
        for vc, kc, flag, val in (('AD', 'AE', '位_是客户', '位_未开票'), ('AF', 'AG', '位_是供应商', '位_欠票')):
            ws[f'{vc}{r}'] = f'=IF(AND(INDEX({flag},{n})=1,ABS(INDEX({val},{n}))>0.005),INDEX({val},{n}),"")'
            ws[f'{kc}{r}'] = f'=IF({vc}{r}="","",COUNTIF(${vc}${IU0}:${vc}${IU1},">"&{vc}{r})*1000+{n})'
            ws[f'{vc}{r}'].font = ws[f'{kc}{r}'].font = F_HELP

    # ── ②③④ 往下排 ──
    for k, c in FH.items():
        ws[f'{c}{FLOW_R0 - 1}'] = k
        ws[f'{c}{FLOW_R0 - 1}'].font = F_HELP
    flow_cols = 'ABCDEFGHIJ'
    # 同一列在 ②③ 和 ④ 里有的是金额、有的是文字：金额用带「* 」填充的格式（数字总是靠右），文字按列的对齐
    MONEY_R = '* #,##0.00;[Red]* -#,##0.00;* "-";@'
    fm = {'A': DATE, 'B': MONEY_R, 'C': MONEY_R, 'D': MONEY_R, 'E': MONEY_R, 'F': DATE, 'G': MONEY_R, 'H': MONEY_R, 'I': MONEY_R}
    al = {'A': 'left', 'B': 'center', 'C': 'center', 'D': 'center', 'E': 'center', 'F': 'center', 'G': 'center', 'H': 'center',
          'I': 'center', 'J': 'left'}
    Ux = lambda nm, ix: f'INDEX({nm},{ix})'
    neg = lambda val, word, extra='': f'IF({val}<-0.005,"{word} "&TEXT(-{val},"#,##0.00")&"{extra}",{val})'
    cutoff = 'TEXT(P_截止,"yyyy-mm-dd")'
    srate = 'IF(P_纳税人=1,P_销项税率,P_征收率)'
    rtxt = lambda x: f'IF(MOD(ROUND({x}*1000,0),10)=0,TEXT({x},"0%"),TEXT({x},"0.0%"))'
    ap_last = ap_last_arrive_col()
    for j in range(FLOW_N):
        r = FLOW_R0 + j
        seg, o, kd, k, ix = (f'${FH[x]}{r}' for x in ('段', '行', '类', '第几', '源'))
        ws[f'{FH["段"]}{r}'] = f'=IF({j}<{S("s3")},2,IF({j}<{S("s4")},3,IF({j}<{S("尾")},4,0)))'
        ws[f'{FH["行"]}{r}'] = f'=IF({seg}=2,{j},IF({seg}=3,{j}-{S("s3")},IF({seg}=4,{j}-{S("s4")},0)))'
        n23 = f'IF({seg}=2,{S("客户显示")},{S("供应商显示")})'
        n4 = S('清单显示')
        ws[f'{FH["类"]}{r}'] = (f'=IF({seg}=0,0,IF({o}=0,1,IF({o}=1,2,IF({o}=2,3,IF({seg}=4,'
                               f'IF({o}=3,4,IF({o}<4+{n4},5,IF({o}=4+{n4},6,0))),'
                               f'IF({o}<3+{n23},5,IF({o}=3+{n23},6,0)))))))')
        ws[f'{FH["第几"]}{r}'] = f'=IF({kd}=5,IF({seg}=4,{o}-3,{o}-2),0)'
        ws[f'{FH["源"]}{r}'] = (f'=IF({k}=0,0,IF({seg}=2,MOD(SMALL({KC_RNG},{k}),1000),IF({seg}=3,MOD(SMALL({KS_RNG},{k}),1000),'
                               f'MOD(SMALL({IVK},{S("清单前")}+{k}),{M}))))')
        for c in FH.values():
            ws[f'{c}{r}'].font = F_HELP
        # ④ 一张票：源 ≤ N_INV ＝ 发票登记第 源 条；否则 ＝ 销售台账第 源－N_INV 条
        isp = f'{ix}<={N_INV}'
        sx = f'{ix}-{N_INV}'
        pv = lambda p_, s_: f'IF({isp},{p_},{s_})'
        tkn = lambda no: f'IF({no}>0,INDEX(票种_名称,{no}),"")'
        p_dz = f'INDEX({inref(SH_INV, INV_COLS["对应单号"])},{ix}+1)&""'
        p_dk = Ux('票_垫付键', ix)
        p_j = f'{p_dz}&IF({p_dk}="","",IF({p_dz}="","","；")&"补垫付 "&MID({p_dk},2,100))'
        s_ht = f'INDEX({inref(SH_SALE, SALE_COLS["合同号"])},{sx}+1)&""'
        s_cp = f'INDEX({inref(SH_SALE, SALE_COLS["车牌号"])},{sx}+1)&""'
        s_j = f'"销售台账"&IF({s_ht}="",""," "&{s_ht})&IF({s_cp}="",""," "&{s_cp})'
        p_rate = f'INDEX(票种_税率,{Ux("票_票种号", ix)})'
        item = {
            'A': (Ux('位_简称', ix), Ux('位_简称', ix), pv(Ux('票_日期', ix), Ux('销_日期', sx))),
            'B': (Ux('位_应开票', ix), Ux('位_应收票', ix), pv(f'IF({Ux("票_方向", ix)}=1,"开出","收到")', '"开出"')),
            'C': (Ux('位_已开票', ix), Ux('位_已收票', ix), pv(Ux('票_单位', ix), Ux('销_客户', sx))),
            'D': (neg(Ux('位_未开票', ix), '多开', '（含建账前的货）'), neg(Ux('位_欠票', ix), '多给'),
                  pv(f'MID({Ux("票_号码键", ix)},2,100)', f'MID({Ux("销_发票号键", sx)},2,100)')),
            'E': (neg(Ux('位_应收', ix), '预收'), neg(Ux('位_应付', ix), '预付'),
                  pv(tkn(Ux('票_票种号', ix)), tkn(Ux('销_票种号', sx)))),
            'F': (f'IF({Ux("位_最后发货", ix)}>0,{Ux("位_最后发货", ix)},"")',
                  f'IF({Ux(ap_last, ix)}>0,{Ux(ap_last, ix)},"")',
                  pv(f'IF({Ux("票_票种号", ix)}>0,{rtxt(p_rate)},"")', rtxt(srate))),
            'G': (Ux('位_联系人', ix), Ux('位_联系人', ix), pv(Ux('票_价税合计', ix), Ux('销_收入含税', sx))),
            'H': (Ux('位_电话', ix), Ux('位_电话', ix), pv(Ux('票_税额', ix), Ux('销_销项', sx))),
            'I': ('""', '""', pv(Ux('票_不含税', ix), Ux('销_收入', sx))),
            'J': ('""', '""', pv(p_j, s_j)),
        }
        title_ = {
            'A': ('"② 客户开票情况"', '"③ 供应商、物流欠票"', f'"④ "&P_年度&"年"&{S("清单月")}&"月发票清单"'),
            'D': (f'"截至 "&{cutoff}', f'"截至 "&{cutoff}', '"按日期排"'),
            'E': ('"没开票多的在前"', '"欠票多的在前"', '""'),
        }
        head = {
            'A': ('"客户"', '"供应商、物流"', '"日期"'), 'B': ('"应开票"', '"应收票"', '"开出/收到"'),
            'C': ('"已开票"', '"已收票"', '"往来单位"'), 'D': ('"未开票"', '"欠票"', '"发票号码"'),
            'E': ('"现在欠款"', '"现在欠他"', '"票种"'), 'F': ('"最后发货"', '"最后到货"', '"税率"'),
            'G': ('"联系人"', '"联系人"', '"价税合计"'), 'H': ('"电话"', '"电话"', '"税额"'),
            'I': ('""', '""', '"不含税"'), 'J': ('""', '""', '"来源／对应单号"'),
        }
        tot = {
            'A': (f'"合计（全部 "&{S("客户数")}&" 家，不含多开）"', f'"合计（全部 "&{S("供应商数")}&" 家，不含多给）"',
                  f'"开出合计（全部 "&{S("开出张数")}&" 张）"'),
            'B': (S('客应开'), S('供应收'), '""'), 'C': (S('客已开'), S('供已收'), '""'),
            'D': (S('客未开'), S('供欠票'), '""'), 'E': (S('客欠款'), S('供欠他'), '""'),
            'G': ('""', '""', S('开价税')), 'H': ('""', '""', S('开税额')), 'I': ('""', '""', S('开不含税')),
        }
        tot2 = {'A': f'"收到合计（全部 "&{S("收到张数")}&" 张）"', 'G': S('收价税'), 'H': S('收税额'), 'I': S('收不含税')}
        trail = {
            'A': (f'IF({S("客户数")}=0,"客户的票都开齐了","共 "&{S("客户数")}&" 家")',
                  f'IF({S("供应商数")}=0,"供应商的票都收齐了","共 "&{S("供应商数")}&" 家")',
                  f'IF({S("清单数")}=0,IF({S("月初")}>P_截止,"这个月在截止日期以后","这个月没有发票"),"共 "&{S("清单数")}&" 张")'),
            'D': (f'IF({S("客户数")}>{N2},"只列了前 {N2} 家（合计是全部的）","")',
                  f'IF({S("供应商数")}>{N3},"只列了前 {N3} 家（合计是全部的）","")',
                  f'IF({S("清单数")}>{N4},"只列了前 {N4} 张（合计是全部的）","")'),
        }

        def by_seg(t):
            if t is None:
                return '""'
            if isinstance(t, str):
                return t
            a, b, c = t
            if a == b == c:
                return a
            if a == b:
                return f'IF({seg}=4,{c},{a})'
            return f'IF({seg}=2,{a},IF({seg}=3,{b},{c}))'

        for c in flow_cols:
            parts = [(5, by_seg(item.get(c))), (1, by_seg(title_.get(c))), (2, by_seg(head.get(c))), (3, by_seg(tot.get(c))),
                     (4, by_seg(tot2.get(c))), (6, by_seg(trail.get(c)))]
            parts = [(x, e) for x, e in parts if e != '""']
            out = '""'
            for x, e in reversed(parts):
                out = f'IF({kd}={x},{e},{out})'
            cl_ = ws[f'{c}{r}']
            cl_.value = '=' + out
            cl_.font = F_TXT
            cl_.alignment = Alignment(horizontal=al[c], vertical='center', shrink_to_fit=(c in 'ACDEJ'))
            if c in fm:
                cl_.number_format = fm[c]
    r0, r1 = FLOW_R0, FLOW_R0 + FLOW_N - 1
    K = lambda: f'${FH["类"]}{r0}'
    SG = lambda: f'${FH["段"]}{r0}'
    st_title = dict(fill=cf_fill(C_WL), font=Font(bold=True, color='FFFFFFFF'))
    st_head = dict(fill=cf_fill('FFF8CBAD'), font=Font(bold=True, color='FF000000'), border=CF_BD)
    st_tot = dict(fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD)
    seg4 = lambda cond: (lambda c: f'AND({cond},{SG()}=4)')
    cf_blocks(ws, r0, r1, [       # ②③ 只用到 H 列，I、J 只给 ④
        ('A', 'H', lambda c: f'{K()}=1', st_title),
        ('I', IV_LAST, seg4(f'{K()}=1'), st_title),
        ('A', 'H', lambda c: f'{K()}=2', st_head),
        ('I', IV_LAST, seg4(f'{K()}=2'), st_head),
        ('A', 'H', lambda c: f'OR({K()}=3,{K()}=4)', st_tot),
        ('I', IV_LAST, seg4(f'OR({K()}=3,{K()}=4)'), st_tot),
        ('A', IV_LAST, lambda c: f'{K()}=6', dict(font=Font(color='FF808080'))),
        ('D', 'D', lambda c: f'AND({K()}=5,{SG()}<4,LEFT($D{r0},1)="多")', dict(font=Font(bold=True, color='FFC65911'), border=CF_BD)),
        ('E', 'E', lambda c: f'AND({K()}=5,{SG()}<4,LEFT($E{r0},1)="预")', dict(font=Font(bold=True, color='FF2F75B5'), border=CF_BD)),
        ('B', 'B', lambda c: f'AND({K()}=5,{SG()}=4,$B{r0}="收到")', dict(font=Font(color='FFC65911'), border=CF_BD)),
        ('A', 'H', lambda c: f'{K()}=5', dict(border=CF_BD)),
        ('I', IV_LAST, seg4(f'{K()}=5'), dict(border=CF_BD)),
    ])

    hide(ws, *[CL(i) for i in range(CI('AA'), CI(TK_M0) + 12)])
    ws.freeze_panes = 'A4'
    print_setup(ws, None, landscape=True)
    ws.oddFooter.center.text = '第 &P 页 共 &N 页'
    q = f"'{ws.title}'"
    ws.defined_names['Print_Area'] = DefinedName(
        'Print_Area', attr_text=f'{q}!$A$1:INDEX({q}!${IV_LAST}$1:${IV_LAST}${r1},{q}!{S("末行")})')
    return S


def build(wb, ctx=None):
    build_party(wb[SH_AR], True)
    build_party(wb[SH_AP], False)
    sc = build_stmt(wb[SH_STMT])
    build_seq(wb, sc)
    build_inv(wb)
