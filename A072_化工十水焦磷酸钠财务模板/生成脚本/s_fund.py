# -*- coding: utf-8 -*-
"""A072 化工 · 查看表（绿）：资金日报、资金月报、费用统计、报销统计、个人垫付现金明细。
只用 layout 的定义名称和本表格子取数（个人垫付现金明细的「备注」另外按第 n 条取资金台帐备注列，插行删行不错位）。

口径（跟 A069 不一样的地方）：
- 可用资金只算现金账户（账户_现金＝1：现金、银行、微信支付宝）；票据（承兑）单列「手上票据」；垫付、备用金账户单列，不进合计。
- 垫付账户的「余额」＝ －欠垫付人的钱：建账前 ⑦ 性质＝垫付 的（出借人＝这个人）＋ 垫付账户自己的行（支出＝垫了钱）
  ＋ 别的账户里「挂到这个人身上」的行（原对方号＝224101、往来＝垫付人：公司还他的钱算这个账户的收入）。
  备用金账户同理：⑧ 期初 122103-这个人 ＋ 别的账户原对方号＝122103、往来＝这个人 的行（发备用金＝这个账户收入，退回＝支出）＋ 自己花的。
  同一个人有两个同类账户时，挂过来的行只算在第一个账户上。
- 「实际收支」只算现金账户，不含归类＝内部转账（现金账户之间倒钱）；承兑到期/贴现入账 归类＝收货款，银行那边算实际收入。
- 月份一律是账套年度里的 1～12（资_月）；资金月报、个人垫付的年月只认本年度。
- 费用统计、报销统计按「费用月」（资_费用月），只算截止日以前付的（资_日期≤P_截止），费用月在截止月以后的不算（C3/提示里另说一句）；
  不走钱的几行（工资计提、折旧、摊销、利息计提、补票冲减）也只到截止月。

【资金日报】C3 日期（空＝截止日）。第 4 行固定格：C4 当天可用资金、E4 手上票据、G4 当天实际收入、I4 当天实际支出。
   ① 各账户（第 7 行起 N_ACC＋3 行）：按 现金→票据→垫付→备用金 排，后面 可用资金合计、手上票据、全公司实际收支。
   ② 当天每一笔（按 资_排序键）。
【资金月报】B3 年月（空＝截止月）。第 4 行：B4 月末可用资金、D4 月末手上票据、F4 本月实际收入、H4 本月实际支出、J4 本月净现金流。
   ① 各账户；② 全年各月按归类汇总（收货款、付货款、费用、其他收入、长期资产、借款投资押金垫付还款；现金账户）；③ 按收支类别（现金账户）。
【费用统计】左：按收支类别（只列归类∈EXPENSE_KINDS＋长期资产、以及 @报销 类别）×1～12 月＋全年，「员工报销」下面展开报销项目（有数的）；
   下面不走钱的：工资计提（4 个部门）、折旧、建厂费用/预付摊销、利息计提、补票冲减；小计、合计。B4 全年合计。
   右：R3 类别（空＝左边这些类别）、U3 月份（空＝全年）→ 每一笔（按 资_排序键，150 笔）。
【报销统计】B3 起月、D3 止月（空＝1～截止月）；报销人 × 报销项目（有数的项目，最多 15 列，多的并「其余项目」）、合计、其中老板垫付的、
   其中备用金的；右边每人 1～12 月（全年）。
【个人垫付现金明细】C3 垫付账户（下拉只列类型＝垫付的；空＝第一个）、F3 年月（空＝截止月）；第 5 行起打印区（照 39.png）；
   已报/未报：这个人到截止日公司已还的钱，按先进先出先冲 ⑦ 建账前垫付、再按 资_排序键 冲垫付账户的行。
   右边（打印区外）：补票跟进（待补票、没被发票登记补掉的，按天数从大到小，最多 100 行）、无票合计。
隐藏列：各表右边；每条资金一行的键（N_CASH 行）：资金月报 CX、费用统计 CX/CZ、报销统计 DX/DY、个人垫付 EX/EY（各自前一列是 n 常量）；
Sc 标量 名字/值 两列。"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break
from common import *
from layout import *                      # layout 的 norm/esc（去不换行空格）盖过 common 的
from s_mirror import unit_match

M = SORT_M
GREEN_H = 'FF70AD47'
_thin = Side(style='thin', color='FFBFBFBF')
CF_BD = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
F_KV = Font(name=YH, sz=10, bold=True, color='FFC00000')
LBL = fill('FFD9E1F2')
MONEY_B = '#,##0.00;[Red]-#,##0.00;""'
DAY_SHOW, FEE_SHOW, ADV_SHOW, BP_SHOW = 80, 150, 150, 100
RE_SLOTS = 15                              # 报销统计：报销项目最多列 15 个，多的并「其余项目」
NOT_XFER = '资_归类,"<>内部转账"'
CASH_OK = '资_有效,1,资_现金,1'
FEE_KINDS_ALL = EXPENSE_KINDS + ['长期资产']
KH224101, KH122103 = kh('224101'), kh('122103')


def cf_fill(rgb):
    return PatternFill('solid', fgColor=rgb, bgColor=rgb)


def R2(x):
    return f'ROUND({x},2)'


class Sc:
    """隐藏标量：cl 列名字、cv 列值（第 r0 行起）"""

    def __init__(self, ws, keys, r0=3, cl='AA', cv='AB'):
        self.ws, self.cv = ws, cv
        self.r = {k: r0 + i for i, k in enumerate(keys)}
        for k, r in self.r.items():
            ws[f'{cl}{r}'] = k
            ws[f'{cl}{r}'].font = F_HELP

    def __call__(self, k):
        return f'${self.cv}${self.r[k]}'

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


def note(ws, coord, last, f, height=None):
    r = coord.lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
    ws.merge_cells(f'{coord}:{last}{r}')
    put(ws, coord, f, F_NOTE, align=ALW, border=False)
    if height:
        ws.row_dimensions[int(r)].height = height


def kpi(ws, lbl_cell, lbl, val_cell, f, fmt=MONEY, merge_to=None):
    put(ws, lbl_cell, lbl, F_KPI_L, LBL, align=AC)
    put(ws, val_cell, '=' + f, F_KV, FILL_AUTO, fmt, AR)
    if merge_to:
        ws.merge_cells(f'{val_cell}:{merge_to}')


def pick(idx, col, r0, n, choices):
    """idx>0：取隐藏表第 idx 行；idx<0：CHOOSE(-idx, 合计行的值…)；0：空"""
    return (f'IF({idx}=0,"",IF({idx}>0,INDEX(${col}${r0}:${col}${r0 + n - 1},{idx}),'
            f'CHOOSE(-{idx},{",".join(choices)})))')


def mhead(ws, row, c1, c2):
    ws.merge_cells(f'{c1}{row}:{c2}{row}')
    for i in range(CI(c1), CI(c2) + 1):
        ws.cell(row=row, column=i).border = BD


def finish(ws, last, hdr_rows, end_cell, cap_row, freeze='A5', brk=None, extra=(), first='A', top=1, landscape=True):
    """打印：A4、一页宽、页脚；打印区域按内容变长（first 列 top 行 到 last 列 end_cell 行）。
       brk：在这一行上面强制分页；extra：再加几块打印区域 (首列, 末列, 容量行, 末行格)"""
    if freeze:
        ws.freeze_panes = freeze
    print_setup(ws, hdr_rows, landscape=landscape)
    ws.oddFooter.center.text = '第 &P 页 共 &N 页'
    if brk:
        ws.row_breaks.append(Break(id=brk - 1))
    q = f"'{ws.title}'"
    areas = [f'{q}!${first}${top}:INDEX({q}!${last}$1:${last}${cap_row},{q}!{end_cell})']
    areas += [f'{q}!${c0}$1:INDEX({q}!${c1}$1:${c1}${cr},{q}!{ec})' for c0, c1, cr, ec in extra]
    ws.defined_names['Print_Area'] = DefinedName('Print_Area', attr_text=','.join(areas))


def ym_parse(sc, cell_, default):
    """年月黄格 → sc 里 输入、认出前（年×100＋月；看不懂 0）、m（本年度的月；空/看不懂/别的年 → default）、认出（0＝看不懂或不是本年度）"""
    t = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM({cell_}&""),"月",""),"年","-"),"/","-")'
    sc.set('输入', f'IF(ISNUMBER({cell_}),{cell_},IFERROR(--{t},0))')
    v = sc('输入')
    sc.set('认出前', (f'IF(AND({v}>=1,{v}<=12,INT({v})={v}),P_年度*100+{v},'
                    f'IF(AND({v}>=200001,{v}<=209912,INT({v})={v},MOD({v},100)>=1,MOD({v},100)<=12),{v},'
                    f'IF(AND({v}>=36526,{v}<73051),YEAR({v})*100+MONTH({v}),0)))'))
    y = sc('认出前')
    sc.set('认出', f'IF(TRIM({cell_}&"")="",1,IF(AND({y}>0,INT({y}/100)=P_年度),1,0))')
    sc.set('m', f'IF({sc("认出")}=0,{default},IF(TRIM({cell_}&"")="",{default},MOD({y},100)))')


# ═══════════════════════════════ 账户表（日报、月报共用） ═══════════════════════════════
AC_ = dict(名称='AD', 类型='AE', 现金='AF', 票据='AG', 显示='AN', 排序='AO', 人='AQ', 首='AR', 期初='AS', 挂='AT', 类型字='AU',
           人字='AV', 说明='AW')


def acct_table(ws, r0):
    """第 a 个账户（第 r0+a-1 行）：名称、显示（有名字、类型认得、第一次出现）、类型号、现金、票据、人（垫付人/备用金人，规范后）、
       首（同类同人的第一个账户）、期初（建账时）、挂（挂过来的行的原对方号：224101/122103；0）、排序键（组×100＋a）、类型字、人字"""
    C = AC_
    for i in range(N_ACC):
        a, r = i + 1, r0 + i
        c = lambda k: f'{C[k]}{r}'
        nm, sh, ty, per, fs = c('名称'), c('显示'), c('类型'), c('人'), c('首')
        hset(ws, nm, f'INDEX(账户_名称,{a})&""')
        hset(ws, sh, f'IF({nm}="",0,IF(INDEX(账户_类型号,{a})=0,0,IF(MATCH({esc(nm)},账户_名称,0)={a},1,0)))')
        hset(ws, ty, f'IF({sh}=1,INDEX(账户_类型号,{a}),0)')
        hset(ws, c('现金'), f'IF(AND({ty}>=1,{ty}<=3),1,0)')
        hset(ws, c('票据'), f'IF({ty}=4,1,0)')
        pn = norm(f'INDEX(账户_人员,{a})')
        hset(ws, per, f'IF(OR({ty}=5,{ty}=6),{unit_match(pn, "简称", pn)},"")')
        hset(ws, fs, (f'IF(AND(OR({ty}=5,{ty}=6),{per}<>""),IF(COUNTIFS(${C["类型"]}${r0}:{ty},{ty},'
                      f'${C["人"]}${r0}:{per},{esc(per)})=1,1,0),0)'))
        lo = f'INDEX(账户_出借人号,{a})'
        op_dz = f'IF({lo}>0,-SUMIFS(借期_金额,借期_出借人号,{lo},借期_性质,3,借期_有效,1),0)'
        op_by = (f'SUMIFS(期初_借,期初_科目号,{KH122103},期初_明细,{esc(per)},期初_有效,1)'
                 f'-SUMIFS(期初_贷,期初_科目号,{KH122103},期初_明细,{esc(per)},期初_有效,1)')
        hset(ws, c('期初'), (f'IF({sh}<>1,0,ROUND(IF({ty}<=4,INDEX(账户_期初,{a}),IF({fs}<>1,0,IF({ty}=5,{op_dz},{op_by}))),2))'))
        hset(ws, c('挂'), f'IF({fs}=1,IF({ty}=5,{KH224101},{KH122103}),0)')
        hset(ws, c('排序'), f'IF({sh}=1,IF({ty}<=3,1,{ty}-2)*100+{a},"")')
        who = f'IF({per}="",{nm},{per})'
        hset(ws, c('人字'), f'IF({sh}=1,{who},"")')
        hset(ws, c('类型字'), (f'IF({sh}<>1,"",CHOOSE({ty},"现金","银行","微信支付宝","票据（承兑）","垫付·"&{who},'
                             f'"备用金·"&{who}))'))


def flow(a, r, fld, conds, C=AC_):
    other = {'收入': '支出', '支出': '收入', '净额': '净额'}[fld]
    sg = '-' if fld == '净额' else '+'
    own = f'SUMIFS(资_{fld},资_账户号,{a},{conds})'
    lk = (f'IF(${C["挂"]}{r}>0,SUMIFS(资_{other},资_原对方号,${C["挂"]}{r},资_账户类型,"<>"&${C["类型"]}{r},'
          f'资_单位,{esc("$" + C["人"] + str(r))},{conds}),0)')
    return f'{own}{sg}{lk}'


def acct_text(ws, r, bal):
    """垫付、备用金账户的一句话（第 r 行；bal＝余额格）"""
    C = AC_
    ty, who = f'${C["类型"]}{r}', f'${C["人字"]}{r}'
    hset(ws, f'{C["说明"]}{r}', (f'IF({ty}=5,IF(-{bal}>0.005,"欠 "&{who}&" 垫付 "&TEXT(-{bal},"#,##0.00")&" 元",'
                                 f'IF({bal}>0.005,{who}&" 的垫付多还了 "&TEXT({bal},"#,##0.00")&" 元","不欠 "&{who}&" 垫付")),'
                                 f'IF({ty}=6,{who}&" 手上备用金 "&TEXT({bal},"#,##0.00")&" 元",""))'))


def acct_rows_disp(ws, r0, nrow, sc_na, extra_idx):
    """显示行第 k 行的账户序号（AP 列）：前 na 行按排序键取账户，后面接合计行（负数）"""
    keys = f'${AC_["排序"]}${r0}:${AC_["排序"]}${r0 + N_ACC - 1}'
    for k in range(1, nrow + 1):
        r = r0 + k - 1
        tail = '0'
        for j in reversed(range(extra_idx)):
            tail = f'IF({k}={sc_na}+{j + 1},-{j + 1},{tail})'
        hset(ws, f'AP{r}', f'IF({k}<={sc_na},MOD(SMALL({keys},{k}),100),{tail})')


# ═══════════════════════════════ 资金日报 ═══════════════════════════════
TIP_DAY = ('★ 每天看一眼钱。黄格选日期（空着＝截止日，截止日在首页改）。① 每个账户：昨日余额、今天收了多少、付了多少、今日余额，'
           '还有这个月 1 号到这天的收支（都含内部转账，余额才接得上）。账户按 现金/银行/微信 → 承兑汇票 → 老板垫付 → 备用金 排；'
           '合计：可用资金（只算现金、银行、微信支付宝）、手上票据（承兑，不能直接花）、全公司当天实际收支（只算现金账户，不含内部转账；承兑到期、贴现进银行算收入）。'
           '垫付账户的余额是负数＝公司欠垫付人的（建账前垫的＋垫的－公司已还的；公司在银行还他的钱算这个账户的收入）；'
           '备用金账户的余额＝这个人手上还有的备用金（领的－退的－花的）。② 当天每一笔，同一天按录入顺序，带这一笔之后本账户的余额。')

D_A0, D_N1 = 7, N_ACC + 3
D_NOTE1 = D_A0 + D_N1
D_SEC2, D_NOTE2, D_H2 = D_NOTE1 + 1, D_NOTE1 + 2, D_NOTE1 + 3
D_L0 = D_H2 + 1
D_NL = DAY_SHOW + 2


def build_day(ws):
    LAST = 'I'
    widths(ws, {'A': 5.5, 'B': 16, 'C': 17, 'D': 14, 'E': 24, 'F': 13.5, 'G': 13.5, 'H': 14, 'I': 13.5})
    title(ws, '=IF(P_公司名称="","",P_公司名称&"　")&"资金日报"', LAST, C_VIEW, TIP_DAY)
    selector(ws, 'B3', '选日期', 'C3', None, fmt=DATE)
    dv_date(ws, 'C3')
    home_link(ws, f'{LAST}3')
    ws.row_dimensions[3].height = 30

    keys = ['d', '认出', '月初', '账户数', '笔数', '前面', '显示数', '末行', '内转净', '垫付说明',
            '全_今收', '全_今支', '全_月收', '全_月支', '全天收入', '全天支出']
    for p in ('可用', '票据'):
        keys += [f'{p}_{k}' for k in ('昨', '今收', '今支', '今余', '月收', '月支')]
    sc = Sc(ws, keys)
    d, m0 = sc('d'), sc('月初')
    sc.set('d', f'IF(AND(ISNUMBER(C3),C3>=36526,C3<73051),INT(C3),P_截止)')
    sc.set('认出', 'IF(TRIM(C3&"")="",1,IF(AND(ISNUMBER(C3),C3>=36526,C3<73051),1,0))')
    sc.set('月初', f'DATE(YEAR({d}),MONTH({d}),1)')
    sc.set('笔数', f'COUNTIFS(资_排序键,">="&{d}*{M},资_排序键,"<"&({d}+1)*{M})')
    sc.set('前面', f'COUNTIF(资_排序键,"<"&{d}*{M})')
    sc.set('显示数', f'MIN({sc("笔数")},{DAY_SHOW})')
    sc.set('末行', f'{D_H2}+{sc("显示数")}+2')
    sc.set('内转净', R2(f'SUMIFS(资_净额,资_归类,"内部转账",资_日期,{d})'))
    mon_ = f'资_日期,">="&{m0},资_日期,"<="&{d}'
    sc.set('全_今收', R2(f'SUMIFS(资_收入,资_日期,{d},{CASH_OK},{NOT_XFER})'))
    sc.set('全_今支', R2(f'SUMIFS(资_支出,资_日期,{d},{CASH_OK},{NOT_XFER})'))
    sc.set('全_月收', R2(f'SUMIFS(资_收入,{mon_},{CASH_OK},{NOT_XFER})'))
    sc.set('全_月支', R2(f'SUMIFS(资_支出,{mon_},{CASH_OK},{NOT_XFER})'))
    sc.set('全天收入', R2(f'SUMIFS(资_收入,资_日期,{d},资_有效,1)'))
    sc.set('全天支出', R2(f'SUMIFS(资_支出,资_日期,{d},资_有效,1)'))

    # ── 隐藏账户表 ──
    acct_table(ws, D_A0)
    C = AC_
    H = dict(昨='AH', 今收='AI', 今支='AJ', 今余='AK', 月收='AL', 月支='AM')
    for i in range(N_ACC):
        a, r = i + 1, D_A0 + i
        z = lambda f: f'IF(${C["显示"]}{r}<>1,0,{f})'
        hset(ws, f'AH{r}', z(R2(f'${C["期初"]}{r}+' + flow(a, r, '净额', f'资_日期,"<"&{d}'))))
        hset(ws, f'AI{r}', z(R2(flow(a, r, '收入', f'资_日期,{d},资_有效,1'))))
        hset(ws, f'AJ{r}', z(R2(flow(a, r, '支出', f'资_日期,{d},资_有效,1'))))
        hset(ws, f'AK{r}', z(R2(f'AH{r}+AI{r}-AJ{r}')))
        hset(ws, f'AL{r}', z(R2(flow(a, r, '收入', f'{mon_},资_有效,1'))))
        hset(ws, f'AM{r}', z(R2(flow(a, r, '支出', f'{mon_},资_有效,1'))))
        acct_text(ws, r, f'AK{r}')
    ks = f'${C["排序"]}${D_A0}:${C["排序"]}${D_A0 + N_ACC - 1}'
    sc.set('账户数', f'COUNT({ks})')
    for p, flag in (('可用', C['现金']), ('票据', C['票据'])):
        for k, col in H.items():
            sc.set(f'{p}_{k}', R2(f'SUMIFS(${col}${D_A0}:${col}${D_A0 + N_ACC - 1},${flag}${D_A0}:${flag}${D_A0 + N_ACC - 1},1)'))
    sc.set('垫付说明', '&'.join(f'IF({C["说明"]}{D_A0 + i}="","",{C["说明"]}{D_A0 + i}&"；")' for i in range(N_ACC)))

    # ── 第 3、4 行 ──
    ws.merge_cells('D3:H3')
    put(ws, 'D3', (f'="实际用的："&TEXT({d},"yyyy-mm-dd")&IF(TRIM(C3&"")="","（空着＝截止日）","")'
                   f'&IF({sc("认出")}=0,"（⚠ 没认出这个日期，先按截止日）","")'
                   f'&IF({d}>P_截止,"　⚠ 晚于截止日 "&TEXT(P_截止,"yyyy-mm-dd")&"：截止日以后的不算（月报、利润表都不算），日报照样按这一天算","")'
                   f'&IF({d}<P_建账日,"　早于建账日 "&TEXT(P_建账日,"yyyy-mm-dd")&"：建账前没有流水，余额都是期初余额","")'),
        F_NOTE, align=ALW, border=False)
    kpi(ws, 'B4', '当天可用资金', 'C4', sc('可用_今余'))
    kpi(ws, 'D4', '手上票据', 'E4', sc('票据_今余'))
    kpi(ws, 'F4', '当天实际收入', 'G4', sc('全_今收'))
    kpi(ws, 'H4', '当天实际支出', 'I4', sc('全_今支'))
    ws.row_dimensions[4].height = 22

    # ── ① 各账户 ──
    section(ws, 5, 'A', LAST, '① 各账户（余额按日期算；收入、支出含内部转账，余额才接得上）', C_VIEW)
    header(ws, 6, [('A', '序号'), ('B', '账户'), ('C', '类型'), ('D', '昨日余额'), ('E', '今日收入'), ('F', '今日支出'),
                   ('G', '今日余额'), ('H', '本月收入\n（1 日～当天）'), ('I', '本月支出\n（1 日～当天）')], GREEN_H, height=34)
    na = sc('账户数')
    acct_rows_disp(ws, D_A0, D_N1, na, 3)
    for k in range(1, D_N1 + 1):
        r = D_A0 + k - 1
        ix = f'$AP{r}'
        cell(ws, f'A{r}', f'IF({ix}>0,{k},"")', F_NOTE)
        cell(ws, f'B{r}', pick(ix, C['名称'], D_A0, N_ACC, ['"可用资金合计"', '"手上票据合计"', '"全公司实际收支"']), align=AL)
        cell(ws, f'C{r}', pick(ix, C['类型字'], D_A0, N_ACC, ['"现金、银行、微信"', '"承兑，不能直接花"', '"现金账户，不含内部转账"']),
             align=AL)
        for col, k_ in zip('DEFGHI', H):
            ch = [sc(f'可用_{k_}'), sc(f'票据_{k_}'), sc(f'全_{k_}') if k_ in ('今收', '今支', '月收', '月支') else '""']
            cell(ws, f'{col}{r}', pick(ix, H[k_], D_A0, N_ACC, ch), fmt=MONEY, align=AR)
    a0, a1 = D_A0, D_A0 + D_N1 - 1
    cf = ws.conditional_formatting
    z0 = f'$AP{a0}'
    rng = f'A{a0}:{LAST}{a1}'
    cf.add(rng, FormulaRule(formula=[f'OR({z0}=-1,{z0}=-2)'], fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD,
                            stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{z0}=-3'], fill=cf_fill('FFDDEBF7'), font=Font(bold=True), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'AND({z0}>0,INDEX(${C["类型"]}${D_A0}:${C["类型"]}${D_A0 + N_ACC - 1},MAX(1,{z0}))>=5)'],
                            font=Font(italic=True, color='FF7F6000'), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{z0}>0'], border=CF_BD))
    note(ws, f'A{D_NOTE1}', LAST,
         f'=IF({sc("内转净")}<>0,"⚠ 当天的内部转账转出、转进差 "&TEXT(ABS({sc("内转净")}),"#,##0.00")&"（只记了一边？手续费要另记一行）　","")'
         f'&IF({sc("垫付说明")}="","","到这天："&{sc("垫付说明")}&"　")'
         f'&"可用资金＝现金、银行、微信支付宝账户的合计；手上票据＝承兑汇票账户；垫付、备用金不是公司的钱，不算可用资金。"'
         f'&"全公司实际收支只算现金账户，不算内部转账（自己的现金账户之间倒钱）；承兑到期、贴现进银行算收货款。"', height=44)
    cf.add(f'A{D_NOTE1}', FormulaRule(formula=[f'LEFT($A${D_NOTE1},1)="⚠"'], font=Font(bold=True, color='FFC00000')))

    # ── ② 当天每一笔 ──
    section(ws, D_SEC2, 'A', LAST, '② 当天每一笔（同一天按录入顺序；余额＝这一笔之后本账户的余额）', C_VIEW)
    N, NN = sc('笔数'), sc('显示数')
    note(ws, f'A{D_NOTE2}', LAST,
         f'=IF({N}=0,TEXT({d},"yyyy-mm-dd")&" 没有收支",TEXT({d},"yyyy-mm-dd")&" 共 "&{N}&" 笔（含内部转账、垫付、备用金）")'
         f'&IF({N}>{DAY_SHOW},"　⚠ 只列前 {DAY_SHOW} 笔（合计是全天的）","")')
    cf.add(f'A{D_NOTE2}', FormulaRule(formula=[f'ISNUMBER(FIND("⚠",$A${D_NOTE2}))'], font=Font(bold=True, color='FFC00000')))
    header(ws, D_H2, [('A', '序号'), ('B', '账户'), ('C', '收支类别'), ('D', '往来单位'), ('E', '摘要'), ('F', '收入'),
                      ('G', '支出'), ('H', '本账户余额'), ('I', '算到哪')], GREEN_H, height=30)
    AR_ = lambda col: f'${col}${D_A0}:${col}${D_A0 + N_ACC - 1}'
    for k in range(1, D_NL + 1):
        r = D_L0 + k - 1
        q, y, t, a = f'$AQ{r}', f'$AR{r}', f'$AS{r}', f'$AU{r}'
        ws[f'AQ{r}'] = k
        ws[f'AQ{r}'].font = F_HELP
        hset(ws, f'AR{r}', f'IF({q}>{NN},"",SMALL(资_排序键,{sc("前面")}+{q}))')
        hset(ws, f'AT{r}', f'IF({y}="",0,MOD({y},{M}))')
        n = f'$AT{r}'
        hset(ws, f'AS{r}', f'IF({n}>0,1,IF({q}={NN}+1,3,IF({q}={NN}+2,4,0)))')
        hset(ws, f'AU{r}', f'IF({t}=1,INDEX(资_账户号,{n}),0)')
        ix = lambda nm: f'INDEX({nm},{n})'
        cell(ws, f'A{r}', f'IF({t}=1,{q},"")', F_NOTE)
        cell(ws, f'B{r}', f'IF({t}=1,INDEX(账户_名称,{a}),IF({t}=3,"当天合计",IF({t}=4,"共 "&{N}&" 笔","")))', align=AL)
        cell(ws, f'C{r}', f'IF({t}=1,INDEX(类别_名称,{ix("资_类别号")}),IF({t}=3,"含内部转账",""))', align=AL)
        cell(ws, f'D{r}', (f'IF({t}=1,IF({ix("资_单位")}<>"",{ix("资_单位")},IF({ix("资_报销人号")}>0,'
                           f'INDEX(人员_姓名,{ix("资_报销人号")}),"")),"")'), align=AL)
        cell(ws, f'E{r}', f'IF({t}=1,{ix("资_显示摘要")},IF(AND({t}=4,{N}>{DAY_SHOW}),"⚠ 只列了前 {DAY_SHOW} 笔（合计是全天的）",""))',
             align=AL)
        cell(ws, f'F{r}', f'IF({t}=1,IF({ix("资_收入")}=0,"",{ix("资_收入")}),IF({t}=3,{sc("全天收入")},""))', fmt=MONEY_B, align=AR)
        cell(ws, f'G{r}', f'IF({t}=1,IF({ix("资_支出")}=0,"",{ix("资_支出")}),IF({t}=3,{sc("全天支出")},""))', fmt=MONEY_B, align=AR)
        lk, ty, per = (f'INDEX({AR_(C[k_])},{a})' for k_ in ('挂', '类型', '人'))
        bal = (f'INDEX({AR_(C["期初"])},{a})+SUMIFS(资_净额,资_账户号,{a},资_排序键,"<="&{y})'
               f'-IF({lk}>0,SUMIFS(资_净额,资_原对方号,{lk},资_账户类型,"<>"&{ty},资_单位,{esc(per)},资_排序键,"<="&{y}),0)')
        cell(ws, f'H{r}', f'IF({t}=1,ROUND({bal},2),"")', fmt=MONEY, align=AR)
        cell(ws, f'I{r}', f'IF({t}=1,IF({ix("资_归类")}="","？没归类",{ix("资_归类")}),"")', align=AC)
    d0, d1 = D_L0, D_L0 + D_NL - 1
    t0 = f'$AS{d0}'
    rng = f'A{d0}:{LAST}{d1}'
    cf.add(f'E{d0}:E{d1}', FormulaRule(formula=[f'AND({t0}=4,LEFT($E{d0},1)="⚠")'], font=Font(bold=True, color='FFC00000'),
                                       stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=3'], fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=4'], font=Font(color='FF808080'), stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'AND({t0}=1,$I{d0}="内部转账")'], fill=cf_fill('FFEAF4FB'), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=1'], border=CF_BD))
    hide(ws, *[CL(i) for i in range(CI('J'), CI('AW') + 1)])
    finish(ws, LAST, f'{D_H2}:{D_H2}', sc('末行'), d1, brk=D_SEC2)


# ═══════════════════════════════ 资金月报 ═══════════════════════════════
TIP_MON = ('★ 黄格选年月（填 9、2026-9 或 202609 都行；只认本年度；空着＝截止日那个月）。只算截止日以前的钱。'
           '① 每个账户：月初余额（上个月底）、本月收入、本月支出（含内部转账）、月末余额（月底，没到月底按截止日）；'
           '垫付账户余额是负数＝欠垫付人的，备用金账户余额＝这个人手上的备用金。'
           '② 这一年各月（只算现金账户）：收货款、付货款、费用（生产成本、管理、销售、财务费用、税金、工资、其他支出）、其他收入、'
           '长期资产（设备、建厂费用）、借款投资等（借款、投资、押金借出、还垫付的净额）、本月净现金流（不含内部转账）、月末可用资金和手上票据。'
           '③ 按收支类别（只算现金账户）：本月、今年 1 月到本月累计的收入和支出；合计不含内部转账，内部转账单列一行。')

MO_A0, MO_N1 = 7, N_ACC + 2
MO_NOTE1 = MO_A0 + MO_N1
MO_SEC2, MO_H2 = MO_NOTE1 + 1, MO_NOTE1 + 2
MO_M0 = MO_H2 + 1                   # 1～12 月，第 13 行全年
MO_NOTE2 = MO_M0 + 13
MO_SEC3, MO_NOTE3, MO_H3 = MO_NOTE2 + 1, MO_NOTE2 + 2, MO_NOTE2 + 3
MO_C0 = MO_H3 + 1
MO_NC = N_CAT + 3
MAT0 = 'BA'                         # ② 隐藏：月末余额矩阵（账户 1..N_ACC 横排）
MO_GRP = [('收货款', ['收货款']), ('付货款', ['付货款']), ('费用', EXPENSE_KINDS), ('其他收入', ['其他收入']),
          ('长期资产', ['长期资产']), ('借款投资等', ['借款', '投资', '押金借出', '垫付还款']), ('内部转账', ['内部转账'])]
MO_K0 = 2                           # 每条资金一行的键（CX 类别月键、CY 组月键）
MO_KC = 'CX'
MO_CM = [CL(CI('BR') + i) for i in range(12)]   # ③ 类别表右边：每个类别 1～12 月净额（现金账户、截止日以前）


def grp_of(x):
    """归类文字 → 资金月报的组号 1～7（0＝认不得）"""
    out = '0'
    for g in reversed(range(len(MO_GRP))):
        out = f'IF(OR({",".join(f"{x}=" + chr(34) + k + chr(34) for k in MO_GRP[g][1])}),{g + 1},{out})'
    return out


def build_mon(ws):
    LAST = 'J'
    widths(ws, dict({'A': 12}, **{c: 13.5 for c in 'BCDEFGHIJ'}))
    title(ws, '=IF(P_公司名称="","",P_公司名称&"　")&"资金月报"', LAST, C_VIEW, TIP_MON)
    selector(ws, 'A3', '选年月', 'B3', None, fmt='General')
    dv = DataValidation(type='custom', formula1='TRUE', allow_blank=True, showErrorMessage=False, showInputMessage=True,
                        promptTitle='提示', prompt='填 9、2026-9 或 202609；只认本年度；空着＝截止日那个月')
    ws.add_data_validation(dv)
    dv.add('B3')
    home_link(ws, f'{LAST}3')
    ws.row_dimensions[3].height = 30

    keys = ['输入', '认出前', 'm', '认出', '月初', '月末', '起日', '末日', '账户数', '类别数', '有没归类',
            '显示行数', '末行', '全_收', '全_支', '全_净', '空月', '差月数', '垫付说明']
    for p in ('可用', '票据'):
        keys += [f'{p}_{k}' for k in ('初', '收', '支', '末')]
    for p in ('合计', '内转', '全部', '没归类'):
        keys += [f'{p}_{k}' for k in ('月收', '月支', '年收', '年支')]
    sc = Sc(ws, keys)
    ym_parse(sc, 'B3', 'P_截止月')
    m = sc('m')
    sc.set('月初', f'DATE(P_年度,{m},1)')
    sc.set('月末', f'DATE(P_年度,{m}+1,0)')
    sc.set('起日', f'MIN({sc("月初")}-1,P_截止)')
    sc.set('末日', f'MIN({sc("月末")},P_截止)')
    sc.set('空月', f'IF({sc("月初")}>P_截止,1,0)')
    upto = '资_日期,"<="&P_截止'
    inm = f'资_有效,1,资_月,{m},{upto}'
    sc.set('全_收', R2(f'SUMIFS(资_收入,{inm},资_现金,1,{NOT_XFER})'))
    sc.set('全_支', R2(f'SUMIFS(资_支出,{inm},资_现金,1,{NOT_XFER})'))
    sc.set('全_净', R2(f'SUMIFS(资_净额,{inm},资_现金,1,{NOT_XFER})'))

    # ── 隐藏账户表 ──
    acct_table(ws, MO_A0)
    C = AC_
    H = dict(初='AH', 收='AI', 支='AJ', 末='AK')
    for i in range(N_ACC):
        a, r = i + 1, MO_A0 + i
        z = lambda f: f'IF(${C["显示"]}{r}<>1,0,{f})'
        hset(ws, f'AH{r}', z(R2(f'${C["期初"]}{r}+' + flow(a, r, '净额', f'资_日期,"<="&{sc("起日")}'))))
        hset(ws, f'AI{r}', z(R2(flow(a, r, '收入', inm))))
        hset(ws, f'AJ{r}', z(R2(flow(a, r, '支出', inm))))
        hset(ws, f'AK{r}', z(R2(f'${C["期初"]}{r}+' + flow(a, r, '净额', f'资_日期,"<="&{sc("末日")}'))))
        acct_text(ws, r, f'AK{r}')
    ks = f'${C["排序"]}${MO_A0}:${C["排序"]}${MO_A0 + N_ACC - 1}'
    sc.set('账户数', f'COUNT({ks})')
    for p, flag in (('可用', C['现金']), ('票据', C['票据'])):
        for k, col in H.items():
            sc.set(f'{p}_{k}', R2(f'SUMIFS(${col}${MO_A0}:${col}${MO_A0 + N_ACC - 1},${flag}${MO_A0}:${flag}${MO_A0 + N_ACC - 1},1)'))

    ws.merge_cells('C3:I3')
    put(ws, 'C3', (f'="＝ "&P_年度&" 年 "&{m}&" 月"&IF(TRIM(B3&"")="","（空着＝截止日那个月）","")'
                   f'&IF({sc("认出")}=0,IF(AND({sc("认出前")}>0,INT({sc("认出前")}/100)<>P_年度),'
                   f'"（⚠ 不是 "&P_年度&" 年：一本账只记一年，先按截止日那个月）","（⚠ 没认出来，先按截止日那个月）"),"")'
                   f'&"；月初余额算到 "&TEXT({sc("起日")},"yyyy-mm-dd")&"，月末余额算到 "&TEXT({sc("末日")},"yyyy-mm-dd")'
                   f'&IF({sc("月末")}>P_截止,"（截止日；截止日以后的不算）","")'
                   f'&IF({sc("月初")}>P_截止,"　⚠ 这个月在截止日 "&TEXT(P_截止,"yyyy-mm-dd")&" 以后，还没有数","")'
                   f'&IF({sc("月末")}<P_建账日,"　⚠ 这个月在建账日以前，没有流水","")'),
        F_NOTE, align=ALW, border=False)
    em = sc('空月')
    kpi(ws, 'A4', '月末可用资金', 'B4', f'IF({em}=1,"",{sc("可用_末")})')
    kpi(ws, 'C4', '月末手上票据', 'D4', f'IF({em}=1,"",{sc("票据_末")})')
    kpi(ws, 'E4', '本月实际收入', 'F4', f'IF({em}=1,"",{sc("全_收")})')
    kpi(ws, 'G4', '本月实际支出', 'H4', f'IF({em}=1,"",{sc("全_支")})')
    kpi(ws, 'I4', '本月净现金流', 'J4', f'IF({em}=1,"",{sc("全_净")})')
    ws.row_dimensions[4].height = 22
    cf = ws.conditional_formatting

    # ── ① 各账户 ──
    section(ws, 5, 'A', LAST, '① 各账户（收入、支出含内部转账；月末＝月底和截止日较早的那天）', C_VIEW)
    header(ws, 6, [('A', '账户'), ('C', '类型'), ('D', '月初余额\n（上月底）'), ('E', '本月收入'), ('F', '本月支出'),
                   ('G', '月末余额'), ('H', '说明')], GREEN_H, height=34)
    mhead(ws, 6, 'A', 'B')
    mhead(ws, 6, 'H', LAST)
    na = sc('账户数')
    acct_rows_disp(ws, MO_A0, MO_N1, na, 2)
    for k in range(1, MO_N1 + 1):
        r = MO_A0 + k - 1
        ix = f'$AP{r}'
        ws.merge_cells(f'A{r}:B{r}')
        cell(ws, f'A{r}', pick(ix, C['名称'], MO_A0, N_ACC, ['"可用资金合计"', '"手上票据合计"']), align=AL)
        cell(ws, f'C{r}', pick(ix, C['类型字'], MO_A0, N_ACC, ['"现金、银行、微信"', '"承兑，不能直接花"']), align=AL)
        for col, k_ in zip('DEFG', H):
            cell(ws, f'{col}{r}', f'IF({em}=1,"",{pick(ix, H[k_], MO_A0, N_ACC, [sc(f"可用_{k_}"), sc(f"票据_{k_}")])})',
                 fmt=MONEY, align=AR)
        ws.merge_cells(f'H{r}:{LAST}{r}')
        tyx = f'INDEX(${C["类型"]}${MO_A0}:${C["类型"]}${MO_A0 + N_ACC - 1},{ix})'
        txt_ = f'INDEX(${C["说明"]}${MO_A0}:${C["说明"]}${MO_A0 + N_ACC - 1},{ix})'
        cell(ws, f'H{r}', (f'IF(OR({ix}<=0,{em}=1),"",IF({tyx}>=5,"月末："&{txt_},IF(N(G{r})<-0.005,"⚠ 余额是负数：漏记了收入，还是记错了账户？",'
                           f'IF(ROUND(N(D{r})+N(E{r})-N(F{r})-N(G{r}),2)<>0,"⚠ 月初＋收入－支出≠月末",""))))'), F_NOTE, align=AL)
    a0, a1 = MO_A0, MO_A0 + MO_N1 - 1
    z0 = f'$AP{a0}'
    rng = f'A{a0}:{LAST}{a1}'
    cf.add(f'H{a0}:{LAST}{a1}', FormulaRule(formula=[f'LEFT($H{a0},1)="⚠"'], font=Font(bold=True, color='FFC00000')))
    cf.add(rng, FormulaRule(formula=[f'{z0}<0'], fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'AND({z0}>0,INDEX(${C["类型"]}${MO_A0}:${C["类型"]}${MO_A0 + N_ACC - 1},MAX(1,{z0}))>=5)'],
                            font=Font(italic=True, color='FF7F6000'), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{z0}>0'], border=CF_BD))
    note(ws, f'A{MO_NOTE1}', LAST, '="可用资金＝现金、银行、微信支付宝账户；手上票据＝承兑汇票账户；垫付、备用金账户不是公司的钱，不算可用资金。'
                                   '各账户收支含内部转账（自己账户之间倒钱），全公司合起来一进一出抵掉。"')

    # ── 每条资金的键（现金账户、有效、截止日以前）：CX 类别号×100＋月（不符合 0） ──
    for i in range(N_CASH):
        r, n = MO_K0 + i, i + 1
        ws[f'CW{r}'] = n
        ws[f'CW{r}'].font = F_HELP
        n = f'$CW{r}'
        x = lambda nm: f'INDEX({nm},{n})'
        hset(ws, f'{MO_KC}{r}', (f'IF(AND({x("资_有效")}=1,{x("资_现金")}=1,{x("资_日期")}<=P_截止),'
                                 f'{x("资_类别号")}*100+{x("资_月")},0)'))
    C0 = MO_C0
    CG = f'$AF${C0}:$AF${C0 + N_CAT - 1}'
    KC = f'${MO_KC}${MO_K0}:${MO_KC}${MO_K0 + N_CASH - 1}'
    for i in range(N_CAT):          # ③ 类别表右边：第 p 个类别每月净额（② 按组号汇总用）
        p, r = i + 1, C0 + i
        for mi, c in enumerate(MO_CM):
            hset(ws, f'{c}{r}', f'SUMIF({KC},{p}*100+{mi + 1},资_净额)')

    # ── ② 全年各月 ──
    section(ws, MO_SEC2, 'A', LAST, '="② "&P_年度&" 年各月（只算现金账户；按付款日期；月末＝月底和截止日较早的；截止日以后的月份空着）"', C_VIEW)
    header(ws, MO_H2, [('A', '月份'), ('B', '收货款\n（净收）'), ('C', '付货款\n（净付）'), ('D', '费用\n（净付）'),
                       ('E', '其他收入'), ('F', '长期资产\n（净付）'), ('G', '借款投资等\n（净额）'), ('H', '本月净现金流\n（不含内部转账）'),
                       ('I', '月末可用资金'), ('J', '月末手上票据')], GREEN_H, height=34)
    mat = [CL(CI(MAT0) + i) for i in range(N_ACC)]
    for i, c in enumerate(mat):
        r_ = MO_A0 + i
        hset(ws, f'{c}{MO_SEC2}', f'${C["现金"]}${r_}')
        hset(ws, f'{c}{MO_H2}', f'${C["票据"]}${r_}')
    fu_a, fu_b = f'${mat[0]}${MO_SEC2}:${mat[-1]}${MO_SEC2}', f'${mat[0]}${MO_H2}:${mat[-1]}${MO_H2}'
    for k in range(1, 14):
        r = MO_M0 + k - 1
        yr = k == 13
        if yr:
            hset(ws, f'AT{r}', 'MIN(P_年末,P_截止)')
            hset(ws, f'AU{r}', 'IF(DATE(P_年度,1,1)>P_截止,0,1)')
        else:
            hset(ws, f'AR{r}', k)
            hset(ws, f'AT{r}', f'MIN(DATE(P_年度,{k}+1,0),P_截止)')
            hset(ws, f'AU{r}', f'IF(OR(DATE(P_年度,{k},1)>P_截止,DATE(P_年度,{k}+1,0)<P_建账日-1),0,1)')
            hset(ws, f'AV{r}', f'IF(OR(DATE(P_年度,{k},1)>P_截止,{k}<P_建账月),1,0)')
        for i, c in enumerate(mat):
            hset(ws, f'{c}{r}', (f'IF(OR($AU{r}=0,{c}${MO_SEC2}+{c}${MO_H2}=0),0,ROUND(INDEX(账户_期初,{i + 1})'
                                 f'+SUMIFS(资_净额,资_账户号,{i + 1},资_日期,"<="&$AT{r}),2))'))
        put(ws, f'A{r}', '全年' if yr else f'{k} 月', F_TXTB if yr else F_TXT, align=AC)
        if yr:
            for col in 'BCDEFGH':
                cell(ws, f'{col}{r}', f'ROUND(SUM({col}{MO_M0}:{col}{MO_M0 + 11}),2)', F_TXTB, MONEY, AR)
        else:
            fu = f'$AV{r}'
            cm = f'${MO_CM[k - 1]}${C0}:${MO_CM[k - 1]}${C0 + N_CAT - 1}'
            g = lambda gi: f'SUMIF({CG},{gi},{cm})'
            fx = {'B': g(1), 'C': '-' + g(2), 'D': '-' + g(3), 'E': g(4), 'F': '-' + g(5), 'G': g(6),
                  'H': f'SUM({cm})-{g(7)}'}
            for col, f in fx.items():
                cell(ws, f'{col}{r}', f'IF({fu}=1,"",{R2(f)})', F_TXTB if col == 'H' else F_TXT, MONEY, AR)
            hset(ws, f'AW{r}', R2(g(0)))         # 归类认不得的钱（只在净现金流里）
        mr = f'${mat[0]}{r}:${mat[-1]}{r}'
        bold = F_TXTB if yr else F_TXT
        cell(ws, f'I{r}', f'IF($AU{r}=0,"",ROUND(SUMPRODUCT({mr},{fu_a}),2))', bold, MONEY, AR)
        cell(ws, f'J{r}', f'IF($AU{r}=0,"",ROUND(SUMPRODUCT({mr},{fu_b}),2))', bold, MONEY, AR)
        for col in 'ABCDEFGHIJ':
            ws[f'{col}{r}'].border = BD
            if yr:
                ws[f'{col}{r}'].fill = FILL_TOT
    m0_, m1_ = MO_M0, MO_M0 + 11
    cf.add(f'A{m0_}:{LAST}{m1_}', FormulaRule(formula=[f'$AR{m0_}={m}'], fill=cf_fill('FFFFF2CC'), font=Font(bold=True)))
    cf.add(f'A{m0_}:{LAST}{m1_}', FormulaRule(formula=[f'$AV{m0_}=1'], font=Font(color='FFA6A6A6')))
    sc.set('差月数', f'COUNTIF($AW${m0_}:$AW${m1_},"<>0")')
    note(ws, f'A{MO_NOTE2}', LAST,
         '="收货款、付货款都是净额（退款已经减掉）；费用＝生产成本、管理费用、销售费用、财务费用、税金、工资、其他支出这几类收支类别的净付'
         '（按付款月份；费用统计、利润表按费用月）；借款投资等＝借款、还借款、投资、退投资、押金借出、发备用金、还老板垫付的净额（正数＝净进来）。"'
         '&"本月净现金流＝现金账户全部收入－全部支出（不含内部转账）＝收货款－付货款－费用＋其他收入－长期资产＋借款投资等。'
         '老板垫付、备用金账户里的钱不是现金，不在这里（看费用统计、个人垫付现金明细）。全年那行的月末＝年底或截止日。选的月份黄底。"'
         f'&IF({sc("差月数")}>0,"　⚠ 有 "&{sc("差月数")}&" 个月有收支类别的「归类」认不得，钱只算在净现金流里：到【基础资料】③收支类别改归类","")',
         height=58)
    cf.add(f'A{MO_NOTE2}', FormulaRule(formula=[f'ISNUMBER(FIND("⚠",$A${MO_NOTE2}))'], font=Font(bold=True, color='FFC00000')))

    # ── ③ 按收支类别 ──
    section(ws, MO_SEC3, 'A', LAST, f'="③ 按收支类别（"&P_年度&" 年 "&{m}&" 月；本年累计＝1 月～"&{m}&" 月；只算现金账户、截止日以前的）"', C_VIEW)
    note(ws, f'A{MO_NOTE3}', LAST,
         f'="按【基础资料】③收支类别的顺序。收入、支出是原数（退款记在哪栏就算哪栏）。合计不含内部转账；内部转账单列一行，转出、转进应该相等'
         f'。"'
         f'&IF({sc("有没归类")}=1,"　⚠ 有些钱的收支类别没算进上面的清单，单列一行（到资金台帐看最右边的校验列）","")')
    cf.add(f'A{MO_NOTE3}', FormulaRule(formula=[f'ISNUMBER(FIND("⚠",$A${MO_NOTE3}))'], font=Font(bold=True, color='FFC00000')))
    header(ws, MO_H3, [('A', '收支类别'), ('C', '算到哪'), ('D', '本月收入'), ('E', '本月支出'), ('F', '本年累计收入\n（1 月～本月）'),
                       ('G', '本年累计支出\n（1 月～本月）'), ('H', '说明')], GREEN_H, height=34)
    mhead(ws, MO_H3, 'A', 'B')
    mhead(ws, MO_H3, 'H', LAST)
    # 隐藏类别表（第 MO_C0 行起，第 i 个类别）：AD 名称、AN 显示、AE 归类、AF 组号、AH..AK 本月收/支、本年收/支、AO 累计显示数
    for i in range(N_CAT):
        p, r = i + 1, C0 + i
        hset(ws, f'AD{r}', f'INDEX(类别_名称,{p})&""')
        hset(ws, f'AN{r}', f'IF(AD{r}="",0,IF(MATCH({esc(f"AD{r}")},类别_名称,0)={p},1,0))')
        hset(ws, f'AE{r}', f'IF(AN{r}=1,INDEX(类别_归类,{p})&"","")')
        hset(ws, f'AF{r}', grp_of(f'INDEX(类别_归类,{p})'))
        for col, fld, yr_ in (('AH', '资_收入', 0), ('AI', '资_支出', 0), ('AJ', '资_收入', 1), ('AK', '资_支出', 1)):
            f = (f'SUMIFS({fld},{KC},">="&({p}*100+1),{KC},"<="&({p}*100+{m}))' if yr_ else f'SUMIF({KC},{p}*100+{m},{fld})')
            hset(ws, f'{col}{r}', f'IF(AN{r}<>1,0,{R2(f)})')
        hset(ws, f'AO{r}', f'AN{r}' + (f'+AO{r - 1}' if i else ''))
    sc.set('类别数', f'$AO${C0 + N_CAT - 1}')
    CK = dict(月收='AH', 月支='AI', 年收='AJ', 年支='AK')
    iny = f'资_有效,1,资_月,">=1",资_月,"<="&{m},{upto}'
    for k, col in CK.items():
        fld = '资_收入' if k.endswith('收') else '资_支出'
        rng_ = inm if k.startswith('月') else iny
        sc.set(f'合计_{k}', R2(f'SUMIFS({fld},{rng_},资_现金,1,{NOT_XFER})'))
        sc.set(f'内转_{k}', R2(f'SUMIFS({fld},{rng_},资_现金,1,资_归类,"内部转账")'))
        sc.set(f'全部_{k}', R2(f'SUMIFS({fld},{rng_},资_现金,1)'))
        sc.set(f'没归类_{k}', R2(f'{sc(f"全部_{k}")}-SUM(${col}${C0}:${col}${C0 + N_CAT - 1})'))
    sc.set('有没归类', 'IF(OR(' + ','.join(f'{sc(f"没归类_{k}")}<>0' for k in CK) + '),1,0)')
    nc, oth = sc('类别数'), sc('有没归类')
    sc.set('显示行数', f'{nc}+2+{oth}')
    sc.set('末行', f'{MO_H3}+{sc("显示行数")}')
    cnt_rng = f'$AO${C0}:$AO${C0 + N_CAT - 1}'
    for k in range(1, MO_NC + 1):
        r = MO_C0 + k - 1
        ix = f'$AP{r}'
        hset(ws, f'AP{r}', f'IF({k}<={nc},MATCH({k},{cnt_rng},0),IF({k}={nc}+1,IF({oth}=1,-1,-2),'
                           f'IF({k}={nc}+2,IF({oth}=1,-2,-3),IF(AND({k}={nc}+3,{oth}=1),-3,0))))')
        ws.merge_cells(f'A{r}:B{r}')
        cell(ws, f'A{r}', pick(ix, 'AD', C0, N_CAT, ['"（类别不在清单里的）"', '"合计（不含内部转账）"', '"其中内部转账"']), align=AL)
        cell(ws, f'C{r}', pick(ix, 'AE', C0, N_CAT, ['"？"', '""', '"内部转账"']), align=AC)
        for col, k_ in zip('DEFG', CK):
            cell(ws, f'{col}{r}', pick(ix, CK[k_], C0, N_CAT, [sc(f'没归类_{k_}'), sc(f'合计_{k_}'), sc(f'内转_{k_}')]),
                 fmt=MONEY, align=AR)
        ws.merge_cells(f'H{r}:{LAST}{r}')
        dif = f'ROUND(N(D{r})-N(E{r}),2)'
        ydif = f'ROUND(N(F{r})-N(G{r}),2)'
        cell(ws, f'H{r}', (f'IF({ix}=0,"",IF({ix}=-3,IF(AND({dif}=0,{ydif}=0),"✓ 转进＝转出",'
                           f'"转进比转出"&IF({dif}>0,"多 ","少 ")&TEXT(ABS({dif}),"#,##0.00")&"（本月）；"'
                           f'&IF({ydif}>0,"多 ","少 ")&TEXT(ABS({ydif}),"#,##0.00")&"（1 月到本月）：只记了一边？"),'
                           f'IF({ix}=-1,"⚠ 到资金台帐看最右边的校验列",IF({ix}=-2,"＝全部收支－内部转账",'
                           f'IF(C{r}="内部转账","不算在合计里","")))))'), F_NOTE, align=AL)
    c0, c1 = MO_C0, MO_C0 + MO_NC - 1
    z0 = f'$AP{c0}'
    rng = f'A{c0}:{LAST}{c1}'
    cf.add(f'H{c0}:{LAST}{c1}', FormulaRule(formula=[f'LEFT($H{c0},1)="⚠"'], font=Font(bold=True, color='FFC00000')))
    cf.add(rng, FormulaRule(formula=[f'{z0}=-2'], fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{z0}=-3'], fill=cf_fill('FFDDEBF7'), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{z0}=-1'], font=Font(italic=True, color='FFC00000'), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'AND({z0}>0,$C{c0}="内部转账")'], font=Font(italic=True, color='FF808080'), border=CF_BD,
                            stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{z0}>0'], border=CF_BD))
    hide(ws, *[CL(i) for i in range(CI('K'), CI(MO_KC) + 1)])
    finish(ws, LAST, f'{MO_H3}:{MO_H3}', sc('末行'), c1, brk=MO_SEC2)


# ═══════════════════════════════ 费用统计 ═══════════════════════════════
TIP_FEE = ('★ 费用按月统计（按「费用月」：资金台帐填了所属月份的算到那个月，比如 9 月付 8 月电费算 8 月；没填按付款月份）。只算截止日以前付的，'
           '截止日那个月以后的月份是空的（提前付的到那个月才算）。上面按你自己的收支类别（【基础资料】③ 的顺序，只列生产成本、管理费用、销售费用、财务费用、'
           '税金、工资、其他支出、长期资产这几类，还有员工报销），金额是不含进项税的、含老板垫付和备用金付的；「员工报销」下面展开报销项目。'
           '下面是不走钱的：工资计提、折旧、建厂费用和预付摊销、借款利息计提、补到发票冲减的进项税。右边 ② 选一个类别、一个月份看每一笔。')

F_R0 = 8                                   # ① 第 8 行起
F_NT = 12                                  # 尾巴：小计、工资计提＋4 部门、折旧、摊销、利息、补票、小计、合计
F_NR = N_CAT + N_RE + 1 + F_NT
F_CT0 = 3                                  # 隐藏类别表第 3 行起（AD..）
F_IT0 = F_CT0 + N_CAT + 2                  # 隐藏报销项目表（j＝1..N_RE，最后一行 j＝0）
F_TT0 = F_IT0 + N_RE + 3                   # 隐藏不走钱表（9 行）＋ 3 行小计/合计
F_L0 = 8
F_NL = FEE_SHOW + 2
F_K0 = 2                                   # 每条资金一行的键：CX H1、CY H2（报销项目）、CZ 明细键
F_H1, F_H3 = 'CX', 'CZ'
FMC = [CL(CI('C') + i) for i in range(12)]          # C～N：1～12 月（显示）
FHM = [CL(CI('AJ') + i) for i in range(12)]         # 隐藏表 12 个月：AJ～AU；AV 全年
TAIL = [  # (标签, 来源说明, 类型 4 主行/5 其中, 每月公式(m))
    ('工资计提（工资表应发）', '工资表', 4, lambda m: '+'.join(f'INDEX(工汇_应发,{d},{m})' for d in range(1, 5))),
] + [(f'　其中 {dp}', '工资表', 5, (lambda d: (lambda m: f'INDEX(工汇_应发,{d},{m})'))(i + 1)) for i, dp in enumerate(DEPTS)] + [
    ('折旧', '固定资产卡片', 4, lambda m: f'SUM(INDEX(卡_月折旧块,0,{m}))'),
    ('建厂费用、预付摊销', '长期待摊、预付', 4, lambda m: f'INDEX(摊_贷1801,{m})+INDEX(摊_贷1123,{m})'),
    ('借款利息计提', '借款按日计息', 4, lambda m: f'SUM(INDEX(息_借本月利息,0,{m}))'),
    ('补票冲减（补到专票的进项税）', '发票登记', 4, lambda m: f'-SUMIF(票_月,{m},票_补票额)'),
]
FEE_PL_NOTE = ('★ 合计跟利润表对不上是正常的：这里上面按收支类别算付出去的钱（不含能抵的进项税，按费用月），里面有买设备、建厂费用、一次付几个月的房租（长期资产、预付），'
               '有发工资、员工借支、交增值税和代扣个税；下面「不走钱的」是月末自动计提的。利润表按会计科目：设备、建厂费用、预付按折旧、摊销分月进，'
               '工资按工资表计提进（发工资只冲应付工资），交增值税、代扣个税、员工借支不进费用。想看某个科目的数，看利润表、科目余额表。')


def build_fee(ws):
    LAST, L2 = 'O', 'W'
    widths(ws, dict({'A': 22, 'B': 11, 'O': 13.5, 'P': 2, 'Q': 11.5, 'R': 11, 'S': 13, 'T': 15, 'U': 24, 'V': 13.5, 'W': 8},
                    **{c: 11 for c in FMC}))
    title(ws, '=IF(P_公司名称="","",P_公司名称&"　")&"费用统计"', LAST, C_VIEW, TIP_FEE)
    ws.merge_cells(f'Q1:{L2}1')
    put(ws, 'Q1', '② 明细', F_TITLE, fill(C_VIEW), align=AC, border=False)
    ws.merge_cells(f'Q2:{L2}2')
    put(ws, 'Q2', '★ 选一个收支类别（空着＝左边列的全部类别）、一个月份（1～12，空着＝全年），列出这一年（按费用月，到截止日那个月为止）的每一笔，按日期排。'
                  '金额＝不含税的费用额（退回来的钱是负数）。打印时 ② 另起一页（没印出来：选中 ② 这块 → 打印 → 打印选定区域）。',
        F_TIP, FILL_TIP, align=ALW, border=False)
    home_link(ws, f'{LAST}3')
    selector(ws, 'Q3', '类别', 'R3', None, '=收支类别列表', prompt='从下拉选收支类别；空着＝左边列的全部类别')
    ws.merge_cells('R3:S3')
    selector(ws, 'T3', '月份', 'U3', None, '"' + ','.join(str(i) for i in range(1, 13)) + '"', fmt='0',
             prompt='选 1～12 月；空着＝全年')
    ws.row_dimensions[3].height = 26

    keys = ['类数', '报销类', '报销位', '项数', '段行数', '末行', '以后', '选类别', '选类别号', '类别在', '类别归类', '月输入', '月',
            '月认出', '笔数', '显示数', '明细合计', '明细末行']
    sc = Sc(ws, keys)
    upto = '资_日期,"<="&P_截止'
    K0, K1 = F_K0, F_K0 + N_CASH - 1
    H1R, H3R = (f'${c}${K0}:${c}${K1}' for c in (F_H1, F_H3))

    # ── 隐藏类别表（第 F_CT0 行起，第 i 个类别）：AD 名称、AE 显示、AF 归类、AG 报销类、AH 列不列、AI 累计、AJ..AU 1～12 月、AV 全年、AW 以后 ──
    CT0, CT1 = F_CT0, F_CT0 + N_CAT - 1
    kinds = 'OR(' + ','.join(f'AF{{r}}="{k}"' for k in FEE_KINDS_ALL) + ')'
    for i in range(N_CAT):
        p, r = i + 1, CT0 + i
        hset(ws, f'AD{r}', f'INDEX(类别_名称,{p})&""')
        hset(ws, f'AE{r}', f'IF(AD{r}="",0,IF(MATCH({esc(f"AD{r}")},类别_名称,0)={p},1,0))')
        hset(ws, f'AF{r}', f'IF(AE{r}=1,INDEX(类别_归类,{p})&"","")')
        hset(ws, f'AG{r}', f'IF(AND(AE{r}=1,OR(INDEX(类别_收特,{p})=5,INDEX(类别_付特,{p})=5)),1,0)')
        hset(ws, f'AH{r}', (f'IF(AE{r}<>1,0,IF(AND(INDEX(类别_收特,{p})=9,INDEX(类别_付特,{p})=9),0,'
                            f'IF(OR({kinds.format(r=r)},AG{r}=1),1,0)))'))
        hset(ws, f'AI{r}', f'AH{r}' + (f'+AI{r - 1}' if i else ''))
        for j, c in enumerate(FHM):
            hset(ws, f'{c}{r}', f'IF(OR(AH{r}<>1,{j + 1}>P_截止月),0,{R2(f"SUMIF({H1R},{p}*100+{j + 1},资_费用额)")})')
        hset(ws, f'AV{r}', R2(f'SUM({FHM[0]}{r}:{FHM[-1]}{r})'))
        later = f'SUMIFS(资_费用额,{H1R},">"&({p}*100+P_截止月),{H1R},"<="&({p}*100+12))'
        hset(ws, f'AW{r}', f'IF(AH{r}<>1,0,{R2(later)})')
    sc.set('类数', f'$AI${CT1}')
    sc.set('报销类', f'IFERROR(MATCH(1,$AG${CT0}:$AG${CT1},0),0)')
    RC = sc('报销类')
    sc.set('报销位', f'IF({RC}=0,{sc("类数")},INDEX($AI${CT0}:$AI${CT1},{RC}))')
    sc.set('以后', R2(f'SUM($AW${CT0}:$AW${CT1})'))

    # ── 隐藏报销项目表（第 F_IT0 行起；第 N_RE+1 行＝没选报销项目）：AD 名称、AE j、AF 有数、AG 累计、AJ..AU 月、AV 全年 ──
    IT0, IT1 = F_IT0, F_IT0 + N_RE
    for i in range(N_RE + 1):
        r = IT0 + i
        j = i + 1 if i < N_RE else 0
        ws[f'AE{r}'] = j
        ws[f'AE{r}'].font = F_HELP
        if j:
            hset(ws, f'AD{r}', f'INDEX(报销_名称,{j})&""')
        else:
            ws[f'AD{r}'] = '（没选报销项目）'
            ws[f'AD{r}'].font = F_HELP
        for mi, c in enumerate(FHM):
            hset(ws, f'{c}{r}', f'IF(OR({RC}=0,{mi + 1}>P_截止月),0,{R2(f"SUMIFS(资_费用额,{H1R},{RC}*100+{mi + 1},资_报销项目号,{j})")})')
        hset(ws, f'AV{r}', R2(f'SUM({FHM[0]}{r}:{FHM[-1]}{r})'))
        hset(ws, f'AF{r}', f'IF(COUNTIF({FHM[0]}{r}:{FHM[-1]}{r},"<>0")>0,1,0)')
        hset(ws, f'AG{r}', f'AF{r}' + (f'+AG{r - 1}' if i else ''))
    sc.set('项数', f'IF({RC}=0,0,$AG${IT1})')

    # ── 隐藏不走钱表（第 F_TT0 行起 9 行）＋ 小计1、小计2、合计 ──
    TT0 = F_TT0
    for i, (lab, src, ty, fm) in enumerate(TAIL):
        r = TT0 + i
        ws[f'AD{r}'] = lab
        ws[f'AD{r}'].font = F_HELP
        ws[f'AE{r}'] = ty
        ws[f'AE{r}'].font = F_HELP
        for mi, c in enumerate(FHM):
            hset(ws, f'{c}{r}', f'IF(OR({mi + 1}>P_截止月,{mi + 1}<P_建账月),0,{R2(fm(mi + 1))})')
        hset(ws, f'AV{r}', R2(f'SUM({FHM[0]}{r}:{FHM[-1]}{r})'))
    TT1 = TT0 + len(TAIL) - 1
    S1, S2, ST = TT1 + 1, TT1 + 2, TT1 + 3
    for r_, lab in ((S1, '小计1'), (S2, '小计2'), (ST, '合计')):
        ws[f'AD{r_}'] = lab
        ws[f'AD{r_}'].font = F_HELP
    for c in FHM + ['AV']:
        hset(ws, f'{c}{S1}', R2(f'SUMIF($AH${CT0}:$AH${CT1},1,{c}{CT0}:{c}{CT1})'))
        hset(ws, f'{c}{S2}', R2(f'SUMIF($AE${TT0}:$AE${TT1},4,{c}{TT0}:{c}{TT1})'))
        hset(ws, f'{c}{ST}', R2(f'{c}{S1}+{c}{S2}'))

    # ── 每条资金一行：H1＝类别号×100＋费用月（有效、截止日以前、费用月≥1），H3＝明细排序键 ──
    sel, mo, selno = sc('选类别'), sc('月'), sc('选类别号')
    inc = f'$AH${CT0}:$AH${CT1}'
    for i in range(N_CASH):
        r, n = K0 + i, i + 1
        ws[f'CW{r}'] = n
        ws[f'CW{r}'].font = F_HELP
        n = f'$CW{r}'
        x = lambda nm: f'INDEX({nm},{n})'
        h1 = f'{F_H1}{r}'
        hset(ws, h1, (f'IF(AND({x("资_有效")}=1,{x("资_日期")}<=P_截止,{x("资_费用月")}>=1),'
                      f'{x("资_类别号")}*100+{x("资_费用月")},0)'))
        ok_cat = f'IF({sel}="",INDEX({inc},INT({h1}/100))=1,INT({h1}/100)={selno})'
        hset(ws, f'{F_H3}{r}', (f'IF({h1}=0,"",IF(AND(MOD({h1},100)<=P_截止月,OR({mo}=0,MOD({h1},100)={mo}),{ok_cat}),'
                                f'{x("资_排序键")},""))'))

    # ── 显示行（第 F_R0 行起）：BE k、BF 类型（1 类别 2 报销项目 3 小计1 4 不走钱 5 其中部门 6 小计2 7 合计 0 空）、BG 引用行号 ──
    nC, posR, nI = sc('类数'), sc('报销位'), sc('项数')
    sc.set('段行数', f'{nC}+{nI}+{F_NT}')
    sc.set('末行', f'{F_R0 - 1}+{sc("段行数")}')
    ccnt = f'$AI${CT0}:$AI${CT1}'
    icnt = f'$AG${IT0}:$AG${IT1}'
    for k in range(1, F_NR + 1):
        r = F_R0 + k - 1
        hset(ws, f'BE{r}', k)
        q, t, g = f'$BE{r}', f'$BF{r}', f'$BG{r}'
        tail = f'{q}-{nC}-{nI}'
        tmap = 'CHOOSE(' + tail + ',3,' + ','.join(str(ty) for _l, _s, ty, _f in TAIL) + ',6,7)'
        hset(ws, f'BF{r}', f'IF({q}<={nC}+{nI},IF(AND({q}>{posR},{q}<={posR}+{nI}),2,1),IF({tail}<={F_NT},{tmap},0))')
        hset(ws, f'BG{r}', (f'IF({t}=1,MATCH(IF({q}<={posR},{q},{q}-{nI}),{ccnt},0),IF({t}=2,MATCH({q}-{posR},{icnt},0),'
                            f'IF(OR({t}=4,{t}=5),{tail}-1,0)))'))
        lab_tail = 'CHOOSE(' + tail + ',"小计：按收支类别付的",' + ','.join(f'"{l_}"' for l_, *_r in TAIL) + \
                   ',"小计：不走钱的（月末自动计提）","合计")'
        cell(ws, f'A{r}', (f'IF({t}=1,INDEX($AD${CT0}:$AD${CT1},{g}),IF({t}=2,"　其中 "&INDEX($AD${IT0}:$AD${IT1},{g}),'
                           f'IF({t}=0,"",{lab_tail})))'), align=AL)
        src_tail = 'CHOOSE(' + tail + ',"",' + ','.join(f'"{s_}"' for _l, s_, *_r in TAIL) + ',"","")'
        cell(ws, f'B{r}', f'IF({t}=1,INDEX($AF${CT0}:$AF${CT1},{g}),IF({t}=2,"报销项目",IF(OR({t}=4,{t}=5),{src_tail},"")))',
             F_NOTE, align=AC)
        for mi, c in enumerate(FMC + ['O']):
            hc = (FHM + ['AV'])[mi]
            f = (f'IF({t}=1,INDEX(${hc}${CT0}:${hc}${CT1},{g}),IF({t}=2,INDEX(${hc}${IT0}:${hc}${IT1},{g}),'
                 f'IF(OR({t}=4,{t}=5),INDEX(${hc}${TT0}:${hc}${TT1},{g}),IF({t}=3,${hc}${S1},IF({t}=6,${hc}${S2},'
                 f'IF({t}=7,${hc}${ST},""))))))')
            cell(ws, f'{c}{r}', f, F_TXTB if c == 'O' else F_TXT, MONEY0, AR)

    # ── 第 3、4、6 行 ──
    ws.merge_cells('A3:N3')
    put(ws, 'A3', (f'="＝ "&P_年度&" 年（按费用月；只算截止日 "&TEXT(P_截止,"yyyy-mm-dd")&" 以前付的，"&P_截止月&" 月以后的月份空着）"'
                   f'&IF(ABS({sc("以后")})>0.005,"；另有已付、算在以后月份的 "&TEXT({sc("以后")},"#,##0.00")&" 元（到那个月才算）","")'),
        F_NOTE, align=ALW, border=False)
    kpi(ws, 'A4', '全年合计', 'B4', f'$AV${ST}', MONEY0, merge_to='C4')
    kpi(ws, 'D4', '按收支类别', 'E4', f'$AV${S1}', MONEY0, merge_to='F4')
    kpi(ws, 'G4', '不走钱的', 'H4', f'$AV${S2}', MONEY0, merge_to='I4')
    kpi(ws, 'J4', '其中员工报销', 'K4', f'IF({RC}=0,0,INDEX($AV${CT0}:$AV${CT1},{RC}))', MONEY0, merge_to='L4')
    ws.row_dimensions[4].height = 22
    section(ws, 5, 'A', LAST, '="① "&P_年度&" 年费用（按费用月；含老板垫付、备用金付的；金额不含能抵的进项税）"', C_VIEW)
    ws.merge_cells(f'A6:{LAST}6')
    put(ws, 'A6', FEE_PL_NOTE, F_NOTE, align=ALW, border=False)
    ws.row_dimensions[6].height = 46
    header(ws, 7, [('A', '项目'), ('B', '归类／来源')] + [(c, f'{i + 1} 月') for i, c in enumerate(FMC)] + [('O', '全年')],
           GREEN_H, height=30)
    r0, r1 = F_R0, F_R0 + F_NR - 1
    cf = ws.conditional_formatting
    t0 = f'$BF{r0}'
    rng = f'A{r0}:{LAST}{r1}'
    cf.add(rng, FormulaRule(formula=[f'{t0}=7'], fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'OR({t0}=3,{t0}=6)'], fill=cf_fill('FFDDEBF7'), font=Font(bold=True), border=CF_BD,
                            stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'OR({t0}=2,{t0}=5)'], font=Font(italic=True, color='FF595959'), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=4'], fill=cf_fill('FFF2F2F2'), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=1'], border=CF_BD))
    for i, c in enumerate(FMC):
        cf.add(f'{c}7', FormulaRule(formula=[f'{i + 1}>P_截止月'], fill=cf_fill('FFBFBFBF')))

    # ── ② 明细（右边 Q～W） ──
    sc.set('选类别', 'TRIM(R3&"")')
    sc.set('选类别号', f'IF({sel}="",0,IFERROR(MATCH({esc(sel)},类别_名称,0),-1))')
    sc.set('类别在', f'IF({selno}<0,0,1)')
    sc.set('类别归类', f'IF({selno}<=0,"",INDEX(类别_归类,{selno})&"")')
    sc.set('月输入', 'IF(ISNUMBER(U3),U3,IFERROR(--SUBSTITUTE(TRIM(U3&""),"月",""),0))')
    mv = sc('月输入')
    sc.set('月', f'IF(AND({mv}>=1,{mv}<=12,INT({mv})={mv}),{mv},0)')
    sc.set('月认出', f'IF(TRIM(U3&"")="",1,IF({mo}>0,1,0))')
    sc.set('笔数', f'COUNT({H3R})')
    sc.set('显示数', f'MIN({sc("笔数")},{FEE_SHOW})')
    sc.set('明细末行', f'{F_L0 - 1}+{sc("显示数")}+2')
    sc.set('明细合计', R2(f'SUMIF({H3R},">0",资_费用额)'))
    N, NN = sc('笔数'), sc('显示数')
    ws.merge_cells(f'V3:{L2}3')
    put(ws, 'V3', (f'=IF({sc("月认出")}=0,"⚠ 月份填 1～12（先按全年）",IF({mo}=0,"＝全年","＝"&{mo}&" 月"))'
                   f'&IF({mo}>P_截止月,"（在截止日以后）","")'
                   f'&IF({sc("类别在")}=0,"　⚠ 类别不在清单里","")'), F_NOTE, align=ALW, border=False)
    ws.merge_cells(f'Q4:{L2}4')
    put(ws, 'Q4', (f'=IF(AND({selno}>0,NOT(ISNUMBER(MATCH({sc("类别归类")},{{{",".join(chr(34) + k + chr(34) for k in FEE_KINDS_ALL)}}},0)))),'
                   f'"「"&{sel}&"」算到「"&{sc("类别归类")}&"」，左边不列：金额＝不含税额（支出＋、收入－）","")'),
        F_NOTE, align=ALW, border=False)
    section(ws, 5, 'Q', L2, (f'="② "&P_年度&" 年"&IF({mo}=0,"全年"," "&{mo}&" 月")&"　"&IF({sel}="","左边全部类别","「"&{sel}&"」")'
                             f'&"（按费用月）共 "&{N}&" 笔"&IF({N}>{FEE_SHOW},"，⚠ 只列前 {FEE_SHOW} 笔","")'), C_VIEW)
    header(ws, 7, [('Q', '日期'), ('R', '账户'), ('S', '往来单位'), ('T', '收支类别'), ('U', '摘要'), ('V', '金额'),
                   ('W', '费用月')], GREEN_H, height=30)
    for k in range(1, F_NL + 1):
        r = F_L0 + k - 1
        ws[f'BK{r}'] = k
        ws[f'BK{r}'].font = F_HELP
        q, y, n, t = f'$BK{r}', f'$BL{r}', f'$BM{r}', f'$BN{r}'
        hset(ws, f'BL{r}', f'IF({q}>{NN},"",SMALL({H3R},{q}))')
        hset(ws, f'BM{r}', f'IF({y}="",0,MOD({y},{M}))')
        hset(ws, f'BN{r}', f'IF({n}>0,1,IF({q}={NN}+1,3,IF({q}={NN}+2,4,0)))')
        ix = lambda nm: f'INDEX({nm},{n})'
        cell(ws, f'Q{r}', f'IF({t}=1,{ix("资_日期")},IF({t}=3,"合计",IF({t}=4,"共 "&{N}&" 笔","")))', fmt=DATE)
        cell(ws, f'R{r}', f'IF({t}=1,INDEX(账户_名称,{ix("资_账户号")}),"")', align=AL)
        cell(ws, f'S{r}', (f'IF({t}=1,IF({ix("资_单位")}<>"",{ix("资_单位")},IF({ix("资_报销人号")}>0,'
                           f'INDEX(人员_姓名,{ix("资_报销人号")}),"")),"")'), align=AL)
        cell(ws, f'T{r}', (f'IF({t}=1,INDEX(类别_名称,{ix("资_类别号")})&IF({ix("资_报销项目号")}>0,"·"&INDEX(报销_名称,'
                           f'{ix("资_报销项目号")}),""),"")'), align=AL)
        cell(ws, f'U{r}', f'IF({t}=1,{ix("资_显示摘要")},IF(AND({t}=4,{N}>{FEE_SHOW}),"⚠ 只列了前 {FEE_SHOW} 笔（合计是全部的）",""))',
             align=AL)
        cell(ws, f'V{r}', f'IF({t}=1,{ix("资_费用额")},IF({t}=3,{sc("明细合计")},""))', fmt=MONEY, align=AR)
        cell(ws, f'W{r}', f'IF({t}=1,{ix("资_费用月")}&" 月","")')
    d0, d1 = F_L0, F_L0 + F_NL - 1
    t0 = f'$BN{d0}'
    rng = f'Q{d0}:{L2}{d1}'
    cf.add(f'U{d0}:U{d1}', FormulaRule(formula=[f'AND({t0}=4,LEFT($U{d0},1)="⚠")'], font=Font(bold=True, color='FFC00000'),
                                       stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=3'], fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=4'], font=Font(color='FF808080'), stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=1'], border=CF_BD))
    cf.add('V3', FormulaRule(formula=['ISNUMBER(FIND("⚠",$V$3))'], font=Font(bold=True, color='FFC00000')))
    hide(ws, *[CL(i) for i in range(CI('X'), CI(F_H3) + 1)])
    finish(ws, LAST, '7:7', sc('末行'), r1, freeze='B8', extra=[('Q', L2, d1, sc('明细末行'))])


# ═══════════════════════════════ 报销统计 ═══════════════════════════════
TIP_REIM = ('★ 员工报销按人、按报销项目统计（资金台帐类别是「员工报销」的行，不管从哪个账户付的：银行、现金、老板垫付、备用金都算）。'
            '黄格选起月、止月（1～12；空着＝1 月到截止日那个月），按费用月、不含能抵的进项税，只算截止日以前付的。'
            '左边：每人每个报销项目的金额（只列有数的项目，最多 15 个，再多的并在「其余项目」）、合计、其中老板垫付的（垫付账户付的，要还给垫付人）、'
            '其中备用金的（从备用金里花的）。右边：每人 1～12 月（全年，不受起止月影响）。')

RM_R0 = 7                                   # 显示第 7 行起
RM_NR = N_PER + 2                           # 人（含「没填报销人」）＋合计
RM_PT0 = 7                                  # 隐藏人员表第 7 行起（p＝1..N_PER，最后一行 p＝0）
RM_IT0 = RM_PT0 + N_PER + 3                 # 隐藏报销项目表（j＝1..N_RE，最后一行 j＝0）
RM_SLOT = [CL(CI('C') + i) for i in range(RE_SLOTS)]     # C～Q
RM_OTH, RM_TOT, RM_DZ, RM_BY = (CL(CI('C') + RE_SLOTS + i) for i in range(4))     # R 其余、S 合计、T 其中垫付、U 其中备用金
RM_LAST = RM_BY
RM_W = CL(CI(RM_LAST) + 2)                  # W：右边 报销人
RM_MC = [CL(CI(RM_W) + 1 + i) for i in range(12)]        # X～AI：1～12 月
RM_YR = CL(CI(RM_MC[-1]) + 1)               # AJ：全年
RM_L2 = RM_YR
RM_K0 = 2
RM_H = dict(R1='DX', R2='DY')                            # 每条资金一行的键


def build_reim(ws):
    LAST = RM_L2
    widths(ws, dict({'A': 12, 'B': 8, RM_OTH: 11, RM_TOT: 12, RM_DZ: 12, RM_BY: 12, CL(CI(RM_LAST) + 1): 2, RM_W: 12, RM_YR: 12},
                    **{c: 10 for c in RM_SLOT}, **{c: 10 for c in RM_MC}))
    title(ws, '=IF(P_公司名称="","",P_公司名称&"　")&"报销统计"', LAST, C_VIEW, TIP_REIM)
    months = '"' + ','.join(str(i) for i in range(1, 13)) + '"'
    selector(ws, 'A3', '起月', 'B3', None, months, fmt='0', prompt='选 1～12 月；空着＝1 月')
    selector(ws, 'C3', '止月', 'D3', None, months, fmt='0', prompt='选 1～12 月；空着＝截止日那个月')
    home_link(ws, f'{LAST}3')
    ws.row_dimensions[3].height = 26

    keys = ['起输入', '止输入', '起', '止填', '止', '认出', '人数', '项数', '末行', '合计', '其中垫付', '其中备用金', '笔数']
    sc = Sc(ws, keys, cl='AL', cv='AM')
    for k_, c_ in (('起输入', 'B3'), ('止输入', 'D3')):
        sc.set(k_, f'IF(ISNUMBER({c_}),{c_},IFERROR(--SUBSTITUTE(TRIM({c_}&""),"月",""),0))')
    ok = lambda v: f'AND({v}>=1,{v}<=12,INT({v})={v})'
    a_, b_ = sc('起输入'), sc('止输入')
    sc.set('起', f'IF({ok(a_)},{a_},1)')
    sc.set('止填', f'IF({ok(b_)},{b_},P_截止月)')
    sc.set('止', f'MIN({sc("止填")},P_截止月)')
    sc.set('认出', f'IF(AND(OR(TRIM(B3&"")="",{ok(a_)}),OR(TRIM(D3&"")="",{ok(b_)})),1,0)')
    m1, m2 = sc('起'), sc('止')

    # ── 每条资金一行：R1＝人号×100＋费用月（报销类别、有效、截止日以前、费用月 1～截止月；不符合 -1）；
    #    R2＝人号×100＋项目号（再限起止月） ──
    K0, K1 = RM_K0, RM_K0 + N_CASH - 1
    HR = {k: f'${c}${K0}:${c}${K1}' for k, c in RM_H.items()}
    for i in range(N_CASH):
        r, n = K0 + i, i + 1
        ws[f'DW{r}'] = n
        ws[f'DW{r}'].font = F_HELP
        n = f'$DW{r}'
        x = lambda nm: f'INDEX({nm},{n})'
        c1 = f'{RM_H["R1"]}{r}'
        cat = x('资_类别号')
        sp = f'IF({x("资_收入")}<>0,INDEX(类别_收特,{cat}),INDEX(类别_付特,{cat}))'
        hset(ws, c1, (f'IF(AND({x("资_有效")}=1,{x("资_日期")}<=P_截止,{x("资_费用月")}>=1,{x("资_费用月")}<=P_截止月),'
                      f'IF({sp}=5,{x("资_报销人号")}*100+{x("资_费用月")},-1),-1)'))
        c2 = f'{RM_H["R2"]}{r}'
        hset(ws, c2, f'IF({c1}<0,-1,IF(AND(MOD({c1},100)>={m1},MOD({c1},100)<={m2}),INT({c1}/100)*100+{x("资_报销项目号")},-1))')

    # ── 隐藏报销项目表（第 RM_IT0 行起）：AO 名称、AP j、AQ 用到、AR 累计、AS 合计 ──
    IT0, IT1 = RM_IT0, RM_IT0 + N_RE
    for i in range(N_RE + 1):
        r = IT0 + i
        j = i + 1 if i < N_RE else 0
        ws[f'AP{r}'] = j
        ws[f'AP{r}'].font = F_HELP
        if j:
            hset(ws, f'AO{r}', f'INDEX(报销_名称,{j})&""')
        else:
            ws[f'AO{r}'] = '（没选报销项目）'
            ws[f'AO{r}'].font = F_HELP
        hset(ws, f'AQ{r}', f'IF(COUNTIFS({HR["R2"]},">=0",资_报销项目号,{j})>0,1,0)')
        hset(ws, f'AR{r}', f'AQ{r}' + (f'+AR{r - 1}' if i else ''))
        hset(ws, f'AS{r}', R2(f'SUMIFS(资_费用额,{HR["R2"]},">=0",资_报销项目号,{j})'))
    sc.set('项数', f'$AR${IT1}')
    nI = sc('项数')
    # 槽 s（1..RE_SLOTS）对应的项目：第 6 行上面（第 5 行隐藏列 AU..）放 j
    slot_j = {}
    for s in range(1, RE_SLOTS + 1):
        c = CL(CI('AU') + s - 1)
        hset(ws, f'{c}5', f'IF({s}>{nI},-1,INDEX($AP${IT0}:$AP${IT1},MATCH({s},$AR${IT0}:$AR${IT1},0)))')
        slot_j[s] = f'{c}$5'

    # ── 隐藏人员表（第 RM_PT0 行起；最后一行 p＝0 没填报销人）：AO 名称、AP p、AQ 显示、AR 累计、AS 部门、
    #    AT 合计、AU..（槽 15 个）、BJ 其余、BK 其中垫付、BL 其中备用金、BM..BX 1～12 月、BY 全年 ──
    PT0, PT1 = RM_PT0, RM_PT0 + N_PER
    SLC = [CL(CI('AU') + s) for s in range(RE_SLOTS)]
    PO, PDZ, PBY = (CL(CI('AU') + RE_SLOTS + i) for i in range(3))
    PMC = [CL(CI(PBY) + 1 + i) for i in range(12)]
    PYR = CL(CI(PMC[-1]) + 1)
    for i in range(N_PER + 1):
        r = PT0 + i
        p = i + 1 if i < N_PER else 0
        ws[f'AP{r}'] = p
        ws[f'AP{r}'].font = F_HELP
        used = f'COUNTIFS({HR["R1"]},">="&({p}*100+1),{HR["R1"]},"<="&({p}*100+12))>0'
        if p:
            hset(ws, f'AO{r}', f'INDEX(人员_姓名,{p})&""')
            hset(ws, f'AQ{r}', f'IF(AO{r}="",0,IF(AND(MATCH({esc(f"AO{r}")},人员_姓名,0)={p},{used}),1,0))')
            hset(ws, f'AS{r}', f'IF(AQ{r}<>1,"",IF(INDEX(人员_部门号,{p})=0,"",CHOOSE(INDEX(人员_部门号,{p}),{",".join(chr(34) + d + chr(34) for d in DEPTS)})))')
        else:
            ws[f'AO{r}'] = '（没填报销人）'
            ws[f'AO{r}'].font = F_HELP
            hset(ws, f'AQ{r}', f'IF({used},1,0)')
            hset(ws, f'AS{r}', '""')
        hset(ws, f'AR{r}', f'AQ{r}' + (f'+AR{r - 1}' if i else ''))
        pr = f'{HR["R2"]},">="&{p * 100},{HR["R2"]},"<"&{(p + 1) * 100}'
        hset(ws, f'AT{r}', R2(f'SUMIFS(资_费用额,{pr})'))
        for s in range(1, RE_SLOTS + 1):
            sv = R2(f'SUMIF({HR["R2"]},{p}*100+{slot_j[s]},资_费用额)')
            hset(ws, f'{SLC[s - 1]}{r}', f'IF({slot_j[s]}<0,0,{sv})')
        hset(ws, f'{PO}{r}', R2(f'AT{r}-SUM({SLC[0]}{r}:{SLC[-1]}{r})'))
        hset(ws, f'{PDZ}{r}', R2(f'SUMIFS(资_费用额,{pr},资_账户类型,5)'))
        hset(ws, f'{PBY}{r}', R2(f'SUMIFS(资_费用额,{pr},资_账户类型,6)'))
        for mi, c in enumerate(PMC):
            hset(ws, f'{c}{r}', R2(f'SUMIF({HR["R1"]},{p}*100+{mi + 1},资_费用额)'))
        hset(ws, f'{PYR}{r}', R2(f'SUM({PMC[0]}{r}:{PMC[-1]}{r})'))
    sc.set('人数', f'$AR${PT1}')
    nP = sc('人数')
    sc.set('末行', f'{RM_R0 - 1}+{nP}+1')
    PT = lambda c: f'${c}${PT0}:${c}${PT1}'
    sc.set('合计', R2(f'SUM({PT("AT")})'))
    sc.set('其中垫付', R2(f'SUM({PT(PDZ)})'))
    sc.set('其中备用金', R2(f'SUM({PT(PBY)})'))
    sc.set('笔数', f'COUNTIF({HR["R2"]},">=0")')

    # ── 第 3、4 行 ──
    ws.merge_cells(f'E3:{RM_LAST}3')
    put(ws, 'E3', (f'="＝ "&P_年度&" 年 "&{m1}&"～"&{m2}&" 月（按费用月；只算截止日 "&TEXT(P_截止,"yyyy-mm-dd")&" 以前付的）"'
                   f'&IF({sc("认出")}=0,"　⚠ 月份填 1～12（没认出来的按默认）","")'
                   f'&IF({sc("止填")}>P_截止月,"　止月在截止日那个月以后，只算到 "&P_截止月&" 月","")'
                   f'&IF({m1}>{m2},"　⚠ 起月比止月晚，没有数","")'), F_NOTE, align=ALW, border=False)
    ws.conditional_formatting.add('E3', FormulaRule(formula=['ISNUMBER(FIND("⚠",$E$3))'], font=Font(bold=True, color='FFC00000')))
    kpi(ws, 'A4', '报销合计', 'B4', sc('合计'), MONEY, merge_to='C4')
    kpi(ws, 'D4', '其中老板垫付的', 'E4', sc('其中垫付'), MONEY, merge_to='F4')
    kpi(ws, 'G4', '其中备用金的', 'H4', sc('其中备用金'), MONEY, merge_to='I4')
    put(ws, 'J4', f'="共 "&{sc("笔数")}&" 笔"', F_NOTE, align=AL, border=False)
    ws.row_dimensions[4].height = 22
    section(ws, 5, 'A', RM_LAST, f'="① 报销人 × 报销项目（"&{m1}&"～"&{m2}&" 月）"', C_VIEW)
    section(ws, 5, RM_W, RM_L2, '="② 每人按月（"&P_年度&" 年 1～"&P_截止月&" 月）"', C_VIEW)
    hd = [('A', '报销人'), ('B', '部门')]
    header(ws, 6, hd + [(RM_OTH, '其余项目'), (RM_TOT, '合计'), (RM_DZ, '其中老板垫付的\n（垫付账户）'), (RM_BY, '其中备用金的'),
                        (RM_W, '报销人')] + [(c, f'{i + 1} 月') for i, c in enumerate(RM_MC)] + [(RM_YR, '全年')],
           GREEN_H, height=44)
    for s, c in enumerate(RM_SLOT, 1):
        hj = slot_j[s]
        put(ws, f'{c}6', f'=IF({hj}<0,"",INDEX($AO${IT0}:$AO${IT1},MATCH({hj},$AP${IT0}:$AP${IT1},0)))', F_HDR, fill(GREEN_H),
            align=ACW)
    cf = ws.conditional_formatting
    cf.add(f'{RM_OTH}6', FormulaRule(formula=[f'{nI}<={RE_SLOTS}'], font=Font(color='FF70AD47')))

    # ── 显示行 ──
    pcnt = f'$AR${PT0}:$AR${PT1}'
    for k in range(1, RM_NR + 1):
        r = RM_R0 + k - 1
        ix = f'$CA{r}'
        hset(ws, f'CA{r}', f'IF({k}<={nP},MATCH({k},{pcnt},0),IF({k}={nP}+1,-1,0))')
        cell(ws, f'A{r}', pick(ix, 'AO', PT0, N_PER + 1, ['"合计"']), align=AL)
        cell(ws, f'B{r}', pick(ix, 'AS', PT0, N_PER + 1, ['""']), F_NOTE, align=AC)
        cols = [(c, SLC[s]) for s, c in enumerate(RM_SLOT)] + [(RM_TOT, 'AT'), (RM_DZ, PDZ), (RM_BY, PBY)]
        for c, hc in cols:
            f = pick(ix, hc, PT0, N_PER + 1, [R2(f'SUM({PT(hc)})')])
            if c in RM_SLOT:
                f = f'IF({slot_j[RM_SLOT.index(c) + 1]}<0,"",{f})'
            cell(ws, f'{c}{r}', f, F_TXTB if c == RM_TOT else F_TXT, MONEY, AR)
        cell(ws, f'{RM_OTH}{r}', f'IF({nI}<={RE_SLOTS},"",{pick(ix, PO, PT0, N_PER + 1, [R2(f"SUM({PT(PO)})")])})', fmt=MONEY,
             align=AR)
        cell(ws, f'{RM_W}{r}', f'A{r}&""', align=AL)
        for c, hc in list(zip(RM_MC, PMC)) + [(RM_YR, PYR)]:
            cell(ws, f'{c}{r}', pick(ix, hc, PT0, N_PER + 1, [R2(f'SUM({PT(hc)})')]), F_TXTB if c == RM_YR else F_TXT, MONEY, AR)
    r0, r1 = RM_R0, RM_R0 + RM_NR - 1
    z0 = f'$CA{r0}'
    for rng in (f'A{r0}:{RM_LAST}{r1}', f'{RM_W}{r0}:{RM_L2}{r1}'):
        cf.add(rng, FormulaRule(formula=[f'{z0}=-1'], fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD, stopIfTrue=True))
        cf.add(rng, FormulaRule(formula=[f'{z0}>0'], border=CF_BD))
    for i, c in enumerate(RM_MC):
        cf.add(f'{c}6', FormulaRule(formula=[f'{i + 1}>P_截止月'], fill=cf_fill('FFBFBFBF')))
    hide(ws, *[CL(i) for i in range(CI(RM_L2) + 1, CI(RM_H['R2']) + 1)])
    finish(ws, RM_LAST, '6:6', sc('末行'), r1, freeze='C7', extra=[(RM_W, RM_L2, r1, sc('末行'))])


# ═══════════════════════════════ 个人垫付现金明细 ═══════════════════════════════
TIP_ADV = ('★ 老板（垫付人）个人掏钱付的，一个月一张，照原来的「个人垫付现金明细」打印、签字。黄格选垫付账户（只列【基础资料】②里类型是「垫付」的；空着＝第一个）、'
           '年月（填 9 或 2026-9；空着＝截止日那个月）。第 5 行往下是打印的部分（打印区到签字那一行）。'
           '「已报/未报」：公司还给他的钱（银行记的「垫付现金已报」，往来单位选垫付人，截止日以前的）按先后顺序先冲建账前垫的（基础资料 ⑦），'
           '再一笔一笔冲垫付账户的行。右边（不打印）：还没补到票的垫付，按过了多少天排；「无票」的合计（汇算清缴要调增）。')

AD_TT, AD_HD = 5, 6                         # 打印标题行、表头行
AD_L0 = 7
AD_NL = ADV_SHOW + 2
AD_BP0 = 7                                  # 右边补票跟进第 7 行起
AD_BPN = BP_SHOW + 2
AD_K0 = 2                                   # 每条资金一行的键：EX 本月行、EY 待补票行
AD_G1, AD_G2 = 'EX', 'EY'
AD_LIST = 'EV'                              # 垫付账户压紧清单（第 2 行起），EU 累计
ADV_TITLE_FILL = 'FF4F81BD'


def build_adv(ws):
    LAST = 'I'
    widths(ws, {'A': 5.5, 'B': 11, 'C': 13, 'D': 20, 'E': 15, 'F': 13, 'G': 9, 'H': 16, 'I': 18, 'J': 2,
                'K': 5.5, 'L': 11, 'M': 13, 'N': 18, 'O': 14, 'P': 12, 'Q': 9, 'R': 14})
    title(ws, '个人垫付现金明细', 'R', C_VIEW, TIP_ADV)
    ws.merge_cells('A3:B3')
    put(ws, 'A3', '垫付账户', F_KPI_L, LBL, align=AC)
    put(ws, 'B3', None, F_KPI_L, LBL, align=AC)
    ws.merge_cells('C3:D3')
    put(ws, 'C3', None, F_SEL, FILL_SEL, align=AC)
    put(ws, 'D3', None, F_SEL, FILL_SEL, align=AC)
    dv = DataValidation(type='list', formula1=f'${AD_LIST}$2:${AD_LIST}${N_ACC + 1}', allow_blank=True, showErrorMessage=False,
                        showInputMessage=True, promptTitle='提示', prompt='选垫付账户（基础资料②里类型是「垫付」的）；空着＝第一个')
    ws.add_data_validation(dv)
    dv.add('C3')
    selector(ws, 'E3', '年月', 'F3', None, fmt='General')
    dv = DataValidation(type='custom', formula1='TRUE', allow_blank=True, showErrorMessage=False, showInputMessage=True,
                        promptTitle='提示', prompt='填 9、2026-9 或 202609；只认本年度；空着＝截止日那个月')
    ws.add_data_validation(dv)
    dv.add('F3')
    home_link(ws, 'R3')
    ws.row_dimensions[3].height = 28

    keys = ['输入', '认出前', 'm', '认出', '垫付数', '选', 'a', '账户认出', '人', '人字', '出借人号', '建账前', '已还', '本月笔数', '前面',
            '显示数', '末行', '本月合计', '本月未报', '累计垫付', '待补数', '待补显示', '待补合计', '无票号', '无票合计', '补票末行']
    sc = Sc(ws, keys, cl='ES', cv='ET')
    ym_parse(sc, 'F3', 'P_截止月')
    m = sc('m')
    # 垫付账户压紧清单
    for i in range(N_ACC):
        a, r = i + 1, 2 + i
        ok = (f'IF(INDEX(账户_类型号,{a})<>5,0,IF(INDEX(账户_名称,{a})="",0,'
              f'IF(MATCH({esc(f"INDEX(账户_名称,{a})")},账户_名称,0)={a},1,0)))')
        hset(ws, f'EU{r}', ok + (f'+EU{r - 1}' if i else ''))
    for i in range(N_ACC):
        r = 2 + i
        hset(ws, f'{AD_LIST}{r}', f'IF({i + 1}>$EU${N_ACC + 1},"",INDEX(账户_名称,MATCH({i + 1},$EU$2:$EU${N_ACC + 1},0)))')
    sc.set('垫付数', f'$EU${N_ACC + 1}')
    sc.set('选', 'TRIM(C3&"")')
    sel = sc('选')
    first = f'IF({sc("垫付数")}=0,0,MATCH(1,$EU$2:$EU${N_ACC + 1},0))'
    found = f'IFERROR(MATCH({esc(sel)},账户_名称,0),0)'
    sc.set('a', f'IF({sel}="",{first},IF({found}=0,{first},IF(INDEX(账户_类型号,{found})<>5,{first},{found})))')
    a = sc('a')
    sc.set('账户认出', f'IF({sel}="",1,IF({found}=0,0,IF(INDEX(账户_类型号,{found})<>5,0,1)))')
    pn = norm(f'INDEX(账户_人员,{a})')
    sc.set('人', f'IF({a}=0,"",{unit_match(pn, "简称", pn)})')
    sc.set('人字', f'IF({a}=0,"",IF(TRIM(INDEX(账户_人员,{a})&"")="",INDEX(账户_名称,{a}),TRIM(INDEX(账户_人员,{a})&"")))')
    per = sc('人')
    sc.set('出借人号', f'IF({a}=0,0,INDEX(账户_出借人号,{a}))')
    lo = sc('出借人号')
    sc.set('建账前', R2(f'IF({lo}>0,SUMIFS(借期_金额,借期_出借人号,{lo},借期_性质,3,借期_有效,1),0)'))
    sc.set('已还', R2(f'IF({per}="",0,-SUMIFS(资_净额,资_原对方号,{KH224101},资_账户类型,"<>5",资_单位,{esc(per)},资_日期,"<="&P_截止))'))
    sc.set('本月笔数', f'COUNT(${AD_G1}${AD_K0}:${AD_G1}${AD_K0 + N_CASH - 1})')
    sc.set('显示数', f'MIN({sc("本月笔数")},{ADV_SHOW})')
    sc.set('末行', f'{AD_HD}+{sc("显示数")}+2')
    sc.set('本月合计', R2(f'IF({a}=0,0,-SUMIFS(资_净额,资_账户号,{a},资_月,{m}))'))
    sc.set('累计垫付', R2(f'{sc("建账前")}+IF({a}=0,0,-SUMIFS(资_净额,资_账户号,{a},资_日期,"<="&P_截止))'))
    sc.set('无票号', 'IFERROR(MATCH("无票",票种_名称,0),0)')
    sc.set('无票合计', R2(f'IF(OR({a}=0,{sc("无票号")}=0),0,-SUMIFS(资_净额,资_账户号,{a},资_票种号,{sc("无票号")},资_日期,"<="&P_截止))'))

    # ── 每条资金一行：G1＝这个账户这个月的排序键；G2＝这个账户待补票、没补掉、截止日以前的排序键 ──
    G1R = f'${AD_G1}${AD_K0}:${AD_G1}${AD_K0 + N_CASH - 1}'
    G2R = f'${AD_G2}${AD_K0}:${AD_G2}${AD_K0 + N_CASH - 1}'
    for i in range(N_CASH):
        r, n = AD_K0 + i, i + 1
        ws[f'EW{r}'] = n
        ws[f'EW{r}'].font = F_HELP
        n = f'$EW{r}'
        x = lambda nm: f'INDEX({nm},{n})'
        hset(ws, f'{AD_G1}{r}', f'IF(AND({a}>0,{x("资_账户号")}={a},{x("资_有效")}=1,{x("资_月")}={m}),{x("资_排序键")},"")')
        hset(ws, f'{AD_G2}{r}', (f'IF(AND({a}>0,{x("资_账户号")}={a},{x("资_待补票")}=1,{x("资_日期")}<=P_截止),'
                                 f'IF(COUNTIFS(票_补票行,{n},票_有效,1)=0,{x("资_排序键")},""),"")'))
    sc.set('待补数', f'COUNT({G2R})')
    sc.set('待补显示', f'MIN({sc("待补数")},{BP_SHOW})')
    sc.set('补票末行', f'{AD_HD}+{sc("待补显示")}+1')
    sc.set('待补合计', R2(f'-SUMIF({G2R},">0",资_净额)'))

    # ── 第 3、4 行说明 ──
    ws.merge_cells('G3:I3')
    put(ws, 'G3', (f'=IF({sc("垫付数")}=0,"⚠ 还没有垫付账户：到【基础资料】②账户加一行，类型选「垫付」",'
                   f'"＝ "&INDEX(账户_名称,{a})&"　"&P_年度&" 年 "&{m}&" 月"&IF({sc("账户认出")}=0,"（⚠ 「"&{sel}&"」不是垫付账户，先看第一个）","")'
                   f'&IF({sc("认出")}=0,"（⚠ 年月没认出来或不是本年度，先看截止日那个月）",""))'),
        F_NOTE, align=ALW, border=False)
    ws.conditional_formatting.add('G3', FormulaRule(formula=['ISNUMBER(FIND("⚠",$G$3))'], font=Font(bold=True, color='FFC00000')))
    ws.merge_cells('A4:I4')
    put(ws, 'A4', (f'=IF({a}=0,"",IF({sc("人")}="","⚠ 这个垫付账户在【基础资料】②里没填人员：公司已还多少对不上（已报/未报都按未报）",'
                   f'IF({lo}=0,"⚠ 「"&{sc("人字")}&"」不在【往来单位】里（类型要选出借人）：建账前垫的、利息都不算","")))'),
        F_NOTE, align=ALW, border=False)
    ws.conditional_formatting.add('A4', FormulaRule(formula=['ISNUMBER(FIND("⚠",$A$4))'], font=Font(bold=True, color='FFC00000')))

    # ── 打印部分 ──
    ws.merge_cells(f'A{AD_TT}:{LAST}{AD_TT}')
    put(ws, f'A{AD_TT}', (f'=IF({a}=0,"个人垫付现金明细",{sc("人字")}&" "&P_年度&"年"&{m}&"月个人垫付现金明细")'),
        Font(name=YH, sz=18, bold=True, color='FFFFFFFF'), fill(ADV_TITLE_FILL), align=AC, border=False)
    ws.row_dimensions[AD_TT].height = 40
    header(ws, AD_HD, [('A', '序号'), ('B', '日期'), ('C', '编号'), ('D', '摘要'), ('E', '往来单位'), ('F', '支出金额'),
                       ('G', '票种'), ('H', '已报/未报'), ('I', '备注')], ADV_TITLE_FILL, height=28)
    N, NN = sc('本月笔数'), sc('显示数')
    P0, RP = sc('建账前'), sc('已还')
    note_col = inref(SH_CASH, CASH_COLS['备注'])
    for k in range(1, AD_NL + 1):
        r = AD_L0 + k - 1
        ws[f'EC{r}'] = k
        ws[f'EC{r}'].font = F_HELP
        q, y, n, t = f'$EC{r}', f'$ED{r}', f'$EE{r}', f'$EF{r}'
        amt, cb, cov, un = f'$EG{r}', f'$EH{r}', f'$EI{r}', f'$EJ{r}'
        hset(ws, f'ED{r}', f'IF({q}>{NN},"",SMALL({G1R},{q}))')
        hset(ws, f'EE{r}', f'IF({y}="",0,MOD({y},{M}))')
        hset(ws, f'EF{r}', f'IF({n}>0,1,IF({q}={NN}+1,3,IF({q}={NN}+2,4,0)))')
        ix = lambda nm: f'INDEX({nm},{n})'
        hset(ws, f'EG{r}', f'IF({t}=1,-{ix("资_净额")},0)')
        hset(ws, f'EH{r}', f'IF({t}=1,{P0}-SUMIFS(资_净额,资_账户号,{a},资_排序键,"<"&{y}),0)')
        hset(ws, f'EI{r}', f'IF({t}=1,ROUND(MAX(0,MIN({amt},{RP}-{cb})),2),0)')
        hset(ws, f'EJ{r}', f'IF(AND({t}=1,{amt}>0),ROUND({amt}-{cov},2),0)')
        cell(ws, f'A{r}', f'IF({t}=1,{q},"")')
        cell(ws, f'B{r}', f'IF({t}=1,{ix("资_日期")},"")', fmt=DATE)
        cell(ws, f'C{r}', f'IF({t}=1,MID({ix("资_编号键")},2,99),"")')
        cell(ws, f'D{r}', f'IF({t}=1,{ix("资_显示摘要")},IF({t}=3,"合计",IF({t}=4,"制表人：","")))', align=AL)
        cell(ws, f'E{r}', (f'IF({t}=1,IF({ix("资_单位")}<>"",{ix("资_单位")},IF({ix("资_报销人号")}>0,'
                           f'INDEX(人员_姓名,{ix("资_报销人号")}),"")),IF({t}=4,"审核：",""))'), align=AL)
        cell(ws, f'F{r}', f'IF({t}=1,{amt},IF({t}=3,{sc("本月合计")},""))', fmt='#,##0.00;[Red]-#,##0.00', align=AR)
        tk = ix('资_票种号')
        cell(ws, f'G{r}', (f'IF({t}<>1,"",IF({tk}=0,"",IF(AND({ix("资_待补票")}=1,COUNTIFS(票_补票行,{n},票_有效,1)>0),"已补票",'
                           f'INDEX(票种_名称,{tk}))))'))
        cell(ws, f'H{r}', (f'IF({t}=1,IF({amt}<=0,"",IF({un}<=0,"已报",IF({cov}<=0,"未报","部分（还差 "&TEXT({un},"#,##0.00")&"）"))),'
                           f'IF({t}=3,IF({N}=0,"",IF({sc("本月未报")}<=0,"全部已报","未报 "&TEXT({sc("本月未报")},"#,##0.00"))),""))'))
        cell(ws, f'I{r}', f'IF({t}=1,INDEX({note_col},{n}+1)&"","")', align=AL)
    d0, d1 = AD_L0, AD_L0 + AD_NL - 1
    sc.set('本月未报', R2(f'SUM($EJ${d0}:$EJ${d1})'))
    cf = ws.conditional_formatting
    t0 = f'$EF{d0}'
    rng = f'A{d0}:{LAST}{d1}'
    cf.add(rng, FormulaRule(formula=[f'{t0}=3'], font=Font(bold=True), border=CF_BD, stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=4'], font=Font(bold=True), stopIfTrue=True))
    cf.add(f'H{d0}:H{d1}', FormulaRule(formula=[f'AND({t0}=1,LEFT($H{d0},1)<>"已")'], font=Font(color='FFC00000'), border=CF_BD,
                                       stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=1'], border=CF_BD))
    for r in range(d0, d1 + 1):
        ws.row_dimensions[r].height = 22

    # ── 右边：补票跟进（不打印） ──
    ws.merge_cells('K3:Q3')
    put(ws, 'K3', (f'=IF({a}=0,"","到截止日 "&TEXT(P_截止,"yyyy-mm-dd")&"："&{sc("人字")}&" 一共垫付 "&TEXT({sc("累计垫付")},"#,##0.00")'
                   f'&" 元（含建账前 "&TEXT({P0},"#,##0.00")&"），公司已还 "&TEXT({RP},"#,##0.00")&" 元，还欠 "'
                   f'&TEXT({sc("累计垫付")}-{RP},"#,##0.00")&" 元")'), F_NOTE, align=ALW, border=False)
    ws.merge_cells('K4:R4')
    put(ws, 'K4', (f'=IF({a}=0,"",IF({sc("无票合计")}=0,"「无票」的垫付：没有",'
                   f'"「无票」的垫付合计 "&TEXT({sc("无票合计")},"#,##0.00")&" 元（到截止日）：没有合规发票，'
                   f'汇算清缴时这些费用不能税前扣除，要纳税调增"))'), F_NOTE, align=ALW, border=False)
    ws.conditional_formatting.add('K4', FormulaRule(formula=['ISNUMBER(FIND("调增",$K$4))'], font=Font(bold=True, color='FFC00000')))
    ws.row_dimensions[4].height = 30
    section(ws, AD_TT, 'K', 'R', (f'="补票跟进（不打印）：待补票、还没在发票登记补上的，共 "&{sc("待补数")}&" 笔 "'
                                  f'&TEXT({sc("待补合计")},"#,##0.00")&" 元"&IF({sc("待补数")}>{BP_SHOW},"，⚠ 只列前 {BP_SHOW} 笔","")'), C_VIEW)
    header(ws, AD_HD, [('K', '序号'), ('L', '日期'), ('M', '编号'), ('N', '摘要'), ('O', '往来单位'), ('P', '金额'),
                       ('Q', '已过天数'), ('R', '收支类别')], GREEN_H, height=28)
    BN, BNN = sc('待补数'), sc('待补显示')
    for k in range(1, AD_BPN + 1):
        r = AD_BP0 + k - 1
        ws[f'EL{r}'] = k
        ws[f'EL{r}'].font = F_HELP
        q, y, n, t = f'$EL{r}', f'$EM{r}', f'$EN{r}', f'$EO{r}'
        hset(ws, f'EM{r}', f'IF({q}>{BNN},"",SMALL({G2R},{q}))')
        hset(ws, f'EN{r}', f'IF({y}="",0,MOD({y},{M}))')
        hset(ws, f'EO{r}', f'IF({n}>0,1,IF({q}={BNN}+1,3,0))')
        ix = lambda nm: f'INDEX({nm},{n})'
        cell(ws, f'K{r}', f'IF({t}=1,{q},"")', F_NOTE)
        cell(ws, f'L{r}', f'IF({t}=1,{ix("资_日期")},IF({t}=3,"合计",""))', fmt=DATE)
        cell(ws, f'M{r}', f'IF({t}=1,MID({ix("资_编号键")},2,99),"")')
        cell(ws, f'N{r}', f'IF({t}=1,{ix("资_显示摘要")},IF({t}=3,"共 "&{BN}&" 笔",""))', align=AL)
        cell(ws, f'O{r}', (f'IF({t}=1,IF({ix("资_单位")}<>"",{ix("资_单位")},IF({ix("资_报销人号")}>0,'
                           f'INDEX(人员_姓名,{ix("资_报销人号")}),"")),"")'), align=AL)
        cell(ws, f'P{r}', f'IF({t}=1,-{ix("资_净额")},IF({t}=3,{sc("待补合计")},""))', fmt=MONEY, align=AR)
        cell(ws, f'Q{r}', f'IF({t}=1,P_截止-{ix("资_日期")},"")', fmt='0')
        cell(ws, f'R{r}', f'IF({t}=1,INDEX(类别_名称,{ix("资_类别号")}),"")', align=AL)
    b0, b1 = AD_BP0, AD_BP0 + AD_BPN - 1
    t0 = f'$EO{b0}'
    rng = f'K{b0}:R{b1}'
    cf.add(rng, FormulaRule(formula=[f'{t0}=3'], fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD, stopIfTrue=True))
    cf.add(f'Q{b0}:Q{b1}', FormulaRule(formula=[f'AND({t0}=1,$Q{b0}>90)'], font=Font(bold=True, color='FFC00000'), border=CF_BD,
                                       stopIfTrue=True))
    cf.add(rng, FormulaRule(formula=[f'{t0}=1'], border=CF_BD))
    hide(ws, *[CL(i) for i in range(CI('S'), CI(AD_G2) + 1)])
    finish(ws, LAST, f'{AD_TT}:{AD_HD}', sc('末行'), d1, freeze=None, top=AD_TT, landscape=False)
    ws.freeze_panes = f'A{AD_L0}'


def build(wb, ctx=None):
    build_day(wb[SH_DAY])
    build_mon(wb[SH_MON])
    build_fee(wb[SH_FEE])
    build_reim(wb[SH_REIM])
    build_adv(wb[SH_ADV])
