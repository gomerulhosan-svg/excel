# -*- coding: utf-8 -*-
"""A072 化工 · 查看表（绿）：进销存、成本费用分析（每吨成本）、借款明细及利息。
只用 layout 的定义名称（隐藏层、引擎 _料 _品 _本 _息）、s_cost 加的 息_*、s_je 加的 矩_汇借/矩_汇贷/矩_汇期初 和本表格子取数；借款明细的「备注」另外按第 n 条取
资金台帐备注列、基础资料 ⑦ 说明列（INDEX(列, n+1)，插行删行不错位）。没有别的表引用这三张表。

【进销存】B3 月份（空＝截止日那个月；认 9、9月、2026-9、202609、日期）。三段：① 原料（类别 原料，以及辅料/包装物里管库存的
   ＝进货即耗用「否」；进货即耗用的辅料本月有数也列出来，免得漏）② 回收料 ③ 产品；每段：标题、表头、每个品种一行、合计；
   最后 ④ 核对：原料＋回收料 月末额 ＝ 1403 科目余额、产品月末额 ＝ 1405（_矩 汇期初＋1～m 月汇借－汇贷，含下级）。负库存 ✗ 标红。
   合计行的数量：这一段全是「吨」才加（单位空着算吨），不然空着。
【成本费用分析】B3 产品（空＝参数默认产品；名称、别名都认）。列＝1～12 月＋本年（建账月以前、截止月以后空着）。
   ① 产量、废料吨、各原料投料吨（＝_本 槽投：生产日报这个产品第 j 列）、投料合计、单耗（投料÷产量，4 位小数）、
      折纯收率（产量×折纯系数×成品含量÷Σ投料×原料含量×100；用到的原料没填含量或产品没填折纯/含量＝「缺含量」）；
   ② 月初在产＋各原料金额（其中进货运费）＋各制造费用明细（本_分到）－废料－月末在产＝完工成本，跟 _本 完工核对（差 0）；
      每吨制造成本、各项每吨（÷产量）。原料金额＝_料 投额按各产品投料量累计舍入分（跟 s_cost 同一个算法，合计正好等于 _本 DM）；
      其中运费＝_料 投运 同样分。
   ③ 销量、不含税收入、均价、销售成本（品_销额）、销货运费（_矩 540102 或 560101，看参数；按各产品销量累计舍入分）、
      期间费用（5601＋5602＋5603，运费记销售费用时扣掉 560101；同样按销量分；本年＝全年合计按全年销量分，不是各月相加）、
      完全成本/吨、毛利/吨（运费进主营成本时扣运费）、
      经营利润/吨、经营利润；核对：全部产品（每吨×销量合计）vs 利润表同口径（主营收入－主营成本－三项费用，_矩）。
   「本年」每吨＝本年金额÷本年产量（销量）。每吨数 ROUND 到 2 位、单耗 4 位（只是显示）。
【借款明细及利息】上半（第 7 行起，动态行数）：每个出借人（_表 出借人清单）一行：类别、借款本金、借款年利率、投资本金、投资年利率、
   本月（截止月）利息、本年利息、已计提未付（期初应付息＋累计利息－已付息，借款＋投资）、到截止日应计未付、到期日（30 天内到期/过期标红）、提醒；合计。
   下半（紧接着）：D3 选出借人（空＝第一个）；按日期列 ⑦ 建账前的＋资金台帐这个出借人的本金事件（截止日以前）：
   日期｜摘要｜性质｜借入（打款）｜归还｜本金余额｜计息天数（截止日＋1－日期）｜年利率｜本笔利息（±金额×年利率÷一年天数×天数）｜来源｜｜备注；
   合计、减已付利息、＝到截止日应计未付，跟上半核对（差 0）。最多列 EV 笔（多了 ⚠，合计只算列出来的）。
   上下两半同一列的格式不一样：利率、日期、天数一律用文字（TEXT），金额列统一金额格式。

隐藏（各表右边，F_HELP 字体）：
   进销存：U.. 候选行表（类型、显示、累计、名称、单位、C..S 各列、neg：U..AQ）；AS/AT 标量；AV/AW/AX 显示行的 k、g、类型。
   成本费用分析：Q/R 标量；T.. 槽表（第 j 槽：料号、含量、名称、显示、投料吨 12＋本年、金额 12＋本年、运费 12＋本年）；
      BP.. 槽投块（第 (j−1)×N_PRD＋q 行：12 月 SUMIF，跟 s_cost 同式）；CD.. 候选行表（第 2 行是月份有效标志）；DC/DD/DE 显示行。
   借款明细及利息：O/P 标量；R..AH 出借人表；AJ 本金事件排序键（⑦ N_LO 行＋资金台帐 N_CASH 行）；AL..BA 事件表（EV 行）；BC/BD/BE 显示行。"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from common import *
from layout import *                      # layout 的 norm/esc 盖过 common 的
from s_mirror import unit_match

MS = range(1, 13)
GREEN_H = 'FF70AD47'
LBL = fill('FFD9E1F2')
F_KV = Font(name=YH, sz=10, bold=True, color='FFC00000')
_thin = Side(style='thin', color='FFBFBFBF')
CF_BD = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
MNY = '#,##0.00;[Red]-#,##0.00;"-"'
QTY = '#,##0.000;[Red]-#,##0.000;"-"'
NUM4 = '#,##0.00##;[Red]-#,##0.00##;"-"'
EV = 120 if SMALL else 400                     # 借款明细下半最多列几笔本金事件
TINY = '0.0000000001'
KH224101 = kh('224101')


def cf_fill(rgb):
    return PatternFill('solid', fgColor=rgb, bgColor=rgb)


class Sc:
    """隐藏标量：cl 列名字、cv 列值（第 r0 行起）"""

    def __init__(self, ws, keys, cl, cv, r0=3):
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


def kpi(ws, lbl_cell, lbl, val_cell, f, fmt=MNY, merge_to=None):
    put(ws, lbl_cell, lbl, F_KPI_L, LBL, align=AC)
    put(ws, val_cell, '=' + f, F_KV, FILL_AUTO, fmt, AR)
    if merge_to:
        ws.merge_cells(f'{val_cell}:{merge_to}')
        for i in range(CI(val_cell.rstrip('0123456789')) + 1, CI(merge_to.rstrip('0123456789')) + 1):
            ws.cell(row=int(val_cell.lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ')), column=i).border = BD


def note(ws, c1, c2, f, row=3):
    ws.merge_cells(f'{c1}{row}:{c2}{row}')
    put(ws, f'{c1}{row}', f, F_NOTE, align=ALW, border=False)


def finish(ws, last, end_cell, cap_row, freeze, titles=None):
    """打印：A4 横向一页宽、页脚页码；打印区域按内容变长（A1 到 last 列 end_cell 行）"""
    ws.freeze_panes = freeze
    print_setup(ws, titles, landscape=True)
    ws.oddFooter.center.text = '第 &P 页 共 &N 页'
    q = f"'{ws.title}'"
    ws.defined_names['Print_Area'] = DefinedName('Print_Area', attr_text=f'{q}!$A$1:INDEX({q}!${last}$1:${last}${cap_row},{q}!{end_cell})')


def ecol(fields, field, m, col0='C'):
    return CL(CI(col0) + fields.index(field) * 12 + m - 1)


def erng(sh, fields, field, m, n):
    """引擎某字段第 m 月那一列（第 2..n+1 行）"""
    c = ecol(fields, field, m)
    return f'{sh}!${c}$2:${c}${n + 1}'


# _矩：s_je 的 矩_借/矩_贷/矩_期初 是每个科目「自己」记的；一级科目连下级的合计在 s_je 加的 矩_汇借/矩_汇贷/矩_汇期初
# （一级＝下级合计＋自己，二级＝自己）。这里一律用 汇 的（只有 acc、bal 两处用到）。
def acc(code, m):
    """科目（含下级）第 m 月 借－贷"""
    kn = kh(code)
    return f'IF({kn}=0,0,INDEX(矩_汇借,{kn},{m})-INDEX(矩_汇贷,{kn},{m}))'


def bal(code, m):
    """科目（含下级）月末余额（借正贷负）＝期初＋1～m 月借－贷"""
    kn = kh(code)
    return (f'IF({kn}=0,0,INDEX(矩_汇期初,{kn})+SUM(INDEX(矩_汇借,{kn},1):INDEX(矩_汇借,{kn},{m}))'
            f'-SUM(INDEX(矩_汇贷,{kn},1):INDEX(矩_汇贷,{kn},{m})))')


def ym_parse(sc, cell_, default):
    """月份黄格 → sc 里 输入、认出前（年×100＋月；看不懂 0）、认出（0＝看不懂或不是本年度）、m（空/看不懂 → default）"""
    t = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM({cell_}&""),"月",""),"年","-"),"/","-")'
    sc.set('输入', f'IF(ISNUMBER({cell_}),{cell_},IFERROR(--{t},0))')
    v = sc('输入')
    sc.set('认出前', (f'IF(AND({v}>=1,{v}<=12,INT({v})={v}),P_年度*100+{v},'
                    f'IF(AND({v}>=200001,{v}<=209912,INT({v})={v},MOD({v},100)>=1,MOD({v},100)<=12),{v},'
                    f'IF(AND({v}>=36526,{v}<73051),YEAR({v})*100+MONTH({v}),0)))'))
    y = sc('认出前')
    sc.set('认出', f'IF(TRIM({cell_}&"")="",1,IF(AND({y}>0,INT({y}/100)=P_年度),1,0))')
    sc.set('m', f'IF({sc("认出")}=0,{default},IF(TRIM({cell_}&"")="",{default},MOD({y},100)))')


# ═══════════════════════════════ 候选行表 ＋ 显示区（进销存、成本费用分析共用） ═══════════════════════════════
class Lines:
    """隐藏的候选行表（第 r0 行起一行一条）：类型、显示（1/0）、累计、标签、单位、各值列、nz（各月 SUMSQ，判断有没有数）。
       显示区第 k 行＝第 k 条要显示的候选行（MATCH(k, 累计, 0)）。值：'=…' 公式、别的字符串＝文字常量、callable(r)。"""

    def __init__(self, ws, r0, col0, keys, nz_keys=()):
        self.ws, self.r0, self.keys, self.nz_keys = ws, r0, keys, list(nz_keys)
        c = CI(col0)
        self.c = {}
        for k in ['类型', '显示', '累计', '标签', '单位'] + keys + (['nz'] if nz_keys else []):
            self.c[k] = CL(c)
            c += 1
        self.end = c - 1
        self.rows = []
        for k in self.c:
            _lab(ws, f'{self.c[k]}{r0 - 1}', k)

    def _w(self, a, v, r):
        if callable(v):
            v = v(r)
        if v is None or v == '':
            return
        self.ws[a] = v
        self.ws[a].font = F_HELP

    def add(self, typ, label, unit='', vals=None, show=1):
        r = self.r0 + len(self.rows)
        self.rows.append(r)
        self.ws[f'{self.c["类型"]}{r}'] = typ
        self.ws[f'{self.c["类型"]}{r}'].font = F_HELP
        self._w(f'{self.c["标签"]}{r}', label, r)
        self._w(f'{self.c["单位"]}{r}', unit, r)
        for k in self.keys:
            self._w(f'{self.c[k]}{r}', (vals or {}).get(k), r)
        if self.nz_keys:
            a, b = self.c[self.nz_keys[0]], self.c[self.nz_keys[-1]]
            self._w(f'{self.c["nz"]}{r}', f'=SUMSQ({a}{r}:{b}{r})', r)
        self.set_show(r, show)
        return r

    def set_show(self, r, show):
        self._w(f'{self.c["显示"]}{r}', show, r)

    def a(self, k, r):
        return f'${self.c[k]}${r}'

    def nzf(self, r):
        return f'{self.c["nz"]}{r}>{TINY}'

    def finish(self):
        S, C = self.c['显示'], self.c['累计']
        for i, r in enumerate(self.rows):
            self._w(f'{C}{r}', f'={S}{r}' + (f'+{C}{r - 1}' if i else ''), r)
        self.last = self.rows[-1]

    def rng(self, k, r1=None, r2=None):
        return f'${self.c[k]}${r1 or self.r0}:${self.c[k]}${r2 or self.last}'

    def total(self):
        return f'${self.c["累计"]}${self.last}'


def _lab(ws, a, t):
    ws[a] = t
    ws[a].font = F_HELP


def show_list(ws, L, d0, nr, colmap, hk, hg, ht):
    """显示区：第 d0 行起 nr 行；hk/hg/ht＝隐藏的 k、候选行号 g、类型 列；colmap：显示列 → 候选表的键"""
    for k in range(1, nr + 1):
        r = d0 + k - 1
        hset(ws, f'{hk}{r}', k)
        g = f'${hg}{r}'
        hset(ws, f'{hg}{r}', f'IF({k}>{L.total()},0,MATCH({k},{L.rng("累计")},0))')
        hset(ws, f'{ht}{r}', f'IF({g}=0,0,INDEX({L.rng("类型")},{g}))')
        for col, key in colmap.items():
            rg = L.rng(key)
            c = ws[f'{col}{r}']
            c.value = f'=IF({g}=0,"",IF(INDEX({rg},{g})="","",INDEX({rg},{g})))'


def type_cf(ws, rng, t0, styles):
    """按类型上色：styles＝[(类型, 底色, 字体, stop)]（先加的优先）"""
    cf = ws.conditional_formatting
    for ty, bg, font, stop in styles:
        kw = dict(formula=[f'{t0}={ty}'], border=CF_BD, stopIfTrue=stop)
        if bg:
            kw['fill'] = cf_fill(bg)
        if font is not None:
            kw['font'] = font
        cf.add(rng, FormulaRule(**kw))


F_SECW = Font(bold=True, color='FFFFFFFF')
F_B = Font(bold=True)
F_IT = Font(italic=True, color='FF595959')


# ═══════════════════════════════ 进销存 ═══════════════════════════════
TIP_STK = ('★ 黄格选月份（1～12、9月、2026-9、202609；空着＝截止日那个月），看这个月每个品种的进、销、存（月末一次加权平均）。'
           '① 原料：类别是原料的，加上辅料、包装物里要管库存的（物料档案「进货即耗用」填否）；② 回收料（废料）；③ 产品。'
           '金额都不含能抵的进项税；进货金额含运费（「其中运费」单列）；待定价的进货按参考单价暂估（⚠）；盘盈亏＝月末盘点实盘－账面。'
           '月末数量是负数＝出得比进得多（✗ 标红）：多半是进货没登、投料或销售登多了。最下面 ④ 核对：原料＋回收料月末金额应等于原材料科目余额、产品月末金额应等于库存商品科目余额。')
STK_LAST = 'S'
STK_D0 = 5
STK_KEYS = [CL(CI('C') + i) for i in range(17)]          # C..S
STK_QTY = ['C', 'E', 'H', 'J', 'L', 'N', 'P']
STK_AMT = ['D', 'F', 'G', 'I', 'K', 'M', 'O', 'Q', 'R']
RM_F = dict(C='期初量', D='期初额', E='进量', F='进额', G='进运', H='回量', I='回额', J='投量', K='投额', L='销量', M='销额',
            N='盘量', O='盘额', P='末量', Q='末额', R='单价')
PR_F = dict(C='期初量', D='期初额', E='完量', F='完额', H='购量', I='购额', L='销量', M='销额', N='盘量', O='盘额', P='末量',
            Q='末额', R='单价')
HDR_RM = dict(C='月初量', D='月初额', E='进货量', F='进货额', G='其中运费', H='回收量', I='回收额', J='投料量', K='投料额',
              L='销售量', M='销售额', N='盘盈亏量', O='盘盈亏额', P='月末量', Q='月末额', R='单价', S='提醒')
HDR_PR = dict(C='月初量', D='月初额', E='完工量', F='完工额', H='外购量', I='外购额', L='销售量', M='销售额', N='盘盈亏量',
              O='盘盈亏额', P='月末量', Q='月末额', R='单价', S='提醒')


def build_stk(ws):
    LAST = STK_LAST
    widths(ws, dict({'A': 24, 'B': 6, 'S': 40, 'T': 2}, **{c: 10.5 for c in STK_QTY}, **{c: 12.5 for c in STK_AMT}))
    title(ws, '=IF(P_公司名称="","",P_公司名称&"　")&"进销存"', LAST, C_VIEW, TIP_STK)
    selector(ws, 'A3', '选月份', 'B3', None, '"' + ','.join(str(i) for i in MS) + '"', fmt='General',
             prompt='选 1～12 月（也认 9月、2026-9、202609）；空着＝截止日那个月')
    ws.merge_cells('B3:C3')
    home_link(ws, f'{LAST}3')
    ws.row_dimensions[3].height = 24

    sc = Sc(ws, ['输入', '认出前', '认出', 'm', 'A数', 'B数', 'C数', '料末额', '品末额', '1403', '1405', '差1', '差2', '负库存',
                 '显示数', '末行'], cl='AS', cv='AT')
    ym_parse(sc, 'B3', 'P_截止月')
    m = sc('m')
    L = Lines(ws, 3, 'U', STK_KEYS + ['neg'])
    T_SEC, T_HDR, T_ITEM, T_SUB, T_BLANK, T_CHK = 1, 2, 3, 4, 0, 6

    def item_rm(k, sect):
        def neg(r):
            return f'=IF(AND({L.c["显示"]}{r}=1,ROUND({L.c["P"]}{r},4)<0),1,0)'
        v = {c: f'=INDEX(料_{f},{k},{m})' for c, f in RM_F.items()}
        v['S'] = (f'=IF(ROUND(INDEX(料_末量,{k},{m}),4)<0,"✗ 负库存：出库比库存多，先补登进货或核对投料、销售",'
                  f'IF(COUNTIFS(进_料月键,{k}*100+{m},进_待定价,1)>0,"⚠ 有待定价进货，金额按参考单价暂估",""))')
        v['neg'] = neg
        ok = f'INDEX(料档_有效,{k})=1'
        kd = f'INDEX(料档_类别号,{k})'
        if sect == 'A':
            show = lambda r: (f'=IF(AND({ok},{kd}<>4,OR({kd}=1,INDEX(料档_耗用,{k})=0,'
                              f'SUMSQ({L.c["C"]}{r}:{L.c["Q"]}{r})>{TINY})),1,0)')
        else:
            show = f'=IF(AND({ok},{kd}=4),1,0)'
        return L.add(T_ITEM, f'=INDEX(料档_名称,{k})&""', f'=IF(TRIM(INDEX(料档_单位,{k})&"")="","吨",TRIM(INDEX(料档_单位,{k})&""))',
                     v, show)

    def item_pr(p):
        v = {c: f'=INDEX(品_{f},{p},{m})' for c, f in PR_F.items()}
        v['S'] = (f'=IF(ROUND(INDEX(品_末量,{p},{m}),4)<0,"✗ 负库存：卖出比库存多，先补登生产日报或核对销售",'
                  f'IF(COUNTIFS(进_品月键,{p}*100+{m},进_待定价,1)>0,"⚠ 有待定价外购，金额按参考单价暂估",""))')
        v['neg'] = lambda r: f'=IF(AND({L.c["显示"]}{r}=1,ROUND({L.c["P"]}{r},4)<0),1,0)'
        return L.add(T_ITEM, f'=INDEX(品档_名称,{p})&""', '吨', v, f'=IF(INDEX(品档_有效,{p})=1,1,0)')

    def subtotal(lab, r1, r2, cnt, skip=()):
        S_ = L.rng('显示', r1, r2)
        U_ = L.rng('单位', r1, r2)
        same = f'COUNTIFS({S_},1,{U_},"<>吨")=0'
        v = {}
        for c in STK_AMT[:-1]:
            if c not in skip:
                v[c] = f'=ROUND(SUMIF({S_},1,{L.rng(c, r1, r2)}),2)'
        for c in STK_QTY:
            if c not in skip:
                v[c] = f'=IF({same},ROUND(SUMIF({S_},1,{L.rng(c, r1, r2)}),6),"")'
        return L.add(T_SUB, lab, f'=IF({same},"吨","")', v, f'=IF({cnt}>0,1,0)')

    sec = {}
    for key, ttl, hdr in (('A', '① 原料', HDR_RM), ('B', '② 回收料', HDR_RM), ('C', '③ 产品', HDR_PR)):
        cnt = sc(f'{key}数')
        show = f'=IF({cnt}>0,1,0)'
        L.add(T_SEC, f'="{ttl}（"&{m}&" 月）"', '', None, show)
        L.add(T_HDR, '名称', '单位', dict(hdr), show)
        r1 = L.r0 + len(L.rows)
        if key == 'C':
            for p in range(1, N_PRD + 1):
                item_pr(p)
        else:
            for k in range(1, N_RM + 1):
                item_rm(k, key)
        r2 = L.r0 + len(L.rows) - 1
        sec[key] = (r1, r2)
        sc.set(f'{key}数', f'SUM({L.rng("显示", r1, r2)})')
        subtotal({'A': '原料合计', 'B': '回收料合计', 'C': '产品合计'}[key], r1, r2, cnt, ('G', 'J', 'K') if key == 'C' else ())
        L.add(T_BLANK, None, '', None, show)
    (a1, a2), (c1, c2) = sec['A'], sec['C']
    sc.set('料末额', f'ROUND(SUM({L.rng("Q", a1, a2)}),2)')
    sc.set('品末额', f'ROUND(SUM({L.rng("Q", c1, c2)}),2)')
    sc.set('1403', f'ROUND({bal("1403", m)},2)')
    sc.set('1405', f'ROUND({bal("1405", m)},2)')
    sc.set('差1', f'ROUND({sc("料末额")}-{sc("1403")},2)')
    sc.set('差2', f'ROUND({sc("品末额")}-{sc("1405")},2)')
    late = f'IF({m}>P_截止月,"截止日以后的月份：月末凭证还没出，先不用核对",IF({m}<P_建账月,"建账月以前，没有数",'
    ok_txt = lambda d: f'IF(ABS({d})<0.005,"✓ 对上了",{late}"✗ 对不上：先看数据校验（凭证、成本计算）")))'
    L.add(T_SEC, f'="④ 核对（"&{m}&" 月末）"')
    L.add(T_CHK, '本表：原料＋回收料 月末额', '元', dict(Q=f'={sc("料末额")}'))
    L.add(T_CHK, '原材料 1403 科目余额', '元', dict(Q=f'={sc("1403")}'))
    L.add(T_CHK, '差（应为 0）', '元', dict(Q=f'={sc("差1")}', S='=' + ok_txt(sc('差1'))))
    L.add(T_CHK, '本表：产品 月末额', '元', dict(Q=f'={sc("品末额")}'))
    L.add(T_CHK, '库存商品 1405 科目余额', '元', dict(Q=f'={sc("1405")}'))
    L.add(T_CHK, '差（应为 0）', '元', dict(Q=f'={sc("差2")}', S='=' + ok_txt(sc('差2'))))
    L.finish()
    sc.set('负库存', f'SUM({L.rng("neg")})')
    NR = len(L.rows)
    sc.set('显示数', L.total())
    sc.set('末行', f'{STK_D0 - 1}+{sc("显示数")}')

    # ── 第 3、4 行 ──
    note(ws, 'D', 'R', (f'="＝"&P_年度&" 年 "&{m}&" 月"&IF({sc("认出")}=0,"　⚠ 月份看不懂或不是本年度（先按截止日那个月）","")'
                        f'&IF({m}>P_截止月,"　（在截止日 "&TEXT(P_截止,"yyyy-mm-dd")&" 以后：盘点、月末凭证还没算）",'
                        f'IF({m}<P_建账月,"　（建账月以前，没有数）",""))'))
    kpi(ws, 'A4', '原料＋回收料月末额', 'B4', sc('料末额'), MNY, merge_to='C4')
    kpi(ws, 'D4', '产品月末（吨）', 'E4', f'ROUND(SUM({L.rng("P", c1, c2)}),3)', QTY, merge_to='F4')
    kpi(ws, 'G4', '产品月末额', 'H4', sc('品末额'), MNY, merge_to='I4')
    kpi(ws, 'J4', '负库存（种）', 'K4', sc('负库存'), '0', merge_to='L4')
    put(ws, 'M4', '核对', F_KPI_L, LBL, align=AC)
    ws.merge_cells('N4:R4')
    put(ws, 'N4', (f'=IF(AND(ABS({sc("差1")})<0.005,ABS({sc("差2")})<0.005),"✓ 原料、产品月末额跟科目余额对上了",'
                   f'IF({m}>P_截止月,"截止日以后的月份，先不用核对","✗ 跟科目余额差 "&TEXT({sc("差1")},"#,##0.00")&"（原料）、"'
                   f'&TEXT({sc("差2")},"#,##0.00")&"（产品），看 ④"))'), F_KV, FILL_AUTO, align=AL)
    ws.row_dimensions[4].height = 22

    # ── 显示区 ──
    colmap = {'A': '标签', 'B': '单位', **{c: c for c in STK_KEYS}}
    show_list(ws, L, STK_D0, NR, colmap, 'AV', 'AW', 'AX')
    d1 = STK_D0 + NR - 1
    for r in range(STK_D0, d1 + 1):
        for col in ['A', 'B'] + STK_KEYS:
            c = ws[f'{col}{r}']
            c.font = F_TXT
            if col in STK_QTY:
                c.number_format, c.alignment = QTY, AR
            elif col in STK_AMT:
                c.number_format, c.alignment = MNY, AR
            else:
                c.alignment = AL if col in ('A', 'S') else AC
    t0 = f'$AX{STK_D0}'
    rng = f'A{STK_D0}:{LAST}{d1}'
    cf = ws.conditional_formatting
    cf.add(f'P{STK_D0}:P{d1}', FormulaRule(formula=[f'AND({t0}=3,N($P{STK_D0})<0)'], fill=cf_fill('FFFFC7CE'),
                                           font=Font(bold=True, color='FFC00000')))
    cf.add(f'S{STK_D0}:S{d1}', FormulaRule(formula=[f'LEFT($S{STK_D0},1)="✗"'], font=Font(bold=True, color='FFC00000')))
    type_cf(ws, rng, t0, [(1, C_VIEW, F_SECW, True), (2, GREEN_H, F_SECW, True), (4, 'FFDDEBF7', F_B, True),
                          (6, 'FFFFF2CC', None, True), (3, None, None, True)])
    hide(ws, *[CL(i) for i in range(CI('U'), CI('AX') + 1)])
    finish(ws, LAST, sc('末行'), d1, 'C5')


# ═══════════════════════════════ 成本费用分析 ═══════════════════════════════
TIP_COST = ('★ 选一个产品（黄格；空着＝基础资料参数「默认产品」），看它 1～12 月和本年的每吨成本（建账月以前、截止日以后的月份空着）。'
            '① 产量、各原料投料吨（生产日报）、单耗（吨原料/吨成品）、折纯收率（产量×折纯系数×成品含量÷Σ投料×原料含量；物料档案没填含量显示「缺含量」）。'
            '② 制造成本：月初在产＋各原料（其中进货运费）＋分到这个产品的各项制造费用－废料－月末在产＝完工成本（跟成本计算核对，差应为 0），÷产量＝每吨制造成本。'
            '③ 按销量：销售成本、销货运费、期间费用（几个产品时按销量分）、完全成本、毛利、经营利润；最后把所有产品乘回销量跟利润表核对'
            '（没有产品销量的月份，运费、费用当月分不出去，差不为 0；本年的运费、期间费用按全年销量分，本年对得上）。'
            '「本年」每吨＝本年金额合计÷本年产量（或销量），不是各月平均。')
COST_MC = [CL(CI('C') + i) for i in range(12)]          # C..N
COST_LAST = 'O'
COST_D0 = 6
SL0 = 3                                                   # 槽表第 j 槽在第 SL0+j−1 行
ST0 = 3                                                   # 槽投块第 (j−1)×N_PRD＋q 条在第 ST0+…−1 行


def _cols(start, n):
    return [CL(CI(start) + i) for i in range(n)]


SLT = dict(j='T', k='U', 含量='V', 名称='W', 显示='X')
SLT_U = _cols('Y', 13)                                    # 投料吨 1～12 月＋本年（Y..AK）
SLT_A = _cols('AL', 13)                                   # 金额（AL..AX）
SLT_F = _cols('AY', 13)                                   # 运费（AY..BK）
STB = dict(j='BM', q='BN')
STB_M = _cols('BO', 12)                                   # BO..BZ


def build_cost(ws):
    LAST = COST_LAST
    widths(ws, dict({'A': 30, 'B': 7.5, 'O': 12.5, 'P': 2}, **{c: 11 for c in COST_MC}))
    title(ws, '=IF(P_公司名称="","",P_公司名称&"　")&"成本费用分析（每吨成本）"', LAST, C_VIEW, TIP_COST)
    selector(ws, 'A3', '选产品', 'B3', None, '=产品列表', prompt='从下拉选产品；空着＝基础资料参数「默认产品」')
    ws.merge_cells('B3:D3')
    home_link(ws, f'{LAST}3')
    ws.row_dimensions[3].height = 24
    sc = Sc(ws, ['规', '找到', '认出', 'p', 'pp', '品名', '折纯', '含量', '显示数', '末行'], cl='Q', cv='R')
    sc.set('规', norm('B3'))
    g = sc('规')
    sc.set('找到', (f'IF({g}="",0,IFERROR(MATCH({esc(g)},品档_名称,0),IFERROR(MATCH("*、"&{esc(g)}&"、*",品档_别名规,0),0)))'))
    sc.set('认出', f'IF({g}="",1,IF({sc("找到")}>0,1,0))')
    sc.set('p', f'IF({sc("找到")}>0,{sc("找到")},N(P_默认品号))')
    p = sc('p')
    sc.set('pp', f'MAX(1,{p})')
    pp = sc('pp')
    sc.set('品名', f'IF({p}=0,"",INDEX(品档_名称,{pp})&"")')
    sc.set('折纯', f'IF({p}=0,0,INDEX(品档_折纯,{pp}))')
    sc.set('含量', f'IF({p}=0,0,INDEX(品档_含量,{pp}))')

    # ── 槽投块（跟 s_cost 同式：第 j 槽、第 q 个产品、第 m 月）──
    for t, c in STB.items():
        _lab(ws, f'{c}2', t)
    for m in MS:
        _lab(ws, f'{STB_M[m - 1]}2', f'槽投{m}')
    for j in range(1, N_SLOT + 1):
        for q in range(1, N_PRD + 1):
            r = ST0 + (j - 1) * N_PRD + q - 1
            hset(ws, f'{STB["j"]}{r}', j)
            hset(ws, f'{STB["q"]}{r}', q)
            for m in MS:
                hset(ws, f'{STB_M[m - 1]}{r}', f'SUMIF(产_品月键,{q}*100+{m},产_槽{j})')

    # ── 候选行表（第 2 行：月份有效标志）──
    L = Lines(ws, 3, 'CD', [f'm{m}' for m in MS] + ['yr'], nz_keys=[f'm{m}' for m in MS])
    MK = {m: L.c[f'm{m}'] for m in MS}
    YK = L.c['yr']
    for m in MS:
        hset(ws, f'{MK[m]}1', f'IF(AND({m}>=P_建账月,{m}<=P_截止月,{p}>0),1,0)')
    fl = lambda m: f'{MK[m]}$1'

    def iv(m, e):
        return f'=IF({fl(m)}=1,{e},"")'

    def msum(r):
        return f'=SUM({MK[1]}{r}:{MK[12]}{r})'

    def ln(r, m):
        return f'${MK[m]}${r}' if m != 'yr' else f'${YK}${r}'

    def per(r_amt, r_q, m, nd=2):
        q = ln(r_q, m)
        e = f'IF(ROUND(N({q}),6)>0,ROUND({ln(r_amt, m)}/{q},{nd}),"")'
        return iv(m, e) if m != 'yr' else '=' + e

    def per_vals(r_amt, r_q, nd=2):
        v = {f'm{m}': per(r_amt, r_q, m, nd) for m in MS}
        v['yr'] = per(r_amt, r_q, 'yr', nd)
        return v

    show_nz = lambda r: f'=IF({L.nzf(r)},1,0)'
    T_SEC, T_ITEM, T_SUB, T_KEY, T_CHK, T_SUBI, T_TON, T_RATIO, T_HID = 1, 3, 4, 5, 6, 7, 8, 9, 0

    # ── 槽表（第 j 槽）──
    for k_, c in SLT.items():
        _lab(ws, f'{c}2', k_)
    for i, c in enumerate(SLT_U):
        _lab(ws, f'{c}2', f'投{i + 1}' if i < 12 else '投年')
    for i, c in enumerate(SLT_A):
        _lab(ws, f'{c}2', f'额{i + 1}' if i < 12 else '额年')
    for i, c in enumerate(SLT_F):
        _lab(ws, f'{c}2', f'运{i + 1}' if i < 12 else '运年')
    SR = lambda c: f'${c}${SL0}:${c}${SL0 + N_SLOT - 1}'
    for j in range(1, N_SLOT + 1):
        r = SL0 + j - 1
        k = f'${SLT["k"]}{r}'
        hset(ws, f'{SLT["j"]}{r}', j)
        hset(ws, f'{SLT["k"]}{r}', f'N(INDEX(槽_料号,{j}))')
        hset(ws, f'{SLT["含量"]}{r}', f'IF({k}>0,INDEX(料档_含量,{k}),0)')
        hset(ws, f'{SLT["名称"]}{r}', f'IF({k}>0,INDEX(料档_名称,{k})&"","")')
        g0 = ST0 + (j - 1) * N_PRD
        g1 = g0 + N_PRD - 1
        for m in MS:
            stc = STB_M[m - 1]
            rj = f'${stc}${g0}:${stc}${g1}'
            hset(ws, f'{SLT_U[m - 1]}{r}', f'IF({fl(m)}=1,INDEX({rj},{pp}),0)')
            S = f'SUM({rj})'
            for cols, fld in ((SLT_A, '投额'), (SLT_F, '投运')):
                T = f'INDEX(料_{fld},{k},{m})'
                t1 = f'ROUND({T}*SUM(INDEX({rj},1):INDEX({rj},{pp}))/{S},2)'
                t0 = f'IF({p}>1,ROUND({T}*SUM(INDEX({rj},1):INDEX({rj},MAX(1,{pp}-1)))/{S},2),0)'
                hset(ws, f'{cols[m - 1]}{r}', f'IF(OR({fl(m)}<>1,{k}=0),0,IF({S}=0,0,{t1}-{t0}))')
        for cols in (SLT_U, SLT_A, SLT_F):
            hset(ws, f'{cols[12]}{r}', f'SUM({cols[0]}{r}:{cols[11]}{r})')
        hset(ws, f'{SLT["显示"]}{r}', (f'IF(AND({k}>0,SUMSQ({SLT_U[0]}{r}:{SLT_U[11]}{r})+SUMSQ({SLT_A[0]}{r}:{SLT_A[11]}{r})>{TINY}),1,0)'))
    slot = lambda c, j: f'${c}${SL0 + j - 1}'

    # ── ① 产量、投料、单耗、收率 ──
    L.add(T_SEC, '① 产量、投料、单耗')
    rq = L.add(T_ITEM, '产量', '吨', dict({f'm{m}': iv(m, f'INDEX(本_产量,{pp},{m})') for m in MS}, yr=msum))
    L.add(T_ITEM, '废料（二次过滤等）', '吨',
          dict({f'm{m}': iv(m, f'SUMIF(产_品月键,{pp}*100+{m},产_废料吨)') for m in MS}, yr=msum))
    for j in range(1, N_SLOT + 1):
        v = {f'm{m}': iv(m, slot(SLT_U[m - 1], j)) for m in MS}
        v['yr'] = msum
        L.add(T_ITEM, f'="投料："&{slot(SLT["名称"], j)}', '吨', v, f'={slot(SLT["显示"], j)}')
    usum = lambda i: f'SUMIF({SR(SLT["k"])},">0",{SR(SLT_U[i])})'
    v = {f'm{m}': iv(m, usum(m - 1)) for m in MS}
    v['yr'] = msum
    L.add(T_SUB, '投料合计', '吨', v, f'=IF(SUM({SR(SLT["显示"])})>0,1,0)')
    for j in range(1, N_SLOT + 1):
        v = {}
        for i, m in enumerate(list(MS) + ['yr']):
            q, u = ln(rq, m), slot(SLT_U[i], j)
            e = f'IF(ROUND(N({q}),6)>0,ROUND({u}/{q},4),"")'
            v[f'm{m}' if m != 'yr' else 'yr'] = iv(m, e) if m != 'yr' else '=' + e
        L.add(T_RATIO, f'="单耗："&{slot(SLT["名称"], j)}', '吨/吨', v, f'={slot(SLT["显示"], j)}')
    v = {}
    zc, hl = sc('折纯'), sc('含量')
    for i, m in enumerate(list(MS) + ['yr']):
        q = ln(rq, m)
        U = SR(SLT_U[i])
        miss = f'OR({zc}=0,{hl}=0,COUNTIFS({U},"<>0",{SR(SLT["含量"])},0,{SR(SLT["k"])},">0")>0)'
        den = f'SUMPRODUCT({U},{SR(SLT["含量"])})'
        e = (f'IF(OR(ROUND(N({q}),6)<=0,ROUND({usum(i)},6)<=0),"",IF({miss},"缺含量",IF({den}<=0,"",'
             f'ROUND({q}*{zc}*{hl}/{den}*100,2))))')
        v[f'm{m}' if m != 'yr' else 'yr'] = iv(m, e) if m != 'yr' else '=' + e
    L.add(T_ITEM, '收率（折纯）', '%', v)

    # ── ② 制造成本 ──
    L.add(T_SEC, '② 制造成本（完工入库）')
    r_wip0 = L.add(T_ITEM, '月初在产', '元', dict({f'm{m}': iv(m, f'INDEX(本_初在产,{pp},{m})') for m in MS},
                                               yr=f'=IF({p}=0,"",INDEX(本_初在产,{pp},P_建账月))'))
    r_amt, r_frt = {}, {}
    for j in range(1, N_SLOT + 1):
        v = {f'm{m}': iv(m, slot(SLT_A[m - 1], j)) for m in MS}
        v['yr'] = msum
        r_amt[j] = L.add(T_ITEM, f'="原料："&{slot(SLT["名称"], j)}', '元', v, f'={slot(SLT["显示"], j)}')
        v = {f'm{m}': iv(m, slot(SLT_F[m - 1], j)) for m in MS}
        v['yr'] = msum
        r_frt[j] = L.add(T_SUBI, '　其中进货运费', '元', v,
                         lambda r, j=j: f'=IF(AND({slot(SLT["显示"], j)}=1,{L.nzf(r)}),1,0)')
    r_oh = {}
    for i in range(1, N_OH + 1):
        no = f'INDEX(制费明细_号,{i})'
        v = {f'm{m}': iv(m, f'INDEX(本_分到,{(i - 1) * N_PRD}+{pp},{m})') for m in MS}
        v['yr'] = msum
        r_oh[i] = L.add(T_ITEM, f'=IF({no}=0,"","制造费用："&INDEX(科目_名称,MAX(1,{no})))', '元', v,
                        lambda r, no=no: f'=IF(AND({no}>0,{L.nzf(r)}),1,0)')
    oh1, oh2 = r_oh[1], r_oh[N_OH]
    r_scr = L.add(T_ITEM, '减：废料（回收料作价）', '元',
                  dict({f'm{m}': iv(m, f'-INDEX(本_废料扣,{pp},{m})') for m in MS}, yr=msum), show_nz)
    r_wip1 = L.add(T_ITEM, '减：月末在产（没产量的料留下月）', '元',
                   dict({f'm{m}': iv(m, f'-INDEX(本_末在产,{pp},{m})') for m in MS},
                        yr=f'=IF({p}=0,"",-INDEX(本_末在产,{pp},P_截止月))'))
    wip_show = f'=IF(OR({L.nzf(r_wip0)},{L.nzf(r_wip1)}),1,0)'
    L.set_show(r_wip0, wip_show)
    L.set_show(r_wip1, wip_show)
    v = {}
    for m in MS:
        e = (f'ROUND({ln(r_wip0, m)}+SUM({SR(SLT_A[m - 1])})+SUM(${MK[m]}${oh1}:${MK[m]}${oh2})+{ln(r_scr, m)}'
             f'+{ln(r_wip1, m)},2)')
        v[f'm{m}'] = iv(m, e)
    v['yr'] = lambda r: f'=ROUND(SUM({MK[1]}{r}:{MK[12]}{r}),2)'
    r_done = L.add(T_KEY, '＝完工成本', '元', v)
    r_eng = L.add(T_CHK, '　核对：成本计算的完工成本', '元',
                  dict({f'm{m}': iv(m, f'INDEX(本_完工,{pp},{m})') for m in MS}, yr=msum))
    v = {f'm{m}': iv(m, f'ROUND({ln(r_done, m)}-{ln(r_eng, m)},2)') for m in MS}
    v['yr'] = f'=ROUND({ln(r_done, "yr")}-{ln(r_eng, "yr")},2)'
    L.add(T_CHK, '　差（应为 0）', '元', v)
    r_ton = L.add(T_KEY, '每吨制造成本', '元/吨', per_vals(r_done, rq))
    L.add(T_TON, '　月初在产/吨', '元/吨', per_vals(r_wip0, rq), wip_show)
    for j in range(1, N_SLOT + 1):
        L.add(T_TON, f'="　原料："&{slot(SLT["名称"], j)}&"/吨"', '元/吨', per_vals(r_amt[j], rq), f'={slot(SLT["显示"], j)}')
        L.add(T_SUBI, '　　其中进货运费/吨', '元/吨', per_vals(r_frt[j], rq), f'={L.c["显示"]}{r_frt[j]}')
    for i in range(1, N_OH + 1):
        L.add(T_TON, f'=IF({L.a("标签", r_oh[i])}="","","　"&{L.a("标签", r_oh[i])}&"/吨")', '元/吨', per_vals(r_oh[i], rq),
              f'={L.c["显示"]}{r_oh[i]}')
    L.add(T_TON, '　减：废料/吨', '元/吨', per_vals(r_scr, rq), f'={L.c["显示"]}{r_scr}')
    L.add(T_TON, '　减：月末在产/吨', '元/吨', per_vals(r_wip1, rq), wip_show)

    # ── ③ 销售口径 ──
    sq_rng = {m: erng(H_PIN, PIN_FIELDS, '销量', m, N_PRD) for m in MS}
    sa_rng = {m: erng(H_PIN, PIN_FIELDS, '销额', m, N_PRD) for m in MS}
    frt = lambda m: f'IF(P_运费进成本=1,{acc("540102", m)},{acc("560101", m)})'
    prd = lambda m: (f'{acc("5601", m)}+{acc("5602", m)}+{acc("5603", m)}-IF(P_运费进成本=1,0,{acc("560101", m)})')
    rF = L.add(T_HID, '（隐藏）销货运费合计', '元', dict({f'm{m}': iv(m, f'ROUND({frt(m)},2)') for m in MS}, yr=msum), 0)
    rE = L.add(T_HID, '（隐藏）期间费用合计', '元', dict({f'm{m}': iv(m, f'ROUND({prd(m)},2)') for m in MS}, yr=msum), 0)
    rS = L.add(T_HID, '（隐藏）全部产品销量', '吨', dict({f'm{m}': iv(m, f'SUM({sq_rng[m]})') for m in MS}, yr=msum), 0)
    rR = L.add(T_HID, '（隐藏）全部产品收入', '元',
               dict({f'm{m}': iv(m, f'ROUND(SUMIFS(销_收入,销_出库,1,销_月,{m}),2)') for m in MS}, yr=msum), 0)
    rC = L.add(T_HID, '（隐藏）全部产品销售成本', '元', dict({f'm{m}': iv(m, f'ROUND(SUM({sa_rng[m]}),2)') for m in MS}, yr=msum), 0)

    # 前 p 个产品的销量（按产品顺序累计）、前 p−1 个：本年＝各月合计（本年的运费、期间费用按全年销量分）
    rK1 = L.add(T_HID, '（隐藏）前 p 个产品销量', '吨',
                dict({f'm{m}': iv(m, f'SUM(INDEX({sq_rng[m]},1):INDEX({sq_rng[m]},{pp}))') for m in MS}, yr=msum), 0)
    rK0 = L.add(T_HID, '（隐藏）前 p−1 个产品销量', '吨',
                dict({f'm{m}': iv(m, f'IF({p}>1,SUM(INDEX({sq_rng[m]},1):INDEX({sq_rng[m]},MAX(1,{pp}-1))),0)') for m in MS},
                     yr=msum), 0)

    def alloc(r_tot, m):
        X, S, c1, c0 = ln(r_tot, m), ln(rS, m), ln(rK1, m), ln(rK0, m)
        e = f'IF(ROUND({S},4)<=0,0,ROUND({X}*{c1}/{S},2)-IF({p}>1,ROUND({X}*{c0}/{S},2),0))'
        return iv(m, e) if m != 'yr' else '=' + e

    L.add(T_SEC, '③ 销售口径（按销量算每吨）')
    rsq = L.add(T_ITEM, '销量', '吨', dict({f'm{m}': iv(m, f'INDEX(品_销量,{pp},{m})') for m in MS}, yr=msum))
    rrev = L.add(T_ITEM, '不含税收入', '元',
                 dict({f'm{m}': iv(m, f'ROUND(SUMIFS(销_收入,销_出库,1,销_品号,{pp},销_月,{m}),2)') for m in MS}, yr=msum))
    r_avg = L.add(T_KEY, '不含税均价', '元/吨', per_vals(rrev, rsq))
    rcost = L.add(T_ITEM, '销售成本（出库成本）', '元', dict({f'm{m}': iv(m, f'INDEX(品_销额,{pp},{m})') for m in MS}, yr=msum))
    L.add(T_TON, '　销售成本/吨', '元/吨', per_vals(rcost, rsq))
    rfr = L.add(T_ITEM, '="销货运费（"&IF(P_运费进成本=1,"主营业务成本","销售费用")&"，按销量分）"', '元',
                dict({f'm{m}': alloc(rF, m) for m in MS}, yr=alloc(rF, 'yr')))
    L.add(T_TON, '　销货运费/吨', '元/吨', per_vals(rfr, rsq))
    rpe = L.add(T_ITEM, '期间费用（销售管理财务，按销量分）', '元', dict({f'm{m}': alloc(rE, m) for m in MS}, yr=alloc(rE, 'yr')))
    L.add(T_TON, '　期间费用/吨', '元/吨', per_vals(rpe, rsq))

    def per_expr(expr_fn, m):
        q = ln(rsq, m)
        e = f'IF(ROUND(N({q}),4)>0,ROUND(({expr_fn(m)})/{q},2),"")'
        return iv(m, e) if m != 'yr' else '=' + e

    def pv(expr_fn):
        v = {f'm{m}': per_expr(expr_fn, m) for m in MS}
        v['yr'] = per_expr(expr_fn, 'yr')
        return v

    full = lambda m: f'{ln(rcost, m)}+{ln(rfr, m)}+{ln(rpe, m)}'
    r_full = L.add(T_KEY, '完全成本/吨', '元/吨', pv(full))
    L.add(T_ITEM, '毛利/吨', '元/吨', pv(lambda m: f'{ln(rrev, m)}-{ln(rcost, m)}-IF(P_运费进成本=1,{ln(rfr, m)},0)'))
    r_op = L.add(T_KEY, '经营利润/吨', '元/吨', pv(lambda m: f'{ln(rrev, m)}-({full(m)})'))
    L.add(T_ITEM, '经营利润（不含其他业务、营业外、所得税）', '元',
          dict({f'm{m}': iv(m, f'ROUND({ln(rrev, m)}-({full(m)}),2)') for m in MS},
               yr=f'=ROUND({ln(rrev, "yr")}-({full("yr")}),2)'))
    allp = lambda m: f'ROUND({ln(rR, m)}-{ln(rC, m)}-IF(ROUND({ln(rS, m)},4)>0,{ln(rF, m)}+{ln(rE, m)},0),2)'
    ra = L.add(T_CHK, '核对：全部产品经营利润（×销量合计）', '元',
               dict({f'm{m}': iv(m, allp(m)) for m in MS}, yr='=' + allp('yr')))
    pl = lambda m: f'-({acc("5001", m)})-({acc("5401", m)})-({acc("5601", m)})-({acc("5602", m)})-({acc("5603", m)})'
    rb = L.add(T_CHK, '　利润表：主营收入－主营成本－三项费用', '元',
               dict({f'm{m}': iv(m, f'ROUND({pl(m)},2)') for m in MS}, yr=msum))
    v = {f'm{m}': iv(m, f'ROUND({ln(ra, m)}-{ln(rb, m)},2)') for m in MS}
    v['yr'] = f'=ROUND({ln(ra, "yr")}-{ln(rb, "yr")},2)'
    L.add(T_CHK, '　差（应为 0；没产品销量的月份运费、费用分不出去）', '元', v)
    L.finish()
    NR = len(L.rows)
    sc.set('显示数', L.total())
    sc.set('末行', f'{COST_D0 - 1}+{sc("显示数")}')

    # ── 第 3、4、5 行 ──
    note(ws, 'E', 'N', (f'="＝"&IF({p}=0,"（没有产品：先在物料档案②建）",{sc("品名")})&"；"&P_年度&" 年 "&P_建账月&"～"&P_截止月'
                        f'&" 月（截止日 "&TEXT(P_截止,"yyyy-mm-dd")&"）"'
                        f'&IF({sc("认出")}=0,"　⚠ 「"&TRIM(B3&"")&"」不在产品清单（先按默认产品）","")'))
    kpi(ws, 'A4', '本年产量（吨）', 'B4', L.a('yr', rq), QTY, merge_to='C4')
    kpi(ws, 'D4', '每吨制造成本', 'E4', L.a('yr', r_ton), MNY, merge_to='F4')
    kpi(ws, 'G4', '每吨完全成本', 'H4', L.a('yr', r_full), MNY, merge_to='I4')
    kpi(ws, 'J4', '不含税均价', 'K4', L.a('yr', r_avg), MNY, merge_to='L4')
    kpi(ws, 'M4', '每吨经营利润', 'N4', L.a('yr', r_op), MNY, merge_to='O4')
    ws.row_dimensions[4].height = 22
    header(ws, 5, [('A', '项目'), ('B', '单位')] + [(c, f'{i + 1} 月') for i, c in enumerate(COST_MC)] + [('O', '本年')],
           GREEN_H, height=26)
    for i, c in enumerate(COST_MC):
        ws.conditional_formatting.add(f'{c}5', FormulaRule(formula=[f'OR({i + 1}<P_建账月,{i + 1}>P_截止月)'],
                                                           fill=cf_fill('FFBFBFBF')))

    # ── 显示区 ──
    colmap = {'A': '标签', 'B': '单位', **{c: f'm{i + 1}' for i, c in enumerate(COST_MC)}, 'O': 'yr'}
    show_list(ws, L, COST_D0, NR, colmap, 'DC', 'DD', 'DE')
    d1 = COST_D0 + NR - 1
    for r in range(COST_D0, d1 + 1):
        for col in ['A', 'B'] + COST_MC + ['O']:
            c = ws[f'{col}{r}']
            c.font = F_TXTB if col == 'O' else F_TXT
            if col == 'A':
                c.alignment = AL
            elif col == 'B':
                c.alignment, c.font = AC, F_NOTE
            else:
                c.number_format, c.alignment = NUM4, AR
    t0 = f'$DE{COST_D0}'
    rng = f'A{COST_D0}:{LAST}{d1}'
    ws.conditional_formatting.add(f'C{COST_D0}:{LAST}{d1}', FormulaRule(
        formula=[f'AND({t0}=6,LEFT($A{COST_D0},3)="　差（",ISNUMBER(C{COST_D0}),ABS(N(C{COST_D0}))>=0.005)'],
        font=Font(bold=True, color='FFC00000'), fill=cf_fill('FFFFC7CE')))
    type_cf(ws, rng, t0, [(1, C_VIEW, F_SECW, True), (5, 'FFFCE4D6', F_B, True), (4, 'FFDDEBF7', F_B, True),
                          (6, 'FFFFF2CC', F_IT, True), (7, None, F_IT, True), (8, None, Font(color='FF404040'), True),
                          (9, None, None, True), (3, None, None, True)])
    hide(ws, *[CL(i) for i in range(CI('Q'), CI('DE') + 1)])
    finish(ws, LAST, sc('末行'), d1, 'C6', '5:5')


# ═══════════════════════════════ 借款明细及利息 ═══════════════════════════════
TIP_LOAN = ('★ 上半：每个出借人（往来单位类型＝出借人）截至截止日的借款本金、投资本金、年利率、本月（截止日那个月）利息、本年利息、'
            '已计提未付（＝应付利息、应付利润里这个人的：期初应付息＋本年计提－已付）、到截止日应计未付（按天算到截止日，减掉已付的利息）、'
            '到期日（30 天内到期、已过期的标红）。下半：黄格选一个出借人（空着＝第一个），按日期列出它每一笔本金：建账前的（基础资料⑦）、'
            '资金台帐里的借款到账、还出、投资款、退还投资款、老板垫付账户的每一笔、还垫付（垫付现金已报）；'
            '本笔利息＝金额×年利率÷一年天数×计息天数（截止日＋1－日期，算头不算尾；归还是负数）；合计减掉已付利息＝到截止日应计未付，跟上半核对。'
            '投资款的「利息」按投资年利率算，记利润分配，不进费用。')
LOAN_LAST = 'L'
LOAN_D0 = 7
LT = dict(l='R', u='S', 名='T', 类别='U', 借本金='V', 投本金='W', 本月='X', 本年='Y', 未付='Z', 应计='AA', 借率字='AB',
          投率字='AC', 到期='AD', 到期字='AE', 红='AF', 提醒='AG')
LK = 'AJ'                                                  # 本金事件排序键（第 3 行起：⑦ N_LO 条，再资金台帐 N_CASH 条）
EVT = dict(e='AL', 键='AM', id='AN', 源='AO', n='AP', 日期='AQ', 额='AR', 类='AS', 余额='AT', 天数='AU', 利率='AV', 利息='AW',
           性质='AX', 摘要='AY', 来源='AZ', 备注='BA')
LD = dict(k='BC', t='BD', i='BE')


def build_loan(ws):
    LAST = LOAN_LAST
    widths(ws, {'A': 11.5, 'B': 30, 'C': 13, 'D': 13, 'E': 13, 'F': 13, 'G': 12.5, 'H': 12.5, 'I': 13.5, 'J': 14, 'K': 11.5,
                'L': 36, 'M': 2})
    title(ws, '=IF(P_公司名称="","",P_公司名称&"　")&"借款明细及利息"', LAST, C_VIEW, TIP_LOAN)
    put(ws, 'A3', '截止日', F_KPI_L, LBL, align=AC)
    put(ws, 'B3', '=P_截止', F_SEL, FILL_AUTO, DATE, AC)
    selector(ws, 'C3', '选出借人', 'D3', None, '=出借人列表', prompt='下半看哪个出借人的逐笔明细；空着＝第一个')
    ws.merge_cells('D3:E3')
    home_link(ws, f'{LAST}3')
    ws.row_dimensions[3].height = 24
    sc = Sc(ws, ['出借数', '规', '找', 'L', '认出', 'L名', 'u', '借率', '投率', '天数', '笔数', '显示笔数', '建账前已付', '资金已付',
                 '借入', '归还', '余额', '利息', '应计', '上半', '差', '显示数', '末行'], cl='O', cv='P')
    sc.set('出借数', 'COUNTIF(出借_单位号,">0")')
    sc.set('规', norm('D3'))
    g = sc('规')
    sc.set('找', f'IF({g}="",0,N({unit_match(g, "出借人号", "0")}))')
    sc.set('L', f'IF({g}="",IF({sc("出借数")}>0,1,0),{sc("找")})')
    Lx = sc('L')
    sc.set('认出', f'IF({g}="",1,IF({sc("找")}>0,1,0))')
    sc.set('L名', f'IF({Lx}=0,"",INDEX(出借_名称,{Lx})&"")')
    sc.set('u', f'IF({Lx}=0,0,N(INDEX(出借_单位号,{Lx})))')
    u = sc('u')
    sc.set('借率', f'IF({u}=0,0,INDEX(位_借款利率,{u}))')
    sc.set('投率', f'IF({u}=0,0,INDEX(位_投资利率,{u}))')
    sc.set('天数', f'IF({u}=0,365,INDEX(位_计息天数,{u}))')

    # ── 出借人表（第 l 个在第 l+2 行）──
    for k_, c in LT.items():
        _lab(ws, f'{c}2', k_)
    for l in range(1, N_LEND + 1):
        r = l + 2
        c = lambda k_: f'${LT[k_]}{r}'
        S = c('u')
        hset(ws, f'{LT["l"]}{r}', l)
        hset(ws, f'{LT["u"]}{r}', f'N(INDEX(出借_单位号,{l}))')
        hset(ws, f'{LT["名"]}{r}', f'IF({S}=0,"",INDEX(出借_名称,{l})&"")')
        hset(ws, f'{LT["类别"]}{r}', (f'IF({S}=0,"",CHOOSE(MAX(1,MIN(3,INDEX(位_出借人类别,{S}))),"银行","个人","合伙人")'
                                     f'&" · 一年按 "&INDEX(位_计息天数,{S})&" 天")'))
        hset(ws, f'{LT["借本金"]}{r}', f'IF({S}=0,0,ROUND(INDEX(息_借本金,{l}),2))')
        hset(ws, f'{LT["投本金"]}{r}', f'IF({S}=0,0,ROUND(INDEX(息_投本金,{l}),2))')
        hset(ws, f'{LT["本月"]}{r}', f'IF({S}=0,0,ROUND(INDEX(息_借本月利息,{l},P_截止月)+INDEX(息_投本月利息,{l},P_截止月),2))')
        hset(ws, f'{LT["本年"]}{r}', f'IF({S}=0,0,ROUND(INDEX(息_借累计利息,{l})+INDEX(息_投累计利息,{l}),2))')
        x = lambda k_: f'INDEX(息_{k_},{l})'
        hset(ws, f'{LT["未付"]}{r}', (f'IF({S}=0,0,ROUND({x("借期初应付息")}+{x("借累计利息")}-{x("借已付息")}'
                                     f'+{x("投期初应付息")}+{x("投累计利息")}-{x("投已付息")},2))'))
        hset(ws, f'{LT["应计"]}{r}', f'IF({S}=0,0,ROUND({x("借到截止应计")}+{x("投到截止应计")},2))')
        for k_, f in (('借率字', '借款利率'), ('投率字', '投资利率')):
            hset(ws, f'{LT[k_]}{r}', f'IF({S}=0,"",IF(INDEX(位_{f},{S})=0,"",TEXT(INDEX(位_{f},{S}),"0.00%")))')
        D = c('到期')
        hset(ws, f'{LT["到期"]}{r}', f'IF({S}=0,0,N(INDEX(位_到期日,{S})))')
        hset(ws, f'{LT["到期字"]}{r}', f'IF({D}>0,TEXT({D},"yyyy-mm-dd"),"")')
        hset(ws, f'{LT["红"]}{r}', f'IF(AND({D}>0,{c("借本金")}+{c("投本金")}>0.005,{D}-P_截止<=30),1,0)')
        hset(ws, f'{LT["提醒"]}{r}', (f'IF({c("红")}=1,IF({D}<P_截止,"✗ 已过期 "&(P_截止-{D})&" 天",IF({D}=P_截止,"⚠ 今天到期",'
                                     f'"⚠ "&({D}-P_截止)&" 天后到期")),IF(AND({c("借本金")}>0.005,INDEX(位_借款利率,MAX(1,{S}))=0),'
                                     f'"ℹ 借款年利率空着：不计提，付的利息直接进财务费用",IF({c("应计")}<-0.005,'
                                     f'"⚠ 已付利息比按天算的多 "&TEXT(-{c("应计")},"#,##0.00")&" 元（对一下利率、付息记录）","")))'))
    LR = lambda k_: f'${LT[k_]}$3:${LT[k_]}${N_LEND + 2}'

    # ── 本金事件排序键 ──
    _lab(ws, f'{LK}2', '排序键')
    for n in range(1, N_LO + 1):
        hset(ws, f'{LK}{n + 2}', (f'IF({Lx}=0,"",IF(AND(INDEX(借期_有效,{n})=1,INDEX(借期_出借人号,{n})={Lx}),'
                                  f'INDEX(借期_日期,{n})*{SORT_M}+{n},""))'))
    for n in range(1, N_CASH + 1):
        hset(ws, f'{LK}{N_LO + n + 2}', (f'IF({Lx}=0,"",IF(AND(INDEX(资_出借人号,{n})={Lx},INDEX(资_本金变动,{n})<>0,'
                                         f'INDEX(资_日期,{n})<=P_截止),INDEX(资_日期,{n})*{SORT_M}+{N_LO + n},""))'))
    KR = f'${LK}$3:${LK}${N_LO + N_CASH + 2}'
    sc.set('笔数', f'COUNT({KR})')
    sc.set('显示笔数', f'MIN({sc("笔数")},{EV})')
    sc.set('建账前已付', f'IF({Lx}=0,0,ROUND(SUMIFS(借期_已付息,借期_出借人号,{Lx},借期_有效,1),2))')
    sc.set('资金已付', f'IF({Lx}=0,0,ROUND(INDEX(息_借已付息,{Lx})+INDEX(息_投已付息,{Lx}),2))')

    # ── 事件表（第 e 笔在第 e+2 行）──
    for k_, c in EVT.items():
        _lab(ws, f'{c}2', k_)
    note_cash = inref(SH_CASH, CASH_COLS['备注'])
    note_lo = inref(SH_BASE, LO_COLS['说明'], block='借期')
    for e in range(1, EV + 1):
        r = e + 2
        c = lambda k_: f'${EVT[k_]}{r}'
        hset(ws, f'{EVT["e"]}{r}', e)
        hset(ws, f'{EVT["键"]}{r}', f'IF({e}>{sc("显示笔数")},"",SMALL({KR},{e}))')
        hset(ws, f'{EVT["id"]}{r}', f'IF({c("键")}="",0,MOD({c("键")},{SORT_M}))')
        hset(ws, f'{EVT["源"]}{r}', f'IF({c("id")}=0,0,IF({c("id")}<={N_LO},1,2))')
        n = c('n')
        hset(ws, f'{EVT["n"]}{r}', f'IF({c("源")}=1,{c("id")},IF({c("源")}=2,{c("id")}-{N_LO},0))')
        hset(ws, f'{EVT["日期"]}{r}', f'IF({c("id")}=0,0,INT({c("键")}/{SORT_M}))')
        hset(ws, f'{EVT["额"]}{r}', f'IF({c("源")}=1,INDEX(借期_金额,{n}),IF({c("源")}=2,INDEX(资_本金变动,{n}),0))')
        hset(ws, f'{EVT["类"]}{r}', f'IF({c("源")}=1,IF(INDEX(借期_性质,{n})=2,2,1),IF({c("源")}=2,INDEX(资_性质,{n}),0))')
        hset(ws, f'{EVT["余额"]}{r}', f'{c("额")}' + (f'+{EVT["余额"]}{r - 1}' if e > 1 else ''))
        hset(ws, f'{EVT["天数"]}{r}', f'IF({c("id")}=0,0,P_截止+1-{c("日期")})')
        hset(ws, f'{EVT["利率"]}{r}', f'IF({c("类")}=2,{sc("投率")},IF({c("类")}=1,{sc("借率")},0))')
        hset(ws, f'{EVT["利息"]}{r}', f'IF({c("id")}=0,0,{c("额")}*{c("利率")}/{sc("天数")}*{c("天数")})')
        hset(ws, f'{EVT["性质"]}{r}', (f'IF({c("源")}=1,CHOOSE(MAX(1,MIN(3,INDEX(借期_性质,{n}))),"借款","投资","垫付"),'
                                      f'IF({c("源")}=2,IF({c("类")}=2,"投资",IF(OR(INDEX(资_账户类型,{n})=5,INDEX(资_原对方号,{n})={KH224101}),'
                                      f'"垫付","借款")),""))'))
        hset(ws, f'{EVT["摘要"]}{r}', f'IF({c("源")}=1,"建账前"&{c("性质")},IF({c("源")}=2,INDEX(资_显示摘要,{n})&"",""))')
        hset(ws, f'{EVT["来源"]}{r}', (f'IF({c("源")}=1,"基础资料⑦",IF({c("源")}=2,'
                                      f'INDEX(账户_名称,MAX(1,INDEX(资_账户号,{n})))&"",""))'))
        hset(ws, f'{EVT["备注"]}{r}', f'IF({c("源")}=1,INDEX({note_lo},{n}+1)&"",IF({c("源")}=2,INDEX({note_cash},{n}+1)&"",""))')
    ER = lambda k_: f'${EVT[k_]}$3:${EVT[k_]}${EV + 2}'
    sc.set('借入', f'SUMIF({ER("额")},">0")')
    sc.set('归还', f'-SUMIF({ER("额")},"<0")')
    sc.set('余额', f'ROUND(SUM({ER("额")}),2)')
    sc.set('利息', f'ROUND(SUMIF({ER("类")},1,{ER("利息")}),2)+ROUND(SUMIF({ER("类")},2,{ER("利息")}),2)')
    sc.set('应计', f'ROUND({sc("利息")}-{sc("建账前已付")}-{sc("资金已付")},2)')
    sc.set('上半', f'IF({Lx}=0,0,ROUND(INDEX(息_借到截止应计,{Lx})+INDEX(息_投到截止应计,{Lx}),2))')
    sc.set('差', f'ROUND({sc("应计")}-{sc("上半")},2)')

    # ── 显示区（第 7 行起）：类型 0 空 1 出借人 2 上半合计 3 下半标题 4 下半表头 5 本金事件 6 合计 7 减已付 8 应计未付 9 上半 10 差 ──
    n1, ne = sc('出借数'), sc('显示笔数')
    NR = N_LEND + EV + 10
    TX = lambda s: '"' + s + '"'
    MM = lambda c: f'TEXT({c},"#,##0.00")'
    for k in range(1, NR + 1):
        r = LOAN_D0 + k - 1
        hset(ws, f'{LD["k"]}{r}', k)
        t, i = f'${LD["t"]}{r}', f'${LD["i"]}{r}'
        hset(ws, f'{LD["t"]}{r}', (f'IF({k}<={n1},1,IF({k}={n1}+1,2,IF({k}={n1}+2,0,IF({k}={n1}+3,3,IF({Lx}=0,0,'
                                   f'IF({k}={n1}+4,4,IF({k}<={n1}+4+{ne},5,IF({k}<={n1}+9+{ne},{k}-{n1}-{ne}+1,0))))))))'))
        hset(ws, f'{LD["i"]}{r}', f'IF({t}=1,{k},IF({t}=5,{k}-{n1}-4,0))')
        up = lambda k_: f'INDEX({LR(k_)},MAX(1,{i}))'
        ev = lambda k_: f'INDEX({ER(k_)},MAX(1,{i}))'
        tot = lambda k_: f'ROUND(SUM({LR(k_)}),2)'
        title3 = (f'IF({n1}=0,"还没有出借人（往来单位类型选「出借人」）",IF({Lx}=0,"⚠ 「"&TRIM(D3&"")&"」不是出借人（黄格 D3 重选）",'
                  f'{sc("L名")}&"：每一笔本金和利息（算到 "&TEXT(P_截止,"yyyy-mm-dd")&"）"))')
        cols = {
            'A': {1: up('名'), 2: TX('合计'), 3: TX('②'), 4: TX('日期'), 5: f'TEXT({ev("日期")},"yyyy-mm-dd")', 6: TX('合计')},
            'B': {1: up('类别'), 3: title3, 4: TX('摘要'), 5: ev('摘要'),
                  6: f'"共 "&{sc("笔数")}&" 笔"&IF({sc("笔数")}>{EV},"（只列了前 {EV} 笔）","")',
                  7: TX('减：已付利息'), 8: TX('＝到截止日应计未付'), 9: TX('上半「到截止日应计未付」'), 10: TX('差（应为 0）')},
            'C': {1: up('借本金'), 2: tot('借本金'), 4: TX('性质'), 5: ev('性质')},
            'D': {1: up('借率字'), 4: TX('借入（打款）'), 5: f'IF({ev("额")}>0,{ev("额")},"")', 6: sc('借入')},
            'E': {1: up('投本金'), 2: tot('投本金'), 4: TX('归还'), 5: f'IF({ev("额")}<0,-{ev("额")},"")', 6: sc('归还')},
            'F': {1: up('投率字'), 4: TX('本金余额'), 5: ev('余额'), 6: sc('余额')},
            'G': {1: up('本月'), 2: tot('本月'), 4: TX('计息天数'), 5: f'{ev("天数")}&" 天"'},
            'H': {1: up('本年'), 2: tot('本年'), 4: TX('年利率'), 5: f'IF({ev("利率")}=0,"",TEXT({ev("利率")},"0.00%"))'},
            'I': {1: up('未付'), 2: tot('未付'), 4: TX('本笔利息'), 5: ev('利息'), 6: sc('利息'),
                  7: f'-({sc("建账前已付")}+{sc("资金已付")})', 8: sc('应计'), 9: sc('上半'), 10: sc('差')},
            'J': {1: up('应计'), 2: tot('应计'), 4: TX('来源'), 5: ev('来源')},
            'K': {1: up('到期字'), 4: TX('')},
            'L': {1: up('提醒'), 4: TX('备注'), 5: ev('备注'),
                  7: f'"建账前已付 "&{MM(sc("建账前已付"))}&"，资金台帐付的 "&{MM(sc("资金已付"))}',
                  10: f'IF(ABS({sc("差")})<0.005,"✓ 跟上半一致","✗ 跟上半不一致")'},
        }
        for col, mp in cols.items():
            opts = ','.join(mp.get(ty, '""') for ty in range(0, 11))
            cell = ws[f'{col}{r}']
            cell.value = f'=CHOOSE({t}+1,{opts})'
    sc.set('显示数', f'IF({Lx}=0,{n1}+3,{n1}+9+{ne})')
    sc.set('末行', f'{LOAN_D0 - 1}+{sc("显示数")}')

    # ── 第 3、4、5、6 行 ──
    note(ws, 'F', 'K', (f'="本月＝"&P_截止月&" 月；下半："&IF({Lx}=0,"（没选到出借人）",{sc("L名")})'
                        f'&IF({sc("认出")}=0,"　⚠ 「"&TRIM(D3&"")&"」不是出借人","")'))
    kpi(ws, 'A4', '借款本金', 'B4', f'ROUND(SUM({LR("借本金")}),2)', MNY)
    kpi(ws, 'C4', '投资本金', 'D4', f'ROUND(SUM({LR("投本金")}),2)', MNY)
    kpi(ws, 'E4', '本年利息', 'F4', f'ROUND(SUM({LR("本年")}),2)', MNY)
    kpi(ws, 'G4', '已计提未付', 'H4', f'ROUND(SUM({LR("未付")}),2)', MNY)
    kpi(ws, 'I4', '到截止应计未付', 'J4', f'ROUND(SUM({LR("应计")}),2)', MNY)
    put(ws, 'K4', '30天内到期', F_KPI_L, LBL, align=AC)
    put(ws, 'L4', f'=IF(SUM({LR("红")})=0,"没有",SUM({LR("红")})&" 个出借人（标红的）")', F_KV, FILL_AUTO, align=AL)
    ws.row_dimensions[4].height = 22
    section(ws, 5, 'A', LAST, '① 各出借人（截至截止日）　＋　② 选的出借人逐笔（往下看）', C_VIEW)
    header(ws, 6, [('A', '出借人'), ('B', '类别'), ('C', '借款本金'), ('D', '借款年利率'), ('E', '投资本金'), ('F', '投资年利率'),
                   ('G', '="本月利息（"&P_截止月&"月）"'), ('H', '本年利息'), ('I', '已计提未付'), ('J', '到截止日应计未付'),
                   ('K', '到期日'), ('L', '提醒')], GREEN_H, height=30)

    d1 = LOAN_D0 + NR - 1
    fmts = {'C': MNY, 'D': MNY, 'E': MNY, 'F': MNY, 'G': MNY, 'H': MNY, 'I': MNY, 'J': MNY}
    for r in range(LOAN_D0, d1 + 1):
        for col in 'ABCDEFGHIJKL':
            c = ws[f'{col}{r}']
            c.font = F_TXT
            if col in fmts:
                c.number_format, c.alignment = fmts[col], AR
            else:
                c.alignment = AL if col in ('B', 'L') else AC
    t0 = f'${LD["t"]}{LOAN_D0}'
    rng = f'A{LOAN_D0}:{LAST}{d1}'
    cf = ws.conditional_formatting
    red = f'AND({t0}=1,INDEX({LR("红")},${LD["i"]}{LOAN_D0})=1)'
    cf.add(f'K{LOAN_D0}:L{d1}', FormulaRule(formula=[red], font=Font(bold=True, color='FFC00000'), fill=cf_fill('FFFFC7CE')))
    cf.add(f'L{LOAN_D0}:L{d1}', FormulaRule(formula=[f'LEFT($L{LOAN_D0},1)="✗"'], font=Font(bold=True, color='FFC00000')))
    cf.add(f'I{LOAN_D0}:I{d1}', FormulaRule(formula=[f'AND({t0}=10,ABS(N($I{LOAN_D0}))>=0.005)'], font=Font(bold=True, color='FFC00000'),
                                            fill=cf_fill('FFFFC7CE')))
    type_cf(ws, rng, t0, [(3, C_VIEW, F_SECW, True), (4, GREEN_H, F_SECW, True), (2, 'FFFCE4D6', F_B, True),
                          (6, 'FFDDEBF7', F_B, True), (8, 'FFFCE4D6', F_B, True), (9, 'FFFFF2CC', F_IT, True),
                          (10, 'FFFFF2CC', F_IT, True), (7, None, F_IT, True), (1, None, None, True), (5, None, None, True)])
    hide(ws, *[CL(i) for i in range(CI('O'), CI(LD['i']) + 1)])
    finish(ws, LAST, sc('末行'), d1, 'B7', '6:6')


def build(wb, ctx=None):
    build_stk(wb[SH_STK])
    build_cost(wb[SH_COST])
    build_loan(wb[SH_LOAN])
