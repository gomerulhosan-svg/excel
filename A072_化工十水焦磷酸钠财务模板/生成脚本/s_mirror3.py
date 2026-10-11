# -*- coding: utf-8 -*-
"""A072 化工 · 隐藏表（业务逐条）：_进（进货物流登记）、_销（销售台账）、_产（生产日报）、_工（工资表）、
   _票（发票登记）、_手（手工凭证）、_盘（月末盘点）。
   第 n 条在第 n+1 行，A 列＝n；取数 INDEX(录入表!$X$4:$X$1048576, n+1)（插行删行排序不错位）。
   「有效」＝这一行有内容且没有 ✗（✗ 只看本行和基础资料，不引用截止日、引擎、往来余额）；⚠/ℹ 不影响有效。
   私有列（别的模块不要引用）：_进 AM 参考价、AN 税点率；_销 AJ 税点率；_产 AF 坏槽、AG 空槽；_工 P 年月；
   _票 X 资行；_手 W 行错文；_盘 N 年月。"""
from layout import *
from common import F_HELP, DATE

M = SORT_M


# ───────────────────────── 公式小工具 ─────────────────────────
def _hdr(ws, cols, priv=None):
    for c, f, _d in cols:
        ws[f'{c}1'] = f
        ws[f'{c}1'].font = F_HELP
    for f, c in (priv or {}).items():
        ws[f'{c}1'] = f
        ws[f'{c}1'].font = F_HELP


def chain(pairs, default='""'):
    out = default
    for c, v in reversed(pairs):
        out = f'IF({c},{v},{out})'
    return out


def _ix(sh, col, r):
    return f'INDEX({inref(sh, col)},$A{r}+1)'


def vdate(x):
    return f'IF(ISNUMBER({x}),IF(AND({x}>=36526,{x}<73051),INT({x}),0),0)'


def txt(x):
    return f'TRIM({x}&"")'


def raw(x):
    return f'TRIM({x}&"")<>""'


def numv(x):
    return f'IF(ISNUMBER({x}),{x},0)'


def badnum(x):
    return f'AND({raw(x)},NOT(ISNUMBER({x})))'


def any_bad(xs):
    return 'OR(' + ','.join(badnum(x) for x in xs) + ')'


def has_any(xs):
    return 'OR(' + ','.join(raw(x) for x in xs) + ')'


def idx_of(items, t, dflt):
    out = str(dflt)
    for i in reversed(range(len(items))):
        out = f'IF({t}="{items[i]}",{i + 1},{out})'
    return out


def code_of(x):
    t = txt(x)
    return f'IFERROR(LEFT({t},FIND(" ",{t})-1),{t})'


def kmno(code):
    return f'IF({code}="",0,IFERROR(MATCH({esc(code)},科目_编码,0),0))'


def unit_no(nm):
    """规范后的名字 → 往来单位第几家（简称、全称、别名都认；0＝不在表里）"""
    e = esc(nm)
    return (f'IF({nm}="",0,IFERROR(MATCH({e},位_简称,0),IFERROR(MATCH({e},位_全称规,0),'
            f'IFERROR(MATCH("*、"&{e}&"、*",位_别名规,0),0))))')


def item_no(nm, pre):
    """规范后的名字 → 物料档案原料（pre＝料档）/产品（品档）第几个（名称、别名都认；0＝没有）"""
    e = esc(nm)
    return (f'IF({nm}="",0,IFERROR(MATCH({e},{pre}_名称,0),IFERROR(MATCH("*、"&{e}&"、*",{pre}_别名规,0),0)))')


def tk_no(t):
    """票种名称（TRIM 后）→ 票种第几个；0＝找不到"""
    return f'IFERROR(MATCH({esc(t)},票种_名称,0),0)'


def ym_any(x):
    """月份格 → 年×100＋月：8 / "8" / "8月" 当账套年度的月；2026-08、2026年8月、202608、真日期都认；空＝-1；看不懂＝0"""
    t = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({txt(x)},"年","-"),"月",""),"/","-"),".","-")'
    tv = (f'IFERROR(IF(ISNUMBER(FIND("-",{t})),VALUE(LEFT({t},4))*100+VALUE(SUBSTITUTE(MID({t},6,2),"-","")),'
          f'IF(AND(VALUE({t})>=1,VALUE({t})<=12,INT(VALUE({t}))=VALUE({t})),P_年度*100+VALUE({t}),VALUE({t}))),0)')
    num = (f'IF(AND({x}>=1,{x}<=12,INT({x})={x}),P_年度*100+{x},IF(AND({x}>=200001,{x}<=209912),{x},'
           f'IF(AND({x}>=36526,{x}<73051),YEAR({x})*100+MONTH({x}),0)))')
    return f'IF({txt(x)}="",-1,IF(ISNUMBER({x}),{num},{tv}))'


def ym_ok(c):
    return f'AND({c}>=200001,{c}<=209912,INT({c})={c},MOD({c},100)>=1,MOD({c},100)<=12)'


TXT_DATE = '"✗ 日期是文字：有「.」的用查找替换换成「-」；复制来的选中日期列 → 数据 → 分列 → 完成"'


def date_errs(d, xd, what='日期'):
    """日期三条 ✗：没填/看不懂、早于建账日、不在本年度"""
    return [
        (f'{d}=0', f'IF({raw(xd)},IF(ISTEXT({xd}),{TXT_DATE},"✗ {what}看不懂（输 2026-8-1 或 2026/8/1）"),"✗ 没填{what}")'),
        (f'{d}<P_建账日', f'"✗ {what}早于建账日（建账前的已经算在期初里了）"'),
        (f'{d}>P_年末', f'"✗ {what}不在本年度（第二年请开新一本）"'),
    ]


def wuxpiao(no):
    """票种号 no 是不是「无票」"""
    return f'AND({no}>0,INDEX(票种_名称,MAX(1,{no}))="无票")'


def _tail(ws, r, c, sh, has, last_key='末行', row_key='录入行'):
    ws[c(row_key)] = f'=ROW({sh}!$A${HDR})+$A{r}'
    ws[c(last_key)] = f'=IF({has},{c(row_key)},0)'


def _fmt(ws, cells, fmt):
    for a in cells:
        ws[a].number_format = fmt


# ───────────────────────── 运费三格（进货、销售共用） ─────────────────────────
def freight(ws, c, x, qty=None, per=None):
    """运费（不含税点）、税点、运费含税：两格有数推第三格；只有运费＝按税点率加；只有运费金额＝按税点率倒推。
       per＝每吨运费格（进货才有）：运费空＝数量×每吨运费。税点率在 c('_率')（运费票种是无票＝0，否则 P_运费税点）"""
    f_in, t_in, a_in = x('运费'), x('税点'), x('运费金额')
    rate = c('_率')
    if per is not None:
        fk = f'OR(ISNUMBER({f_in}),ISNUMBER({per}))'
        fv = f'IF(ISNUMBER({f_in}),ROUND({f_in},2),ROUND({qty}*{per},2))'
    else:
        fk, fv = f'ISNUMBER({f_in})', f'ROUND({f_in},2)'
    ws[c('运费')] = (f'=IF({fk},{fv},IF(ISNUMBER({a_in}),IF(ISNUMBER({t_in}),ROUND({a_in}-{t_in},2),'
                    f'ROUND({a_in}/(1+{rate}),2)),0))')
    ws[c('税点')] = (f'=IF(ISNUMBER({t_in}),ROUND({t_in},2),IF(ISNUMBER({a_in}),ROUND({a_in}-{c("运费")},2),'
                    f'ROUND({c("运费")}*{rate},2)))')
    ws[c('运费含税')] = f'=IF(ISNUMBER({a_in}),ROUND({a_in},2),ROUND({c("运费")}+{c("税点")},2))'
    ws[c('运费税率')] = f'=IF(AND(P_纳税人=1,{c("运费票种号")}>0),INDEX(票种_税率,{c("运费票种号")}),0)'
    ws[rate] = f'=IF({wuxpiao(c("运费票种号"))},0,P_运费税点)'


def freight_warns(c, x):
    return [
        (f'ABS({c("运费含税")}-{c("运费")}-{c("税点")})>0.005', '"⚠ 运费金额≠运费＋税点（按填的运费金额记账）"'),
        (f'AND({wuxpiao(c("运费票种号"))},{c("税点")}<>0)', '"⚠ 运费票种是无票，却加了税点（无票一般不加，税点填 0）"'),
    ]


# ───────────────────────── _进：进货物流登记 ─────────────────────────
JIN_PRIV = dict(参考价='AM', 税点率='AN')


def build_jin(ws):
    _hdr(ws, JIN_COLS, JIN_PRIV)
    Z, C = JIN, PUR_COLS
    for i in range(N_PUR):
        r, n = i + 2, i + 1
        ws[f'A{r}'] = n
        x = lambda k: _ix(SH_PUR, C[k], r)
        priv = {'_参考价': f'{JIN_PRIV["参考价"]}{r}', '_率': f'{JIN_PRIV["税点率"]}{r}'}
        c = lambda k: priv[k] if k in priv else f'{Z[k]}{r}'
        d, k, p = c('日期'), c('料号'), c('品号')
        ws[d] = f'={vdate(x("日期"))}'
        nm = norm(x('物料'))
        ws[k] = f'={item_no(nm, "料档")}'
        ws[p] = f'=IF({k}>0,0,{item_no(nm, "品档")})'
        un = norm(x('供应商'))
        ws[c('供应商号')] = f'={unit_no(un)}'
        ws[c('供应商')] = f'=IF({c("供应商号")}>0,INDEX(位_简称,{c("供应商号")}),{un})'
        lt = norm(x('物流单位'))
        hasfr = 'OR(' + ','.join(f'N({x(f)})<>0' for f in ('每吨运费', '运费', '税点', '运费金额')) + ')'
        ws[c('物流号')] = f'=IF({lt}<>"",{unit_no(lt)},IF({hasfr},{c("供应商号")},0))'
        ws[c('物流单位')] = (f'=IF({lt}<>"",IF({c("物流号")}>0,INDEX(位_简称,{c("物流号")}),{lt}),'
                         f'IF({hasfr},{c("供应商")},""))')
        q = c('数量')
        ws[q] = f'={numv(x("数量"))}'
        ws[c('单价')] = f'={numv(x("单价"))}'
        amt = x('货款金额')
        base = f'IF(ISNUMBER({amt}),ROUND({amt},2),ROUND({q}*{c("单价")},2))'
        ws[c('货款含税')] = f'=IF(P_进货含税=1,{base},ROUND({base}*(1+{c("货款税率")}),2))'
        ws[c('待定价')] = f'=IF(AND({q}<>0,NOT({raw(x("单价"))}),NOT({raw(amt)})),1,0)'
        ws[c('_参考价')] = (f'=IF({k}>0,INDEX(料档_参考单价,{k}),IF({p}>0,INDEX(品档_期初单价,{p}),0))')
        ws[c('暂估额')] = f'=IF({c("待定价")}=1,ROUND({q}*{c("_参考价")},2),0)'
        t1 = txt(x('货款票种'))
        ws[c('货款票种号')] = (f'=IF({t1}<>"",{tk_no(t1)},IF({c("供应商号")}>0,INDEX(位_默认票种号,{c("供应商号")}),0))')
        ws[c('货款税率')] = f'=IF(AND(P_纳税人=1,{c("货款票种号")}>0),INDEX(票种_税率,{c("货款票种号")}),0)'
        t2 = txt(x('运费票种'))
        ws[c('运费票种号')] = (f'=IF({t2}<>"",{tk_no(t2)},IF({c("物流号")}>0,INDEX(位_默认票种号,{c("物流号")}),0))')
        freight(ws, c, x, qty=q, per=x('每吨运费'))
        # —— 有效之后 ——
        ok = c('有效')
        ws[ok] = f'=IF(AND({c("末行")}>0,{c("错")}=""),1,0)'
        ws[c('月')] = f'={mon(d)}'
        ws[c('货款计')] = f'=IF({ok}=1,IF({c("待定价")}=1,{c("暂估额")},{c("货款含税")}),0)'
        ws[c('货税')] = f'=ROUND({c("货款计")}/(1+{c("货款税率")})*{c("货款税率")},2)'
        ws[c('运费计')] = f'=IF({ok}=1,{c("运费含税")},0)'
        ws[c('运税')] = f'=ROUND({c("运费计")}/(1+{c("运费税率")})*{c("运费税率")},2)'
        ws[c('入账额')] = f'=ROUND({c("货款计")}-{c("货税")}+{c("运费计")}-{c("运税")},2)'
        ws[c('运费入账')] = f'=ROUND({c("运费计")}-{c("运税")},2)'
        ws[c('借方号')] = (f'=IF({k}>0,IF(INDEX(料档_耗用,{k})=1,INDEX(料档_计入号,{k}),{kh("1403")}),'
                         f'IF({p}>0,{kh("1405")},0))')
        ws[c('进库')] = f'=IF({k}>0,IF(INDEX(料档_耗用,{k})=1,0,1),IF({p}>0,2,0))'
        ws[c('料月键')] = f'=IF(AND({ok}=1,{c("进库")}=1),{k}*100+{c("月")},0)'
        ws[c('品月键')] = f'=IF(AND({ok}=1,{c("进库")}=2),{p}*100+{c("月")},0)'
        mat = txt(x('物料'))
        ws[c('错')] = '=' + chain([
            (f'{c("末行")}=0', '""')] + date_errs(d, x('日期')) + [
            (f'{mat}=""', '"✗ 没填物料"'),
            (f'AND({k}=0,{p}=0)', f'"✗ 「"&{mat}&"」不在【物料档案】里，先去加一行"'),
            (f'{c("借方号")}=0', f'"✗ 「"&{mat}&"」是进货即耗用，物料档案里「计入明细」没设科目"'),
            (f'{txt(x("供应商"))}=""', '"✗ 没填供应商"'),
            (f'{c("供应商号")}=0', f'"✗ 「"&{txt(x("供应商"))}&"」不在【往来单位】里，先去加一行"'),
            (any_bad([x(f) for f in ('数量', '单价', '货款金额', '每吨运费', '运费', '税点', '运费金额')]),
             '"✗ 数量、单价、金额、运费要填数字"'),
            (f'AND({q}=0,{c("货款含税")}=0)', '"✗ 没填数量（扣款、折让：数量空着、货款金额填负数）"'),
            (f'AND({t1}<>"",{c("货款票种号")}=0)', '"✗ 货款票种不在【基础资料】⑥票种里"'),
            (f'AND({t2}<>"",{c("运费票种号")}=0)', '"✗ 运费票种不在【基础资料】⑥票种里"'),
            (f'AND({c("运费含税")}<>0,{c("物流号")}=0)',
             f'"✗ 物流单位「"&{c("物流单位")}&"」不在【往来单位】里，先去加一行（类型选物流）"'),
        ])
        ws[c('提醒')] = '=' + chain([
            (f'{c("末行")}=0', '""'),
            (f'AND({c("待定价")}=1,{c("_参考价")}=0)', '"⚠ 还没定价，物料档案也没填参考单价：先按 0 暂估，定了价回这一行补单价"'),
            (f'{c("待定价")}=1', '"⚠ 还没定价：按物料档案的参考单价暂估，定了价回这一行补单价（日期别改）"'),
            (f'IF(AND(ISNUMBER({amt}),ISNUMBER({x("单价")})),AND({q}<>0,ABS(ROUND({amt},2)-ROUND({q}*{c("单价")},2))>0.5),FALSE)',
             '"⚠ 货款金额跟数量×单价对不上（按填的金额算）"'),
            (f'IF(AND(ISNUMBER({x("运费")}),ISNUMBER({x("每吨运费")})),ABS(ROUND({x("运费")},2)-ROUND({q}*{x("每吨运费")},2))>0.5,FALSE)',
             '"⚠ 运费跟数量×每吨运费对不上（按填的运费算）"'),
        ] + freight_warns(c, x) + [
            (f'AND({c("供应商号")}>0,INDEX(位_是供应商,MAX(1,{c("供应商号")}))=0)',
             f'"⚠ 「"&{c("供应商")}&"」在往来单位里不是供应商（类型改成供应商或客户和供应商）"'),
            (f'AND({c("物流号")}>0,INDEX(位_是供应商,MAX(1,{c("物流号")}))=0)',
             f'"⚠ 物流单位「"&{c("物流单位")}&"」在往来单位里不是物流、供应商"'),
            (f'{q}<0', '"ℹ 数量是负数：退货"'),
            (f'AND({q}=0,{c("货款含税")}<0)', '"ℹ 扣款、折让：只冲金额，不动库存数量"'),
            (f'{p}>0', '"ℹ 选的是产品：算外购成品，进库存商品"'),
        ])
        has = has_any([x(f) for f in C if f != '备注'])
        _tail(ws, r, c, SH_PUR, has)
        ws[c('排序键')] = f'=IF({ok}=1,{d}*{M}+$A{r},"")'
        ws[c('到货键')] = f'=IF(AND({ok}=1,{q}>0),{c("供应商号")}*{M}+{d},"")'
        _fmt(ws, [d], DATE)
        _fmt(ws, [c('月'), c('排序键'), c('到货键')], '0')


# ───────────────────────── _销：销售台账 ─────────────────────────
XS_PRIV = dict(税点率='AJ')


def prd_default(t):
    """产品格空＝默认产品（P_默认品号，那个产品要有效）；否则按名称、别名找产品"""
    dp = 'MAX(1,P_默认品号)'
    return f'IF({t}="",IF(INDEX(品档_有效,{dp})=1,{dp},0),{item_no(t, "品档")})'


def build_xs(ws):
    _hdr(ws, XS_COLS, XS_PRIV)
    Z, C = XS, SALE_COLS
    for i in range(N_SALE):
        r, n = i + 2, i + 1
        ws[f'A{r}'] = n
        x = lambda k: _ix(SH_SALE, C[k], r)
        priv = {'_率': f'{XS_PRIV["税点率"]}{r}'}
        c = lambda k: priv[k] if k in priv else f'{Z[k]}{r}'
        d, p, k = c('日期'), c('品号'), c('料号')
        xh, xd = x('提货日期'), x('日期')
        ws[d] = f'=IF({raw(xh)},{vdate(xh)},{vdate(xd)})'
        un = norm(x('客户'))
        ws[c('客户号')] = f'={unit_no(un)}'
        ws[c('客户')] = f'=IF({c("客户号")}>0,INDEX(位_简称,{c("客户号")}),{un})'
        t = norm(x('产品'))
        ws[p] = f'={prd_default(t)}'
        ws[k] = f'=IF(AND({p}=0,{t}<>""),{item_no(t, "料档")},0)'
        q = c('吨位')
        ws[q] = f'={numv(x("吨位"))}'
        ws[c('单价')] = f'={numv(x("单价"))}'
        amt = x('金额')
        rate = 'IF(P_纳税人=1,P_销项税率,P_征收率)'
        base = f'IF(ISNUMBER({amt}),ROUND({amt},2),ROUND({q}*{c("单价")},2))'
        ws[c('金额含税')] = f'=IF(P_销售含税=1,{base},ROUND({base}*(1+{rate}),2))'
        ws[c('待定价')] = f'=IF(AND({q}<>0,NOT({raw(x("单价"))}),NOT({raw(amt)})),1,0)'
        t1 = txt(x('票种'))
        cu = c('客户号')
        dflt = (f'IF({cu}=0,0,IF(INDEX(位_默认票种号,{cu})>0,INDEX(位_默认票种号,{cu}),'
                f'IF(INDEX(位_要开票,{cu})=1,{tk_no(chr(34) + "专票13%" + chr(34))},{tk_no(chr(34) + "无票" + chr(34))})))')
        ws[c('票种号')] = f'=IF({t1}<>"",{tk_no(t1)},{dflt})'
        fp = norm(x('发票号'))
        ws[c('发票号键')] = f'=IF({fp}="","","#"&{fp})'
        lt = norm(x('物流单位'))
        ws[c('物流号')] = f'={unit_no(lt)}'
        ws[c('物流单位')] = f'=IF({lt}="","",IF({c("物流号")}>0,INDEX(位_简称,{c("物流号")}),{lt}))'
        t2 = txt(x('运费票种'))
        ws[c('运费票种号')] = (f'=IF({t2}<>"",{tk_no(t2)},IF({c("物流号")}>0,INDEX(位_默认票种号,{c("物流号")}),0))')
        freight(ws, c, x)
        # —— 有效之后 ——
        ok = c('有效')
        ws[ok] = f'=IF(AND({c("末行")}>0,{c("错")}=""),1,0)'
        ws[c('月')] = f'={mon(d)}'
        ws[c('收入含税')] = f'=IF(AND({ok}=1,{c("待定价")}=0),{c("金额含税")},0)'
        ws[c('销项')] = f'=ROUND({c("收入含税")}/(1+{rate})*{rate},2)'
        ws[c('收入')] = f'=ROUND({c("收入含税")}-{c("销项")},2)'
        ws[c('运费计')] = f'=IF({ok}=1,{c("运费含税")},0)'
        ws[c('运税')] = f'=ROUND({c("运费计")}/(1+{c("运费税率")})*{c("运费税率")},2)'
        ws[c('运费入账')] = f'=ROUND({c("运费计")}-{c("运税")},2)'
        ws[c('出库')] = f'=IF({p}>0,1,IF({k}>0,2,0))'
        ws[c('出库键')] = f'=IF({ok}=1,IF({c("出库")}=1,{p}*100+{c("月")},IF({c("出库")}=2,{k}*100+{c("月")},0)),0)'
        cust, prd = txt(x('客户')), txt(x('产品'))
        ws[c('错')] = '=' + chain([
            (f'{c("末行")}=0', '""'),
            (f'{d}=0', f'IF({raw(xh)},IF(ISTEXT({xh}),{TXT_DATE},"✗ 提货日期看不懂（输 2026-8-1 或 2026/8/1）"),'
                       f'IF({raw(xd)},IF(ISTEXT({xd}),{TXT_DATE},"✗ 日期看不懂（输 2026-8-1 或 2026/8/1）"),"✗ 没填日期"))'),
            (f'{d}<P_建账日', '"✗ 提货日期早于建账日（建账前发的货已经在期初应收里了）"'),
            (f'{d}>P_年末', '"✗ 提货日期不在本年度（第二年请开新一本）"'),
            (f'{cust}=""', '"✗ 没填客户"'),
            (f'{cu}=0', f'"✗ 「"&{cust}&"」不在【往来单位】里，先去加一行"'),
            (f'AND({p}=0,{k}=0)', f'IF({prd}="","✗ 默认产品在【物料档案】里找不到（先在物料档案②产品加一行）",'
                                  f'"✗ 「"&{prd}&"」不在【物料档案】里")'),
            (any_bad([x(f) for f in ('单价', '吨位', '金额', '运费', '税点', '运费金额')]), '"✗ 单价、吨位、金额、运费要填数字"'),
            (f'AND({q}=0,{c("金额含税")}=0)', '"✗ 没填吨位（客户扣款、折让：吨位空着、金额填负数）"'),
            (f'AND({t1}<>"",{c("票种号")}=0)', '"✗ 票种不在【基础资料】⑥票种里"'),
            (f'AND({t2}<>"",{c("运费票种号")}=0)', '"✗ 运费票种不在【基础资料】⑥票种里"'),
            (f'AND({c("运费含税")}<>0,{c("物流号")}=0)',
             f'IF({lt}="","✗ 填了运费要填物流单位（运费付给谁）","✗ 物流单位「"&{c("物流单位")}&"」不在【往来单位】里，先去加一行")'),
        ])
        ws[c('提醒')] = '=' + chain([
            (f'{c("末行")}=0', '""'),
            (f'{c("待定价")}=1', '"⚠ 还没定价：出库照算，收入等补了价才记（定了价回这一行补单价，日期别改）"'),
            (f'IF(AND(ISNUMBER({amt}),ISNUMBER({x("单价")})),AND({q}<>0,ABS(ROUND({amt},2)-ROUND({q}*{c("单价")},2))>0.5),FALSE)',
             '"⚠ 金额跟吨位×单价对不上（按填的金额算）"'),
        ] + freight_warns(c, x) + [
            (f'AND({cu}>0,INDEX(位_是客户,MAX(1,{cu}))=0)', f'"⚠ 「"&{c("客户")}&"」在往来单位里不是客户（类型改成客户或客户和供应商）"'),
            (f'AND({k}>0,INDEX(料档_耗用,MAX(1,{k}))=1)', '"⚠ 这种料是进货即耗用的，库存里没有数（卖出去成本按 0 算）"'),
            (f'{q}<0', '"ℹ 吨位是负数：退货"'),
            (f'AND({q}=0,{c("金额含税")}<0)', '"ℹ 扣款、折让：只冲收入，不动库存"'),
            (f'{k}>0', '"ℹ 卖的是原料（废料）：记其他业务收入"'),
        ])
        has = has_any([x(f) for f in C if f != '备注'])
        _tail(ws, r, c, SH_SALE, has)
        ws[c('排序键')] = f'=IF({ok}=1,{d}*{M}+$A{r},"")'
        ws[c('发货键')] = f'=IF(AND({ok}=1,{q}>0),{cu}*{M}+{d},"")'
        _fmt(ws, [d], DATE)
        _fmt(ws, [c('月'), c('排序键'), c('发货键')], '0')


# ───────────────────────── _产：生产日报 ─────────────────────────
CHAN_PRIV = dict(坏槽='AF', 空槽='AG')


def build_chan(ws):
    _hdr(ws, CHAN_COLS, CHAN_PRIV)
    Z, C = CHAN, PROD_COLS
    for i in range(N_PROD):
        r, n = i + 2, i + 1
        ws[f'A{r}'] = n
        x = lambda k: _ix(SH_PROD, C[k], r)
        xs = lambda j: _ix(SH_PROD, PROD_SLOT_COLS[j - 1], r)
        priv = {'_坏槽': f'{CHAN_PRIV["坏槽"]}{r}', '_空槽': f'{CHAN_PRIV["空槽"]}{r}'}
        c = lambda k: priv[k] if k in priv else f'{Z[k]}{r}'
        d, p, h = c('日期'), c('品号'), c('回收料号')
        ws[d] = f'={vdate(x("日期"))}'
        t = norm(x('产品'))
        ws[p] = f'={prd_default(t)}'
        ws[c('产量袋')] = f'={numv(x("产量"))}'
        ws[c('每袋吨')] = f'=IF({p}>0,INDEX(品档_每袋吨,{p}),0)'
        ws[c('成品吨')] = f'=ROUND({c("产量袋")}*{c("每袋吨")},3)'
        ws[c('废料袋')] = f'={numv(x("废料"))}'
        ws[h] = f'=IF({p}>0,INDEX(品档_回收料号,{p}),0)'
        ws[c('废料吨')] = f'=ROUND({c("废料袋")}*IF({h}>0,INDEX(料档_每袋吨,{h}),1),3)'
        for j in range(1, N_SLOT + 1):
            ws[c(f'槽{j}')] = f'={numv(xs(j))}'
        s1, s12 = c('槽1'), c(f'槽{N_SLOT}')
        ws[c('投料合计')] = f'=SUM({s1}:{s12})'
        ws[c('_坏槽')] = '=' + chain([(badnum(xs(j)), str(j)) for j in range(1, N_SLOT + 1)], '0')
        ws[c('_空槽')] = '=' + chain([(f'AND({c(f"槽{j}")}<>0,N(INDEX(槽_料号,{j}))=0)', str(j)) for j in range(1, N_SLOT + 1)], '0')
        ok = c('有效')
        ws[ok] = f'=IF(AND({c("末行")}>0,{c("错")}=""),1,0)'
        ws[c('月')] = f'={mon(d)}'
        ws[c('品月键')] = f'=IF({ok}=1,{p}*100+{c("月")},0)'
        ws[c('回月键')] = f'=IF(AND({ok}=1,{h}>0),{h}*100+{c("月")},0)'
        prd = txt(x('产品'))
        ws[c('错')] = '=' + chain([
            (f'{c("末行")}=0', '""')] + date_errs(d, x('日期')) + [
            (f'{p}=0', f'IF({prd}="","✗ 默认产品在【物料档案】里找不到（先在物料档案②产品加一行）",'
                       f'"✗ 「"&{prd}&"」不是【物料档案】②里的产品")'),
            (any_bad([x('产量'), x('废料')]), '"✗ 产量、废料要填数字（袋数）"'),
            (f'{c("_坏槽")}>0', f'"✗ 原料第 "&{c("_坏槽")}&" 列不是数字"'),
            (f'{c("_空槽")}>0', f'"✗ 原料第 "&{c("_空槽")}&" 列在【物料档案】里没指定是哪种料（「生产日报第几列」），投的数没法算"'),
            (f'AND({c("废料袋")}<>0,{h}=0)', '"✗ 废料归到哪种回收料找不到（物料档案产品的「废料归入」或基础资料「默认回收料」）"'),
            (f'AND({c("产量袋")}=0,{c("废料袋")}=0,{c("投料合计")}=0)', '"✗ 没填产量、废料、投料"'),
        ])
        ws[c('提醒')] = '=' + chain([
            (f'{c("末行")}=0', '""'),
            (f'MIN({c("产量袋")},{c("废料袋")},{s1}:{s12})<0', '"⚠ 有负数（冲减前面登多了的）"'),
            (f'AND({c("产量袋")}>0,{c("投料合计")}=0)', '"⚠ 有产量没投料（投料登在哪一天都行，月底合计对就行）"'),
            (f'AND({p}>0,COUNTIFS(产_日期,{d},产_品号,{p})>1)', '"⚠ 这一天这个产品登了不止一行（都算）"'),
            (f'AND({c("产量袋")}=0,{c("投料合计")}<>0)', '"ℹ 产量是 0：投的料留在在产品，等有产量一起转"'),
        ])
        has = has_any([x(f) for f in ('日期', '产品', '产量', '废料')] + [xs(j) for j in range(1, N_SLOT + 1)])
        _tail(ws, r, c, SH_PROD, has)
        ws[c('排序键')] = f'=IF({ok}=1,{d}*{M}+$A{r},"")'
        _fmt(ws, [d], DATE)
        _fmt(ws, [c('月'), c('排序键')], '0')


# ───────────────────────── _工：工资表 ─────────────────────────
GONG_PRIV = dict(年月='P')


def build_gong(ws):
    _hdr(ws, GONG_COLS, GONG_PRIV)
    Z, C = GONG, WAGE_COLS
    for i in range(N_WAGE):
        r, n = i + 2, i + 1
        ws[f'A{r}'] = n
        x = lambda k: _ix(SH_WAGE, C[k], r)
        priv = {'_年月': f'{GONG_PRIV["年月"]}{r}'}
        c = lambda k: priv[k] if k in priv else f'{Z[k]}{r}'
        ym, m, pn = c('_年月'), c('月'), c('人员号')
        ws[ym] = f'={ym_any(x("月份"))}'
        ws[m] = f'=IF(AND({ym_ok(ym)},INT({ym}/100)=P_年度),MOD({ym},100),0)'
        nm = txt(x('姓名'))
        ws[pn] = f'=IF({nm}="",0,IFERROR(MATCH({esc(nm)},人员_姓名,0),0))'
        dp = txt(x('部门'))
        ws[c('部门号')] = f'=IF({dp}<>"",{idx_of(DEPTS, dp, 0)},IF({pn}>0,INDEX(人员_部门号,{pn}),0))'
        four = 'ROUND(' + '+'.join(numv(x(f)) for f in ('基本工资', '计件提成', '加班', '补贴')) + ',2)'
        ws[c('应发')] = f'=IF(ISNUMBER({x("应发合计")}),ROUND({x("应发合计")},2),{four})'
        for f in ('扣社保', '扣借支', '其他扣款'):
            ws[c(f)] = f'=ROUND({numv(x(f))},2)'
        calc = f'ROUND({c("应发")}-{c("扣社保")}-{c("扣借支")}-{c("其他扣款")},2)'
        ws[c('实发')] = f'=IF(ISNUMBER({x("实发")}),ROUND({x("实发")},2),{calc})'
        ok = c('有效')
        ws[ok] = f'=IF(AND({c("末行")}>0,{c("错")}=""),1,0)'
        ws[c('部门月键')] = f'=IF({ok}=1,{c("部门号")}*100+{m},0)'
        money = ('基本工资', '计件提成', '加班', '补贴', '应发合计', '扣社保', '扣借支', '其他扣款', '实发')
        ws[c('错')] = '=' + chain([
            (f'{c("末行")}=0', '""'),
            (f'{ym}=-1', '"✗ 没填月份"'),
            (f'NOT({ym_ok(ym)})', '"✗ 月份看不懂（填 8、2026-8 或 2026年8月）"'),
            (f'INT({ym}/100)<>P_年度', '"✗ 不是本年度的工资（第二年请开新一本）"'),
            (f'{m}<P_建账月', '"✗ 月份在建账月以前（建账前欠的工资填【基础资料】⑧期初）"'),
            (f'{nm}=""', '"✗ 没填姓名"'),
            (f'{pn}=0', f'"✗ 「"&{nm}&"」不在【基础资料】⑤人员里，先去加一行"'),
            (f'AND({dp}<>"",{c("部门号")}=0)', '"✗ 部门只能选：生产、车间管理、管理、销售"'),
            (f'{c("部门号")}=0', '"✗ 没填部门（【基础资料】⑤人员里也没有）"'),
            (any_bad([x(f) for f in money]), '"✗ 工资各项要填数字"'),
            (f'{c("应发")}=0', '"✗ 应发是 0（基本工资、计件、加班、补贴都没填）"'),
        ])
        ws[c('提醒')] = '=' + chain([
            (f'{c("末行")}=0', '""'),
            (f'AND(ISNUMBER({x("实发")}),ABS({c("实发")}-{calc})>0.005)', '"⚠ 实发跟应发－扣款对不上（按填的实发算）"'),
            (f'AND(ISNUMBER({x("应发合计")}),{four}<>0,ABS({c("应发")}-{four})>0.005)',
             '"⚠ 应发合计跟基本＋计件＋加班＋补贴对不上（按填的应发合计算）"'),
            (f'AND({pn}>0,{m}>0,COUNTIFS(工_人员号,{pn},工_月,{m})>1)', '"⚠ 这个人这个月登了不止一行（都算）"'),
            (f'{c("实发")}<0', '"⚠ 实发是负数"'),
        ])
        has = has_any([x(f) for f in C if f != '备注'])
        _tail(ws, r, c, SH_WAGE, has)
        _fmt(ws, [ym, m], '0')


# ───────────────────────── _票：发票登记 ─────────────────────────
PIAO_PRIV = dict(资行='X')
FEE1 = lambda v: 'OR(' + ','.join(f'{v}={kh(cd)}' for cd in FEE_PARENTS) + ')'


def build_piao(ws):
    _hdr(ws, PIAO_COLS, PIAO_PRIV)
    Z, C = PIAO, INV_COLS
    for i in range(N_INV):
        r, n = i + 2, i + 1
        ws[f'A{r}'] = n
        x = lambda k: _ix(SH_INV, C[k], r)
        priv = {'_资行': f'{PIAO_PRIV["资行"]}{r}'}
        c = lambda k: priv[k] if k in priv else f'{Z[k]}{r}'
        d, u, tk = c('日期'), c('单位号'), c('票种号')
        ws[d] = f'={vdate(x("日期"))}'
        un = norm(x('往来单位'))
        ws[u] = f'={unit_no(un)}'
        ws[c('单位')] = f'=IF({u}>0,INDEX(位_简称,{u}),{un})'
        dr = txt(x('方向'))
        ws[c('方向')] = (f'=IF({dr}="开出",1,IF({dr}="收到",2,IF({dr}="",IF(AND({u}>0,INDEX(位_类型,MAX(1,{u}))="客户"),1,2),0)))')
        fp = norm(x('发票号码'))
        ws[c('号码键')] = f'=IF({fp}="","","#"&{fp})'
        t1 = txt(x('票种'))
        ws[tk] = f'=IF({t1}<>"",{tk_no(t1)},IF({u}>0,INDEX(位_默认票种号,{u}),0))'
        ws[c('价税合计')] = f'=ROUND({numv(x("价税合计"))},2)'
        rate = f'IF({tk}>0,INDEX(票种_税率,{tk}),0)'
        calc = f'ROUND({c("价税合计")}/(1+{rate})*{rate},2)'
        ws[c('税额')] = f'=IF(ISNUMBER({x("税额")}),ROUND({x("税额")},2),{calc})'
        ws[c('不含税')] = f'=ROUND({c("价税合计")}-{c("税额")},2)'
        dk = norm(x('垫付编号'))
        ws[c('垫付键')] = f'=IF({dk}="","","#"&{dk})'
        ok = c('有效')
        ws[ok] = f'=IF(AND({c("末行")}>0,{c("错")}=""),1,0)'
        ws[c('月')] = f'={mon(d)}'
        key, zr = c('垫付键'), c('_资行')
        ws[zr] = f'=IF({key}="",0,IFERROR(MATCH({esc(key)},资_编号键,0),0))'
        z = c('补票行')
        ws[z] = (f'=IF(OR({zr}=0,{c("方向")}<>2),0,IF(AND(INDEX(资_有效,{zr})=1,INDEX(资_待补票,{zr})=1),{zr},0))')
        cat = f'INDEX(资_类别号,{z})'
        re_ = f'INDEX(资_报销项目号,{z})'
        nod = (f'OR(INDEX(类别_不抵,MAX(1,{cat}))=1,IF({re_}>0,INDEX(报销_不抵,MAX(1,{re_})),0)=1)')
        ws[c('补票额')] = (f'=IF({z}=0,0,IF(AND({ok}=1,P_纳税人=1,{tk}>0,INDEX(票种_税率,MAX(1,{tk}))>0,'
                         f'{FEE1(f"INDEX(资_原对方一级,{z})")},NOT({nod})),{c("税额")},0))')
        ws[c('补票目标号')] = f'=IF({z}>0,INDEX(资_原对方号,{z}),0)'
        ws[c('补票键')] = f'=IF({c("补票额")}<>0,{c("补票目标号")}*100+{c("月")},0)'
        hk = c('号码键')
        ws[c('销售重复')] = f'=IF({hk}="",0,IF(COUNTIF(销_发票号键,{esc(hk)})>0,1,0))'
        unit_t = txt(x('往来单位'))
        ws[c('错')] = '=' + chain([
            (f'{c("末行")}=0', '""')] + date_errs(d, x('日期')) + [
            (f'{unit_t}=""', '"✗ 没填往来单位"'),
            (f'{c("方向")}=0', '"✗ 方向只能填开出或收到（空着＝按往来单位类型）"'),
            (any_bad([x('价税合计'), x('税额')]), '"✗ 价税合计、税额要填数字"'),
            (f'{c("价税合计")}=0', '"✗ 没填价税合计"'),
            (f'AND({t1}<>"",{tk}=0)', '"✗ 票种不在【基础资料】⑥票种里"'),
        ])
        ws[c('提醒')] = '=' + chain([
            (f'{c("末行")}=0', '""'),
            (f'AND({key}<>"",{c("方向")}=1)', '"⚠ 开出的发票不用填垫付编号（不算补票）"'),
            (f'AND({key}<>"",{zr}=0)', '"⚠ 垫付编号在【资金台帐】里找不到（不算补票）"'),
            (f'AND({key}<>"",{z}=0)', '"⚠ 资金台帐那一笔不是「待补票」（或那一行有问题），不算补票"'),
            (f'AND({key}<>"",COUNTIF(票_垫付键,{esc(key)})>1)', '"⚠ 这个垫付编号登了不止一张票（都算）"'),
            (f'AND({hk}<>"",COUNTIF(票_号码键,{esc(hk)})>1)', '"⚠ 发票号码重复了"'),
            (f'AND(ISNUMBER({x("税额")}),ABS({c("税额")}-{calc})>1)', '"⚠ 税额跟价税合计×税率算的差 1 元以上"'),
            (f'AND({z}>0,{c("补票额")}=0)', '"ℹ 补到票了，但这张票不抵进项（普票、不得抵扣、小规模或不是费用）：只消掉待补票"'),
            (f'{z}>0', '"ℹ 补票：这张票的进项在登票这个月抵扣"'),
            (f'{c("销售重复")}=1', '"ℹ 这个发票号在销售台账里已经填了（已开票不重复算）"'),
            (f'{u}=0', f'"⚠ 「"&{unit_t}&"」不在【往来单位】里（不算进哪一家的已收票、已开票）"'),
            (f'AND({c("方向")}=1,INDEX(位_是客户,MAX(1,{u}))=0)', '"⚠ 开出给不是客户的单位？方向是不是反了"'),
            (f'AND({c("方向")}=2,INDEX(位_类型,MAX(1,{u}))="客户")', '"⚠ 从客户那收到发票？方向是不是反了"'),
            (f'{c("价税合计")}<0', '"ℹ 负数：红字发票"'),
        ])
        has = has_any([x(f) for f in C if f != '备注'])
        _tail(ws, r, c, SH_INV, has)
        ws[c('排序键')] = f'=IF({ok}=1,{d}*{M}+$A{r},"")'
        _fmt(ws, [d], DATE)
        _fmt(ws, [c('月'), c('排序键')], '0')


# ───────────────────────── _手：手工凭证 ─────────────────────────
SHOU_PRIV = dict(行错文='W')
MJ_BAN_CASH = ['1001', '1002', '1012', '1121']
MJ_BAN_STOCK = ['1403', '1405', '4001']
MJ_WARN2 = ['224101', '224102']
MJ_WARN1 = ['2001', '2501', '3001', '2231', '2232']


def build_shou(ws):
    _hdr(ws, SHOU_COLS, SHOU_PRIV)
    Z, C = SHOU, MJ_COLS
    N = N_MJ + 1
    for i in range(N_MJ):
        r, n = i + 2, i + 1
        ws[f'A{r}'] = n
        x = lambda k: _ix(SH_MJ, C[k], r)
        priv = {'_行错文': f'{SHOU_PRIV["行错文"]}{r}'}
        c = lambda k: priv[k] if k in priv else f'{Z[k]}{r}'
        d, m, key, kn, k1 = c('日期'), c('月'), c('组键'), c('科目号'), c('一级号')
        ws[d] = f'={vdate(x("日期"))}'
        ws[m] = f'={mon(d)}'
        g = txt(x('凭证组号'))
        dig = g
        for ch in '0123456789':
            dig = f'SUBSTITUTE({dig},"{ch}","")'
        ws[c('组号')] = f'=IF({g}="","",IF({dig}="",VALUE({g})&"",{g}))'
        ws[key] = f'=IF({c("组号")}="","",{c("组号")}&"|"&{m})'
        cd = code_of(x('科目'))
        ws[kn] = f'={kmno(cd)}'
        ws[c('借')] = f'=ROUND({numv(x("借方"))},2)'
        ws[c('贷')] = f'=ROUND({numv(x("贷方"))},2)'
        ws[k1] = f'=IF({kn}>0,INDEX(科目_一级号,{kn}),0)'
        dn = norm(x('往来明细'))
        ws[c('单位号')] = f'=IF(AND({k1}>0,OR({k1}={kh("1122")},{k1}={kh("2202")})),{unit_no(dn)},0)'
        ws[c('明细')] = f'=IF({c("单位号")}>0,INDEX(位_简称,{c("单位号")}),{dn})'
        subj = txt(x('科目'))
        in1 = lambda codes: 'OR(' + ','.join(f'{k1}={kh(cc)}' for cc in codes) + ')'
        ws[c('_行错文')] = '=' + chain([
            (f'{c("末行")}=0', '""')] + date_errs(d, x('日期')) + [
            (f'{c("组号")}=""', '"✗ 没填凭证组号（同一张凭证的几行填同一个号）"'),
            (f'{subj}=""', '"✗ 没填科目"'),
            (f'{kn}=0', f'"✗ 科目「"&{subj}&"」不在【会计科目】里"'),
            (f'INDEX(科目_有效,{kn})<>1', f'"✗ 科目「"&{subj}&"」在【会计科目】里有问题（看那一行的校验）"'),
            (in1(MJ_BAN_CASH), '"✗ 现金、银行、微信、承兑的钱请记【资金台帐】"'),
            (in1(MJ_BAN_STOCK), '"✗ 原材料、库存商品、生产成本不能手工调（走进货物流登记、月末盘点）"'),
            (any_bad([x('借方'), x('贷方')]), '"✗ 借方、贷方要填数字"'),
            (f'AND({c("借")}<>0,{c("贷")}<>0)', '"✗ 借方、贷方只能填一个（分两行）"'),
            (f'AND({c("借")}=0,{c("贷")}=0)', '"✗ 没填金额"'),
            (f'AND({c("单位号")}=0,OR({k1}={kh("1122")},{k1}={kh("2202")}))',
             '"✗ 应收账款、应付账款的往来/明细要填【往来单位】里的单位"'),
        ])
        err = c('行错')
        ws[err] = f'=IF({c("_行错文")}<>"",1,0)'
        ek = esc(key)
        cnt = f'COUNTIF(手_组键,{ek})'
        bad_in = f'COUNTIFS(手_组键,{ek},手_行错,1)'
        same_d = f'COUNTIFS(手_组键,{ek},手_日期,{d})'
        sdr, scr = f'SUMIF(手_组键,{ek},手_借)', f'SUMIF(手_组键,{ek},手_贷)'
        ws[c('组平')] = (f'=IF(OR({key}="",{err}=1),0,IF(AND({bad_in}=0,{same_d}={cnt},ROUND({sdr}-{scr},2)=0),1,0))')
        ws[c('记账')] = f'=IF({c("组平")}=1,1,0)'
        ws[c('组首')] = f'=IF({key}="",0,IFERROR(MATCH({ek},手_组键,0),0))'
        ws[c('组内序')] = f'=IF({c("记账")}=1,COUNTIFS(${Z["组键"]}$2:{key},{ek},${Z["记账"]}$2:{c("记账")},1),0)'
        ws[c('印键')] = f'=IF({c("记账")}=1,{c("组首")}*1000+{c("组内序")},0)'
        ws[c('键')] = f'=IF({c("记账")}=1,{kn}*100+{m},0)'
        grp = (f'IF({bad_in}>0,"✗ 同一组别的行有问题，整组不记账",IF({same_d}<{cnt},"✗ 同一组的日期不一样（同一张凭证日期要相同）",'
               f'"✗ 这一组借贷不平：借 "&TEXT({sdr},"#,##0.00")&"，贷 "&TEXT({scr},"#,##0.00")&"，整组不记账"))')
        ws[c('错')] = f'=IF({c("_行错文")}<>"",{c("_行错文")},IF(AND({c("末行")}>0,{c("组平")}=0),{grp},""))'
        lv = f'INDEX(科目_级次,MAX(1,{kn}))'
        ws[c('提醒')] = '=' + chain([
            (f'OR({c("末行")}=0,{kn}=0)', '""'),
            (f'OR({kn}={kh("224101")},{kn}={kh("224102")},' + ','.join(f'{k1}={kh(cc)}' for cc in MJ_WARN1) + ')',
             '"⚠ 垫付、借款、实收资本、应付利息/利润平时由资金台帐和⑦自动出：手工调了，借款明细、首页欠款跟着变，确认一下"'),
            (f'AND({lv}=1,COUNTIF(科目_一级号,{kn})>1)', f'"⚠ 「"&{subj}&"」下面有明细科目，最好记到明细"'),
        ])
        has = has_any([x(f) for f in C if f != '备注'])
        _tail(ws, r, c, SH_MJ, has)
        _fmt(ws, [d], DATE)
        _fmt(ws, [m, c('键')], '0')


# ───────────────────────── _盘：月末盘点 ─────────────────────────
PAN_PRIV = dict(年月='N')


def build_pan(ws):
    _hdr(ws, PAN_COLS, PAN_PRIV)
    Z, C = PAN, CNT_COLS
    for i in range(N_CNT):
        r, n = i + 2, i + 1
        ws[f'A{r}'] = n
        x = lambda k: _ix(SH_CNT, C[k], r)
        priv = {'_年月': f'{PAN_PRIV["年月"]}{r}'}
        c = lambda k: priv[k] if k in priv else f'{Z[k]}{r}'
        ym, m, k, p = c('_年月'), c('月'), c('料号'), c('品号')
        ws[ym] = f'={ym_any(x("月份"))}'
        ws[m] = f'=IF(AND({ym_ok(ym)},INT({ym}/100)=P_年度),MOD({ym},100),0)'
        nm = norm(x('物料'))
        ws[k] = f'={item_no(nm, "料档")}'
        ws[p] = f'=IF({k}>0,0,{item_no(nm, "品档")})'
        xq = x('实盘数量')
        ws[c('实盘')] = f'={numv(xq)}'
        ok = c('有效')
        ws[ok] = f'=IF(AND({c("末行")}>0,{c("错")}=""),1,0)'
        ws[c('键')] = f'=IF(AND({ok}=1,{k}>0),{k}*100+{m},0)'
        ws[c('品键')] = f'=IF(AND({ok}=1,{p}>0),{p}*100+{m},0)'
        g, hh = Z['键'], Z['品键']
        ws[c('重复')] = (f'=IF({ok}<>1,0,IF({k}>0,IF(COUNTIF(${g}$1:{g}{r - 1},{g}{r})>0,1,0),'
                       f'IF(COUNTIF(${hh}$1:{hh}{r - 1},{hh}{r})>0,1,0)))')
        mat = txt(x('物料'))
        ws[c('错')] = '=' + chain([
            (f'{c("末行")}=0', '""'),
            (f'{ym}=-1', '"✗ 没填月份"'),
            (f'NOT({ym_ok(ym)})', '"✗ 月份看不懂（填 8 或 2026-8）"'),
            (f'INT({ym}/100)<>P_年度', '"✗ 不是本年度的盘点（第二年请开新一本）"'),
            (f'{m}<P_建账月', '"✗ 月份在建账月以前（建账日的数填物料档案的期初数量）"'),
            (f'{mat}=""', '"✗ 没填物料"'),
            (f'AND({k}=0,{p}=0)', f'"✗ 「"&{mat}&"」不在【物料档案】里"'),
            (f'AND({k}>0,INDEX(料档_耗用,MAX(1,{k}))=1)', f'"✗ 「"&{mat}&"」是进货即耗用的（不管库存），不用盘"'),
            (f'NOT({raw(xq)})', '"✗ 没填实盘数量（盘光了填 0）"'),
            (badnum(xq), '"✗ 实盘数量要填数字"'),
            (f'{c("实盘")}<0', '"✗ 实盘数量不能是负数"'),
        ])
        ws[c('提醒')] = '=' + chain([
            (f'{c("末行")}=0', '""'),
            (f'{c("重复")}=1', '"⚠ 同一个月这种料上面已经盘过了，这一行不算（只认上面那条）"'),
        ])
        has = has_any([x(f) for f in C if f != '备注'])
        _tail(ws, r, c, SH_CNT, has)
        _fmt(ws, [ym, m], '0')


BUILDERS = {H_JIN: build_jin, H_XS: build_xs, H_CHAN: build_chan, H_GONG: build_gong, H_PIAO: build_piao,
            H_SHOU: build_shou, H_PAN: build_pan}


def build(wb, ctx=None):
    for name, fn in BUILDERS.items():
        fn(wb[name])
