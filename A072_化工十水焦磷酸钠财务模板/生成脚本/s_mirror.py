# -*- coding: utf-8 -*-
"""A072 化工 · 隐藏表（基础镜像）：_参（全书参数）、_表（基础资料镜像＋压紧下拉清单）、_科（会计科目镜像＋期初）、
   _物（物料、产品镜像＋生产日报槽表）、_固（固定支出）、_卡（固定资产卡片＋12 个月折旧）、_位（往来单位逐家＋往来余额、发票）。
   全部按「第 n 条」取数：INDEX(表!$X$表头行:$X$1048576, n+1)，插行删行不错位。别的模块只用 layout.NAMES 里的名称。
   私有列（别的模块不要引用）：_表 CE..CG（⑦ 单位号、银行、息净）；_科 T、U（费用/制费明细累计数）、W..Y（期初来源小表）；
   _物 AL2（原料块行数）；_固 U、V（开始/结束年月）；_卡 Q（处置日）、AD（本年合计）；_参 C..H（截止日、投产月的辅助格）。"""
from layout import *
from common import F_HELP, DATE

M = SORT_M


def _hdr(ws, cols):
    for c, f, _d in cols:
        ws[f'{c}1'] = f
        ws[f'{c}1'].font = F_HELP


def _lab(ws, cell, text):
    ws[cell] = text
    ws[cell].font = F_HELP


def _fmt(ws, cells, fmt):
    for a in cells:
        ws[a].number_format = fmt


def chain(pairs, default='""'):
    out = default
    for c, v in reversed(pairs):
        out = f'IF({c},{v},{out})'
    return out


def _ix(sh, col, n_ref, block=None):
    return f'INDEX({inref(sh, col, block=block)},{n_ref}+1)'


def vdate(x):
    """真日期（2000～2099 年）取整，不然 0"""
    return f'IF(ISNUMBER({x}),IF(AND({x}>=36526,{x}<73051),INT({x}),0),0)'


def txt(x):
    return f'TRIM({x}&"")'


def raw(x):
    return f'TRIM({x}&"")<>""'


def numv(x):
    return f'IF(ISNUMBER({x}),{x},0)'


def badnum(x):
    return f'AND({raw(x)},NOT(ISNUMBER({x})))'


def pct(x, dflt):
    """税率、利率：填 13 当 13%；空或不是数字＝默认"""
    return f'IF(ISNUMBER({x}),IF({x}>=1,{x}/100,{x}),{dflt})'


def in_list(items, x):
    return 'OR(' + ','.join(f'{x}="{k}"' for k in items) + ')'


def idx_of(items, t, dflt):
    """文字在固定清单里第几个（1 起）；不在＝dflt"""
    out = str(dflt)
    for i in reversed(range(len(items))):
        out = f'IF({t}="{items[i]}",{i + 1},{out})'
    return out


def code_of(x):
    """科目格：「编码 名称」截编码；数字编码转文字"""
    t = txt(x)
    return f'IFERROR(LEFT({t},FIND(" ",{t})-1),{t})'


def kmno(code):
    """科目编码（文字）→ 科目号（会计科目第几行）；0＝空或找不到"""
    return f'IF({code}="",0,IFERROR(MATCH({esc(code)},科目_编码,0),0))'


def unit_match(nm, attr='简称', dflt='0', key=None):
    """规范后的名字 → 往来单位那一家的某个属性（简称、全称、别名都认）；找不到＝dflt"""
    e = esc(nm)
    return (f'IF({nm}="",{dflt},IFERROR(INDEX(位_{attr},MATCH({e},位_简称,0)),IFERROR(INDEX(位_{attr},MATCH({e},位_全称规,0)),'
            f'IFERROR(INDEX(位_{attr},MATCH("*、"&{e}&"、*",位_别名规,0)),{dflt}))))')


def unit_no(nm):
    """规范后的名字 → 往来单位第几家（0＝不在表里）"""
    e = esc(nm)
    return (f'IF({nm}="",0,IFERROR(MATCH({e},位_简称,0),IFERROR(MATCH({e},位_全称规,0),'
            f'IFERROR(MATCH("*、"&{e}&"、*",位_别名规,0),0))))')


def alias(x):
    """别名：各种分隔符统一成「、」，前后加「、」（MATCH "*、名、*" 用）"""
    al = (f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({txt(x)},"，","、"),",","、"),"；","、"),'
          f'";","、"),"/","、"),"／","、")')
    return f'IF({txt(x)}="","","、"&{norm(al)}&"、")'


def ym_any(x, blank='-1'):
    """月份格 → 年×100＋月：1～12 当账套年度的月；2026-07 / 2026年7月 / 202607 / 真日期都认；空＝blank；看不懂＝0"""
    t = (f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({txt(x)},"年","-"),"月",""),"/","-"),".","-")')
    tv = (f'IFERROR(IF(ISNUMBER(FIND("-",{t})),VALUE(LEFT({t},4))*100+VALUE(SUBSTITUTE(MID({t},6,2),"-","")),'
          f'VALUE({t})),0)')
    num = (f'IF(AND({x}>=1,{x}<=12,INT({x})={x}),P_年度*100+{x},IF(AND({x}>=200001,{x}<=209912),{x},'
           f'IF(AND({x}>=36526,{x}<73051),YEAR({x})*100+MONTH({x}),0)))')
    return f'IF({txt(x)}="",{blank},IF(ISNUMBER({x}),{num},{tv}))'


def ym_ok(c):
    """ym_any 的结果格 c 是不是像样的年月"""
    return f'AND({c}>=200001,{c}<=209912,INT({c})={c},MOD({c},100)>=1,MOD({c},100)<=12)'


TXT_DATE = '"✗ 日期是文字：有「.」的用查找替换换成「-」；复制来的选中日期列 → 数据 → 分列 → 完成"'


# ───────────────────────── _参 ─────────────────────────
def build_par(ws):
    P = lambda k: f'{SH_BASE}!${BASE_PAR_VAL}${PAR_ROW[k]}'
    b = lambda k: PAR[k]
    home = f'{SH_HOME}!${HOME_END[0]}${HOME_END[1:]}'
    c = P('建账日')
    rows = {
        '公司名称': f'=TRIM({P("公司名称")}&"")',
        '建账日': (f'=IF(ISNUMBER({c}),IF(AND({c}>=200001,{c}<=209912),IF(AND(INT({c})={c},MOD({c},100)>=1,MOD({c},100)<=12),'
                 f'DATE(INT({c}/100),MOD({c},100),1),DATE(2026,1,1)),IF(AND({c}>=36526,{c}<73051),INT({c}),DATE(2026,1,1))),'
                 f'IFERROR(INT(DATEVALUE(TRIM({c}&""))),DATE(2026,1,1)))'),
        '年度': f'=YEAR({b("建账日")})',
        '建账月': f'=MONTH({b("建账日")})',
        '纳税人': f'=IF(TRIM({P("纳税人类型")}&"")="小规模",0,1)',
        '销项税率': f'={pct(P("销项税率"), 0.13)}',
        '征收率': f'={pct(P("征收率"), 0.01)}',
        '销售含税': f'=IF(TRIM({P("销售含税")}&"")="否",0,1)',
        '进货含税': f'=IF(TRIM({P("进货含税")}&"")="否",0,1)',
        '运费税点': f'={pct(P("运费税点"), 0.07)}',
        '运费进成本': f'=IF(TRIM({P("销货运费记入")}&"")="销售费用",0,1)',
        # 投产月：G2＝年月（空 -1，看不懂 0）；空/看不懂/早于年度＝1，晚于年度＝13
        '投产月': (f'=IF(NOT({ym_ok("$G$2")}),1,IF(INT($G$2/100)<{b("年度")},1,IF(INT($G$2/100)>{b("年度")},13,'
                f'MOD($G$2,100))))'),
        '摊销月数': f'=IF(ISNUMBER({P("摊销月数")}),IF({P("摊销月数")}>=1,ROUND({P("摊销月数")},0),36),36)',
        '回款天数': f'=IF(ISNUMBER({P("回款天数")}),IF({P("回款天数")}>0,{P("回款天数")},60),60)',
        '对账工作日': f'=IF(ISNUMBER({P("对账工作日")}),IF({P("对账工作日")}>=0,{P("对账工作日")},5),5)',
        '付款天数': f'=IF(ISNUMBER({P("付款天数")}),IF({P("付款天数")}>=0,{P("付款天数")},30),30)',
        '凭证号打印': f'=IF(TRIM({P("凭证号打印")}&"")="是",1,0)',
        '日期上限': f'={b("年末")}',
        '挡后日期': '=MAX($E$2:$E$5)',
        '近月笔数': '=SUM($F$2:$F$5)',
        '截止被挡': f'=IF(AND({b("挡后日期")}>0,{b("近月笔数")}<=2),1,0)',
        '最后日期': f'=MAX({b("建账日")},IF({b("截止被挡")}=1,{b("挡后日期")},MAX($D$2:$D$5)))',
        '截止': f'=MIN({b("年末")},MAX({b("建账日")},IF(ISNUMBER({home}),INT({home}),{b("最后日期")})))',
        '截止月': f'=MONTH({b("截止")})',
        '月初': f'=DATE(YEAR({b("截止")}),MONTH({b("截止")}),1)',
        '年末': f'=DATE({b("年度")},12,31)',
    }
    for k in ('会计主管', '记账', '出纳', '审核', '制单'):
        rows[k] = f'=TRIM({P(k)}&"")'
    # 默认产品：名称、别名；找不到＝第一个有效产品
    nm = norm(P('默认产品'))
    first = 'IFERROR(MATCH(1,品档_有效,0),1)'
    rows['默认品号'] = (f'=IF({nm}="",{first},IFERROR(MATCH({esc(nm)},品档_名称,0),'
                     f'IFERROR(MATCH("*、"&{esc(nm)}&"、*",品档_别名规,0),{first})))')
    nm = norm(P('默认回收料'))
    first = 'IFERROR(MATCH(4,料档_类别号,0),0)'
    rows['默认回收料号'] = (f'=IF({nm}="",{first},IFERROR(MATCH({esc(nm)},料档_名称,0),'
                       f'IFERROR(MATCH("*、"&{esc(nm)}&"、*",料档_别名规,0),{first})))')
    # 截止日：D 各表不超过上限的最晚日期；E 比 L（D 的最大）早 32 天以上的最晚日期；F L 前 31 天内（含 L）的笔数（照 A069）
    _lab(ws, 'C1', '来源')
    _lab(ws, 'D1', '最晚')
    _lab(ws, 'E1', '挡后')
    _lab(ws, 'F1', '近月笔数')
    for i, nm in enumerate(('资_日期', '进_日期', '销_日期', '产_日期')):
        r = 2 + i
        _lab(ws, f'C{r}', nm)
        ws[f'D{r}'] = f'=IFERROR(LARGE({nm},COUNTIF({nm},">"&{b("日期上限")})+1),0)'
        ws[f'E{r}'] = f'=IFERROR(LARGE({nm},COUNTIF({nm},">"&(MAX($D$2:$D$5)-32))+1),0)'
        ws[f'F{r}'] = f'=COUNTIFS({nm},">"&(MAX($D$2:$D$5)-31),{nm},"<="&MAX($D$2:$D$5))'
    _lab(ws, 'G1', '投产年月')
    ws['G2'] = f'={ym_any(P("投产月份"))}'
    for i, (key, _t, _tr, hr, cap) in enumerate(BASE_BLOCKS):
        if i + 1 < len(BASE_BLOCKS):
            nxt = BASE_BLOCKS[i + 1][1]
            rows[f'{key}位'] = f'=IFERROR(MATCH("{nxt}",{SH_BASE}!$A:$A,0)-ROW({SH_BASE}!$A${hr})-1,{cap})'
        else:
            rows[f'{key}位'] = f'={cap + 300}'
    for k, pre in (('资金', '资'), ('进货', '进'), ('销售', '销'), ('生产', '产'), ('工资', '工'), ('发票', '票'), ('盘点', '盘'),
                   ('手工', '手'), ('往来', '位'), ('固定', '固'), ('卡片', '卡')):
        rows[f'{k}末行'] = f'=MAX(MAX({pre}_末行),{HDR + 1})+1'
    rows['下一垫付号'] = (f'=TEXT(YEAR({b("截止")})*100+MONTH({b("截止")}),"000000")&"-X"&TEXT(MAX(资_本月垫付号)+1,"00")')
    # 列数错：每张录入表、基础表 COLUMNS(首列:最后一列) ≠ 设计值（中间插了列/删了列）
    sheets = [(sh, spec[2]) for sh, spec in INPUT_SPEC.items()] + [
        (SH_BASE, 'M'), (SH_UNIT, UNIT_LAST), (SH_ITEM, ITEM_SHOW_RM['校验']), (SH_FA, FA_LAST), (SH_FIX, FIX_LAST),
        (SH_COA, COA_LAST)]
    parts = '&'.join(f'IF(COLUMNS({sh}!$A${HDR}:${last}${HDR})<>{CI(last)},"、{sh}","")' for sh, last in sheets)
    rows['列数错'] = f'=MID({parts},2,300)'
    ws['A1'] = '全书共用的参数（自动，别改）'
    ws['A1'].font = F_HELP
    for k, cell in PAR.items():
        r = int(cell[1:])
        ws[f'A{r}'] = k
        ws[cell] = rows[k]
    _fmt(ws, [PAR[k] for k in ('建账日', '最后日期', '挡后日期', '截止', '月初', '日期上限', '年末')], DATE)
    _fmt(ws, [f'{c}{r}' for c in 'DE' for r in range(2, 6)], DATE)


# ───────────────────────── _表 ─────────────────────────
def _blk_counter(ws, cnt, n, cond_range, crit):
    """计数列（50 行一段）：第 i 条＝前 i 条里满足 COUNTIF(区域, crit) 的个数"""
    for i in range(n):
        r = i + 2
        bs = 2 + (i // BLK) * BLK
        base = f'{cnt}{bs - 1}+' if bs > 2 else ''
        ws[f'{cnt}{r}'] = f'={base}COUNTIF({cond_range(bs, r)},{crit})'


def _cum(ws, cnt, n, flag):
    """累计数列：第 i 行＝前 i 条 flag 之和"""
    for i in range(n):
        r = i + 2
        ws[f'{cnt}{r}'] = f'={flag(r)}' + (f'+{cnt}{r - 1}' if i else '')


LO_PRIV = dict(单位号='CE', 银行='CF', 息净='CG')          # ⑦ 私有列


def build_list(ws):
    for lst in (LIST_ACC, LIST_CAT, LIST_RE, LIST_PER, LIST_TK, LIST_LEND, LIST_LO, LIST_OP, LO_PRIV):
        for k, c in lst.items():
            _lab(ws, f'{c}1', k)
    tk = lambda t: f'IF({t}="",0,IFERROR(MATCH({esc(t)},票种_名称,0),0))'
    for i in range(LIST_N):
        r, n = i + 2, i + 1
        if n <= N_ACC:
            g = lambda k: _ix(SH_BASE, ACC_COLS[k], n, '账户')
            L = lambda k: f'{LIST_ACC[k]}{r}'
            ws[L('名称')] = f'=IF({n}>P_账户位,"",TRIM({g("名称")}&""))'
            t = txt(g('类型'))
            ws[L('类型号')] = f'=IF({L("名称")}="",0,IF({t}="",2,{idx_of(ACC_TYPES, t, 0)}))'
            ws[L('期初')] = f'=IF(AND({L("类型号")}>=1,{L("类型号")}<=4),ROUND(N({g("期初余额")}),2),0)'
            ws[L('人员')] = f'=IF({L("名称")}="","",TRIM({g("人员")}&""))'
            ws[L('现金')] = f'=IF(AND({L("类型号")}>=1,{L("类型号")}<=3),1,0)'
            ws[L('科目号')] = (f'=IF({L("类型号")}=0,0,CHOOSE({L("类型号")},{kh("1001")},{kh("1002")},{kh("1012")},'
                             f'{kh("1121")},{kh("224101")},{kh("122103")}))')
            ws[L('出借人号')] = f'=IF({L("类型号")}<>5,0,{unit_match(norm(L("人员")), "出借人号")})'
        if n <= N_CAT:
            g = lambda k: _ix(SH_BASE, CAT_COLS[k], n, '类别')
            L = lambda k: f'{LIST_CAT[k]}{r}'
            a = L('名称')
            ws[a] = f'=IF({n}>P_类别位,"",TRIM({g("名称")}&""))'
            ws[L('归类')] = f'=IF({a}="","",TRIM({g("归类")}&""))'
            for side, kc in (('收', '收钱科目'), ('付', '付钱科目')):
                cd = code_of(g(kc))
                sp = '+'.join(f'({cd}="{k}")*{v}' for k, v in SPECIAL.items())
                ws[L(side + '特')] = f'=IF({a}="",0,{sp})'
                ws[L(side + '号')] = f'=IF(OR({a}="",{L(side + "特")}>0),0,{kmno(cd)})'
            ws[L('现流收')] = f'=IF({a}="",0,N({g("现流收")}))'
            ws[L('现流付')] = f'=IF({a}="",0,N({g("现流付")}))'
            ws[L('票种号')] = f'=IF({a}="",0,{tk(txt(g("默认票种")))})'
            ws[L('不抵')] = f'=IF(AND({a}<>"",TRIM({g("不得抵扣")}&"")="是"),1,0)'
            v = g('分摊月数')
            ws[L('分摊月数')] = f'=IF({a}="",0,IF(ISNUMBER({v}),IF({v}>0,ROUND({v},0),0),0))'
            ws[L('摊进')] = f'=IF(TRIM({g("摊进")}&"")="管理",2,1)'
            ws[L('投产后号')] = f'=IF({a}="",0,{kmno(code_of(g("投产后科目")))})'
            un = norm(g('默认往来'))
            ws[L('默认往来')] = f'=IF({a}="","",{unit_match(un, "简称", un)})'
        if n <= N_RE:
            g = lambda k: _ix(SH_BASE, RE_COLS[k], n, '报销')
            L = lambda k: f'{LIST_RE[k]}{r}'
            a = L('名称')
            ws[a] = f'=IF({n}>P_报销位,"",TRIM({g("名称")}&""))'
            ws[L('科目号')] = f'=IF({a}="",0,{kmno(code_of(g("科目")))})'
            ws[L('票种号')] = f'=IF({a}="",0,{tk(txt(g("默认票种")))})'
            ws[L('不抵')] = f'=IF(AND({a}<>"",TRIM({g("不得抵扣")}&"")="是"),1,0)'
            ws[L('现流')] = f'=IF({a}="",0,N({g("现流")}))'
        if n <= N_PER:
            g = lambda k: _ix(SH_BASE, PER_COLS[k], n, '人员')
            L = lambda k: f'{LIST_PER[k]}{r}'
            ws[L('姓名')] = f'=IF({n}>P_人员位,"",TRIM({g("姓名")}&""))'
            ws[L('部门号')] = f'=IF({L("姓名")}="",0,{idx_of(DEPTS, txt(g("部门")), 0)})'
        if n <= N_TK:
            g = lambda k: _ix(SH_BASE, TK_COLS[k], n, '票种')
            L = lambda k: f'{LIST_TK[k]}{r}'
            ws[L('名称')] = f'=IF({n}>P_票种位,"",TRIM({g("名称")}&""))'
            ws[L('税率')] = f'=IF({L("名称")}="",0,{pct(g("可抵税率"), 0)})'
        if n <= N_LEND:
            L = lambda k: f'{LIST_LEND[k]}{r}'
            ws[L('单位号')] = f'=IFERROR(MATCH({n},位_出借人号,0),0)'
            ws[L('名称')] = f'=IF({L("单位号")}=0,"",INDEX(位_简称,{L("单位号")}))'
    # ⑦ 建账前借款、投资、垫付（第 n 条在第 n+1 行）
    lend_u = f'${LIST_LEND["单位号"]}$2:${LIST_LEND["单位号"]}${N_LEND + 1}'
    for i in range(N_LO):
        r, n = i + 2, i + 1
        g = lambda k: _ix(SH_BASE, LO_COLS[k], n, '借期')
        L = lambda k: f'{LIST_LO[k]}{r}'
        Q = lambda k: f'{LO_PRIV[k]}{r}'
        off = f'{n}>P_借期位'
        ws[L('出借人号')] = f'=IF({off},0,{unit_match(norm(g("出借人")), "出借人号")})'
        ws[L('日期')] = f'=IF({off},0,{vdate(g("日期"))})'
        ws[L('日期')].number_format = DATE
        ws[L('金额')] = f'=IF({off},0,ROUND({numv(g("金额"))},2))'
        ws[L('性质')] = f'=IF({off},0,{idx_of(LO_KINDS, txt(g("性质")), 0)})'
        ws[L('已付息')] = f'=IF({off},0,ROUND({numv(g("已付利息"))},2))'
        ws[L('有效')] = (f'=IF(AND({L("出借人号")}>0,{L("日期")}>0,{L("日期")}<=P_建账日,{L("金额")}<>0,{L("性质")}>0),1,0)')
        ws[Q('单位号')] = f'=IF({L("出借人号")}>0,INDEX({lend_u},{L("出借人号")}),0)'
        ws[Q('银行')] = f'=IF({Q("单位号")}>0,IF(INDEX(位_出借人类别,{Q("单位号")})=1,1,0),0)'
        u = Q('单位号')
        rate = f'IF({L("性质")}=2,INDEX(位_投资利率,{u}),INDEX(位_借款利率,{u}))'
        ws[L('应计息')] = (f'=IF({L("有效")}<>1,0,ROUND({L("金额")}*{rate}/INDEX(位_计息天数,{u})*(P_建账日-{L("日期")}),2))')
        ws[Q('息净')] = f'=IF({L("有效")}=1,ROUND({L("应计息")}-{L("已付息")},2),0)'
    # ⑧ 其他期初余额
    fam = ['1001', '1002', '1012', '1121', '1122', '2202', '1403', '1405', '1601', '1602', '2001', '3001', '2231', '2232', '4001']
    exact = ['224101', '224102', '310415']
    for i in range(N_OP):
        r, n = i + 2, i + 1
        g = lambda k: _ix(SH_BASE, OP_COLS[k], n, '期初')
        L = lambda k: f'{LIST_OP[k]}{r}'
        off = f'{n}>P_期初位'
        no = L('科目号')
        ws[no] = f'=IF({off},0,{kmno(code_of(g("科目")))})'
        ws[L('明细')] = f'=IF({off},"",TRIM({g("明细")}&""))'
        ws[L('借')] = f'=IF({off},0,ROUND({numv(g("借方"))},2))'
        ws[L('贷')] = f'=IF({off},0,ROUND({numv(g("贷方"))},2))'
        v = g('剩余月数')
        ws[L('剩余月数')] = f'=IF({off},0,IF(ISNUMBER({v}),IF({v}>0,ROUND({v},0),0),0))'
        ws[L('摊进')] = f'=IF(TRIM({g("摊进")}&"")="管理",2,1)'
        f1 = f'INDEX(科目_一级号,MAX(1,{no}))'
        bad = 'OR(' + ','.join([f'{f1}={kh(c)}' for c in fam] + [f'{no}={kh(c)}' for c in exact]) + ')'
        # 不认：② ⑦ 往来 物料 卡片 的科目（各有自己的来源）、4001 家族、310415（倒挤）、损益类（表结法：建账前的利润都在倒挤里）
        ws[L('有效')] = (f'=IF({no}=0,0,IF(AND(OR({L("借")}<>0,{L("贷")}<>0),INDEX(科目_有效,{no})=1,'
                       f'INDEX(科目_类别号,{no})<>5,NOT({bad})),1,0))')
    # 压紧的清单（下拉用）：第 1 行＝条数
    for key, (out, src, cap) in COMPACT.items():
        cnt = COMPACT_CNT[key]
        for i in range(cap):
            r = i + 2
            prev = f'+{cnt}{r - 1}' if i else ''
            ws[f'{cnt}{r}'] = f'=IF({src}{r}<>"",1,0){prev}'
            ws[f'{out}{r}'] = (f'=IF({i + 1}>${out}$1,"",INDEX(${src}$2:${src}${cap + 1},'
                               f'MATCH({i + 1},${cnt}$2:${cnt}${cap + 1},0)))')
        ws[f'{out}1'] = f'={cnt}{cap + 1}'
    # 往来单位下拉：全部 / 客户 / 供应商 / 出借人（来源 _位）
    for key, flag, crit in (('往来单位', '有效', 1), ('客户', '是客户', 1), ('供应商', '是供应商', 1), ('出借人', '出借人号', '">0"')):
        out, cnt, fc = UNIT_DV[key], COMPACT_CNT[key], WEI[flag]
        _blk_counter(ws, cnt, N_UNIT, lambda bs, r, fc=fc: f'{H_WEI}!${fc}${bs}:${fc}{r}', crit)
        for i in range(N_UNIT):
            r = i + 2
            ws[f'{out}{r}'] = (f'=IF({i + 1}>${out}$1,"",INDEX(位_简称,MATCH({i + 1},${cnt}$2:${cnt}${N_UNIT + 1},0)))')
        ws[f'{out}1'] = f'={cnt}{N_UNIT + 1}'
    # 物料、产品下拉（来源 _物）
    cq, cr, cs = COMPACT_CNT['原料'], COMPACT_CNT['产品'], COMPACT_CNT['物料产品']
    _cum(ws, cq, N_RM, lambda r: f'{H_WU}!${WRM["有效"]}${r}')
    _cum(ws, cr, N_PRD, lambda r: f'{H_WU}!${WPRD["有效"]}${r}')
    nq, npr = f'${cq}${N_RM + 1}', f'${cr}${N_PRD + 1}'
    for i in range(N_RM):
        r = i + 2
        o = ITEM_DV['原料']
        ws[f'{o}{r}'] = f'=IF({i + 1}>${o}$1,"",INDEX(料档_名称,MATCH({i + 1},${cq}$2:{nq},0)))'
    ws[f'{ITEM_DV["原料"]}1'] = f'={nq}'
    for i in range(N_PRD):
        r = i + 2
        o = ITEM_DV['产品']
        ws[f'{o}{r}'] = f'=IF({i + 1}>${o}$1,"",INDEX(品档_名称,MATCH({i + 1},${cr}$2:{npr},0)))'
    ws[f'{ITEM_DV["产品"]}1'] = f'={npr}'
    o = ITEM_DV['物料产品']
    for i in range(N_RM + N_PRD):
        r = i + 2
        ws[f'{cs}{r}'] = f'={cq}{r}' if i < N_RM else f'={nq}+{cr}{r - N_RM}'
        k = i + 1
        ws[f'{o}{r}'] = (f'=IF({k}>${o}$1,"",IF({k}<={nq},INDEX(料档_名称,MATCH({k},${cq}$2:{nq},0)),'
                         f'INDEX(品档_名称,MATCH({k}-{nq},${cr}$2:{npr},0))))')
    ws[f'{o}1'] = f'={nq}+{npr}'
    # 科目下拉「编码 名称」
    ct = COMPACT_CNT['科目']
    _cum(ws, ct, N_COA, lambda r: f'{H_KM}!${KM_COLS["有效"]}${r}')
    for i in range(N_COA):
        r = i + 2
        m = f'MATCH({i + 1},${ct}$2:${ct}${N_COA + 1},0)'
        ws[f'{KM_DV}{r}'] = f'=IF({i + 1}>${KM_DV}$1,"",INDEX(科目_编码,{m})&" "&INDEX(科目_名称,{m}))'
    ws[f'{KM_DV}1'] = f'=${ct}${N_COA + 1}'
    for c in COMPACT_CNT.values():
        _lab(ws, f'{c}1', '计数')


# ───────────────────────── _科：会计科目镜像 ─────────────────────────
OPEN_SRC = ['1001', '1002', '1012', '1121', '1122', '2202', '1403', '1405', '400101', '1601', '1602',
            '2001', '224102', '3001', '224101', '2231', '2232', '310415']
KM_PRIV = dict(费用累计='T', 制费累计='U', 来源号='W', 来源额='X', 来源='Y')


def build_km(ws):
    K = KM_COLS
    for k, c in K.items():
        _lab(ws, f'{c}1', k)
    for k, c in {**{f'固定{k}': c for k, c in KM_FIXED.items()}, **{f'费用{k}': c for k, c in KM_FEE.items()},
                 **{f'制费{k}': c for k, c in KM_OH.items()}, **KM_PRIV}.items():
        _lab(ws, f'{c}1', k)
    src_n = f'${KM_PRIV["来源号"]}$2:${KM_PRIV["来源号"]}${len(OPEN_SRC) + 1}'
    src_v = f'${KM_PRIV["来源额"]}$2:${KM_PRIV["来源额"]}${len(OPEN_SRC) + 1}'
    fee_par = 'OR(' + ','.join(f'{{e}}={kh(c)}' for c in FEE_PARENTS) + ')'
    for i in range(N_COA):
        r, n = i + 2, i + 1
        x = lambda k: _ix(SH_COA, COA_COLS[k], n)
        c = lambda k: f'{K[k]}{r}'
        a = c('编码')
        ws[a] = f'={txt(x("编码"))}'
        ws[c('名称')] = f'={txt(x("名称"))}'
        t = txt(x('类别'))
        ws[c('类别号')] = f'=IF({a}="",0,{idx_of(COA_CLASSES, t, f"IFERROR(FIND(LEFT({a},1),{chr(34)}12345{chr(34)}),0)")})'
        t = txt(x('方向'))
        ws[c('方向')] = f'=IF(OR({t}="借",{t}="贷"),{t},IF(OR({c("类别号")}=2,{c("类别号")}=3),"贷","借"))'
        ws[c('一级号')] = f'=IF({a}="",0,IFERROR(MATCH({esc(f"LEFT({a},4)")},科目_编码,0),0))'
        ws[c('级次')] = f'=IF({a}="",0,IF(LEN({a})=4,1,IF(LEN({a})=6,2,0)))'
        ws[c('全称')] = (f'=IF(AND({c("级次")}=2,{c("一级号")}>0),INDEX(科目_名称,{c("一级号")})&"-"&{c("名称")},{c("名称")})')
        ws[c('成本项目')] = f'={txt(x("成本项目"))}'
        ws[c('有效')] = (f'=IF({a}="",0,IFERROR(IF(AND({c("级次")}>0,{c("一级号")}>0,{c("类别号")}>0,ISNUMBER(VALUE({a})),'
                       f'MATCH({esc(a)},科目_编码,0)={n}),1,0),0))')
        ws[c('期初')] = (f'=IF({c("有效")}<>1,0,ROUND(SUMIF({src_n},{n},{src_v})+SUMIFS(期初_借,期初_科目号,{n},期初_有效,1)'
                       f'-SUMIFS(期初_贷,期初_科目号,{n},期初_有效,1),2))')
        ok2 = f'AND({c("有效")}=1,{c("级次")}=2)'
        fee = f'IF(AND({ok2},{fee_par.format(e=c("一级号"))}),1,0)'
        oh = f'IF(AND({ok2},{c("一级号")}={kh("4101")}),1,0)'
        pf, po = KM_PRIV['费用累计'], KM_PRIV['制费累计']
        ws[f'{pf}{r}'] = f'={fee}' + (f'+{pf}{r - 1}' if i else '')
        ws[f'{po}{r}'] = f'={oh}' + (f'+{po}{r - 1}' if i else '')
    for i, code in enumerate(FIXED_CODES):
        r = i + 2
        ws[f'{KM_FIXED["代码"]}{r}'] = code
        ws[f'{KM_FIXED["号"]}{r}'] = f'=IFERROR(MATCH("{code}",科目_编码,0),0)'
    pf, po = KM_PRIV['费用累计'], KM_PRIV['制费累计']
    for j in range(1, N_FEE + 1):
        r = j + 1
        ws[f'{KM_FEE["j"]}{r}'] = j
        ws[f'{KM_FEE["号"]}{r}'] = f'=IFERROR(MATCH({j},${pf}$2:${pf}${N_COA + 1},0),0)'
    for j in range(1, N_OH + 1):
        r = j + 1
        ws[f'{KM_OH["i"]}{r}'] = j
        ws[f'{KM_OH["号"]}{r}'] = f'=IFERROR(MATCH({j},${po}$2:${po}${N_COA + 1},0),0)'
    # 期初来源小表（借正贷负）：每个科目只有一个来源；310415＝倒挤
    card = f'卡_有效,1,卡_取得日,"<"&P_建账日'
    lo = '借期_有效,1,借期_性质'
    bank = f'{H_TAB}!${LO_PRIV["银行"]}$2:${LO_PRIV["银行"]}${N_LO + 1}'
    xi = f'{H_TAB}!${LO_PRIV["息净"]}$2:${LO_PRIV["息净"]}${N_LO + 1}'
    amt = {
        '1001': 'SUMIF(账户_类型号,1,账户_期初)', '1002': 'SUMIF(账户_类型号,2,账户_期初)',
        '1012': 'SUMIF(账户_类型号,3,账户_期初)', '1121': 'SUMIF(账户_类型号,4,账户_期初)',
        '1122': 'SUM(位_期初应收)', '2202': '-SUM(位_期初应付)',
        '1403': 'SUMIF(料档_有效,1,料档_期初额)', '1405': 'SUMIF(品档_有效,1,品档_期初额)', '400101': 'SUMIF(品档_有效,1,品档_期初在产)',
        '1601': f'SUMIFS(卡_原值,{card})', '1602': f'-SUMIFS(卡_建账前已提,{card})',
        '2001': f'-SUMIFS(借期_金额,{lo},1,{bank},1)', '224102': f'-SUMIFS(借期_金额,{lo},1,{bank},0)',
        '3001': f'-SUMIFS(借期_金额,{lo},2)', '224101': f'-SUMIFS(借期_金额,{lo},3)',
        '2231': f'-SUMIFS({xi},{lo},1)-SUMIFS({xi},{lo},3)', '2232': f'-SUMIFS({xi},{lo},2)',
    }
    pn, pv, pl = KM_PRIV['来源号'], KM_PRIV['来源额'], KM_PRIV['来源']
    last = len(OPEN_SRC)
    for i, code in enumerate(OPEN_SRC):
        r = i + 2
        ws[f'{pn}{r}'] = f'={kh(code)}'
        ws[f'{pl}{r}'] = code
        if code == '310415':
            ws[f'{pv}{r}'] = (f'=-ROUND(SUM(${pv}$2:${pv}${r - 1})+SUMIFS(期初_借,期初_有效,1)-SUMIFS(期初_贷,期初_有效,1),2)')
        else:
            ws[f'{pv}{r}'] = f'=ROUND({amt[code]},2)'
    assert OPEN_SRC[-1] == '310415' and last == len(OPEN_SRC)


# ───────────────────────── _物：物料、产品镜像＋槽表 ─────────────────────────
WU_RMCAP = 'AL2'                                          # 私有：原料块现在有几行（插删行跟着变）


def build_wu(ws):
    for cols in (WU_RM, WU_PRD):
        _hdr(ws, cols)
    for k, c in WU_SLOT.items():
        _lab(ws, f'{c}1', '槽' + k)
    _lab(ws, 'AL1', '原料块行数')
    ws[WU_RMCAP] = f'=ROW({SH_ITEM}!$A${ITEM_PRD_HDR})-ROW({SH_ITEM}!$A${ITEM_RM_HDR})-4'
    cap = f'${WU_RMCAP[:2]}${WU_RMCAP[2:]}'
    W = WRM
    for i in range(N_RM):
        r, k = i + 2, i + 1
        ws[f'A{r}'] = k
        x = lambda f: _ix(SH_ITEM, RM_COLS[f], f'$A{r}', block='原料')
        c = lambda f: f'{W[f]}{r}'
        ok = f'{c("有效")}<>1'
        ws[c('名称')] = f'=IF($A{r}>{cap},"",{norm(x("名称"))})'
        ws[c('别名规')] = f'=IF({c("名称")}="","",{alias(x("别名"))})'
        ws[c('有效')] = f'=IF({c("名称")}="",0,IFERROR(IF(MATCH({esc(c("名称"))},料档_名称,0)=$A{r},1,0),0))'
        ws[c('类别号')] = f'=IF({ok},0,{idx_of(RM_KINDS, txt(x("类别")), 1)})'
        ws[c('单位')] = f'=IF({ok},"",{txt(x("单位"))})'
        t = txt(x('进货即耗用'))
        ws[c('耗用')] = f'=IF({ok},0,IF({t}="是",1,IF({t}="否",0,IF(OR({c("类别号")}=2,{c("类别号")}=3),1,0))))'
        cd = code_of(x('计入明细'))
        ws[c('计入号')] = (f'=IF({ok},0,IF({cd}<>"",{kmno(cd)},IF({c("类别号")}=2,{kh("410110")},'
                         f'IF({c("类别号")}=3,{kh("410113")},0))))')
        v = x('生产日报列')
        ws[c('槽')] = f'=IF({ok},0,IF(ISNUMBER({v}),IF(AND({v}>=1,{v}<={N_SLOT},INT({v})={v}),{v},0),0))'
        v = x('含量')
        ws[c('含量')] = f'=IF({ok},0,IF(ISNUMBER({v}),IF({v}>1,{v}/100,{v}),0))'
        ws[c('参考单价')] = f'=IF({ok},0,{numv(x("参考单价"))})'
        ws[c('期初量')] = f'=IF({ok},0,{numv(x("期初数量"))})'
        v = x('期初金额')
        ws[c('期初额')] = f'=IF({ok},0,IF(ISNUMBER({v}),ROUND({v},2),ROUND({c("期初量")}*{numv(x("期初单价"))},2)))'
        ws[c('期初单价')] = f'=IF({ok},0,IF({c("期初量")}<>0,ROUND({c("期初额")}/{c("期初量")},6),{c("参考单价")}))'
        ws[c('定额')] = f'=IF({ok},0,{numv(x("定额单价"))})'
        v = x('每袋吨数')
        ws[c('每袋吨')] = f'=IF(ISNUMBER({v}),IF({v}>0,{v},1),1)'
    W = WPRD
    p0 = WU_PRD0
    rm_no = lambda nm: (f'IFERROR(MATCH({esc(nm)},料档_名称,0),IFERROR(MATCH("*、"&{esc(nm)}&"、*",料档_别名规,0),0))')
    for i in range(N_PRD):
        r, p = i + 2, i + 1
        ws[f'{p0}{r}'] = p
        x = lambda f: _ix(SH_ITEM, PRD_COLS[f], f'${p0}{r}', block='产品')
        c = lambda f: f'{W[f]}{r}'
        ok = f'{c("有效")}<>1'
        ws[c('名称')] = f'={norm(x("名称"))}'
        ws[c('别名规')] = f'=IF({c("名称")}="","",{alias(x("别名"))})'
        ws[c('有效')] = f'=IF({c("名称")}="",0,IFERROR(IF(MATCH({esc(c("名称"))},品档_名称,0)={p},1,0),0))'
        v = x('每袋吨数')
        ws[c('每袋吨')] = f'=IF(ISNUMBER({v}),IF({v}>0,{v},1),1)'
        v = x('分摊系数')
        ws[c('系数')] = f'=IF({ok},0,IF(ISNUMBER({v}),IF({v}>=0,{v},1),1))'
        nm = norm(x('废料归入'))
        ws[c('回收料号')] = f'=IF({ok},0,IF({nm}="",P_默认回收料号,{rm_no(nm)}))'
        ws[c('折纯')] = f'=IF({ok},0,{numv(x("折纯系数"))})'
        v = x('成品含量')
        ws[c('含量')] = f'=IF({ok},0,IF(ISNUMBER({v}),IF({v}>1,{v}/100,{v}),0))'
        ws[c('期初量')] = f'=IF({ok},0,{numv(x("期初数量"))})'
        v = x('期初金额')
        ws[c('期初额')] = f'=IF({ok},0,IF(ISNUMBER({v}),ROUND({v},2),ROUND({c("期初量")}*{numv(x("期初单价"))},2)))'
        ws[c('期初单价')] = f'=IF({ok},0,IF({c("期初量")}<>0,ROUND({c("期初额")}/{c("期初量")},6),{numv(x("期初单价"))}))'
        ws[c('期初在产')] = f'=IF({ok},0,ROUND({numv(x("期初在产"))},2))'
    for j in range(1, N_SLOT + 1):
        r = j + 1
        ws[f'{WU_SLOT["j"]}{r}'] = j
        no = f'{WU_SLOT["料号"]}{r}'
        ws[no] = f'=IF(COUNTIF(料档_槽,{j})=1,MATCH({j},料档_槽,0),0)'
        ws[f'{WU_SLOT["名称"]}{r}'] = f'=IF({no}>0,INDEX(料档_名称,{no}),"")'


# ───────────────────────── _固：固定支出（沿用 A069，月份＝账套年度 1～12） ─────────────────────────
GU_PRIV = dict(开始年月='U', 结束年月='V')


def build_gu(ws):
    _hdr(ws, GU_COLS)
    for k, c in GU_PRIV.items():
        _lab(ws, f'{c}1', k)
    G, C = GU, FIX_COLS
    for i in range(N_FIX):
        r, n = i + 2, i + 1
        ws[f'A{r}'] = n
        x = lambda k: _ix(SH_FIX, C[k], f'$A{r}')
        c = lambda k: f'{G[k]}{r}'
        q = lambda k: f'{GU_PRIV[k]}{r}'
        ws[c('编号')] = f'={txt(x("编号"))}'
        ws[c('编号键')] = f'=IF({c("编号")}="","","#"&{norm(c("编号"))})'
        ws[c('项目')] = f'={txt(x("项目"))}'
        ws[c('类别')] = f'={txt(x("类别"))}'
        un = norm(x('收款单位'))
        ws[c('收款单位')] = f'={unit_match(un, "简称", un)}'
        ws[c('每月金额')] = f'={numv(x("每月金额"))}'
        ws[c('每月几号')] = f'=N({x("每月几号")})'
        ws[c('付款账户')] = f'={txt(x("付款账户"))}'
        ws[q('开始年月')] = f'={ym_any(x("开始月份"))}'
        ws[q('结束年月')] = f'={ym_any(x("结束月份"))}'
        s, e = q('开始年月'), q('结束年月')
        # 开始月：空＝1；看不懂＝0；早于年度＝1；晚于年度＝13。结束月：空＝12；看不懂＝-1；早于年度＝0；晚于年度＝12
        ws[c('开始月')] = (f'=IF({s}=-1,1,IF(NOT({ym_ok(s)}),0,IF(INT({s}/100)<P_年度,1,IF(INT({s}/100)>P_年度,13,'
                        f'MOD({s},100)))))')
        ws[c('结束月')] = (f'=IF({e}=-1,12,IF(NOT({ym_ok(e)}),-1,IF(INT({e}/100)<P_年度,0,IF(INT({e}/100)>P_年度,12,'
                        f'MOD({e},100)))))')
        ws[c('匹配键')] = f'={c("类别")}&"|"&{c("收款单位")}'
        ws[c('有效')] = f'=IF(AND({c("类别")}<>"",{c("每月金额")}>0),1,0)'
        ws[c('本月生效')] = (f'=IF(AND({c("有效")}=1,{c("开始月")}>=1,{c("开始月")}<=P_截止月,{c("结束月")}>=P_截止月),1,0)')
        prev = f'COUNTIFS($L$1:L{r - 1},{esc(c("匹配键"))},$M$1:M{r - 1},1)'
        ws[c('先占')] = f'=IF({c("有效")}<>1,0,IF({prev}=0,1,0))'
        prev_m = f'COUNTIFS($L$1:L{r - 1},{esc(c("匹配键"))},$O$1:O{r - 1},1)'
        ws[c('本月先占')] = f'=IF({c("本月生效")}<>1,0,IF({prev_m}=0,1,0))'
        when = '资_月,P_截止月,资_日期,"<="&P_截止'
        k1 = f'IF({c("编号键")}="",0,-SUMIFS(资_净额,资_编号键,{esc(c("编号键"))},{when}))'
        cat = f'IFERROR(MATCH({esc(c("类别"))},类别_名称,0),-1)'
        base = f'资_净额,资_关联固定行,0,资_类别号,{cat},{when}'
        k2 = (f'IF({c("本月先占")}=1,-IF({c("收款单位")}="",SUMIFS({base}),SUMIFS({base},资_单位,{esc(c("收款单位"))})),0)')
        ws[c('本月已付')] = f'=IF({c("有效")}<>1,0,ROUND({k1}+{k2},2))'
        has = f'OR({c("编号")}<>"",{c("项目")}<>"",{c("类别")}<>"",{raw(x("每月金额"))})'
        cat_t = c('类别')
        ws[c('校验')] = '=' + chain([
            (f'NOT({has})', '""'),
            (f'{cat_t}=""', '"✗ 没选收支类别"'),
            (f'COUNTIF(类别_名称,{esc(cat_t)})=0', '"✗ 收支类别不在【基础资料】③收支类别里"'),
            (f'OR({badnum(x("每月金额"))},{c("每月金额")}<=0)', '"✗ 没填每月金额"'),
            (f'AND({c("编号")}<>"",COUNTIF(固_编号键,{esc(c("编号键"))})>1)', '"✗ 编号重复了"'),
            (f'{c("开始月")}=0', '"✗ 开始月份看不懂（填 8 或 2026-8 都行）"'),
            (f'{c("结束月")}=-1', '"✗ 结束月份看不懂（填 12 或 2026-12 都行）"'),
            (f'AND({c("结束月")}>=1,{c("开始月")}<=12,{c("开始月")}>{c("结束月")})', '"✗ 开始月份比结束月份晚"'),
            (f'{c("结束月")}=0', '"⚠ 结束月份在本年度以前（今年不算）"'),
            (f'{c("开始月")}=13', '"⚠ 开始月份在本年度以后（今年不算）"'),
            (f'{c("先占")}=0', '"⚠ 类别、收款单位跟上面一条一样：没填编号的付款只算给上面那条（付款时编号填它就能分开）"'),
            (f'{c("编号")}=""', '"⚠ 没填编号（付款时资金台帐编号填它才对得准，不填只能按类别、收款单位对）"'),
            (f'OR({c("每月几号")}<0,{c("每月几号")}>31)', '"⚠ 每月几号要填 1～31"'),
            (f'AND({c("付款账户")}<>"",COUNTIF(账户_名称,{esc(c("付款账户"))})=0)', '"⚠ 付款账户不在【基础资料】②账户里"'),
        ])
        ws[c('录入行')] = f'=ROW({SH_FIX}!$A${HDR})+$A{r}'
        ws[c('末行')] = f'=IF({has},{c("录入行")},0)'


# ───────────────────────── _卡：固定资产卡片 × 12 个月折旧 ─────────────────────────
KA_DISP = 'Q'                                             # 私有：处置日（真日期；空＝0）


def build_ka(ws):
    _hdr(ws, KA_COLS)
    _lab(ws, f'{KA_DISP}1', '处置日')
    for m in range(1, 13):
        _lab(ws, f'{CL(CI(KA_MON) + m - 1)}1', f'{m}月')
    tot = CL(CI(KA_MON) + 12)
    _lab(ws, f'{tot}1', '本年合计')
    K, C = KA, FA_COLS
    for i in range(N_FA):
        r, n = i + 2, i + 1
        ws[f'A{r}'] = n
        x = lambda k: _ix(SH_FA, C[k], f'$A{r}')
        c = lambda k: f'{K[k]}{r}'
        dd = f'{KA_DISP}{r}'
        ok = f'{c("有效")}<>1'
        ws[c('名称')] = f'={txt(x("名称"))}'
        t = txt(x('使用部门'))
        ws[c('部门号')] = f'=IF({t}="管理",2,IF({t}="销售",3,1))'
        ws[c('取得日')] = f'={vdate(x("取得日期"))}'
        ws[c('原值')] = f'=ROUND({numv(x("原值"))},2)'
        ws[c('残值率')] = f'={pct(x("残值率"), 0.05)}'
        v = x('折旧月数')
        ws[c('月数')] = f'=IF(ISNUMBER({v}),IF({v}>=1,INT({v}),0),0)'
        ws[dd] = f'={vdate(x("处置日期"))}'
        ws[c('有效')] = (f'=IF(AND({c("取得日")}>0,{c("取得日")}<=P_年末,{c("原值")}>0,{c("月数")}>0,{c("残值率")}>=0,'
                       f'{c("残值率")}<1,OR({dd}=0,{dd}>=P_建账日)),1,0)')
        ws[c('月折旧')] = f'=IF({ok},0,ROUND({c("原值")}*(1-{c("残值率")})/{c("月数")},2))'
        ws[c('应提总额')] = f'=IF({ok},0,ROUND({c("原值")}*(1-{c("残值率")}),2))'
        v = x('建账前已提')
        mons = f'MAX(0,(YEAR(P_建账日)*12+MONTH(P_建账日))-(YEAR({c("取得日")})*12+MONTH({c("取得日")}))-1)'
        ws[c('建账前已提')] = (f'=IF(OR({ok},{c("取得日")}>=P_建账日),0,IF(ISNUMBER({v}),ROUND({v},2),'
                           f'MIN({c("应提总额")},ROUND({c("月折旧")}*{mons},2))))')
        ws[c('开始月序')] = f'=IF({c("取得日")}=0,0,(YEAR({c("取得日")})-P_年度)*12+MONTH({c("取得日")})+1)'
        ws[c('处置月')] = f'=IF({dd}=0,99,IF(YEAR({dd})<P_年度,0,IF(YEAR({dd})>P_年度,99,MONTH({dd}))))'
        has = f'OR({raw(x("编号"))},{c("名称")}<>"",{raw(x("取得日期"))},{raw(x("原值"))},{raw(x("折旧月数"))})'
        v0 = x('建账前已提')
        ws[c('校验')] = '=' + chain([
            (f'NOT({has})', '""'),
            (f'{c("取得日")}=0', 'IF(' + raw(x('取得日期')) + ',IF(ISTEXT(' + x('取得日期') + '),' + TXT_DATE
             + ',"✗ 取得日期看不懂（输 2026-5-20 或 2026/5/20）"),"✗ 没填取得日期")'),
            (f'OR({badnum(x("原值"))},{c("原值")}<=0)', '"✗ 原值要填正数（不含税）"'),
            (f'{c("月数")}<=0', '"✗ 没填折旧月数（机器设备一般 120，运输设备 48，电子设备 36）"'),
            (f'OR({c("残值率")}<0,{c("残值率")}>=1)', '"✗ 残值率填 5% 这样的数（空＝5%）"'),
            (f'{c("取得日")}>P_年末', '"✗ 取得日期在账套年度以后（请开新一本）"'),
            (f'AND({raw(x("处置日期"))},{dd}=0)', '"✗ 处置日期看不懂"'),
            (f'AND({dd}>0,{dd}<{c("取得日")})', '"✗ 处置日期比取得日期早"'),
            (f'AND({dd}>0,{dd}<P_建账日)', '"✗ 建账以前就处置了：这张卡片不算（删掉）"'),
            (f'{badnum(v0)}', '"✗ 建账前已提要填数字（空着＝按月数自动算）"'),
            (f'AND({raw(v0)},{c("取得日")}>=P_建账日)', '"⚠ 建账以后才取得的，建账前已提不算"'),
            (f'AND({c("有效")}=1,{c("建账前已提")}>{c("应提总额")})', '"⚠ 建账前已提比应提总额（原值×(1－残值率)）还多"'),
            (f'NOT({in_list(FA_DEPTS + ["车间管理", ""], txt(x("使用部门")))})', '"⚠ 使用部门要选生产、管理、销售（按生产算）"'),
        ])
        ws[c('录入行')] = f'=ROW({SH_FA}!$A${HDR})+$A{r}'
        ws[c('末行')] = f'=IF({has},{c("录入行")},0)'
        # 闭式：第 m 月＝在提的月份里 MIN(月折旧, 应提总额－建账前已提－月折旧×前面已提的月数)；截止月以后＝0
        s = f'MAX(${K["开始月序"]}{r},P_建账月)'
        for m in range(1, 13):
            col = CL(CI(KA_MON) + m - 1)
            ws[f'{col}{r}'] = (f'=IF(AND(${K["有效"]}{r}=1,{m}>={s},{m}<=${K["处置月"]}{r},{m}<=P_截止月),'
                               f'MAX(0,MIN(${K["月折旧"]}{r},ROUND(${K["应提总额"]}{r}-${K["建账前已提"]}{r}'
                               f'-${K["月折旧"]}{r}*({m}-{s}),2))),0)')
        ws[f'{tot}{r}'] = f'=SUM({KA_MON}{r}:{CL(CI(KA_MON) + 11)}{r})'
        _fmt(ws, [c('取得日'), dd], DATE)


# ───────────────────────── _位：往来单位逐家＋往来余额 ─────────────────────────
def build_wei(ws):
    _hdr(ws, WEI_COLS)
    W, C = WEI, UNIT_COLS
    upto = lambda nm: f'{nm},"<="&P_截止'
    for i in range(N_UNIT):
        r, n = i + 2, i + 1
        ws[f'A{r}'] = n
        x = lambda k: _ix(SH_UNIT, C[k], f'$A{r}')
        c = lambda k: f'{W[k]}{r}'
        A = f'$A{r}'
        nm, ty = c('简称'), c('类型')
        ws[nm] = f'={norm(x("简称"))}'
        ws[ty] = f'={txt(x("类型"))}'
        ws[c('全称规')] = f'=IF({txt(x("全称"))}="","",{norm(x("全称"))})'
        ws[c('别名规')] = f'={alias(x("别名"))}'
        ws[c('联系人')] = f'={txt(x("联系人"))}'
        ws[c('电话')] = f'={txt(x("电话"))}'
        ws[c('有效')] = f'=IF({nm}<>"",1,0)'
        nl = f'AND({c("有效")}=1,{ty}<>"出借人")'
        for k in ('期初应收', '期初应付', '期初未开票', '期初欠票'):
            ws[c(k)] = f'=IF({nl},ROUND(N({x(k)}),2),0)'
        ws[c('要开票')] = f'=IF({txt(x("开票"))}="不要",0,1)'
        tk = txt(x('默认票种'))
        ws[c('默认票种号')] = f'=IF({tk}="",0,IFERROR(MATCH({esc(tk)},票种_名称,0),0))'
        isl = f'AND({c("有效")}=1,{ty}="出借人")'
        lk = txt(x('出借人类别'))
        ws[c('出借人类别')] = f'=IF({isl},IF({lk}="银行",1,IF({lk}="合伙人",3,2)),0)'
        ws[c('借款利率')] = f'=IF({isl},{pct(x("借款年利率"), 0)},0)'
        ws[c('投资利率')] = f'=IF({isl},{pct(x("投资年利率"), 0)},0)'
        v = x('计息天数')
        ws[c('计息天数')] = f'=IF(ISNUMBER({v}),IF({v}>0,{v},IF({c("出借人类别")}=1,360,365)),IF({c("出借人类别")}=1,360,365))'
        ws[c('到期日')] = f'={vdate(x("到期日"))}'
        ws[c('是客户')] = f'=IF(AND({c("有效")}=1,OR({ty}="客户",{ty}="客户和供应商")),1,0)'
        ws[c('是供应商')] = f'=IF(AND({c("有效")}=1,OR({ty}="供应商",{ty}="物流",{ty}="客户和供应商")),1,0)'
        cnt = f'COUNTIFS(${W["类型"]}$2:{ty},"出借人",${W["有效"]}$2:{c("有效")},1)'
        ws[c('出借人号')] = f'=IF({isl},IF({cnt}<={N_LEND},{cnt},0),0)'
        v1 = f'{c("有效")}<>1'
        # 余额、发生额（截至截止日）
        sale = f'SUMIFS(销_收入含税,销_客户号,{A},{upto("销_日期")})'
        ws[c('发货')] = f'=IF({v1},0,ROUND({sale},2))'
        ws[c('回款')] = f'=IF({v1},0,-ROUND(SUMIFS(资_应收变动,资_单位号,{A},{upto("资_日期")}),2))'
        arr = (f'SUMIFS(进_货款计,进_供应商号,{A},{upto("进_日期")})+SUMIFS(进_运费计,进_物流号,{A},{upto("进_日期")})'
               f'+SUMIFS(销_运费计,销_物流号,{A},{upto("销_日期")})')
        ws[c('到货')] = f'=IF({v1},0,ROUND({arr},2))'
        ws[c('付款')] = f'=IF({v1},0,-ROUND(SUMIFS(资_应付变动,资_单位号,{A},{upto("资_日期")}),2))'
        mj = lambda side, code: (f'SUMIFS(手_{side},手_单位号,{A},手_一级号,{kh(code)},手_记账,1,{upto("手_日期")})')
        ws[c('应收')] = (f'=IF({v1},0,ROUND({c("期初应收")}+{c("发货")}-{c("回款")}+{mj("借", "1122")}-{mj("贷", "1122")},2))')
        ws[c('应付')] = (f'=IF({v1},0,ROUND({c("期初应付")}+{c("到货")}-{c("付款")}+{mj("贷", "2202")}-{mj("借", "2202")},2))')
        nt = 'IFERROR(MATCH("无票",票种_名称,0),-1)'
        ws[c('已开票')] = (f'=IF({v1},0,ROUND(SUMIFS(销_收入含税,销_客户号,{A},销_发票号键,"#?*",{upto("销_日期")})'
                         f'+SUMIFS(票_价税合计,票_单位号,{A},票_方向,1,票_销售重复,0,票_有效,1,{upto("票_日期")}),2))')
        ws[c('应开票')] = (f'=IF({v1},0,ROUND({c("期初未开票")}+{c("发货")}'
                         f'-SUMIFS(销_收入含税,销_客户号,{A},销_票种号,{nt},{upto("销_日期")}),2))')
        ws[c('未开票')] = f'=ROUND({c("应开票")}-{c("已开票")},2)'
        ws[c('已收票')] = (f'=IF({v1},0,ROUND(SUMIFS(票_价税合计,票_单位号,{A},票_方向,2,票_补票行,0,票_有效,1,'
                         f'{upto("票_日期")}),2))')
        need = (f'SUMIFS(进_货款计,进_供应商号,{A},进_货款票种号,"<>"&{nt},{upto("进_日期")})'
                f'+SUMIFS(进_运费计,进_物流号,{A},进_运费票种号,"<>"&{nt},{upto("进_日期")})'
                f'+SUMIFS(销_运费计,销_物流号,{A},销_运费票种号,"<>"&{nt},{upto("销_日期")})')
        ws[c('应收票')] = f'=IF({v1},0,ROUND({c("期初欠票")}+{need},2))'
        ws[c('欠票')] = f'=ROUND({c("应收票")}-{c("已收票")},2)'
        for k, key in (('最后发货', '销_发货键'), ('最后回款', '资_回款键'), ('最后到货', '进_到货键'), ('最后付款', '资_付款键')):
            v = f'IFERROR(LARGE({key},COUNTIF({key},">"&({A}*{M}+P_截止))+1),0)'
            ws[c(k)] = f'=IF({v1},0,IF({v}>{A}*{M},{v}-{A}*{M},0))'
        fs = f'IFERROR(SMALL(销_发货键,COUNTIF(销_发货键,"<="&{A}*{M})+1)-{A}*{M},0)'
        ws[c('首次发货')] = f'=IF({v1},0,IF(AND({fs}>0,{fs}<=P_截止),{fs},0))'
        ws[c('待定价发货')] = f'=IF({v1},0,COUNTIFS(销_客户号,{A},销_待定价,1,销_有效,1,{upto("销_日期")}))'
        ws[c('待定价到货')] = f'=IF({v1},0,COUNTIFS(进_供应商号,{A},进_待定价,1,进_有效,1,{upto("进_日期")}))'
        start = f'MAX({c("最后回款")},IF({c("期初应收")}>0.005,P_建账日,{c("首次发货")}))'
        ws[c('没回款天数')] = f'=IF(AND({c("有效")}=1,{c("应收")}>0.005,{start}>0),P_截止-{start},"")'
        has = (f'OR({nm}<>"",{ty}<>"",{c("全称规")}<>"",{raw(x("期初应收"))},{raw(x("期初应付"))},'
               f'{raw(x("期初未开票"))},{raw(x("期初欠票"))})')
        lend_raw = f'OR({raw(x("出借人类别"))},{raw(x("借款年利率"))},{raw(x("投资年利率"))})'
        ws[c('校验')] = '=' + chain([
            (f'NOT({has})', '""'),
            (f'{nm}=""', '"✗ 没填简称（这一行不算）"'),
            (f'COUNTIF(位_简称,{esc(nm)})>1', '"✗ 简称重复了"'),
            (f'NOT({in_list(UNIT_TYPES, ty)})', '"✗ 类型要选：客户、供应商、物流、客户和供应商、出借人、其他"'),
            (f'AND({c("全称规")}<>"",{c("全称规")}<>{nm},COUNTIF(位_简称,{esc(c("全称规"))})>0)', '"✗ 全称跟别家的简称一样"'),
            (f'AND({c("全称规")}<>"",COUNTIF(位_全称规,{esc(c("全称规"))})>1)', '"✗ 全称跟别家重复了"'),
            (f'NOT({in_list(["要", "不要", ""], txt(x("开票")))})', '"✗ 开票只能填「要」或「不要」（空＝要）"'),
            (f'AND({tk}<>"",{c("默认票种号")}=0)', '"✗ 默认票种不在【基础资料】⑥票种里"'),
            (f'OR({badnum(x("借款年利率"))},{badnum(x("投资年利率"))},{badnum(x("计息天数"))})', '"✗ 利率、计息天数要填数字"'),
            (f'AND({isl},{c("出借人号")}=0)', f'"✗ 出借人超过 {N_LEND} 家：这一家的借款、利息不算"'),
            (f'AND({ty}="出借人",OR(N({x("期初应收")})<>0,N({x("期初应付")})<>0))',
             '"✗ 出借人不填期初应收、应付（建账前借的钱、投的钱、垫的钱填【基础资料】⑦）"'),
            (f'AND({isl},NOT({in_list(LENDER_KINDS, lk)}))', '"⚠ 出借人类别没选（按个人算：借款记其他应付款-个人及合伙人借款）"'),
            (f'AND({ty}<>"出借人",{lend_raw})', '"⚠ 不是出借人：出借人类别、利率不算"'),
            (f'AND({raw(x("到期日"))},{c("到期日")}=0)', '"⚠ 到期日看不懂"'),
            (f'COUNTIF(位_别名规,"*、"&{esc(nm)}&"、*")>0', '"⚠ 别家的别名里也有这个简称"'),
            (f'AND({c("是供应商")}=1,{c("应付")}<-0.005)',
             '"⚠ 预付 "&TEXT(-' + c('应付') + ',"#,##0.00")&"：是先付的款，还是到了货没登进货物流登记？"'),
            (f'AND({c("是客户")}=1,{c("应收")}<-0.005)', '"ℹ 预收 "&TEXT(-' + c('应收') + ',"#,##0.00")'),
        ])
        ws[c('录入行')] = f'=ROW({SH_UNIT}!$A${HDR})+$A{r}'
        ws[c('末行')] = f'=IF({has},{c("录入行")},0)'
        _fmt(ws, [c(k) for k in ('到期日', '最后发货', '最后回款', '最后到货', '最后付款', '首次发货')], DATE)
        _fmt(ws, [c('没回款天数')], '0')


BUILDERS = {H_PAR: build_par, H_TAB: build_list, H_KM: build_km, H_WU: build_wu, H_GU: build_gu, H_KA: build_ka,
            H_WEI: build_wei}


def build(wb, ctx=None):
    for name, fn in BUILDERS.items():
        fn(wb[name])
