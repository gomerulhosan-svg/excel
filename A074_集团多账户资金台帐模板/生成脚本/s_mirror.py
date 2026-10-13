# -*- coding: utf-8 -*-
"""A074 集团多账户资金台帐 · 隐藏计算层（设计.md §2）：
  _参  全书参数（B 列，名称 P_xx）
  _表  基础资料镜像（公司/账户/收费项/大类/收付人/规则）＋两张压紧倒序的规则表（项规则、人规则）＋下拉压紧＋收支分类候选键＋公司→账户两级候选键
  _流  资金台帐逐条（第 n 条在第 n+1 行）＋块表（每 BLK 行一块：各账户净额/块前累计/最晚日期，窗口用的月统计）＋规则推断自检格

防环（设计 §9.1，make/cyclecheck 查）：
  _流「原始」列（B..L，到 月码）只看本行台帐格、_表、PAR_SAFE；正常日 N 只引用本块「前最晚」那一格（直接写格子）；
  块表只引用 _流 的 C、L、M、N、V；余额、问题、配对、提醒 可以引用块表（它们是叶子）。
性能（设计 §8.4）：块表区域直接写 $M$bs:$M$be；规则推断只在收费项/收付人空着时算，各自只扫自己那张压紧表的条数；
  贵的检查（重复粘贴的文字比较、「好几种收费项」的 LOOKUP、本块日期顺序）都用 IF 挡在便宜条件后面（AND/OR 不短路）。
"""
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI
from layout import (
    H_PAR, H_TAB, H_FLOW, SH_BASE, SH_LEDGER, SH_HOME, SH_SPAN, PAR, PAR_KEYS, PAR_ROW, BASE_PAR_VAL, BASE_COLS,
    LIST_CO, LIST_ACC, LIST_ITEM, LIST_GROUP, LIST_PARTY, LIST_RULE,
    LIST_RULE_I, LIST_RULE_P, DV_COLS, DV_CNT, CAND_ROWS, CAND_TWO, CAND_TWO_LAST, _TAB, _TAB_NEXT, HOME_SEL, SPAN_SEL,
    FLOW, FLOW_COLS, FLOW_BLK, FLOW_PROBE, N_LEDGER, BLK, N_BLK, N_CO, N_ACC, N_ITEM, N_PARTY, N_RULE, N_RULE_P, N_GROUP,
    N_YEARS, JUMP_DAYS, PAIR_DAYS, DUP_WIN, ATTR_NO, ATTR_UNKNOWN, ACC_STATUS, RULE_DIRS, RULE_WHERE, ITEM_PENDING,
    K_DAY, K_YM, K_PAIR, q, inref, lref, lrange, esc, norm, raw, name2no, vdate, numv, numv_bad, money_in, money_out, chain,
    dsel, win_first, win_last, win, ym_of, k_aypm, k_igypm, LEDGER_COLS)

R1 = 2                                  # 第 1 条在第 2 行
RN = N_LEDGER + 1                       # 最后一条所在行
HAS_FIELDS = ('日期', '账户', '收费项', '收付人', '摘要', '收入', '支出', '对方账户', '归属公司')   # 「有内容」看这几格（备注、原表位置不算）
ATTR_LIST = list(ATTR_NO)               # 收入 支出 投资建设 往来 待分类 内部转账
ATTR_NUMS = [ATTR_NO[a] for a in ATTR_LIST]
# 私有帮助列（_表 里 layout 列之后；不进 NAMES，只给 _流 规则推断用）：能在「收付人」里找的项规则（查找位置 收付人/空）另压一张倒序小表，
# 规则推断拆成「摘要那一半（扫全部项规则）＋收付人那一半（只扫这张小表）」，取用户清单里靠上的那条：
# LibreOffice 实测 1,000 行 × 454 条：3.2 秒 → 1.7 秒（收付人那一半原来对每条规则都 SEARCH 一遍）。超过 Q_CAP 条自动退回整表写法。
Q_CAP = 100
_q0 = CI(_TAB_NEXT) + 1
QPRIV = dict(计数=CL(_q0), 源=CL(_q0 + 2), 词=CL(_q0 + 3), 向=CL(_q0 + 4), 范围=CL(_q0 + 5), 项号=CL(_q0 + 6))
Q_N = f"{H_TAB}!${QPRIV['源']}$1"                         # 小表条数（不封顶；>Q_CAP 时退回整表）


def _bref(block, field):
    """基础资料某块某列的锚定区域（表头行 ～ BASE_LAST）"""
    return inref(SH_BASE, BASE_COLS[block][field], block=block)


def _bx(block, field):
    """_表 第 k 行（k＝ROW()-1）对应基础资料该块第 k 个：INDEX(区域, ROW())（区域从表头行开始，所以第 k 个是第 k+1 格）"""
    return f'INDEX({_bref(block, field)},ROW())'


def _alias(x):
    """别名规：别名空＝""；否则 "、"&norm(别名，把 ，, ;； 换成 、)&"、" """
    s = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({x},"，","、"),",","、"),"；","、"),";","、")'
    return f'"、"&{norm(s)}&"、"'


# ═════════════════════════════ _表 ═════════════════════════════
def build_tab(ws, ctx):
    T = _TAB
    for grp, cols in T.items():
        for f, c in cols.items():
            ws[f'{c}1'] = f'{grp}·{f}'
    # —— 公司 ——
    C = LIST_CO
    for r in range(2, N_CO + 2):
        x = _bx('公司', '公司简称')
        ws[f'{C["有效"]}{r}'] = f'=IF({raw(x)},1,0)'
        ws[f'{C["名称"]}{r}'] = f'=IF({C["有效"]}{r}=1,TRIM({x}&""),"")'
        ws[f'{C["名称规"]}{r}'] = f'=IF({C["有效"]}{r}=1,{norm(x)},CHAR(1))'
    # —— 账户 ——
    A = LIST_ACC
    for r in range(2, N_ACC + 2):
        v = f'{A["有效"]}{r}'
        x = lambda f: _bx('账户', f)
        ws[v] = f'=IF({raw(x("账户名称"))},1,0)'
        ws[f'{A["名称"]}{r}'] = f'=IF({v}=1,TRIM({x("账户名称")}&""),"")'
        ws[f'{A["名称规"]}{r}'] = f'=IF({v}=1,{norm(x("账户名称"))},CHAR(1))'
        ws[f'{A["别名规"]}{r}'] = f'=IF(AND({v}=1,{raw(x("别名"))}),{_alias(x("别名"))},"")'
        ws[f'{A["子表名规"]}{r}'] = f'=IF(AND({v}=1,{raw(x("子表名"))}),{norm(x("子表名"))},CHAR(1))'
        ws[f'{A["公司号"]}{r}'] = f'=IF({v}=1,{name2no(x("公司"), "公司")},0)'
        t = f'TRIM({x("户别")}&"")'
        has = lambda w: f'ISNUMBER(SEARCH("{w}",{t}))'
        ws[f'{A["类型号"]}{r}'] = (f'=IF({v}<>1,0,IF(OR({has("对公")},{has("收单")},{has("基本户")},{has("一般户")}),1,'
                                  f'IF({has("私")},2,IF({has("现金")},3,4))))')
        ws[f'{A["期初"]}{r}'] = f'=IF({v}=1,{numv(x("期初余额"))},0)'
        st = f'TRIM({x("状态")}&"")'
        arr = '{' + ','.join(f'"{s}"' for s in ACC_STATUS) + '}'
        ws[f'{A["状态号"]}{r}'] = f'=IF({v}<>1,0,IF({st}="",1,IFERROR(MATCH({st},{arr},0),1)))'
        ws[f'{A["停用日"]}{r}'] = f'=IF({v}=1,{dsel(x("停用日"), "0")},0)'
    # —— 收费项 ——
    I = LIST_ITEM
    attr_arr = '{' + ','.join(f'"{a}"' for a in ATTR_LIST) + '}'
    for r in range(2, N_ITEM + 2):
        v = f'{I["有效"]}{r}'
        x = lambda f: _bx('收费项', f)
        ws[v] = f'=IF({raw(x("收费项"))},1,0)'
        ws[f'{I["名称"]}{r}'] = f'=IF({v}=1,TRIM({x("收费项")}&""),"")'
        ws[f'{I["名称规"]}{r}'] = f'=IF({v}=1,{norm(x("收费项"))},CHAR(1))'
        ws[f'{I["别名规"]}{r}'] = f'=IF(AND({v}=1,{raw(x("别名"))}),{_alias(x("别名"))},"")'
        ws[f'{I["属性号"]}{r}'] = (f'=IF({v}<>1,0,IFERROR(CHOOSE(MATCH(TRIM({x("属性")}&""),{attr_arr},0),'
                                  f'{",".join(str(n) for n in ATTR_NUMS)}),0))')
        g = f'{I["大类"]}{r}'
        ws[g] = f'=IF({v}=1,TRIM({x("大类")}&""),"")'
        fc = I['大类首现']
        ws[f'{fc}{r}'] = (f'=IF({g}="","",IF(MATCH({esc(g)},项_大类,0)=ROW()-1,MAX(${fc}$1:{fc}{r - 1})+1,""))')
        gn = f'{I["大类号"]}{r}'
        ws[gn] = (f'=IF({v}<>1,0,IF({g}="",99,IF(INDEX(项_大类首现,MATCH({esc(g)},项_大类,0))>{N_GROUP},99,'
                  f'INDEX(项_大类首现,MATCH({esc(g)},项_大类,0)))))')
        an = f'{I["属性号"]}{r}'
        for p in ATTR_NUMS:
            cond = f'OR({an}={p},{an}=0)' if p in (ATTR_NO['收入'], ATTR_NO['支出']) else f'{an}={p}'
            ws[f'{I[f"序键{p}"]}{r}'] = f'=IF(AND({v}=1,{cond}),{gn}*1000+ROW()-1,"")'
    for p in ATTR_NUMS:
        c = I[f'序键{p}']
        ws[f'{c}1'] = f'=COUNT({c}2:{c}{N_ITEM + 1})'
    # —— 大类压紧清单 ——
    gc = LIST_GROUP['名称']
    ws[f'{gc}1'] = '=MAX(项_大类首现)'
    for r in range(2, N_GROUP + 2):
        ws[f'{gc}{r}'] = f'=IF(ROW()-1>${gc}$1,"",INDEX(项_大类,MATCH(ROW()-1,项_大类首现,0)))'
    # —— 收付人 ——
    P = LIST_PARTY
    for r in range(2, N_PARTY + 2):
        v = f'{P["有效"]}{r}'
        x = lambda f: _bx('收付人', f)
        ws[v] = f'=IF({raw(x("收付人"))},1,0)'
        ws[f'{P["名称"]}{r}'] = f'=IF({v}=1,TRIM({x("收付人")}&""),"")'
        ws[f'{P["名称规"]}{r}'] = f'=IF({v}=1,{norm(x("收付人"))},CHAR(1))'
        ws[f'{P["别名规"]}{r}'] = f'=IF(AND({v}=1,{raw(x("别名"))}),{_alias(x("别名"))},"")'
        ws[f'{P["类型"]}{r}'] = f'=IF({v}=1,TRIM({x("类型")}&""),"")'
    # —— 规则（用户顺序）——
    G = LIST_RULE
    for r in range(2, N_RULE + 2):
        x = lambda f: _bx('规则', f)
        c = lambda f: f'{G[f]}{r}'
        kwt = 'TRIM(' + x('关键词') + '&"")'
        ws[c('词')] = '=' + esc(kwt)
        w = f'TRIM({x("查找位置")}&"")'
        ws[c('位')] = f'=IF({w}="",0,IFERROR(MATCH({w},{{"{RULE_WHERE[0]}","{RULE_WHERE[1]}"}},0),-1))'
        d = f'TRIM({x("方向")}&"")'
        ws[c('向')] = (f'=IF(OR({d}="",{d}="{RULE_DIRS[2]}"),0,IF({d}="{RULE_DIRS[0]}",1,IF({d}="{RULE_DIRS[1]}",2,-1)))')
        s = x('适用范围')
        ws[c('范围')] = (f'=IF(NOT({raw(s)}),0,IF({name2no(s, "公司")}>0,{name2no(s, "公司")},'
                         f'IF({name2no(s, "账户")}>0,100+{name2no(s, "账户")},-1)))')
        it = x('收费项')
        ws[c('项号')] = f'=IF(NOT({raw(it)}),0,IF({name2no(it, "项")}>0,{name2no(it, "项")},-1))'
        pt = x('收付人')
        ws[c('人号')] = f'=IF(NOT({raw(pt)}),0,IF({name2no(pt, "人")}>0,{name2no(pt, "人")},-1))'
        ws[c('有效')] = (f'=IF(AND({c("词")}<>"",{c("位")}>=0,{c("向")}>=0,{c("范围")}>=0,{c("项号")}>=0,{c("人号")}>=0,'
                         f'{c("项号")}+{c("人号")}>0),1,0)')
        one_i = f'IF(AND({c("有效")}=1,{c("项号")}>0),1,0)'
        one_p = f'IF(AND({c("有效")}=1,{c("人号")}>0),1,0)'
        ws[c('项计数')] = f'={one_i}' if r == 2 else f'={G["项计数"]}{r - 1}+{one_i}'
        ws[c('人计数')] = f'={one_p}' if r == 2 else f'={G["人计数"]}{r - 1}+{one_p}'
    # —— 压紧倒序：规项（第 j 行＝第 (项规则数+1−j) 条项规则）、规人 ——
    for tbl, cnt, par, n, res in ((LIST_RULE_I, '规_项计数', 'P_项规则数', N_RULE, '项号'),
                                  (LIST_RULE_P, '规_人计数', 'P_人规则数', N_RULE_P, '人号')):
        for r in range(2, n + 2):
            src = f'{tbl["源"]}{r}'
            ws[src] = f'=IF(ROW()-1>{par},0,IFERROR(MATCH({par}+2-ROW(),{cnt},0),0))'
            for f in ('词', '位', '向', '范围', res):
                ws[f'{tbl[f]}{r}'] = f'=IF({src}=0,"",INDEX(规_{f},{src}))'
    # —— 下拉压紧（第 1 行＝个数）——
    for k, pre, n in (('账户', '账户', N_ACC), ('公司', '公司', N_CO), ('收费项', '项', N_ITEM), ('收付人', '人', N_PARTY)):
        out, cnt = DV_COLS[k], DV_CNT[k]
        for r in range(2, n + 2):
            one = f'IF(INDEX({pre}_有效,ROW()-1)=1,1,0)'
            ws[f'{cnt}{r}'] = f'={one}' if r == 2 else f'={cnt}{r - 1}+{one}'
            ws[f'{out}{r}'] = f'=IF(ROW()-1>${out}$1,"",INDEX({pre}_名称,MATCH(ROW()-1,${cnt}$2:${cnt}${n + 1},0)))'
        ws[f'{out}1'] = f'={cnt}{n + 1}'
        ws[f'{cnt}1'] = f'下拉计数·{k}'
    # —— 收支分类候选键（属性 p 一列）——
    i0, g0, u0 = CAND_ROWS['项首'], CAND_ROWS['组首'], CAND_ROWS['认不出']
    for p in ATTR_NUMS:
        c = _TAB['分类候选'][f'p{p}']
        key = f'项_序键{p}'
        ws[f'{c}{CAND_ROWS["段首"]}'] = p * 10 ** 6
        for r in range(i0, i0 + N_ITEM):
            i = f'ROW()-{i0 - 1}'
            ws[f'{c}{r}'] = f'=IF(INDEX({key},{i})="","",{p * 10 ** 6}+INDEX({key},{i}))'
        for r in range(g0, g0 + N_ITEM):
            i = f'ROW()-{g0 - 1}'
            ws[f'{c}{r}'] = (f'=IF(INDEX({key},{i})="","",IF(COUNTIFS({key},">="&INDEX(项_大类号,{i})*1000,{key},"<"&INDEX({key},{i}))=0,'
                             f'{p * 10 ** 6}+INDEX(项_大类号,{i})*1000,""))')
        ws[f'{c}{u0}'] = p * 10 ** 6 + 999999
    # —— 公司→账户两级候选键 ——
    c = CAND_TWO
    for r in range(2, N_CO + 2):
        ws[f'{c}{r}'] = f'=IF(COUNTIFS(账户_公司号,ROW()-1,账户_状态号,1)>0,(ROW()-1)*100,"")'
    a0 = N_CO + 2
    for r in range(a0, a0 + N_ACC):
        a = f'(ROW()-{a0 - 1})'
        ws[f'{c}{r}'] = (f'=IF(INDEX(账户_有效,{a})<>1,"",IF(INDEX(账户_状态号,{a})>1,9900+{a},'
                         f'IF(INDEX(账户_公司号,{a})>0,INDEX(账户_公司号,{a})*100+{a},9800+{a})))')
    ws[f'{c}{a0 + N_ACC}'] = '=IF(COUNTIFS(账户_有效,1,账户_状态号,1,账户_公司号,0)>0,9800,"")'
    ws[f'{c}{a0 + N_ACC + 1}'] = '=IF(COUNTIFS(账户_有效,1,账户_状态号,">1")>0,9900,"")'
    assert a0 + N_ACC + 1 == CAND_TWO_LAST
    # —— 私有：能在收付人里找的项规则（位≠1）压紧倒序小表 ——
    G = LIST_RULE
    qc = QPRIV['计数']
    ws[f'{qc}1'] = '私·收付人项规则计数'
    for r in range(2, N_RULE + 2):
        one = f'IF(AND({G["有效"]}{r}=1,{G["项号"]}{r}>0,{G["位"]}{r}<>1),1,0)'
        ws[f'{qc}{r}'] = f'={one}' if r == 2 else f'={qc}{r - 1}+{one}'
    src = QPRIV['源']
    ws[f'{src}1'] = f'={qc}{N_RULE + 1}'
    for k in ('词', '向', '范围', '项号'):
        ws[f'{QPRIV[k]}1'] = f'私·收付人项规则·{k}'
    for r in range(2, Q_CAP + 2):
        ws[f'{src}{r}'] = f'=IF(ROW()-1>${src}$1,0,IFERROR(MATCH(${src}$1+2-ROW(),${qc}$2:${qc}${N_RULE + 1},0),0))'
        for k in ('词', '向', '范围', '项号'):
            ws[f'{QPRIV[k]}{r}'] = f'=IF({src}{r}=0,"",INDEX(规_{k},{src}{r}))'


# ═════════════════════════════ 规则推断（_流 I/J/AB 和自检格共用一套写法） ═════════════════════════════
# LOOKUP(1,0/(条件),结果)：数组表达式里不写 IF（只用乘、加、ISNUMBER(SEARCH)）；挑写法的 IF 都在数组外面（单值）。
def _tail(R, dirno, cono, accno):
    d, sc = R('向'), R('范围')
    return f'*(({d}=0)+({d}={dirno})>0)*(({sc}=0)+({sc}={cono})+({sc}=100+{accno})>0)'


def rule_cond(R, memo, party, dirno, cono, accno, both=True):
    """整表条件：both＝摘要、收付人都找（收费项规则）；False＝只找摘要（收付人规则：收付人本来就空着）"""
    kw, pos = R('词'), R('位')
    if both:
        hit = f'((ISNUMBER(SEARCH({kw},{memo}))*({pos}<>2)+ISNUMBER(SEARCH({kw},{party}))*({pos}<>1))>0)'
    else:
        hit = f'ISNUMBER(SEARCH({kw},{memo}))*({pos}<>2)'
    return hit + _tail(R, dirno, cono, accno)


def _memo_cond(R, memo, dirno, cono, accno):
    return f'ISNUMBER(SEARCH({R("词")},{memo}))*({R("位")}<>2)' + _tail(R, dirno, cono, accno)


def _party_cond(Q, party, dirno, cono, accno):
    return f'ISNUMBER(SEARCH({Q("词")},{party}))' + _tail(Q, dirno, cono, accno)


def item_lookup(R, Q, qn, user_ino, memo, party, dirno, cono, accno):
    """收费项规则推断：用户清单里最上面那条命中的项号；没有＝0。
       R＝压紧倒序的全部项规则（规项_*），Q＝其中能在收付人里找的（私有小表），qn＝小表条数，user_ino(m)＝第 m 条（用户顺序）的项号"""
    full = f'IFERROR(LOOKUP(1,0/({rule_cond(R, memo, party, dirno, cono, accno)}),{R("项号")}),0)'
    a = f'IF({raw(memo)},IFERROR(LOOKUP(1,0/({_memo_cond(R, memo, dirno, cono, accno)}),{R("源")}),10^6),10^6)'
    b = f'IF(AND({raw(party)},{qn}>0),IFERROR(LOOKUP(1,0/({_party_cond(Q, party, dirno, cono, accno)}),{Q("源")}),10^6),10^6)'
    return f'IF({qn}>{Q_CAP},{full},IFERROR(INDEX({user_ino},MIN({a},{b})),0))'


def item_multi(R, Q, qn, memo, party, dirno, cono, accno, ino):
    """「摘要能对上好几种收费项」：还有别的命中规则的项号≠ino → 那个项号（>0）；否则 0"""
    full = (f'IFERROR(LOOKUP(1,0/({rule_cond(R, memo, party, dirno, cono, accno)}*({R("项号")}<>{ino})),{R("项号")}),0)')
    a = (f'IF({raw(memo)},IFERROR(LOOKUP(1,0/({_memo_cond(R, memo, dirno, cono, accno)}*({R("项号")}<>{ino})),'
         f'{R("项号")}),0),0)')
    b = (f'IF(AND({raw(party)},{qn}>0),IFERROR(LOOKUP(1,0/({_party_cond(Q, party, dirno, cono, accno)}*({Q("项号")}<>{ino})),'
         f'{Q("项号")}),0),0)')
    return f'IF({qn}>{Q_CAP},{full},MAX({a},{b}))'


def _rng(pre, cnt):
    """压紧表某列的前 cnt 条：INDEX(规项_x,1):INDEX(规项_x,cnt)"""
    return lambda f: f'INDEX({pre}_{f},1):INDEX({pre}_{f},{cnt})'


def _qrng(f):
    """私有小表某列的前 Q_N 条（Q_N>Q_CAP 时不会用到）"""
    c = QPRIV[f]
    rg = f'{H_TAB}!${c}$2:${c}${Q_CAP + 1}'
    return f'INDEX({rg},1):INDEX({rg},{Q_N})'


# ═════════════════════════════ _流 ═════════════════════════════
def _blk_cols(k):
    """块表某组（各 N_ACC 列）的第 a 列"""
    return lambda a: CL(CI(FLOW_BLK[k]) + a - 1)


def build_flow(ws, ctx):
    F = FLOW
    for c, f, _d in FLOW_COLS:
        ws[f'{c}1'] = f
    pre_late = FLOW_BLK['前最晚']
    RI = _rng('规项', 'P_项规则数')
    RP = _rng('规人', 'P_人规则数')
    for r in range(R1, RN + 1):
        n = r - 1
        b = (n - 1) // BLK + 1                  # 第几块
        bs = (b - 1) * BLK + 2                  # 本块首行
        first = (r == bs)
        A = f'$A{r}'
        x = lambda f: lref(f, A)
        c = lambda k: f'{F[k]}{r}'
        B_, C_, D_, E_, F_, G_, H_, I_, J_, K_, L_, M_, N_, O_ = (c(k) for k in (
            '日期', '账户号', '收', '支', '有效', '对方号', '归属号', '项号', '人号', '属性号', '月码', '净', '正常日', '余额'))
        V_, Y_, Z_ = c('账日键'), c('转账键'), c('配对')
        ws[f'A{r}'] = n
        # ── 原始（B..L）：只看本行台帐格、_表、PAR_SAFE ──
        ws[B_] = '=' + vdate(x('日期'))
        ws[C_] = '=' + name2no(x('账户'), '账户')
        ws[D_] = '=' + money_in(x('收入'), x('支出'))
        ws[E_] = '=' + money_out(x('收入'), x('支出'))
        ws[F_] = f'=IF(AND({B_}>=P_开账日,{B_}<=P_日期上限,{C_}>0,OR({D_}<>0,{E_}<>0)),1,0)'
        ws[G_] = '=' + name2no(x('对方账户'), '账户')
        co = name2no(x('归属公司'), '公司')
        ws[H_] = f'=IF({C_}=0,0,IF({co}>0,{co},INDEX(账户_公司号,{C_})))'
        dirno = f'(2-({D_}>0))'
        cono = f'INDEX(账户_公司号,{C_})'
        memo, party = x('摘要'), x('收付人')
        ws[I_] = (f'=IF({raw(x("收费项"))},{name2no(x("收费项"), "项")},'
                  f'IF(AND({G_}>0,{G_}<>{C_},P_内部号>0),P_内部号,'
                  f'IF(AND({C_}>0,P_项规则数>0,OR({raw(memo)},{raw(party)})),'
                  f'{item_lookup(RI, _qrng, Q_N, "规_项号", memo, party, dirno, cono, C_)},0)))')
        cond_p = rule_cond(RP, memo, None, dirno, cono, C_, both=False)
        ws[J_] = (f'=IF({raw(party)},{name2no(party, "人")},'
                  f'IF(AND({C_}>0,P_人规则数>0,{raw(memo)}),IFERROR(LOOKUP(1,0/({cond_p}),{RP("人号")}),0),0))')
        unk = f'IF({D_}>0,1,2)' if ATTR_UNKNOWN == 0 else str(ATTR_UNKNOWN)
        ws[K_] = (f'=IF({I_}>0,IF(INDEX(项_属性号,{I_})>0,INDEX(项_属性号,{I_}),IF({D_}>0,1,2)),'
                  f'IF(AND({G_}>0,{G_}<>{C_}),{ATTR_NO["内部转账"]},{unk}))')
        ws[L_] = f'=IF({F_}=1,YEAR({B_})*100+MONTH({B_}),0)'
        # ── 往下可以引用块表（叶子）──
        ws[M_] = f'=IF({F_}=1,ROUND({D_}-{E_},2),0)'
        prev = f'${pre_late}${b + 1}' if first else f'MAX(${pre_late}${b + 1},MAX({F["正常日"]}${bs}:{F["正常日"]}{r - 1}))'
        ws[N_] = f'=IF({F_}<>1,0,IF(AND({prev}>=P_开账日,{B_}>{prev}+{JUMP_DAYS}),0,{B_}))'
        ws[O_] = (f'=IF({F_}<>1,"",ROUND(INDEX(账户_期初,{C_})+INDEX(块_前,{b},{C_})'
                  f'+SUMIFS({F["净"]}${bs}:{F["净"]}{r},{F["账户号"]}${bs}:{F["账户号"]}{r},{C_}),2))')
        yr, mo = f'YEAR({B_})', f'MONTH({B_})'
        ws[c('账年属月键')] = f'=IF({F_}=1,{k_aypm(C_, yr, K_, mo)},0)'
        ws[c('归年属月键')] = f'=IF({F_}=1,{k_aypm(H_, yr, K_, mo)},0)'
        ws[c('项年属月键')] = f'=IF({F_}=1,{k_aypm(I_, yr, K_, mo)},0)'
        ws[c('人年属月键')] = f'=IF({F_}=1,{k_aypm(J_, yr, K_, mo)},0)'
        ws[c('项归年属月键')] = f'=IF({F_}=1,{k_igypm(I_, H_, yr, K_, mo)},0)'
        ws[V_] = f'=IF({F_}=1,{C_}*{K_DAY}+{B_},0)'
        ws[c('账年键')] = f'=IF({F_}=1,{C_}*10000+{yr},0)'
        ws[c('账月键')] = f'=IF({F_}=1,{C_}*{K_YM}+{L_},0)'
        ws[Y_] = (f'=IF(AND({F_}=1,{K_}={ATTR_NO["内部转账"]},{G_}>0,{G_}<>{C_}),'
                  f'(MIN({C_},{G_})*{K_PAIR}+MAX({C_},{G_}))*{K_DAY}+{B_},0)')
        lo = win_first(ym_of(f'({B_}-{PAIR_DAYS})'))
        hi = win_last(ym_of(f'({B_}+{PAIR_DAYS})'), lo)
        wy, wm = win('流_转账键', lo, hi), win('流_净', lo, hi)
        ws[Z_] = (f'=IF({Y_}=0,"",IF(ROUND(SUMIF({wy},{Y_},{wm}),2)=0,0,'
                  f'IF(COUNTIFS({wy},">="&({Y_}-{PAIR_DAYS}),{wy},"<="&({Y_}+{PAIR_DAYS}),{wm},-{M_})'
                  f'>=COUNTIFS({wy},">="&({Y_}-{PAIR_DAYS}),{wy},"<="&({Y_}+{PAIR_DAYS}),{wm},{M_}),0,1)))')
        ws[c('期人键')] = f'=IF(AND({F_}=1,{B_}>=P_期起,{B_}<=P_期止),{J_}*10+{K_},0)'
        # ── △ 提醒（只在有效行算；没有 ※ 的行一定有效）──
        ws[c('提醒')] = '=' + _remind(r, A, x, c, RI)
        # ── ※ 问题 ──
        ws[c('问题')] = '=' + _problem(r, A, x, c)
    # ── 块表 ──
    net, pre, late, plate = (_blk_cols(k) for k in ('净', '前', '账最晚', '账前最晚'))
    for k, t in (('净', '块内净额'), ('前', '块前累计'), ('账最晚', '块内最晚日'), ('账前最晚', '块前最晚日')):
        for a in range(1, N_ACC + 1):
            ws[f'{_blk_cols(k)(a)}1'] = f'{t}·{a}'
    for k in ('最晚', '前最晚', '最晚月', '至最晚月', '最早月', '后最早月'):
        ws[f'{FLOW_BLK[k]}1'] = f'块·{k}'
    Cc, Mc, Nc, Lc, Vc = (F[k] for k in ('账户号', '净', '正常日', '月码', '账日键'))
    for b in range(1, N_BLK + 1):
        r = b + 1
        bs, be = (b - 1) * BLK + 2, b * BLK + 1
        rg = lambda col: f'${col}${bs}:${col}${be}'
        for a in range(1, N_ACC + 1):
            ws[f'{net(a)}{r}'] = f'=SUMIFS({rg(Mc)},{rg(Cc)},{a})'
            ws[f'{pre(a)}{r}'] = '=0' if b == 1 else f'={pre(a)}{r - 1}+{net(a)}{r - 1}'
            ws[f'{late(a)}{r}'] = (f'=MAX(0,IFERROR(LARGE({rg(Vc)},COUNTIF({rg(Vc)},">="&{(a + 1) * K_DAY})+1)-{a * K_DAY},0))')
            ws[f'{plate(a)}{r}'] = '=0' if b == 1 else f'=MAX({plate(a)}{r - 1},{late(a)}{r - 1})'
        cl = lambda k: FLOW_BLK[k]
        ws[f'{cl("最晚")}{r}'] = f'=MAX({rg(Nc)})'
        ws[f'{cl("前最晚")}{r}'] = '=0' if b == 1 else f'=MAX({cl("前最晚")}{r - 1},{cl("最晚")}{r - 1})'
        ws[f'{cl("最晚月")}{r}'] = f'=MAX({rg(Lc)})'
        ws[f'{cl("至最晚月")}{r}'] = (f'={cl("最晚月")}{r}' if b == 1 else f'=MAX({cl("至最晚月")}{r - 1},{cl("最晚月")}{r})')
        ws[f'{cl("最早月")}{r}'] = f'=IFERROR(SMALL({rg(Lc)},COUNTIF({rg(Lc)},0)+1),999999)'
        ws[f'{cl("后最早月")}{r}'] = (f'={cl("最早月")}{r}' if b == N_BLK else f'=MIN({cl("最早月")}{r},{cl("后最早月")}{r + 1})')
    _probe(ws)


def _has(x):
    return 'OR(' + ','.join(raw(x(f)) for f in HAS_FIELDS) + ')'


def _problem(r, A, x, c):
    """※ 链（只报第一个；没有 ※ 就等于 △ 提醒）"""
    B_, C_, D_, E_, F_, G_, I_, K_, N_ = (c(k) for k in ('日期', '账户号', '收', '支', '有效', '对方号', '项号', '属性号', '正常日'))
    nm = lambda f: f'TRIM({x(f)}&"")'
    bno = f'(INT(({A}-1)/{BLK})+1)'
    inblk = (f'IF(MOD({A}-1,{BLK})=0,0,MAX(INDEX(流_正常日,{A}-MOD({A}-1,{BLK})):INDEX(流_正常日,{A}-1)))')
    prev = f'MAX(INDEX(块_前最晚,{bno}),{inblk})'
    jump = f'"※ 比前面最晚的一笔（"&MONTH({prev})&"/"&DAY({prev})&"）晚 "&({B_}-{prev})&" 天：年份写错了？"'
    xin, xout = x('收入'), x('支出')
    pairs = [
        (f'NOT({raw(x("日期"))})', '"※ 没填日期"'),
        (f'{B_}=0', '"※ 日期看不懂：要写 2026/9/1 这样的日期"'),
        (f'{B_}<P_开账日', '"※ 早于开账日：上一年的账记在上一本"'),
        (f'{B_}>P_日期上限', '"※ 日期晚于 "&YEAR(P_日期上限)&" 年底：年份写错了？"'),
        (f'AND({F_}=1,{N_}=0)', jump),
        (f'NOT({raw(x("账户"))})', '"※ 没选账户"'),
        (f'{C_}=0', f'"※ 账户「"&{nm("账户")}&"」不在基础资料③：改了名就把旧名加进别名"'),
        (f'OR({numv_bad(xin)},{numv_bad(xout)})', '"※ 金额看不懂：10000+200 要写成 =10000+200"'),
        (f'AND({D_}=0,{E_}=0)', '"※ 没填金额"'),
        (f'AND({numv(xin)}<>0,{numv(xout)}<>0)', '"※ 收入、支出都填了：分成两行记"'),
        (f'IF({raw(x("收费项"))},{I_}=0,FALSE)', f'"※ 收费项「"&{nm("收费项")}&"」不在④：从下拉选，或去④加（改了名就加别名）"'),
        (f'AND({K_}={ATTR_NO["内部转账"]},NOT({raw(x("对方账户"))}))', '"※ 内部转账要选转账的另一户；对方不在本册就改成往来款、借款等"'),
        (f'IF({raw(x("对方账户"))},{G_}=0,FALSE)', f'"※ 转账的另一户「"&{nm("对方账户")}&"」不在③：外单位、个人请填在收付人"'),
        (f'AND({G_}>0,{G_}={C_})', '"※ 转账的另一户就是本账户"'),
        (f'AND({G_}>0,{K_}<>{ATTR_NO["内部转账"]})', '"※ 填了转账的另一户，收费项要选内部转账（原性质写摘要）"'),
        (f'IF({raw(x("归属公司"))},{name2no(x("归属公司"), "公司")}=0,FALSE)', f'"※ 归属公司「"&{nm("归属公司")}&"」不在②"'),
    ]
    return f'IF(NOT({_has(x)}),"",{chain(pairs, c("提醒"))})'


def _remind(r, A, x, c, RI):
    """△ 链（只报第一个；无效行＝""）。贵的检查都用 IF 挡在便宜条件后面"""
    B_, C_, D_, E_, F_, G_, I_, J_, K_, O_, V_, Z_ = (c(k) for k in (
        '日期', '账户号', '收', '支', '有效', '对方号', '项号', '人号', '属性号', '余额', '账日键', '配对'))
    lo, hi = f'MAX(1,{A}-{DUP_WIN})', f'MIN({N_LEDGER},{A}+{DUP_WIN})'
    wf = lambda nm_: f'INDEX({nm_},{lo}):INDEX({nm_},{hi})'
    wl = lambda f: f'INDEX({inref(SH_LEDGER, LEDGER_COLS[f])},{lo}+1):INDEX({inref(SH_LEDGER, LEDGER_COLS[f])},{hi}+1)'
    num = f'{wf("流_账日键")},{V_},{wf("流_收")},{D_},{wf("流_支")},{E_}'
    txt = ','.join(wl(f) + ',"="&' + esc(x(f) + '&""') for f in ('摘要', '备注', '原表位置'))
    dup = f'IF(COUNTIFS({num})>1,IFERROR(COUNTIFS({num},{txt}),0),0)>1'
    bno = f'(INT(({A}-1)/{BLK})+1)'
    kb = f'INDEX(流_账日键,{A}-MOD({A}-1,{BLK})):INDEX(流_账日键,{A}-1)'
    order = (f'IF({B_}<INDEX(块_账前最晚,{bno},{C_}),1,IF(MOD({A}-1,{BLK})=0,0,'
             f'COUNTIFS({kb},">"&{V_},{kb},"<"&({C_}+1)*{K_DAY})))>0')
    guessed = f'AND(NOT({raw(x("收费项"))}),{I_}>0,{G_}=0)'
    multi = item_multi(RI, _qrng, Q_N, x('摘要'), x('收付人'), f'(2-({D_}>0))', f'INDEX(账户_公司号,{C_})', C_, I_)
    iname = f'INDEX(项_名称,{I_})'
    xin, xout = x('收入'), x('支出')
    textc = f'OR(ISTEXT({x("日期")}),AND(ISTEXT({xin}),{raw(xin)}),AND(ISTEXT({xout}),{raw(xout)}))'
    negc = f'OR({numv(xin)}<0,{numv(xout)}<0)'
    me = f'{x("摘要")}&{x("备注")}'
    refund = f'OR(ISNUMBER(SEARCH("退",{me})),ISNUMBER(SEARCH("冲",{me})),ISNUMBER(SEARCH("返还",{me})))'
    paused = (f'AND(INDEX(账户_状态号,{C_})>1,OR(INDEX(账户_停用日,{C_})=0,{B_}>INDEX(账户_停用日,{C_})))')
    pairs = [
        (f'{Z_}=1', '"△ 内部转账没配上：漏录一头、金额或日期对不上"'),
        (dup, f'"△ 重复粘贴了？前后 {DUP_WIN} 行有一笔一模一样的"'),
        (order, '"△ 本账户上面有日期更晚的：按日期排一下序"'),
        (f'AND({I_}>0,{I_}=P_待定号)', '"△ 收费项是「待定」：定了性质请改过来"'),
        (f'IF({guessed},{multi},0)>0', f'"△ 摘要能对上好几种收费项，按「"&{iname}&"」算了：请在收费项里选"'),
        (f'AND(NOT({raw(x("收费项"))}),{I_}=0)', '"△ 收费项没填，摘要也认不出：请选收费项"'),
        (guessed, f'"△ 收费项按摘要认成「"&{iname}&"」：对就在收费项选上"'),
        (textc, '"△ 日期或金额是文字，已照认：最好改成真日期、数字"'),
        (negc, '"△ 金额是负数，已当另一栏算"'),
        (f'IF(AND({K_}=1,{E_}>0),NOT({refund}),FALSE)', '"△ 收入类记在支出栏，按冲减收入算：是退款就在摘要写「退」"'),
        (f'IF(AND({K_}=2,{D_}>0),NOT({refund}),FALSE)', '"△ 支出类记在收入栏，按冲减支出算：是退款就在摘要写「退」"'),
        (f'IF({raw(x("收付人"))},{J_}=0,FALSE)', f'"△ 收付人「"&TRIM({x("收付人")}&"")&"」不在⑤：报表算没填"'),
        (f'AND({J_}=0,OR({K_}=1,{K_}=2))', '"△ 收付人没填，摘要也认不出"'),
        (f'{O_}<0', '"△ 这一笔之后本账户余额是负数：漏录收入或顺序不对？"'),
        (paused, '"△ 账户已暂停/销户，还在记账"'),
    ]
    return f'IF({F_}<>1,"",{chain(pairs)})'


# —— 规则推断自检（设计 §1.5.4 perf#8）：常数规则 6 条（倒序存，跟压紧表一样）＋3 个同样写法的 LOOKUP 格 ——
PROBE_RULES = [  # 用户顺序：词（esc 前）, 位, 向, 范围, 项号
    ('手续费', 1, 2, 0, 11), ('A*B', 0, 0, 0, 12), ('租金', 0, 1, 3, 13), ('顺丰', 2, 0, 0, 14), ('费', 1, 0, 105, 15), ('租', 0, 0, 0, 16)]
PROBE_CASES = [  # 摘要, 收付人, 方向号, 公司号, 账户号, 期望（用户清单里最上面那条命中的项号）
    ('A105 3月租金', '', 1, 3, 7, 13), ('AxB 租金手续费', '顺丰速运', 1, 2, 5, 14), ('退场地费', '', 2, 1, 5, 15)]
PROBE_HDR = ['词', '位', '向', '范围', '项号', '源', '用户项号', 'Q源', 'Q词', 'Q向', 'Q范围', 'Q项号',
             '摘要', '收付人', '方向号', '公司号', '账户号', '期望', '结果', '不一致']


def probe_col(k):
    return CL(CI(FLOW_PROBE['首列']) + PROBE_HDR.index(k))


def _probe(ws):
    """跟 _流 收费项规则推断一模一样的写法（item_lookup：摘要一半＋收付人一半取靠上的）算 3 个例子，期望值生成时写死"""
    n = len(PROBE_RULES)
    assert n == FLOW_PROBE['行数']
    for k in PROBE_HDR:
        ws[f'{probe_col(k)}1'] = '自检·' + k
    qrules = [k for k in range(n, 0, -1) if PROBE_RULES[k - 1][1] != 1]      # 能在收付人里找的（倒序）
    for j in range(1, n + 1):                   # 第 j 行＝第 (n+1−j) 条（倒序）
        src = n + 1 - j
        w, pos, d, sc, it = PROBE_RULES[src - 1]
        r = j + 1
        ws[f'{probe_col("词")}{r}'] = '=' + esc(f'"{w}"')
        ws[f'{probe_col("位")}{r}'], ws[f'{probe_col("向")}{r}'] = pos, d
        ws[f'{probe_col("范围")}{r}'], ws[f'{probe_col("项号")}{r}'] = sc, it
        ws[f'{probe_col("源")}{r}'] = src
        ws[f'{probe_col("用户项号")}{r}'] = PROBE_RULES[j - 1][4]
    for j, src in enumerate(qrules, 1):
        w, pos, d, sc, it = PROBE_RULES[src - 1]
        r = j + 1
        ws[f'{probe_col("Q源")}{r}'] = src
        ws[f'{probe_col("Q词")}{r}'] = '=' + esc(f'"{w}"')
        ws[f'{probe_col("Q向")}{r}'], ws[f'{probe_col("Q范围")}{r}'], ws[f'{probe_col("Q项号")}{r}'] = d, sc, it
    rg = lambda k, m: f'${probe_col(k)}$2:${probe_col(k)}${n + 1}'
    R = lambda f: f'INDEX({rg(f, n)},1):INDEX({rg(f, n)},{n})'
    Q = lambda f: f'INDEX({rg("Q" + f, n)},1):INDEX({rg("Q" + f, n)},{len(qrules)})'
    for i, (memo, party, d, co, acc, exp) in enumerate(PROBE_CASES):
        r = i + 2
        cc = lambda k: f'{probe_col(k)}{r}'
        ws[cc('摘要')], ws[cc('收付人')] = memo, party or None
        ws[cc('方向号')], ws[cc('公司号')], ws[cc('账户号')], ws[cc('期望')] = d, co, acc, exp
        ws[cc('结果')] = '=' + item_lookup(R, Q, str(len(qrules)), rg('用户项号', n), cc('摘要'), cc('收付人'),
                                         cc('方向号'), cc('公司号'), cc('账户号'))
        ws[cc('不一致')] = f'=IF({cc("结果")}={cc("期望")},0,1)'


# ═════════════════════════════ _参 ═════════════════════════════
def build_par(ws, ctx):
    bp = lambda k: f'{q(SH_BASE)}!${BASE_PAR_VAL}${PAR_ROW[k]}'
    open_d = ctx['params'].get('开账日')
    dflt = f'DATE({open_d.year},{open_d.month},{open_d.day})' if open_d else 'DATE(2026,1,1)'
    mg = bp('管理年度')
    ps = bp('报表暂停')
    pst = f'TRIM({ps}&"")'
    pc = probe_col('不一致')
    home_sel = f'{q(SH_HOME)}!${HOME_SEL["截止"][0]}${HOME_SEL["截止"][1:]}'
    span = lambda k: f'{q(SH_SPAN)}!${SPAN_SEL[k][0]}${SPAN_SEL[k][1:]}'
    rows = {
        '集团名称': (f'=TRIM({bp("集团名称")}&"")', '基础资料①（空着：标题自己写「集团」）'),
        '开账日': ('=' + dsel(bp('开账日'), dflt), '基础资料①（dsel）'),
        '开账年': ('=YEAR(P_开账日)', ''),
        '日期上限': (f'=DATE(P_开账年+{N_YEARS - 1},12,31)', '更晚的 ※ 不算'),
        '最后日期': ('=MAX(P_开账日,MAX(块_最晚))', '有效、没跳期的流水里最晚的一天'),
        '管理年度': (f'=IF(ISNUMBER({mg}),IF(AND({mg}>=2000,{mg}<=2100),INT({mg}),YEAR(P_最后日期)),YEAR(P_最后日期))',
                  '基础资料①填了 2000～2100 用它；否则最后日期的年份'),
        '截止': ('=' + dsel(home_sel, 'IF(P_管理年度<YEAR(P_最后日期),DATE(P_管理年度,12,31),P_最后日期)'), '首页 C3'),
        '截止年': ('=YEAR(P_截止)', ''),
        '截止月码': ('=YEAR(P_截止)*100+MONTH(P_截止)', ''),
        '期起': ('=' + dsel(span('开始'), 'DATE(YEAR(P_最后日期),1,1)'), '任意时间段 B3'),
        '期止': ('=' + dsel(span('结束'), 'P_最后日期'), '任意时间段 B4'),
        '已用行数': (f'=MAX(COUNTA({lrange("日期")}),COUNTA({lrange("账户")}))', '资金台帐 日期或账户有内容的行数'),
        '期初合计': ('=SUM(账户_期初)', ''),
        '待定号': (f'=IFERROR(MATCH("{ITEM_PENDING}",项_名称规,0),0)', ''),
        '内部号': (f'=IFERROR(MATCH({ATTR_NO["内部转账"]},项_属性号,0),0)', '第一个属性「内部转账」的收费项'),
        '账户数': ('=COUNTIF(账户_有效,1)', ''),
        '公司数': ('=COUNTIF(公司_有效,1)', ''),
        '收费项数': ('=COUNTIF(项_有效,1)', ''),
        '收付人数': ('=COUNTIF(人_有效,1)', ''),
        '规则数': ('=COUNTIF(规_有效,1)', '有效规则条数'),
        '项规则数': ('=MAX(规_项计数)', ''),
        '人规则数': (f'=MIN({N_RULE_P},MAX(规_人计数))', f'超过 {N_RULE_P} 条核对表 ※'),
        '报表暂停': (f'=IF(OR({pst}="是",{pst}="1",{pst}="暂停"),1,0)', '基础资料①'),
        '自检': (f'=SUM({H_FLOW}!${pc}$2:${pc}${len(PROBE_CASES) + 1})', '规则推断自检不一致的个数（0＝正常）'),
    }
    assert set(rows) == set(PAR_KEYS), set(PAR_KEYS) ^ set(rows)
    ws['A1'], ws['B1'], ws['C1'] = '参数', '值', '说明'
    for k, cell in PAR.items():
        r = int(cell[1:])
        f, note = rows[k]
        ws[f'A{r}'] = k
        ws[cell] = f
        if note:
            ws[f'C{r}'] = note
    for k in ('开账日', '日期上限', '最后日期', '截止', '期起', '期止'):
        ws[PAR[k]].number_format = 'yyyy/mm/dd'
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 40


def build(wb, ctx):
    build_tab(wb[H_TAB], ctx)
    build_flow(wb[H_FLOW], ctx)
    build_par(wb[H_PAR], ctx)
