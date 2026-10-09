# -*- coding: utf-8 -*-
"""隐藏计算表：_参（全书期间）、_表（基础资料逐行镜像＋压紧的下拉清单）、_收（收支登记逐条）、_往（应收应付登记逐条）。
   全部按「第 n 条」取数：INDEX(录入表!$X$表头行:$X$1048576, n+1)，插行删行不错位。查看表只用 layout.NAMES 里的名称。"""
from layout import *
from common import F_HELP


def _hdr(ws, cols):
    for c, f, _d in cols:
        ws[f'{c}1'] = f
        ws[f'{c}1'].font = F_HELP


def chain(pairs, default='""'):
    out = default
    for c, v in reversed(pairs):
        out = f'IF({c},{v},{out})'
    return out


def _ix(sh, col, n_ref, block=None):
    return f'INDEX({inref(sh, col, block=block)},{n_ref}+1)'


# ───────────────────────── _参 ─────────────────────────
def build_par(ws):
    P = lambda k: f'{SH_BASE}!$C${PAR_ROW[k]}'
    first = ('IF(COUNTIF(收_日期,">0")+COUNTIF(往_日期,">0")=0,0,'
             'MIN(IF(COUNTIF(收_日期,">0")=0,99999,SMALL(收_日期,COUNTIF(收_日期,0)+1)),'
             'IF(COUNTIF(往_日期,">0")=0,99999,SMALL(往_日期,COUNTIF(往_日期,0)+1))))')
    rows = {
        '单位名称': f'=TRIM({P("单位名称")}&"")',
        '建账日': f'=IF(ISNUMBER({P("建账日")}),INT({P("建账日")}),IF({first}>0,{first},DATE(2026,1,1)))',
        '账期': f'=IF(N({P("账期")})>0,N({P("账期")}),30)',
        '最后日期': '=IF(MAX(MAX(收_日期),MAX(往_日期))>0,MAX(MAX(收_日期),MAX(往_日期)),IF(B3>0,B3,DATE(2026,1,1)))',
        '截止': f'=IF(ISNUMBER({SH_HOME}!${HOME_END[0]}${HOME_END[1:]}),INT({SH_HOME}!${HOME_END[0]}${HOME_END[1:]}),B5)',
        '年度': '=YEAR(B6)',
        '年初': '=DATE(B7,1,1)',
        '默认账户': (f'=IF(ISNUMBER(MATCH("现金",账户_名称,0)),"现金",IF({H_LIST}!${COMPACT["账户"][0]}$2="","现金",'
                 f'{H_LIST}!${COMPACT["账户"][0]}$2))'),
        '截止年月': '=YEAR(B6)*100+MONTH(B6)',
    }
    for i, (key, _t, _tr, hr, cap) in enumerate(BASE_BLOCKS):
        if i + 1 < len(BASE_BLOCKS):
            nxt = BASE_BLOCKS[i + 1][1]
            rows[f'{key}位'] = f'=IFERROR(MATCH("{nxt}",{SH_BASE}!$A:$A,0)-ROW({SH_BASE}!$A${hr})-1,{cap})'
        else:
            rows[f'{key}位'] = f'={cap + 300}'
    ws['A1'] = '全书共用的参数（自动，别改）'
    for k, cell in PAR.items():
        r = int(cell[1:])
        ws[f'A{r}'] = k
        ws[cell] = rows[k]


# ───────────────────────── _表 ─────────────────────────
def build_list(ws):
    ws['A1'] = '板块'
    ws['F1'] = '账户'
    ws['I1'] = '收支项目'
    ws['N1'] = '经手人'
    ws['P1'] = '往来单位'
    for i in range(LIST_N):
        r, n = i + 2, i + 1
        if n <= N_SEG:
            g = lambda k: _ix(SH_BASE, SEG_COLS[k], n, '板块')
            a = f'{LIST_SEG["名称"]}{r}'
            ws[a] = f'=IF({n}>P_板块位,"",TRIM({g("名称")}&""))'
            for k in ('期初结余', '期初存量', '期初固定资产'):
                ws[f'{LIST_SEG[k]}{r}'] = f'=IF({a}="",0,N({g(k)}))'
        if n <= N_ACC:
            g = lambda k: _ix(SH_BASE, ACC_COLS[k], n, '账户')
            a = f'{LIST_ACC["名称"]}{r}'
            ws[a] = f'=IF({n}>P_账户位,"",TRIM({g("名称")}&""))'
            ws[f'{LIST_ACC["期初余额"]}{r}'] = f'=IF({a}="",0,N({g("期初余额")}))'
        if n <= N_ITEM:
            g = lambda k: _ix(SH_BASE, ITEM_COLS[k], n, '项目')
            a = f'{LIST_ITEM["名称"]}{r}'
            ws[a] = f'=IF({n}>P_项目位,"",TRIM({g("名称")}&""))'
            ws[f'{LIST_ITEM["方向"]}{r}'] = f'=IF({a}="","",IF(TRIM({g("方向")}&"")="","双向",TRIM({g("方向")}&"")))'
            ws[f'{LIST_ITEM["用途"]}{r}'] = f'=IF({a}="","",IF(TRIM({g("用途")}&"")="","普通",TRIM({g("用途")}&"")))'
            ws[f'{LIST_ITEM["库存"]}{r}'] = f'=IF({a}="","",TRIM({g("库存")}&""))'
        if n <= N_PER:
            ws[f'{LIST_PER["姓名"]}{r}'] = f'=IF({n}>P_经手人位,"",TRIM({_ix(SH_BASE, PER_COLS["姓名"], n, "经手人")}&""))'
        if n <= N_UNIT:
            g = lambda k: _ix(SH_UNIT, UNIT_COLS[k], n)
            a = f'{LIST_UNIT["名称"]}{r}'
            ws[a] = f'=TRIM({g("名称")}&"")'
            ws[f'{LIST_UNIT["类型"]}{r}'] = f'=IF({a}="","",TRIM({g("类型")}&""))'
            ws[f'{LIST_UNIT["账期"]}{r}'] = f'=IF({a}="",0,IF(N({g("账期")})>0,N({g("账期")}),P_账期))'
            ws[f'{LIST_UNIT["联系人"]}{r}'] = f'=IF({a}="","",TRIM({g("联系人")}&""))'
            ws[f'{LIST_UNIT["电话"]}{r}'] = f'=IF({a}="","",TRIM({g("电话")}&""))'
            ws[f'{LIST_UNIT["备注"]}{r}'] = f'=IF({a}="","",TRIM({g("备注")}&""))'
    # 压紧的清单（下拉用）：计数列逐行累加，清单第 k 个＝计数第一次到 k 的那行
    for key, (out, src, cap) in COMPACT.items():
        cnt = COMPACT_CNT[key]
        for i in range(cap):
            r = i + 2
            prev = f'+{cnt}{r - 1}' if i else ''
            ws[f'{cnt}{r}'] = f'=IF({src}{r}<>"",1,0){prev}'
            ws[f'{out}{r}'] = (f'=IF({i + 1}>${out}$1,"",INDEX(${src}$2:${src}${cap + 1},'
                               f'MATCH({i + 1},${cnt}$2:${cnt}${cap + 1},0)))')
        ws[f'{out}1'] = f'={cnt}{cap + 1}'
    for c in list(COMPACT_CNT.values()):
        ws[f'{c}1'] = '计数'


# ───────────────────────── _收 ─────────────────────────
def build_shou(ws):
    _hdr(ws, SHOU_COLS)
    S = SHOU
    C = CASH_COLS
    for i in range(N_CASH):
        r = i + 2
        n = f'$A{r}'
        ix = lambda k: _ix(SH_CASH, C[k], n)
        f = {'n': i + 1}
        f['日期'] = f'=IF(ISNUMBER({ix("日期")}),INT({ix("日期")}),0)'
        for k in ('板块', '收支项目', '往来单位', '经手人'):
            f[k] = f'=TRIM({ix(k)}&"")'
        f['摘要'] = f'={ix("摘要")}&""'
        f['备注'] = f'={ix("备注")}&""'
        f['数量'] = f'=N({ix("数量")})'
        f['单价'] = f'=N({ix("单价")})'
        D = f'D{r}'
        f['方向'] = f'=IF({D}="","",IFERROR(INDEX(项目_方向,MATCH({esc(D)},项目_名称,0))&"",""))'
        f['用途'] = f'=IF({D}="","普通",IFERROR(INDEX(项目_用途,MATCH({esc(D)},项目_名称,0))&"","普通"))'
        f['库存'] = f'=IF({D}="","",IFERROR(INDEX(项目_库存,MATCH({esc(D)},项目_名称,0))&"",""))'
        rin, rout = f'N({ix("收入")})', f'N({ix("支出")})'
        qp = f'ROUND(G{r}*H{r},2)'
        f['有效'] = (f'=IF(AND(B{r}>0,OR({rin}<>0,{rout}<>0,AND({qp}<>0,OR(P{r}="收入",P{r}="支出")))),1,0)')
        f['收入'] = f'=IF(O{r}=1,IF({rin}<>0,{rin},IF(AND({rout}=0,P{r}="收入"),{qp},0)),0)'
        f['支出'] = f'=IF(O{r}=1,IF({rout}<>0,{rout},IF(AND({rin}=0,P{r}="支出"),{qp},0)),0)'
        f['净额'] = f'=I{r}-J{r}'
        f['账户'] = f'=IF(TRIM({ix("账户")}&"")="",P_默认账户,TRIM({ix("账户")}&""))'
        f['入库量'] = f'=IF(AND(O{r}=1,R{r}="入库"),G{r},0)'
        f['出库量'] = f'=IF(AND(O{r}=1,R{r}="出库"),G{r},0)'
        f['冲应收'] = f'=IF(AND(O{r}=1,Q{r}="冲应收"),I{r}-J{r},0)'
        f['冲应付'] = f'=IF(AND(O{r}=1,Q{r}="冲应付"),J{r}-I{r},0)'
        f['固定资产'] = f'=IF(AND(O{r}=1,Q{r}="固定资产"),J{r}-I{r},0)'
        f['内部转账'] = f'=IF(AND(O{r}=1,Q{r}="内部转账"),1,0)'
        f['板块收入'] = f'=IF(X{r}=1,0,I{r})'
        f['板块支出'] = f'=IF(X{r}=1,0,J{r})'
        f['录入行'] = f'=ROW({ix("日期")})'
        f['年月'] = f'=IF(B{r}=0,0,YEAR(B{r})*100+MONTH(B{r}))'
        f['显示摘要'] = f'=IF(E{r}<>"",E{r},D{r})'
        f['排序键'] = f'=IF(O{r}=1,B{r}*10000+A{r},"")'
        f['账户有效'] = f'=IF(ISNUMBER(MATCH({esc(f"L{r}")},账户_名称,0)),1,0)'
        F = f'F{r}'
        f['最新收付'] = (f'=IF(AND(O{r}=1,OR(Q{r}="冲应收",Q{r}="冲应付"),{F}<>"",B{r}<=P_截止),'
                       f'IF(COUNTIFS(收_往来单位,{esc(F)},收_用途,Q{r},收_有效,1,收_日期,">"&B{r},收_日期,"<="&P_截止)'
                       f'+COUNTIFS(收_往来单位,{esc(F)},收_用途,Q{r},收_有效,1,收_日期,B{r},收_n,">"&A{r})=0,1,0),0)')
        blank = f'AND(B{r}=0,C{r}="",D{r}="",{rin}=0,{rout}=0,G{r}=0)'
        f['校验'] = '=' + chain([
            (blank, '""'),
            (f'B{r}=0', '"✗ 没填日期（或不是真日期）"'),
            (f'D{r}=""', '"✗ 没选收支项目"'),
            (f'ISNA(MATCH({esc(f"D{r}")},项目_名称,0))', f'"✗ 收支项目「"&D{r}&"」不在【基础资料】④里"'),
            (f'O{r}=0', '"✗ 没填金额（也没有数量×单价）"'),
            (f'AND({rin}<>0,{rout}<>0)', '"⚠ 收入、支出都填了"'),
            (f'AND(C{r}<>"",ISNA(MATCH({esc(f"C{r}")},板块_名称,0)))', f'"✗ 业务板块「"&C{r}&"」不在【基础资料】②里"'),
            (f'AND(C{r}="",X{r}=0)', '"⚠ 没选业务板块（不进任何板块表）"'),
            (f'AG{r}=0', f'"✗ 账户「"&L{r}&"」不在【基础资料】③里"'),
            (f'AND(OR(Q{r}="冲应收",Q{r}="冲应付"),{F}="")', '"✗ 收回/支付欠款要选往来单位"'),
            (f'AND({F}<>"",ISNA(MATCH({esc(F)},单位_名称,0)))', f'"⚠ 往来单位「"&{F}&"」不在【往来单位】里"'),
            (f'AND(R{r}<>"",G{r}=0)', '"⚠ 粮食购进/销售没填数量（存量不变）"'),
            (f'AND(M{r}<>"",ISNA(MATCH({esc(f"M{r}")},经手人_姓名,0)))', f'"⚠ 经手人「"&M{r}&"」不在【基础资料】⑤里"'),
            (f'B{r}<P_建账日', '"⚠ 日期在建账日以前（已经算进期初；期初余额里如果已经有了就重复了）"'),
            (f'OR(YEAR(B{r})>YEAR(P_建账日)+10,YEAR(B{r})<2000)', '"⚠ 年份看着不对（打错了？）"'),
            (f'AND(P{r}="收入",{rout}<>0,{rin}=0)', '"⚠ 收入类的项目填在了支出栏"'),
            (f'AND(P{r}="支出",{rin}<>0,{rout}=0)', '"⚠ 支出类的项目填在了收入栏"'),
            (f'AND(O{r}=1,{rin}=0,{rout}=0)', f'"金额没填，按数量×单价＝"&TEXT(I{r}+J{r},"#,##0.00")&" 计"'),
        ])
        for k, v in f.items():
            ws[f'{S[k]}{r}'] = v


# ───────────────────────── _往 ─────────────────────────
def build_wang(ws):
    _hdr(ws, WANG_COLS)
    Wn = WANG
    C = WL_COLS
    for i in range(N_WL):
        r = i + 2
        n = f'$A{r}'
        ix = lambda k: _ix(SH_WL, C[k], n)
        f = {'n': i + 1}
        f['日期'] = f'=IF(ISNUMBER({ix("日期")}),INT({ix("日期")}),0)'
        for k in ('板块', '往来单位', '类型', '收支项目', '经手人'):
            f[k] = f'=TRIM({ix(k)}&"")'
        f['摘要'] = f'={ix("摘要")}&""'
        f['备注'] = f'={ix("备注")}&""'
        f['数量'] = f'=N({ix("数量")})'
        f['单价'] = f'=N({ix("单价")})'
        f['金额'] = f'=IF(N({ix("金额")})<>0,N({ix("金额")}),ROUND(H{r}*I{r},2))'
        D, E, F = f'D{r}', f'E{r}', f'F{r}'
        term = f'IFERROR(INDEX(单位_账期,MATCH({esc(D)},单位_名称,0)),P_账期)'
        f['约定日期'] = f'=IF(ISNUMBER({ix("约定日期")}),INT({ix("约定日期")}),IF(B{r}>0,B{r}+{term},0))'
        f['有效'] = f'=IF(AND(B{r}>0,{D}<>"",OR({E}="应收",{E}="应付"),J{r}<>0),1,0)'
        f['应收额'] = f'=IF(AND(N{r}=1,{E}="应收"),J{r},0)'
        f['应付额'] = f'=IF(AND(N{r}=1,{E}="应付"),J{r},0)'
        f['库存'] = f'=IF({F}="","",IFERROR(INDEX(项目_库存,MATCH({esc(F)},项目_名称,0))&"",""))'
        f['入库量'] = f'=IF(AND(N{r}=1,Q{r}="入库"),H{r},0)'
        f['出库量'] = f'=IF(AND(N{r}=1,Q{r}="出库"),H{r},0)'
        f['固定资产'] = f'=IF(AND(N{r}=1,IFERROR(INDEX(项目_用途,MATCH({esc(F)},项目_名称,0)),"")="固定资产"),J{r},0)'
        f['正额'] = f'=IF(AND(N{r}=1,J{r}>0,B{r}<=P_截止),J{r},0)'
        f['累计前'] = (f'=IF(U{r}=0,0,SUMIFS(往_正额,往_往来单位,{esc(D)},往_类型,{E},往_日期,"<"&B{r})'
                    f'+SUMIFS(往_正额,往_往来单位,{esc(D)},往_类型,{E},往_日期,B{r},往_n,"<"&A{r}))')
        settled = (f'IF({E}="应收",SUMIFS(收_冲应收,收_往来单位,{esc(D)},收_日期,"<="&P_截止),'
                   f'SUMIFS(收_冲应付,收_往来单位,{esc(D)},收_日期,"<="&P_截止))'
                   f'-SUMIFS(往_金额,往_往来单位,{esc(D)},往_类型,{E},往_有效,1,往_金额,"<0",往_日期,"<="&P_截止)')
        f['未结'] = f'=IF(U{r}=0,0,ROUND(MAX(0,MIN(U{r},V{r}+U{r}-({settled}))),2))'
        f['逾期天数'] = f'=IF(AND(W{r}>0,K{r}<P_截止),P_截止-K{r},0)'
        f['最早未结'] = (f'=IF(W{r}>0,IF(COUNTIFS(往_往来单位,{esc(D)},往_类型,{E},往_未结,">0",往_约定日期,"<"&K{r})'
                     f'+COUNTIFS(往_往来单位,{esc(D)},往_类型,{E},往_未结,">0",往_约定日期,K{r},往_n,"<"&A{r})=0,1,0),0)')
        f['最新发生'] = (f'=IF(AND(N{r}=1,J{r}>0,B{r}<=P_截止),IF(COUNTIFS(往_往来单位,{esc(D)},往_类型,{E},往_有效,1,往_金额,">0",往_日期,">"&B{r},'
                     f'往_日期,"<="&P_截止)+COUNTIFS(往_往来单位,{esc(D)},往_类型,{E},往_有效,1,往_金额,">0",往_日期,B{r},往_n,">"&A{r})=0,1,0),0)')
        f['录入行'] = f'=ROW({ix("日期")})'
        f['年月'] = f'=IF(B{r}=0,0,YEAR(B{r})*100+MONTH(B{r}))'
        f['显示摘要'] = f'=IF(G{r}<>"",G{r},{F})'
        f['排序键'] = f'=IF(N{r}=1,B{r}*10000+{N_CASH}+A{r},"")'
        blank = f'AND(B{r}=0,C{r}="",{D}="",{E}="",J{r}=0)'
        f['校验'] = '=' + chain([
            (blank, '""'),
            (f'B{r}=0', '"✗ 没填日期（或不是真日期）"'),
            (f'{D}=""', '"✗ 没选往来单位"'),
            (f'AND({E}<>"应收",{E}<>"应付")', '"✗ 应收/应付没选"'),
            (f'J{r}=0', '"✗ 没填金额（也没有数量×单价）"'),
            (f'AND(C{r}<>"",ISNA(MATCH({esc(f"C{r}")},板块_名称,0)))', f'"✗ 业务板块「"&C{r}&"」不在【基础资料】②里"'),
            (f'C{r}=""', '"⚠ 没选业务板块（板块表里看不到这笔）"'),
            (f'ISNA(MATCH({esc(D)},单位_名称,0))', f'"⚠ 往来单位「"&{D}&"」不在【往来单位】里"'),
            (f'AND({F}<>"",ISNA(MATCH({esc(F)},项目_名称,0)))', f'"⚠ 业务内容「"&{F}&"」不在收支项目里"'),
            (f'AND(TRIM({ix("约定日期")}&"")<>"",NOT(ISNUMBER({ix("约定日期")})))', '"⚠ 约定日期不是真日期（按账期算了）"'),
            (f'AND(L{r}<>"",ISNA(MATCH({esc(f"L{r}")},经手人_姓名,0)))', f'"⚠ 经手人「"&L{r}&"」不在【基础资料】⑤里"'),
            (f'B{r}<P_建账日', '"⚠ 日期在建账日以前（老欠款就这么记，没问题）"'),
            (f'OR(YEAR(B{r})>YEAR(P_建账日)+10,YEAR(B{r})<2000)', '"⚠ 年份看着不对（打错了？）"'),
        ])
        for k, v in f.items():
            ws[f'{Wn[k]}{r}'] = v


BUILDERS = {H_PAR: build_par, H_LIST: build_list, H_SHOU: build_shou, H_WANG: build_wang}


def build(wb, ctx=None):
    for name, fn in BUILDERS.items():
        fn(wb[name])
