# -*- coding: utf-8 -*-
"""隐藏计算表：_参数、_表（基本信息逐行镜像）、_收（收支登记逐条）、_应（应收登记）、_付（应付登记）、_考（考勤）、_票（发票登记）。
   全部按「第 n 条」取数：INDEX(录入表!$X$表头行:$X$20000, n+1)，插行删行都不会错位。报表只用 layout.NAMES 里的定义名称。"""
from layout import *
from common import F_HELP, esc

GS = '"' + '","'.join(GUISHU) + '"'


def _hdr(ws, cols):
    for c, f, _d in cols:
        ws[f'{c}1'] = f
        ws[f'{c}1'].font = F_HELP


def chain(pairs, default='""'):
    """IF(c1,v1,IF(c2,v2,…,default))：括号自动配好"""
    out = default
    for c, v in reversed(pairs):
        out = f'IF({c},{v},{out})'
    return out


def _ix(sh, col, n_ref):
    """录入表第 n 条的某列"""
    return f'INDEX({inref(sh, col)},{n_ref}+1)'


def _mx(sh, col, n_ref):
    return f'INDEX({inref(sh, col)},{n_ref}+1)'


# ───────────────────────── _参数 ─────────────────────────
def build_par(ws):
    P = lambda k: f'{SH_BASE}!$C${PAR_ROW[k]}'
    rows = {
        '最后流水日': f'=IF(MAX(收_日期)>0,MAX(收_日期),IF(ISNUMBER({P("建账日")}),{P("建账日")},DATE(2026,1,1)))',
        '年度': f'=IF(ISNUMBER({SH_HOME}!${HOME_YEAR[0]}${HOME_YEAR[1:]}),INT({SH_HOME}!${HOME_YEAR[0]}${HOME_YEAR[1:]}),YEAR(B5))',
        '截止': f'=IF(ISNUMBER({SH_HOME}!${HOME_END[0]}${HOME_END[1:]}),INT({SH_HOME}!${HOME_END[0]}${HOME_END[1:]}),IF(B2=YEAR(B5),B5,DATE(B2,12,31)))',
        '年初': '=DATE(B2,1,1)',
        '建账日': f'=IF(ISNUMBER({P("建账日")}),{P("建账日")},DATE(2023,5,1))',
        '考勤起算': f'=IF(ISNUMBER({P("考勤起算")}),DATE(YEAR({P("考勤起算")}),MONTH({P("考勤起算")}),1),B6)',
        '起算年月': '=YEAR(B7)*100+MONTH(B7)',
        '收入口径': f'=IF({P("收入口径")}="","开票与收款取大",{P("收入口径")})',
        '人工口径': f'=IF({P("人工口径")}="","考勤应发",{P("人工口径")})',
        '分摊依据': f'=IF({P("分摊依据")}="","施工费",{P("分摊依据")})',
        '纳税人': f'=IF({P("纳税人")}="","一般纳税人",{P("纳税人")})',
        '销项税率': f'=IF(N({P("销项税率")})>0,{P("销项税率")},0.09)',
        '附加税率': f'=N({P("附加税率")})',
        '征收率': f'=IF(N({P("征收率")})>0,{P("征收率")},0.03)',
        '公司': f'={P("公司")}&""',
        '截止年月': '=YEAR(B3)*100+MONTH(B3)',
        '年初年月': '=B2*100+1',
        '税号': f'=TRIM({P("税号")}&"")',
    }
    for k, cell in PAR.items():
        r = int(cell[1:])
        ws[f'A{r}'] = k
        ws[cell] = rows[k]
    ws['A1'] = '全书共用的期间和参数（自动，别改）'


# ───────────────────────── _表 ─────────────────────────
def build_list(ws):
    ws['A1'] = '项目'
    for i in range(max(LIST_N, N_RATE)):
        r = i + 2
        n = i + 1
        if n <= N_RATE:
            C = LIST_RATE
            a = f'{C["姓名"]}{r}'
            g = lambda k: _mx(SH_RATE, RATE_COLS[k], n)
            ws[a] = f'=TRIM({g("姓名")}&"")'
            ws[f'{C["生效"]}{r}'] = f'=IF({a}="",0,IF(ISNUMBER({g("生效")}),INT({g("生效")}),1))'
            ws[f'{C["日单价"]}{r}'] = f'=IF({a}="",0,N({g("日单价")}))'
            ws[f'{C["月薪"]}{r}'] = f'=IF({a}="",0,N({g("月薪")}))'
            nm_, ef_ = C['姓名'], C['生效']
            ws[f'{C["序"]}{r}'] = (f'=IF({a}="",0,COUNTIFS(${nm_}$2:${nm_}${N_RATE + 1},{esc(a)},${ef_}$2:${ef_}${N_RATE + 1},"<"&{ef_}{r})'
                                  f'+COUNTIFS(${nm_}$2:{nm_}{r},{esc(a)},${ef_}$2:{ef_}{r},{ef_}{r}))')
        if i >= LIST_N:
            continue
        if n <= N_PJ:
            C = LIST_PJ
            nm = f'TRIM({_mx(SH_PJ, PJ_COLS["名称"], n)}&"")'
            ws[f'{C["名称"]}{r}'] = f'={nm}'
            g = lambda k: _mx(SH_PJ, PJ_COLS[k], n)
            a = f'{C["名称"]}{r}'
            ws[f'{C["类型"]}{r}'] = f'=IF({a}="","",IF(TRIM({g("类型")}&"")="","工程",TRIM({g("类型")}&"")))'
            ws[f'{C["甲方"]}{r}'] = f'=IF({a}="","",TRIM({g("甲方")}&""))'
            for k in ('合同', '保证金', '质保比例', '质保金额'):
                ws[f'{C[k]}{r}'] = f'=IF({a}="",0,N({g(k)}))'
            for k in ('质保到期', '开工', '完工'):
                ws[f'{C[k]}{r}'] = f'=IF({a}="",0,IF(ISNUMBER({g(k)}),INT({g(k)}),0))'
            ws[f'{C["状态"]}{r}'] = f'=IF({a}="","",TRIM({g("状态")}&""))'
            ws[f'{C["税率"]}{r}'] = f'=IF({a}="",0,IF(N({g("税率")})>0,N({g("税率")}),P_销项税率))'
            # 开工有效：填了用填的；没填＝建账日
            ws[f'{C["开工有效"]}{r}'] = f'=IF({a}="",0,IF({C["开工"]}{r}>0,{C["开工"]}{r},IF(ISNUMBER(P_建账日),INT(P_建账日),0)))'
        if n <= N_SUP:
            C = LIST_SUP
            a = f'{C["名称"]}{r}'
            ws[a] = f'=TRIM({_mx(SH_SUP, SUP_COLS["名称"], n)}&"")'
            ws[f'{C["类型"]}{r}'] = f'=IF({a}="","",IF(TRIM({_mx(SH_SUP, SUP_COLS["类型"], n)}&"")="","其他",TRIM({_mx(SH_SUP, SUP_COLS["类型"], n)}&"")))'
            ws[f'{C["税率"]}{r}'] = f'=IF({a}="",0,N({_mx(SH_SUP, SUP_COLS["税率"], n)}))'
            ws[f'{C["发票类型"]}{r}'] = f'=IF({a}="","",TRIM({_mx(SH_SUP, SUP_COLS["发票类型"], n)}&""))'
        if n <= N_PER:
            C = LIST_PER
            a = f'{C["姓名"]}{r}'
            ws[a] = f'=TRIM({_mx(SH_PER, PER_COLS["姓名"], n)}&"")'
            ws[f'{C["类别"]}{r}'] = f'=IF({a}="","",TRIM({_mx(SH_PER, PER_COLS["类别"], n)}&""))'
            ws[f'{C["计薪"]}{r}'] = f'=IF({a}="","",IF(TRIM({_mx(SH_PER, PER_COLS["计薪"], n)}&"")="月薪","月薪","日薪"))'
            ws[f'{C["过账"]}{r}'] = f'=IF(TRIM({_mx(SH_PER, PER_COLS["过账"], n)}&"")="是",1,0)'
            ws[f'{C["期初欠薪"]}{r}'] = f'=IF({a}="",0,N({_mx(SH_PER, PER_COLS["期初欠薪"], n)}))'
        if n <= N_ACC:
            C = LIST_ACC
            src = lambda col: f'INDEX({SH_BASE}!${col}${BASE_ACC_HDR}:${col}${BASE_ACC_R0 + N_ACC - 1},{n}+1)'
            a = f'{C["名称"]}{r}'
            ws[a] = f'=TRIM({src(ACC_COLS["名称"])}&"")'
            ws[f'{C["类型"]}{r}'] = f'=IF({a}="","",IF(TRIM({src(ACC_COLS["类型"])}&"")="","银行",TRIM({src(ACC_COLS["类型"])}&"")))'
            ws[f'{C["所属人"]}{r}'] = f'=IF({a}="","",TRIM({src(ACC_COLS["所属人"])}&""))'
            ws[f'{C["期初"]}{r}'] = f'=IF({a}="",0,N({src(ACC_COLS["期初"])}))'
        if n <= N_CUS:
            ws[f'{LIST_CUS["名称"]}{r}'] = f'=TRIM({_mx(SH_CUS, CUS_COLS["名称"], n)}&"")'
        if n <= N_INC:
            src = lambda col: f'INDEX({SH_BASE}!${col}${BASE_INC_HDR}:${col}${BASE_INC_R0 + N_INC - 1},{n}+1)'
            ws[f'{LIST_INC["名称"]}{r}'] = f'=TRIM({src(INC_COLS["名称"])}&"")'
            ws[f'{LIST_INC["归类"]}{r}'] = f'=TRIM({src(INC_COLS["归类"])}&"")'
        if n <= N_EXP:
            src = lambda col: f'INDEX({SH_BASE}!${col}${BASE_EXP_HDR}:${col}${BASE_EXP_R0 + N_EXP - 1},{n}+1)'
            ws[f'{LIST_EXP["名称"]}{r}'] = f'=TRIM({src(EXP_COLS["名称"])}&"")'
            ws[f'{LIST_EXP["归类"]}{r}'] = f'=TRIM({src(EXP_COLS["归类"])}&"")'
            cat_tab = f'{SH_BASE}!$M${BASE_CAT_R0}:$N${BASE_CAT_R0 + len(CATS) - 1}'
            ws[f'{LIST_EXP["归属"]}{r}'] = (f'=IF(TRIM({src(EXP_COLS["归属"])}&"")<>"",TRIM({src(EXP_COLS["归属"])}&""),'
                                            f'IFERROR(VLOOKUP({LIST_EXP["归类"]}{r},{cat_tab},2,FALSE)&"",""))')


# ───────────────────────── _收 ─────────────────────────
def build_shou(ws):
    _hdr(ws, SHOU_COLS)
    S = SHOU
    CC = CASH_COLS
    for i in range(N_CASH):
        r = i + 2
        n = f'$A{r}'
        ix = lambda k: _ix(SH_CASH, CC[k], n)
        f = {}
        f['n'] = i + 1
        f['日期'] = f'=IF(ISNUMBER({ix("日期")}),INT({ix("日期")}),0)'
        f['年月'] = f'=IF(B{r}=0,0,YEAR(B{r})*100+MONTH(B{r}))'
        for k in ('类别', '账户', '收支项目'):
            f[k] = f'=TRIM({ix(k)}&"")'
        f['摘要'] = f'={ix("摘要")}&""'
        f['净额原'] = f'=ROUND(N({ix("收入")})-N({ix("支出")}),2)'
        f['有效'] = f'=IF(AND(B{r}>0,E{r}<>"",AH{r}<>0),IF(ISNUMBER(MATCH(E{r},账户_名称,0)),1,0),0)'
        f['收入'] = f'=IF(S{r}=1,N({ix("收入")}),0)'
        f['支出'] = f'=IF(S{r}=1,N({ix("支出")}),0)'
        f['净额'] = f'=IF(S{r}=1,AH{r},0)'
        f['项目原'] = f'=TRIM({ix("项目")}&"")'
        f['供应商'] = f'=TRIM({ix("供应商")}&"")'
        f['人员原'] = f'=TRIM({ix("人员")}&"")'
        f['对方账户'] = f'=TRIM({ix("对方账户")}&"")'
        f['归属原'] = f'=TRIM({ix("费用归属")}&"")'
        f['已开票'] = f'=IF(S{r}=1,N({ix("已开票")}),0)'
        f['开票日期'] = f'=IF(ISNUMBER({ix("开票日期")}),INT({ix("开票日期")}),IF(Q{r}<>0,B{r},0))'
        f['开票年月'] = f'=IF(R{r}=0,0,YEAR(R{r})*100+MONTH(R{r}))'
        f['归类原'] = (f'=IF(F{r}="","",IF(D{r}="收入",IFERROR(INDEX(收入项目_归类,MATCH(F{r},收入项目_名称,0)),'
                     f'IFERROR(INDEX(支出项目_归类,MATCH(F{r},支出项目_名称,0)),"")),IFERROR(INDEX(支出项目_归类,MATCH(F{r},支出项目_名称,0)),'
                     f'IFERROR(INDEX(收入项目_归类,MATCH(F{r},收入项目_名称,0)),""))))')
        f['归类'] = (f'=IF(S{r}=0,"",IF(AND(AG{r}="内部转账",OR(O{r}="",O{r}=E{r},ISNA(MATCH(O{r},账户_名称,0)))),"",'
                   f'IF(ISNA(MATCH(AG{r},{{"' + '","'.join(CAT_NAMES) + f'"}},0)),"",AG{r})))')
        f['账户类型'] = f'=IFERROR(INDEX(账户_类型,MATCH(E{r},账户_名称,0))&"","")'
        f['账户人'] = f'=IFERROR(INDEX(账户_所属人,MATCH(E{r},账户_名称,0))&"","")'
        f['项目'] = f'=IF(K{r}="","",IF(IFERROR(INDEX(项目_类型,MATCH(K{r},项目_名称,0)),"")="工程",K{r},""))'
        f['客户'] = (f'=IF(TRIM({ix("客户")}&"")<>"",TRIM({ix("客户")}&""),IF(W{r}="","",'
                   f'IFERROR(INDEX(项目_甲方,MATCH(W{r},项目_名称,0))&"","")))')
        f['人员'] = (f'=IF(N{r}<>"",N{r},IF(AND(AG{r}="个人借款",ISNUMBER(FIND("-",F{r}))),TRIM(MID(F{r},FIND("-",F{r})+1,50)),""))')
        f['冲应付'] = (f'=IF(S{r}=1,IF(OR(T{r}="材料款",T{r}="分包款",T{r}="机械费"),IF(M{r}<>"",'
                    f'IF(COUNTIF(付_供应商,{esc(f"M{r}")})>0,1,0),0),0),0)')
        f['过账'] = (f'=IF(S{r}=1,IF(X{r}<>"",IF(OR(T{r}="工资发放",T{r}="代收代付"),'
                   f'IF(IFERROR(INDEX(人员_过账,MATCH(X{r},人员_姓名,0)),0)=1,1,0),0),0),0)')
        f['冲工资'] = (f'=IF(AND(S{r}=1,P_人工口径="考勤应发"),IF(T{r}="工资发放",IF(AA{r}=0,IF(X{r}<>"",IF(B{r}>=P_考勤起算,'
                    f'IF(COUNTIF(考_姓名,{esc(f"X{r}")})>0,1,0),0),0),0),0),0)')
        nature = (f'IF(T{r}="材料款","材料",IF(T{r}="分包款","分包",IF(T{r}="机械费","机械",IF(T{r}="工资发放","人工",'
                  f'IF(T{r}="项目其他费用","其他直接费",IF(T{r}="管理费用","管理费用",IF(T{r}="财务费用","财务费用",'
                  f'IF(T{r}="营业外支出","营业外支出",IF(T{r}="","未分类","")))))))))')
        f['成本类'] = f'=IF(S{r}=0,"",IF(OR(Y{r}=1,Z{r}=1,AA{r}=1),"",{nature}))'
        d0 = f'IFERROR(INDEX(支出项目_归属,MATCH(F{r},支出项目_名称,0))&"","")'
        g1 = f'IF(OR(P{r}="项目成本",P{r}="待摊费用",P{r}="公司费用"),P{r},IF({d0}<>"",{d0},IF(AB{r}="管理费用","待摊费用","项目成本")))'
        f['归属'] = (f'=IF(AB{r}="","",IF(OR(AB{r}="财务费用",AB{r}="营业外支出",AB{r}="未分类"),"公司费用",'
                   f'IF(AND({g1}="项目成本",W{r}=""),"待摊费用",{g1})))')
        f['往来对象'] = f'=IF(X{r}<>"",X{r},IF(M{r}<>"",M{r},L{r}))'
        f['基数'] = (f'=IF(AND(AC{r}="项目成本",W{r}<>""),IF(P_分摊依据="直接成本",-J{r},'
                   f'IF(OR(AB{r}="人工",AB{r}="分包",AB{r}="机械"),-J{r},0)),0)')
        f['摊费'] = f'=IF(AJ{r}=0,0,AJ{r}*{rate_of(f"B{r}")})'
        f['录入行'] = f'=ROW({_ix(SH_CASH, "A", n)})'
        blank = f'AND(B{r}=0,E{r}="",F{r}="",AH{r}=0,G{r}="")'
        f['校验'] = '=' + chain([
            (blank, '""'), (f'B{r}=0', '"✗ 没填日期（或不是真日期）"'), (f'E{r}=""', '"✗ 没选账户"'),
            (f'ISNA(MATCH(E{r},账户_名称,0))', f'"✗ 账户「"&E{r}&"」不在【基本信息】账户表里"'),
            (f'AH{r}=0', '"⚠ 没填金额"'), (f'F{r}=""', '"✗ 没选收支项目"'),
            (f'AG{r}=""', f'"✗ 收支项目「"&F{r}&"」不在【基本信息】③④里（或没选归类）"'),
            (f'AND(AG{r}="内部转账",T{r}="")', '"✗ 内部转账要填对方账户（另一个账户）"'),
            (f'AND(K{r}<>"",ISNA(MATCH(K{r},项目_名称,0)))', f'"✗ 项目「"&K{r}&"」不在【项目档案】里"'),
            (f'AND(T{r}="个人借款",X{r}="")', '"✗ 个人借款没填人员"'),
            (f'AND(T{r}="工程款收款",W{r}="")', '"⚠ 工程款没选项目（应收冲不到项目上）"'),
            (f'AND(M{r}<>"",ISNA(MATCH(M{r},供应商_名称,0)))', f'"⚠ 供应商「"&M{r}&"」不在【供应商信息】里"'),
            (f'AND(N{r}<>"",ISNA(MATCH(N{r},人员_姓名,0)))', f'"⚠ 人员「"&N{r}&"」不在【人员信息】里"'),
            (f'AND(OR(T{r}="报销还款",T{r}="备用金"),X{r}="")', '"⚠ 报销/备用金没填人员"'),
            (f'AND(N({ix("收入")})<>0,N({ix("支出")})<>0)', '"⚠ 收入、支出都填了"'),
        ])
        for k, v in f.items():
            ws[f'{S[k]}{r}'] = v


# ───────────────────────── _应 ─────────────────────────
def build_ys(ws):
    _hdr(ws, YS_COLS)
    Y = YS
    A = AR_COLS
    types = '{"' + '","'.join(AR_TYPES) + '"}'
    for i in range(N_AR):
        r = i + 2
        n = f'$A{r}'
        ix = lambda k: _ix(SH_AR_IN, A[k], n)
        f = {'n': i + 1}
        f['日期'] = f'=IF(ISNUMBER({ix("日期")}),INT({ix("日期")}),0)'
        f['年月'] = f'=IF(B{r}=0,0,YEAR(B{r})*100+MONTH(B{r}))'
        pj = f'TRIM({ix("项目")}&"")'
        f['项目'] = f'=IF({pj}="","",IF(IFERROR(INDEX(项目_类型,MATCH({pj},项目_名称,0)),"")="工程",{pj},""))'
        f['客户'] = f'=IF(TRIM({ix("客户")}&"")<>"",TRIM({ix("客户")}&""),IF(D{r}="","",IFERROR(INDEX(项目_甲方,MATCH(D{r},项目_名称,0))&"","")))'
        f['类型'] = f'=TRIM({ix("类型")}&"")'
        f['有效'] = f'=IF(AND(D{r}<>"",ISNUMBER(MATCH(F{r},{types},0)),N({ix("金额")})<>0),1,0)'
        f['金额'] = f'=IF(M{r}=1,N({ix("金额")}),0)'
        f['税率'] = f'=IF(N({ix("税率")})>0,N({ix("税率")}),IF(D{r}="",P_销项税率,IFERROR(INDEX(项目_税率,MATCH(D{r},项目_名称,0)),P_销项税率)))'
        f['应收额'] = f'=IF(OR(F{r}="补充协议",F{r}="签证",F{r}="结算调整"),G{r},IF(F{r}="扣款",-G{r},0))'
        f['开票额'] = f'=IF(F{r}="开票",G{r},0)'
        f['产值额'] = f'=IF(F{r}="确认产值",G{r},0)'
        f['销项税'] = f'=IF(J{r}=0,0,ROUND(J{r}*IF(P_纳税人="小规模纳税人",P_征收率,H{r})/(1+IF(P_纳税人="小规模纳税人",P_征收率,H{r})),2))'
        f['录入行'] = f'=ROW({_ix(SH_AR_IN, "A", n)})'
        f['校验'] = '=' + chain([
            (f'AND(TRIM({ix("项目")}&"")="",F{r}="",N({ix("金额")})=0)', '""'), (f'{pj}=""', '"✗ 没选项目"'),
            (f'D{r}=""', f'"✗ 项目「"&{pj}&"」不在【项目档案】里（或不是工程项目）"'),
            (f'ISNA(MATCH(F{r},{types},0))', '"✗ 类型要从下拉选"'), (f'N({ix("金额")})=0', '"⚠ 没填金额"'),
            (f'AND(OR(F{r}="开票",F{r}="确认产值"),B{r}=0)', '"✗ 开票/确认产值要填日期"'),
        ])
        for k, v in f.items():
            ws[f'{Y[k]}{r}'] = v


# ───────────────────────── _付 ─────────────────────────
def build_yf(ws):
    _hdr(ws, YF_COLS)
    Y = YF
    A = AP_COLS
    for i in range(N_AP):
        r = i + 2
        n = f'$A{r}'
        ix = lambda k: _ix(SH_AP_IN, A[k], n)
        f = {'n': i + 1}
        f['日期'] = f'=IF(ISNUMBER({ix("日期")}),INT({ix("日期")}),0)'
        f['年月'] = f'=IF(B{r}=0,0,YEAR(B{r})*100+MONTH(B{r}))'
        pj = f'TRIM({ix("项目")}&"")'
        f['项目'] = f'=IF({pj}="","",IF(IFERROR(INDEX(项目_类型,MATCH({pj},项目_名称,0)),"")="工程",{pj},""))'
        f['供应商'] = f'=TRIM({ix("供应商")}&"")'
        f['类型'] = (f'=IF(TRIM({ix("类型")}&"")<>"",TRIM({ix("类型")}&""),IF(E{r}="","",'
                   f'IFERROR(INDEX(供应商_类型,MATCH(E{r},供应商_名称,0))&"","其他")))')
        amt = f'IF(N({ix("应付")})<>0,N({ix("应付")}),ROUND(N({ix("数量")})*N({ix("单价")}),2))'
        f['有效'] = f'=IF(AND(E{r}<>"",OR({amt}<>0,N({ix("已开票")})<>0)),1,0)'
        f['应付额'] = f'=IF(N{r}=1,{amt},0)'
        f['已开票'] = f'=IF(N{r}=1,N({ix("已开票")}),0)'
        f['开票日期'] = f'=IF(ISNUMBER({ix("开票日期")}),INT({ix("开票日期")}),IF(H{r}<>0,B{r},0))'
        f['税率'] = f'=IF(N({ix("税率")})>0,N({ix("税率")}),IFERROR(INDEX(供应商_税率,MATCH(E{r},供应商_名称,0)),0))'
        f['发票类型'] = (f'=IF(TRIM({ix("发票类型")}&"")<>"",TRIM({ix("发票类型")}&""),'
                     f'IFERROR(INDEX(供应商_发票类型,MATCH(E{r},供应商_名称,0))&"",""))')
        f['进项税'] = f'=IF(AND(K{r}="专票",H{r}<>0,P_纳税人<>"小规模纳税人"),ROUND(H{r}*J{r}/(1+J{r}),2),0)'
        f['开票年月'] = f'=IF(I{r}=0,0,YEAR(I{r})*100+MONTH(I{r}))'
        f['录入行'] = f'=ROW({_ix(SH_AP_IN, "A", n)})'
        f['摘要'] = f'={ix("摘要")}&""'
        f['基数'] = (f'=IF(AND(N{r}=1,D{r}<>""),IF(P_分摊依据="直接成本",G{r},'
                   f'IF(OR(F{r}="分包",F{r}="机械"),G{r},0)),0)')
        f['摊费'] = f'=IF(R{r}=0,0,R{r}*{rate_of(f"B{r}")})'
        f['校验'] = '=' + chain([
            (f'AND(E{r}="",{pj}="",{amt}=0,N({ix("已开票")})=0)', '""'), (f'E{r}=""', '"✗ 没选供应商"'),
            (f'ISNA(MATCH(E{r},供应商_名称,0))', f'"✗ 供应商「"&E{r}&"」不在【供应商信息】里"'),
            (f'AND({pj}<>"",D{r}="")', f'"✗ 项目「"&{pj}&"」不在【项目档案】里"'), (f'B{r}=0', '"✗ 没填日期"'),
            (f'AND({amt}=0,N({ix("已开票")})=0)', '"⚠ 没填金额"'), (f'{pj}=""', '"⚠ 没选项目（成本进待摊）"'),
        ])
        for k, v in f.items():
            ws[f'{Y[k]}{r}'] = v


# ───────────────────────── _考 ─────────────────────────
def build_kq(ws):
    _hdr(ws, KQ_FIXED)
    for i in range(NPAIR):
        ws[f'{KQ_PJ[i]}1'], ws[f'{KQ_DAY[i]}1'], ws[f'{KQ_AMT[i]}1'] = f'项目{i + 1}', f'天数{i + 1}', f'金额{i + 1}'
        ws[f'{KQ_TAN[i]}1'] = f'摊{i + 1}'
    ws[f'{KQ_JT}1'], ws[f'{KQ_BASE}1'] = '计提', '基数'
    K = KQ
    for i in range(N_ATT):
        r = i + 2
        n = f'$A{r}'
        ix = lambda col: _ix(SH_ATT, col, n)
        f = {'n': i + 1}
        f['月'] = f'=IF(ISNUMBER({ix(ATT_MON)}),DATE(YEAR({ix(ATT_MON)}),MONTH({ix(ATT_MON)}),1),0)'
        f['年月'] = f'=IF(B{r}=0,0,YEAR(B{r})*100+MONTH(B{r}))'
        f['姓名'] = f'=TRIM({ix(ATT_NAME)}&"")'
        f['过账'] = f'=IF(D{r}="",0,IFERROR(INDEX(人员_过账,MATCH(D{r},人员_姓名,0)),0))'
        f['计薪'] = f'=IF(D{r}="","",IFERROR(INDEX(人员_计薪,MATCH(D{r},人员_姓名,0))&"","日薪"))'
        f['总天数'] = '=' + '+'.join(f'{KQ_DAY[j]}{r}' for j in range(NPAIR))
        f['有效'] = f'=IF(AND(B{r}>0,D{r}<>""),1,0)'
        f['人月天数'] = f'=IF(K{r}=0,0,SUMIFS(考_总天数,考_姓名,{esc(f"D{r}")},考_年月,C{r}))'
        me = f'DATE(YEAR(B{r}),MONTH(B{r})+1,0)'
        std = std_rate(f'D{r}', me, f'F{r}="月薪"')
        f['单价'] = f'=IF(K{r}=0,0,IF(N({ix(ATT_RATE)})<>0,N({ix(ATT_RATE)}),IFERROR({std},0)))'
        cnt = f'COUNTIFS(考_姓名,{esc(f"D{r}")},考_年月,C{r})'
        gross = (f'IF(F{r}="月薪",I{r}*IF(H{r}>0,G{r}/H{r},1/MAX(1,{cnt})),I{r}*G{r})'
                 f'+N({ix(ATT_EXTRA)})-N({ix(ATT_DED)})')
        f['应发'] = f'=IF(OR(K{r}=0,E{r}=1),0,ROUND({gross},2))'
        f['过账应发'] = f'=IF(AND(K{r}=1,E{r}=1),ROUND({gross},2),0)'
        # 8 对：项目（原样）、天数、金额（只给工程项目按天数分；其余进未分摊）
        for j in range(NPAIR):
            pjc, dyc, amc = KQ_PJ[j], KQ_DAY[j], KQ_AMT[j]
            f[f'_p{j}'] = (pjc, f'=TRIM({ix(ATT_PJS[j])}&"")')
            f[f'_d{j}'] = (dyc, f'=N({ix(ATT_DAYS[j])})')
            f[f'_a{j}'] = (amc, f'=IF(OR(J{r}=0,G{r}=0,{pjc}{r}=""),0,IF(IFERROR(INDEX(项目_类型,MATCH({pjc}{r},项目_名称,0)),"")="工程",'
                                f'J{r}*{dyc}{r}/G{r},0))')
        f['未分摊'] = f'=J{r}-(' + '+'.join(f'{KQ_AMT[j]}{r}' for j in range(NPAIR)) + ')'
        f['计提'] = f'=IF(AND(K{r}=1,E{r}=0,C{r}>=P_起算年月,P_人工口径="考勤应发"),1,0)'
        f['基数'] = f'={KQ_JT}{r}*(' + '+'.join(f'{KQ_AMT[j]}{r}' for j in range(NPAIR)) + ')'
        for j in range(NPAIR):
            f[f'_t{j}'] = (KQ_TAN[j], f'=IF({KQ_JT}{r}*{KQ_AMT[j]}{r}=0,0,{KQ_AMT[j]}{r}*{rate_of(f"B{r}")})')
        f['录入行'] = f'=ROW({_ix(SH_ATT, "A", n)})'
        f['校验'] = '=' + chain([
            (f'AND(D{r}="",B{r}=0,G{r}=0)', '""'), (f'D{r}=""', '"✗ 没填姓名"'), (f'B{r}=0', '"✗ 月份不是真日期"'),
            (f'ISNA(MATCH(D{r},人员_姓名,0))', f'"⚠ 「"&D{r}&"」不在【人员信息】里（按日薪算）"'),
            (f'AND(I{r}=0,E{r}=0)', '"✗ 没有单价（【工资标准】里没有这个人这个月有效的单价，也没手填）"'),
            (f'G{r}>31', f'"⚠ 一个月填了 "&G{r}&" 天"'),
        ] + [(f'AND({KQ_PJ[j]}{r}<>"",ISNA(MATCH({KQ_PJ[j]}{r},项目_名称,0)))', f'"✗ 项目「"&{KQ_PJ[j]}{r}&"」不在【项目档案】里"')
             for j in range(NPAIR)])
        for k, v in f.items():
            if k.startswith('_'):
                col, val = v
                ws[f'{col}{r}'] = val
            else:
                ws[f'{K[k]}{r}'] = v


# ───────────────────────── _票（发票登记） ─────────────────────────
def build_piao(ws):
    _hdr(ws, PIAO_COLS)
    from common import date_parse, num
    P = PIAO
    I = INV_COLS
    for i in range(N_INV):
        r = i + 2
        n = f'$A{r}'
        ix = lambda col: _ix(SH_INV, col, n)
        f = {'n': i + 1}
        f['日期'] = f'=IFERROR(N({date_parse(ix(I["日期"]))}),0)'
        f['年月'] = f'=IF(B{r}=0,0,YEAR(B{r})*100+MONTH(B{r}))'
        sell_tax, sell = f'TRIM({ix(I["销方税号"])}&"")', f'TRIM({ix(I["销方"])}&"")'
        f['方向'] = (f'=IF(OR(AND(P_税号<>"",{sell_tax}=P_税号),AND(P_公司<>"",{sell}=P_公司)),"销项","进项")')
        f['对方'] = f'=IF(D{r}="销项",TRIM({ix(I["购方"])}&""),{sell})'
        f['金额'] = f'={num(ix(I["金额"]))}'
        f['税额'] = f'={num(ix(I["税额"]))}'
        f['价税合计'] = f'=IF({num(ix(I["价税合计"]))}<>0,{num(ix(I["价税合计"]))},F{r}+G{r})'
        f['税率'] = f'=IF(F{r}=0,0,ROUND(G{r}/F{r},2))'
        f['项目'] = f'=TRIM({ix(I["项目"])}&"")'
        st = f'{ix(I["状态"])}&""'
        f['有效'] = f'=IF(AND(B{r}>0,H{r}<>0,ISERROR(FIND("作废",{st}))),1,0)'
        f['录入行'] = f'=ROW({_ix(SH_INV, "A", n)})'
        for k, v in f.items():
            ws[f'{P[k]}{r}'] = v
        for k in ('金额', '税额', '价税合计'):
            ws[f'{P[k]}{r}'] = f'=IF(K{r}=1,{ws[f"{P[k]}{r}"].value[1:]},0)' if k != '价税合计' else ws[f'{P[k]}{r}'].value
    # 价税合计、金额的「有效」依赖 B/H，避免循环：有效只看 B 和 价税合计原值
    for i in range(N_INV):
        r = i + 2
        n = f'$A{r}'
        st = f'{_ix(SH_INV, I["状态"], n)}&""'
        raw = f'IF({num(_ix(SH_INV, I["价税合计"], n))}<>0,{num(_ix(SH_INV, I["价税合计"], n))},{num(_ix(SH_INV, I["金额"], n))}+{num(_ix(SH_INV, I["税额"], n))})'
        ws[f'{P["有效"]}{r}'] = f'=IF(AND(B{r}>0,{raw}<>0,ISERROR(FIND("作废",{st}))),1,0)'
        ws[f'{P["价税合计"]}{r}'] = f'=IF(K{r}=1,{raw},0)'
        ws[f'{P["金额"]}{r}'] = f'=IF(K{r}=1,{num(_ix(SH_INV, I["金额"], n))},0)'
        ws[f'{P["税额"]}{r}'] = f'=IF(K{r}=1,{num(_ix(SH_INV, I["税额"], n))},0)'


BUILDERS = {H_PAR: build_par, H_LIST: build_list, H_SHOU: build_shou, H_YS: build_ys, H_YF: build_yf, H_KQ: build_kq,
            H_PIAO: build_piao}


def build(wb, ctx=None):
    for name, fn in BUILDERS.items():
        fn(wb[name])
