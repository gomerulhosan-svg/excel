# -*- coding: utf-8 -*-
"""【数据校验】：两张登记表逐行问题（✗ 要改、⚠ 要看）、基础资料重名、对数检查，最后列出有问题的行（表名＋行号＋问题）。
   首页「提醒」引用 CHK_X / CHK_W / CHECK_ROWS。"""
from openpyxl.formatting.rule import FormulaRule
from layout import *
from common import *

CHK_X, CHK_W = 'C4', 'F4'
CHECK_ROWS = {}


def _over(sh, hdr, cap, col='A'):
    """录入表第 cap 条以后还有没有东西（区域从表头行锚定，插删行跟着动）"""
    return (f'COUNTA(INDEX({sh}!${col}:${col},ROW({sh}!${col}${hdr})+{cap + 1}):INDEX({sh}!${col}:${col},1048576))')


def _over_base():
    parts = []
    for i, (key, ttl, tr, hr, cap) in enumerate(BASE_BLOCKS):
        a = f'INDEX({SH_BASE}!$A:$A,ROW({SH_BASE}!$A${hr})+{cap + 1})'
        if i + 1 < len(BASE_BLOCKS):
            nxt = BASE_BLOCKS[i + 1][1]
            b = f'INDEX({SH_BASE}!$A:$A,MAX(ROW({SH_BASE}!$A${hr})+{cap + 1},IFERROR(MATCH("{nxt}",{SH_BASE}!$A:$A,0)-1,0)))'
        else:
            b = f'INDEX({SH_BASE}!$A:$A,ROW({SH_BASE}!$A${hr})+{cap + 300})'
        parts.append(f'COUNTA({a}:{b})')
    return '=' + '+'.join(parts)


CAP_CASH = '=' + _over(SH_CASH, CASH_HDR, N_CASH)
CAP_WL = '=' + _over(SH_WL, WL_HDR, N_WL)
CAP_UNIT = '=' + _over(SH_UNIT, UNIT_HDR, N_UNIT)
CAP_BASE = _over_base()
ARR_DIR = '{"' + '","'.join(ITEM_DIRS) + '"}'
ARR_USE = '{"' + '","'.join(ITEM_USES) + '"}'
ARR_STK = '{"' + '","'.join(ITEM_STOCK) + '"}'


def build(wb, ctx=None):
    ws = wb[SH_CHK]
    widths(ws, {'A': 6, 'B': 48, 'C': 14, 'D': 7, 'E': 12, 'F': 14, 'G': 52})
    title(ws, '数 据 校 验', 'G', C_CHK,
          '💡 录完数看一眼这张：✗＝一定要改（不改报表会少算），⚠＝看一下对不对。先看每张表有几条问题，再看对数检查，'
          '最后按「表＋行号」回到那张表找到改掉。这张表全是自动的。')
    home_link(ws, 'G3')
    put(ws, 'B4', '要改的（✗）合计', F_KPI_L, fill('FFD9E1F2'), align=AC)
    put(ws, 'E4', '要看的（⚠）合计', F_KPI_L, fill('FFD9E1F2'), align=AC)
    xs, wsum = [], []

    # ① 两张登记表
    r = 6
    section(ws, r, 'A', 'G', '① 两张登记表逐行检查（每行最右边「这一行的问题」汇总）', C_CHK)
    header(ws, r + 1, [('A', '#'), ('B', '表'), ('C', '✗ 要改'), ('D', ''), ('E', '⚠ 要看'), ('F', ''), ('G', '常见原因')], C_CHK)
    for i, (sh, nm, why) in enumerate([(SH_CASH, '收_校验', '没填日期或金额、收支项目/板块/账户不在基础资料里、收回/支付欠款没选往来单位'),
                                        (SH_WL, '往_校验', '没填日期/金额、没选往来单位或应收应付、板块不在基础资料里')]):
        rr = r + 2 + i
        put(ws, f'A{rr}', i + 1, F_TXT, align=AC)
        link(put(ws, f'B{rr}', sh, Font(name=YH, sz=10, bold=True, color='FF0563C1', underline='single'), align=AL), sh)
        put(ws, f'C{rr}', f'=COUNTIF({nm},"✗*")', F_AUTOB, fmt='0', align=AC)
        put(ws, f'D{rr}', f'=IF(C{rr}>0,"✗","✓")', F_TXTB, align=AC)
        put(ws, f'E{rr}', f'=COUNTIF({nm},"⚠*")', F_AUTOB, fmt='0', align=AC)
        put(ws, f'F{rr}', f'=IF(E{rr}>0,"⚠","✓")', F_TXTB, align=AC)
        put(ws, f'G{rr}', why, F_NOTE, align=ALW)
        xs.append(f'C{rr}')
        wsum.append(f'E{rr}')
        CHECK_ROWS[sh] = rr
    r += 5

    # ② 基础资料（隐藏列逐行算重名）
    section(ws, r, 'A', 'G', '② 基础资料检查', C_CHK)
    header(ws, r + 1, [('A', '#'), ('B', '检查'), ('C', '个数'), ('D', '状态'), ('G', '怎么改')], C_CHK)

    def helper(col, n, f):
        for i in range(n):
            ws[f'{col}{i + 1}'] = f(i + 1)
            ws[f'{col}{i + 1}'].font = F_HELP
        hide(ws, col)
        return f'${col}$1:${col}${n}'

    dup = lambda nm: (lambda k: f'=IF(INDEX({nm},{k})="",0,IF(COUNTIF({nm},{esc(f"INDEX({nm},{k})")})>1,1,0))')
    h = {'板块': helper('Z', N_SEG, dup('板块_名称')), '账户': helper('AA', N_ACC, dup('账户_名称')),
         '项目': helper('AB', N_ITEM, dup('项目_名称')), '单位': helper('AC', N_UNIT, dup('单位_名称')),
         '属性': helper('AD', N_ITEM, lambda k: (
             f'=IF(INDEX(项目_名称,{k})="",0,IF(OR(ISNA(MATCH(INDEX(项目_方向,{k}),{ARR_DIR},0)),'
             f'ISNA(MATCH(INDEX(项目_用途,{k}),{ARR_USE},0)),'
             f'AND(INDEX(项目_库存,{k})<>"",ISNA(MATCH(INDEX(项目_库存,{k}),{ARR_STK},0)))),1,0))'))}
    checks = [('业务板块名称重复', f'=SUM({h["板块"]})', '✗', '同一个板块只留一行'),
              ('账户名称重复', f'=SUM({h["账户"]})', '✗', '同一个账户只留一行'),
              ('收支项目名称重复', f'=SUM({h["项目"]})', '✗', '同一个收支项目只留一行'),
              ('往来单位名称重复', f'=SUM({h["单位"]})', '✗', '同一个单位只留一行（重复了数会算到一起）'),
              ('收支项目的方向/用途/库存填得不认识', f'=SUM({h["属性"]})', '✗', '从下拉选：方向 收入/支出/双向；用途 普通/冲应收/冲应付/固定资产/内部转账；库存 入库/出库')]
    r += 2
    for i, (lab, f, kind, how) in enumerate(checks):
        rr = r + i
        put(ws, f'A{rr}', i + 1, F_TXT, align=AC)
        put(ws, f'B{rr}', lab, F_TXT, align=AL)
        put(ws, f'C{rr}', f, F_AUTOB, fmt='0', align=AC)
        put(ws, f'D{rr}', f'=IF(C{rr}>0,"{kind}","✓")', F_TXTB, align=AC)
        put(ws, f'G{rr}', how, F_NOTE, align=ALW)
        (xs if kind == '✗' else wsum).append(f'IF(C{rr}>0,1,0)')
        CHECK_ROWS[lab] = rr
    r += len(checks) + 1

    # ③ 对数检查（截至首页截止日）
    section(ws, r, 'A', 'G', '③ 对数检查（截至首页截止日）', C_CHK)
    header(ws, r + 1, [('A', '#'), ('B', '检查'), ('C', '金额/个数'), ('D', '状态'), ('G', '说明')], C_CHK)
    d = 'P_截止'
    acc_total = f'(SUM(账户_期初余额)+SUMIFS(收_净额,收_日期,"<="&{d}))'
    seg_total = f'(SUM(板块_期初结余)+SUM($AI$1:$AI${N_SEG}))'
    for i in range(N_SEG):                          # 名单里每个板块：收支合计（AG）、赊账合计（AH）、到截止日的净额（AI）
        rr = i + 1
        sg = f'INDEX(板块_名称,{rr})'
        ws[f'AG{rr}'] = f'=IF({sg}="",0,SUMIFS(收_板块收入,收_板块,{esc(sg)})+SUMIFS(收_板块支出,收_板块,{esc(sg)}))'
        ws[f'AH{rr}'] = f'=IF({sg}="",0,SUMIFS(往_应收额,往_板块,{esc(sg)})+SUMIFS(往_应付额,往_板块,{esc(sg)}))'
        ws[f'AI{rr}'] = (f'=IF({sg}="",0,SUMIFS(收_板块收入,收_板块,{esc(sg)},收_日期,"<="&{d})'
                         f'-SUMIFS(收_板块支出,收_板块,{esc(sg)},收_日期,"<="&{d}))')
        for c in ('AG', 'AH', 'AI'):
            ws[f'{c}{rr}'].font = F_HELP
    hide(ws, 'AG', 'AH', 'AI')
    num = [
        ('各账户余额合计 − 各板块结余合计', f'=ROUND({acc_total}-{seg_total},2)', 'ABS(C{r})>=0.01', '⚠',
         '不为 0：基础资料里账户期初余额合计和板块期初结余合计不一样，或者有收支没选板块（板块名不对），或者内部转账只记了一边'),
        ('内部转账没配对（一出一进合计不为 0）', '=ROUND(SUMIFS(收_净额,收_内部转账,1),2)', 'ABS(C{r})>=0.01', '⚠',
         '现金存银行、银行取现要记两行（一行支出、一行收入），金额一样'),
        ('收支没选业务板块或板块名不对的金额（收入＋支出，不含内部转账）',
         f'=ROUND(SUM(收_板块收入)+SUM(收_板块支出)-SUM($AG$1:$AG${N_SEG}),2)', 'ABS(C{r})>=0.01', '⚠',
         '这些钱不进任何板块表；在收支登记里补上业务板块（改过板块名的，用查找替换把旧名字改掉）'),
        ('应收应付没选业务板块或板块名不对的金额', f'=ROUND(SUM(往_应收额)+SUM(往_应付额)-SUM($AH$1:$AH${N_SEG}),2)',
         'ABS(C{r})>=0.01', '⚠', '这些赊账不进任何板块表'),
        ('应收为负（多收了）的往来单位家数', None, 'C{r}>0', '⚠', '收回的比赊出去的多：可能是预收款，或者收款记错了单位'),
        ('应付为负（多付了）的往来单位家数', None, 'C{r}>0', '⚠', '付出的比赊进来的多：可能是预付款，或者付款记错了单位'),
        ('收支登记超过 5000 笔（多出来的不算）', CAP_CASH, 'C{r}>0', '✗', '一本账最多 5000 笔收支：建议一年一本，年底另起一本（期初余额填上一年的期末）'),
        ('应收应付登记超过 2000 笔（多出来的不算）', CAP_WL, 'C{r}>0', '✗', '同上'),
        ('往来单位超过 500 个（多出来的不算）', CAP_UNIT, 'C{r}>0', '✗', '删掉不用的单位'),
        ('基础资料某一块超了（多出来的不算）', CAP_BASE, 'C{r}>0', '✗', '业务板块最多 15 个、账户 20 个、收支项目 120 个、经手人 50 个；块里不要留太多空行'),
        ('基础资料没填建账日期', f'=IF(ISNUMBER({SH_BASE}!$C${PAR_ROW["建账日"]}),0,1)', 'C{r}>0', '⚠', '没填就按最早一笔算；填上更清楚（期初余额都是这一天的）'),
        ('应收已逾期金额', f'=ROUND(SUMIFS(往_未结,往_类型,"应收",往_逾期天数,">0"),2)', 'C{r}>0', '⚠', '看【应收应付跟进】，该催款了'),
        ('应付已逾期金额', f'=ROUND(SUMIFS(往_未结,往_类型,"应付",往_逾期天数,">0"),2)', 'C{r}>0', '⚠', '看【应收应付跟进】'),
    ]
    # 按单位算余额（隐藏列 AE：应收余额，AF：应付余额）
    for i in range(N_UNIT):
        rr = i + 1
        u = f'INDEX(单位_名称,{rr})'
        ws[f'AE{rr}'] = (f'=IF({u}="",0,IF(SUMIFS(往_应收额,往_往来单位,{esc(u)},往_日期,"<="&{d})'
                         f'-SUMIFS(收_冲应收,收_往来单位,{esc(u)},收_日期,"<="&{d})<-0.005,1,0))')
        ws[f'AF{rr}'] = (f'=IF({u}="",0,IF(SUMIFS(往_应付额,往_往来单位,{esc(u)},往_日期,"<="&{d})'
                         f'-SUMIFS(收_冲应付,收_往来单位,{esc(u)},收_日期,"<="&{d})<-0.005,1,0))')
        ws[f'AE{rr}'].font = ws[f'AF{rr}'].font = F_HELP
    hide(ws, 'AE', 'AF')
    r += 2
    for i, (lab, f, bad, kind, how) in enumerate(num):
        rr = r + i
        if f is None:
            f = f'=SUM($AE$1:$AE${N_UNIT})' if '应收' in lab else f'=SUM($AF$1:$AF${N_UNIT})'
        put(ws, f'A{rr}', i + 1, F_TXT, align=AC)
        put(ws, f'B{rr}', lab, F_TXT, align=ALW)
        put(ws, f'C{rr}', f, F_AUTOB, fmt='0' if ('家数' in lab or '超' in lab or '没填' in lab) else MONEY, align=AR)
        put(ws, f'D{rr}', f'=IF({bad.format(r=rr)},"{kind}","✓")', F_TXTB, align=AC)
        put(ws, f'G{rr}', how, F_NOTE, align=ALW)
        ws.row_dimensions[rr].height = 30
        (xs if kind == '✗' else wsum).append(f'IF({bad.format(r=rr)},1,0)')
        CHECK_ROWS[lab] = rr
    r += len(num) + 1
    put(ws, CHK_X, '=' + '+'.join(xs), F_KPI_V, fill('FFFFFFFF'), '0', AC)
    put(ws, CHK_W, '=' + '+'.join(wsum), F_KPI_V, fill('FFFFFFFF'), '0', AC)

    # ④ 有问题的行
    section(ws, r, 'A', 'G', '④ 有问题的行（回到那张表按行号找；只列 ✗ 和 ⚠）', C_CHK)
    header(ws, r + 1, [('A', '第几条'), ('B', '问题'), ('C', '表'), ('D', '行号'), ('E', '日期'), ('F', '金额'), ('G', '摘要 / 往来单位')], C_CHK)
    r += 2
    SHOW = 150
    lists = [(SH_CASH, N_CASH, '收_校验', '收_录入行', '收_日期', 'INDEX(收_收入,{k})-INDEX(收_支出,{k})',
              'INDEX(收_板块,{k})&"｜"&INDEX(收_显示摘要,{k})', 'AH'),
             (SH_WL, N_WL, '往_校验', '往_录入行', '往_日期', 'INDEX(往_金额,{k})',
              'INDEX(往_往来单位,{k})&"｜"&INDEX(往_显示摘要,{k})', 'AI')]
    for sh, n, chk, row_nm, date_nm, amt, desc, cc in lists:
        counter(ws, cc, 1, n, lambda i, chk=chk: f'OR(LEFT(INDEX({chk},{i + 1}),1)="✗",LEFT(INDEX({chk},{i + 1}),1)="⚠")')
        hide(ws, cc)
        total = f'${cc}${n}'
        put(ws, f'B{r}', f'="{sh}：共 "&{total}&" 行有问题"&IF({total}>{SHOW},"（只列前 {SHOW} 行）","")', F_TXTB, FILL_SUB, align=AL)
        ws.merge_cells(f'B{r}:G{r}')
        r += 1
        for k in range(1, SHOW + 1):
            rr = r + k - 1
            ix = f'$A{rr}'
            ws[f'A{rr}'] = f'=IF({kth(k, cc, 1, n)}=0,"",{kth(k, cc, 1, n)})'
            ws[f'A{rr}'].font = F_HELP
            g = lambda e: f'=IF({ix}="","",{e.format(k=ix)})'
            put(ws, f'B{rr}', g(f'INDEX({chk},{{k}})'), F_TXT, align=AL)
            put(ws, f'C{rr}', f'=IF({ix}="","","{sh}")', F_TXT, align=AC)
            put(ws, f'D{rr}', g(f'INDEX({row_nm},{{k}})'), F_TXTB, fmt='0', align=AC)
            put(ws, f'E{rr}', g(f'IF(INDEX({date_nm},{{k}})=0,"",INDEX({date_nm},{{k}}))'), F_TXT, fmt=DATE, align=AC)
            put(ws, f'F{rr}', g(amt), F_TXT, fmt=MONEY, align=AR)
            put(ws, f'G{rr}', g(desc), F_NOTE, align=AL)
        r += SHOW + 1
    for col in ('D', 'F'):
        ws.conditional_formatting.add(f'{col}1:{col}{r}', FormulaRule(formula=[f'{col}1="✗"'], font=F_RED))
        ws.conditional_formatting.add(f'{col}1:{col}{r}', FormulaRule(formula=[f'{col}1="⚠"'],
                                                                     font=Font(name=YH, sz=10, bold=True, color='FFC65911')))
    ws.freeze_panes = 'A5'
    print_setup(ws, '1:4', landscape=False)
