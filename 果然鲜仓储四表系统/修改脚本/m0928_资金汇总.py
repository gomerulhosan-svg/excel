# -*- coding: utf-8 -*-
"""《03》资金与各类汇总：
   ① 资金与往来余额表：改回原来两段式（资金账户结存＋往来单位余额），带合计和差额核对，加年份/月份/起止日选择；
   ② 收入月度汇总改成「年份＋起止日」，新增 支出月度汇总 / 费用月度支出汇总 / 家用月度支出汇总（按月矩阵）；
   ③ 费用月度明细汇总改成自带时段选择＋合计，新增 家用月度明细汇总、国外费用汇总、借款汇总；
   ④ 删掉 月度资金明细，资金日记账里只给它和旧汇总用的辅助列清掉。"""
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font
from common0928 import *

JD = lambda c: f'资金日记账!${c}$1:${c}${JR1}'      # INDEX 用（从第 1 行起）
GREEN_TAB = '70AD47'


def jr(c, r):
    return f'资金日记账!${c}{r}'


# ─────────────────────────────── ① 资金与往来余额表 ───────────────────────────────
def fund_balance(wb):
    ws = wb['资金与往来余额表']
    wipe(ws)
    title(ws, '资 金 账 户 与 往 来 单 位 余 额 表（全部自动计算）', 'A', 'L')
    p = period(ws, 2, 'A', month=True, year=2026, mon=9)
    s, e = p['start'], p['end']
    note(ws, '只填第 2 行：年份＋月份＝查某个月；月份留空＝查全年；起始日/截止日填了就按起止日。'
             '期初＝建账期初＋开始日之前的收支，期末＝期初＋本期收入－本期支出。', 'A', 'L', 3, 24)
    section(ws, '一、资金账户结存（取自【资金日记账】）', 'A', 'I', 4)
    header(ws, 5, 'A', ['资金账户', '期初余额', '本期收入', '本期支出', '期末余额', '其中：互转转入', '其中：互转转出'], 26)
    A0, A1 = 6, 25
    for i, r in enumerate(range(A0, A1 + 1)):
        a = f'$A{r}'
        cell(ws, f'A{r}', f'=IF(科目表!$K${4 + i}="","",科目表!$K${4 + i})', F_AUTOB, FL_AUTO, ACN)
        cell(ws, f'B{r}', f'=IF({a}="","",ROUND(SUMIF(期初余额!$H$4:$H$43,{a},期初余额!$I$4:$I$43)'
                          f'+SUMIFS({J("H")},{J("C")},{a},{J("B")},"<"&{s})-SUMIFS({J("I")},{J("C")},{a},{J("B")},"<"&{s}),2))',
             al=AR, fmt=MONEY)
        cell(ws, f'C{r}', f'=IF({a}="","",ROUND(SUMIFS({J("H")},{J("C")},{a},{in_period(J("B"), p)}),2))', al=AR, fmt=MONEY)
        cell(ws, f'D{r}', f'=IF({a}="","",ROUND(SUMIFS({J("I")},{J("C")},{a},{in_period(J("B"), p)}),2))', al=AR, fmt=MONEY)
        cell(ws, f'E{r}', f'=IF({a}="","",ROUND(B{r}+C{r}-D{r},2))', al=AR, fmt=MONEY)
        cell(ws, f'F{r}', f'=IF({a}="","",ROUND(SUMIFS({J("H")},{J("C")},{a},{J("F")},"账户互转",{in_period(J("B"), p)}),2))',
             al=AR, fmt=MONEY)
        cell(ws, f'G{r}', f'=IF({a}="","",ROUND(SUMIFS({J("I")},{J("C")},{a},{J("F")},"账户互转",{in_period(J("B"), p)}),2))',
             al=AR, fmt=MONEY)
    T = A1 + 1
    cell(ws, f'A{T}', '合　计', F_TOT, FL_TOT, ACN)
    for c in 'BCDEFG':
        cell(ws, f'{c}{T}', f'=ROUND(SUM({c}{A0}:{c}{A1}),2)', F_TOTN, FL_TOT, AR, MONEY)
    D = T + 1
    ws.merge_cells(f'A{D}:D{D}')
    cell(ws, f'A{D}', '账面货币资金（记账凭证 · 截至截止日）', F_TOT, FL_NONE, AL)
    for c in 'BCD':
        ws[f'{c}{D}'].border = BOX
    cash = '{"库存现金","银行存款","其他货币资金"}'
    cell(ws, f'E{D}', f'=ROUND(SUMPRODUCT(SUMIFS(期初余额!$D$4:$D$81,期初余额!$B$4:$B$81,{cash}))'
                      f'-SUMPRODUCT(SUMIFS(期初余额!$E$4:$E$81,期初余额!$B$4:$B$81,{cash}))'
                      f'+SUMIFS(记账凭证!$H$4:$H$3609,记账凭证!$G$4:$G$3609,"货币资金",记账凭证!$B$4:$B$3609,"<="&{e})'
                      f'-SUMIFS(记账凭证!$I$4:$I$3609,记账凭证!$G$4:$G$3609,"货币资金",记账凭证!$B$4:$B$3609,"<="&{e}),2)',
         F_TOTN, FL_AUTO, AR, MONEY)
    ws.merge_cells(f'F{D}:I{D}')
    cell(ws, f'F{D}', f'=IF(ROUND($E{D}-$E{T},2)=0,"✔ 与日记账一致","✘ 差额 "&TEXT($E{D}-$E{T},"#,##0.00")'
                      f'&"（多半是【期初余额】右侧资金账户期初没同步到左侧科目期初）")', F_CHK, FL_NONE, AL)
    for c in 'GHI':
        ws[f'{c}{D}'].border = BOX
    K1, K2 = D + 1, D + 2
    ext_h = f'SUMIFS({J("H")},{J("F")},"<>账户互转",{in_period(J("B"), p)})'
    ext_i = f'SUMIFS({J("I")},{J("F")},"<>账户互转",{in_period(J("B"), p)})'
    trn_h = f'SUMIFS({J("H")},{J("F")},"账户互转",{in_period(J("B"), p)})'
    trn_i = f'SUMIFS({J("I")},{J("F")},"账户互转",{in_period(J("B"), p)})'
    kp = [(K1, '本期外部收入（不含互转）', f'=ROUND({ext_h},2)', '本期外部支出（不含互转）', f'=ROUND({ext_i},2)',
           '外部收支净额', f'=ROUND(B{K1}-D{K1},2)'),
          (K2, '账户互转 · 转入', f'=ROUND({trn_h},2)', '账户互转 · 转出', f'=ROUND({trn_i},2)',
           '互转核对', f'=IF(ROUND(B{K2}-D{K2},2)=0,"✔ 转入＝转出","✘ 差 "&TEXT(B{K2}-D{K2},"#,##0.00"))')]
    for r, l1, v1, l2, v2, l3, v3 in kp:
        cell(ws, f'A{r}', l1, F_LBL, FL_LBL, AC)
        cell(ws, f'B{r}', v1, F_AUTOB, FL_AUTO, AR, MONEY)
        cell(ws, f'C{r}', l2, F_LBL, FL_LBL, AC)
        cell(ws, f'D{r}', v2, F_AUTOB, FL_AUTO, AR, MONEY)
        cell(ws, f'E{r}', l3, F_LBL, FL_LBL, AC)
        ws.merge_cells(f'F{r}:G{r}')
        cell(ws, f'F{r}', v3, F_AUTOB if r == K1 else F_CHK, FL_AUTO, AR if r == K1 else AC, MONEY)
        ws[f'G{r}'].border = BOX
        ws.row_dimensions[r].height = 24
    # 二、往来单位余额
    S2 = K2 + 2
    section(ws, '二、往来单位余额（取自【往来业务明细】· 应收＝销售/仓储装卸周转费/物料/筐子/借出；'
                '应付＝采购/公司购买果品/借入/定金押金）', 'A', 'I', S2)
    header(ws, S2 + 1, 'A', ['往来单位', '期初应收', '本期应收发生', '资金实收', '应收余额', '期初应付', '本期应付发生',
                             '资金实付', '应付余额'])
    B0 = S2 + 2
    B1 = B0 + 149
    for r in range(B0, B1 + 1):
        n = f'$A{r}'
        k = f'ROW()-{B0 - 1}'
        cell(ws, f'A{r}', f'=IF({k}>往来业务明细!$Q${WR1},"",INDEX({W("D")},MATCH({k},{W("Q")},0)))', F_AUTOB, FL_AUTO, ACN)
        for d, (c1, c2, c3, c4) in (('应收', 'BCDE'), ('应付', 'FGHI')):
            base = f'{W("D")},{n},{W("E")},"{d}"'
            cell(ws, f'{c1}{r}', f'=IF({n}="","",ROUND(SUMIFS({W("M")},{base},{W("C")},">0",{W("C")},"<"&{s}),2))',
                 al=AR, fmt=MONEY)
            cell(ws, f'{c2}{r}', f'=IF({n}="","",ROUND(SUMIFS({W("K")},{base},{in_period(W("C"), p)})'
                                 f'+SUMIFS({W("K")},{base},{W("C")},0),2))', al=AR, fmt=MONEY)
            cell(ws, f'{c3}{r}', f'=IF({n}="","",-ROUND(SUMIFS({W("L")},{base},{in_period(W("C"), p)})'
                                 f'+SUMIFS({W("L")},{base},{W("C")},0),2))', al=AR, fmt=MONEY)
            cell(ws, f'{c4}{r}', f'=IF({n}="","",ROUND({c1}{r}+{c2}{r}-{c3}{r},2))', al=AR, fmt=MONEY)
    T2 = B1 + 1
    cell(ws, f'A{T2}', '合　计', F_TOT, FL_TOT, ACN)
    for c in 'BCDEFGHI':
        cell(ws, f'{c}{T2}', f'=ROUND(SUM({c}{B0}:{c}{B1}),2)', F_TOTN, FL_TOT, AR, MONEY)
    widths(ws, {'A': 24, 'B': 15, 'C': 15, 'D': 15, 'E': 15, 'F': 15, 'G': 15, 'H': 15, 'I': 15, 'J': 12, 'K': 5, 'L': 12})
    ws.freeze_panes = 'B4'
    return p


# ─────────────────────────────── ② 按月矩阵 ───────────────────────────────
def month_block(ws, top, label, keycol, val, crit, pool_cols, slots, p, ycell, pool_key):
    """top 行＝各月小计；top+1＝表头；下面 slots 行明细；再下一行合计。返回下一可用行。
       crit：资金日记账上的额外 SUMIFS 条件（字符串，已带逗号分隔）"""
    s, e = p['start'], p['end']
    keys, cnts, tot = pool(ws, *pool_cols, key_expr=pool_key)
    h = top + 1
    b0, b1 = h + 1, h + slots
    # 各月小计
    cell(ws, f'A{top}', '各 月 小 计', F_SUBL, FL_SUB, ACN)
    for i in range(13):
        c = L(2 + i)
        cell(ws, f'{c}{top}', f'=ROUND(SUM({c}{b0}:{c}{b1}),2)', F_SUBN, FL_SUB, AR, MONEY)
    # 表头
    cell(ws, f'A{h}', f'{label} / 月份', F_HDR, FL_HDR, AC)
    for i in range(12):
        cell(ws, f'{L(2 + i)}{h}', f'=DATE(YEAR({ycell}),MONTH({ycell})+{i},1)', F_HDR, FL_HDR, AC, MONTH)
    cell(ws, f'N{h}', '合　计', F_HDR, FL_HDR, AC)
    ws.row_dimensions[h].height = 30
    for r in range(b0, b1 + 1):
        a = f'$A{r}'
        cell(ws, f'A{r}', '=' + list_item(keys, cnts, tot, f'ROW()-{b0 - 1}'), F_AUTOB, FL_AUTO, AL)
        for i in range(12):
            c = L(2 + i)
            lo = f'MAX({c}${h},{s})'
            hi = f'MIN(DATE(YEAR({c}${h}),MONTH({c}${h})+1,0),{e})'
            cell(ws, f'{c}{r}', f'=IF({a}="","",ROUND(SUMIFS({J(val)},{J(keycol)},{a}{crit},'
                                f'{J("B")},">="&{lo},{J("B")},"<="&{hi}),2))', al=AR, fmt=MONEY)
        cell(ws, f'N{r}', f'=IF({a}="","",ROUND(SUM(B{r}:M{r}),2))', F_AUTOB, FL_AUTO, AR, MONEY)
    t = b1 + 1
    cell(ws, f'A{t}', '合　计', F_TOT, FL_TOTM, ACN)
    for i in range(13):
        c = L(2 + i)
        cell(ws, f'{c}{t}', f'=ROUND(SUM({c}{b0}:{c}{b1}),2)', F_TOTN, FL_TOTM, AR, MONEY)
    return t


def _matrix_sheet(ws, ttl, note_txt):
    wipe(ws)
    ws.sheet_properties.tabColor = GREEN_TAB
    title(ws, ttl, 'A', 'N')
    p = period(ws, 2, 'A', month=False, year=2026, clamp=12)
    note(ws, note_txt + ' 12 个月从起始月排起：只填年份＝1～12 月；只填起始日 2026/8/1、截止日不填，就排 2026年8月～2027年7月（一个果季）。',
         'A', 'N', 3, 30)
    widths(ws, {'A': 18, **{L(i): 12.5 for i in range(2, 14)}, 'N': 14})
    return p, p['start']


def _check(ws, row, label, total_expr, table_total_ref):
    cell(ws, f'A{row}', label, F_LBL, FL_LBL, AC)
    ws.merge_cells(f'B{row}:C{row}')
    cell(ws, f'B{row}', f'=ROUND({total_expr},2)', F_AUTOB, FL_AUTO, AR, MONEY)
    ws[f'C{row}'].border = BOX
    ws.merge_cells(f'D{row}:G{row}')
    cell(ws, f'D{row}', f'=IF(ROUND(B{row}-{table_total_ref},2)=0,"✔ 与表内合计一致","✘ 差 "&TEXT(B{row}-{table_total_ref},"#,##0.00")'
                        f'&"（有收支没填往来单位/费用项目）")', F_CHK, FL_NONE, AL)
    for c in 'EFG':
        ws[f'{c}{row}'].border = BOX
    ws.row_dimensions[row].height = 22


def income_expense_matrices(wb):
    per = lambda p, r: f'{jr("B", r)}>={p["start"]},{jr("B", r)}<={p["end"]}'
    # 收入月度汇总
    ws = wb['收入月度汇总']
    p, y = _matrix_sheet(ws, '收 入 月 度 汇 总（往来单位 × 月份 · 全自动）',
                         '只填第 2 行：年份（排 1～12 月），起始日/截止日可再缩小范围。口径＝【资金日记账】收入金额，按 E 列往来单位汇总，'
                         '账户互转不算。')
    t = month_block(ws, 5, '往来单位', 'E', 'H', f',{J("F")},"<>账户互转"', ('X', 'Y', 'Z'), 60, p, y,
                    '=IF(AND({e}<>"",N({h})<>0,{f}<>"账户互转",{per}),{e},"")'.format(
                        e=jr('E', '{r}'), h=jr('H', '{r}'), f=jr('F', '{r}'), per=per(p, '{r}'))[1:])
    _check(ws, 4, '资金日记账外部收入', f'SUMIFS({J("H")},{J("F")},"<>账户互转",{in_period(J("B"), p)})', f'$N${t}')
    ws.freeze_panes = 'B7'
    # 支出月度汇总
    ws = wb['支出月度汇总']
    p, y = _matrix_sheet(ws, '支 出 月 度 汇 总（往来单位 × 月份 · 全自动）',
                         '只填第 2 行：年份（排 1～12 月），起始日/截止日可再缩小范围。口径＝【资金日记账】支出金额，按 E 列往来单位汇总，'
                         '账户互转不算；「果然鲜」一行就是公司费用，按费用项目拆开看【费用月度支出汇总】。')
    t = month_block(ws, 5, '往来单位', 'E', 'I', f',{J("F")},"<>账户互转"', ('X', 'Y', 'Z'), 60, p, y,
                    '=IF(AND({e}<>"",N({i})<>0,{f}<>"账户互转",{per}),{e},"")'.format(
                        e=jr('E', '{r}'), i=jr('I', '{r}'), f=jr('F', '{r}'), per=per(p, '{r}'))[1:])
    _check(ws, 4, '资金日记账外部支出', f'SUMIFS({J("I")},{J("F")},"<>账户互转",{in_period(J("B"), p)})', f'$N${t}')
    ws.freeze_panes = 'B7'
    # 费用月度支出汇总
    ws = wb['费用月度支出汇总']
    p, y = _matrix_sheet(ws, '果 然 鲜 成 本 费 用 月 度 汇 总（费用项目 × 月份 · 全自动）',
                         '只填第 2 行。口径同【费用月度明细汇总】：往来单位填「果然鲜」的支出（不含账户互转），按费用项目（没填费用项目就按业务类型）汇总。')
    t = month_block(ws, 5, '费用项目', 'Y', 'I', f',{J("Z")},1', ('X', 'Y', 'Z'), 80, p, y,
                    '=IF(AND(N({z})=1,N({i})>0,{yy}<>"",{per}),{yy},"")'.format(
                        z=jr('Z', '{r}'), i=jr('I', '{r}'), yy=jr('Y', '{r}'), per=per(p, '{r}'))[1:])
    _check(ws, 4, '果然鲜费用支出合计', f'SUMIFS({J("I")},{J("Z")},1,{in_period(J("B"), p)})', f'$N${t}')
    ws.freeze_panes = 'B7'
    # 家用月度支出汇总（两段）
    ws = wb['家用月度支出汇总']
    p, y = _matrix_sheet(ws, '家 用 月 度 支 出 汇 总（全自动）',
                         '只填第 2 行。口径＝【资金日记账】业务类型选「家用」的行：上段按往来单位（家小赫、家小满……）看每月花了多少，'
                         '下段按费用项目看；家用里的收入（红包、还款等）在第 4 行单列。')
    hin = f'SUMIFS({J("H")},{J("F")},"家用",{in_period(J("B"), p)})'
    hout = f'SUMIFS({J("I")},{J("F")},"家用",{in_period(J("B"), p)})'
    for c, lab, v, fmt_ in (('A', '家用收入合计', f'=ROUND({hin},2)', MONEY), ('D', '家用支出合计', f'=ROUND({hout},2)', MONEY),
                            ('G', '净支出（支出－收入）', f'=ROUND(E4-B4,2)', MONEY)):
        cell(ws, f'{c}4', lab, F_LBL, FL_LBL, AC)
        ws.merge_cells(f'{L(CI(c) + 1)}4:{L(CI(c) + 2)}4')
        cell(ws, f'{L(CI(c) + 1)}4', v, F_AUTOB, FL_AUTO, AR, fmt_)
        ws[f'{L(CI(c) + 2)}4'].border = BOX
    ws.row_dimensions[4].height = 22
    section(ws, '一、按往来单位 · 支出', 'A', 'N', 6)
    t1 = month_block(ws, 7, '往来单位', 'E', 'I', f',{J("F")},"家用"', ('P', 'Q', 'R'), 20, p, y,
                     '=IF(AND({f}="家用",{e}<>"",N({i})>0,{per}),{e},"")'.format(
                         f=jr('F', '{r}'), e=jr('E', '{r}'), i=jr('I', '{r}'), per=per(p, '{r}'))[1:])
    section(ws, '二、按费用项目 · 支出', 'A', 'N', t1 + 2)
    t2 = month_block(ws, t1 + 3, '费用项目', 'Y', 'I', f',{J("F")},"家用"', ('S', 'T', 'U'), 60, p, y,
                     '=IF(AND({f}="家用",{yy}<>"",N({i})>0,{per}),{yy},"")'.format(
                         f=jr('F', '{r}'), yy=jr('Y', '{r}'), i=jr('I', '{r}'), per=per(p, '{r}'))[1:])
    ws.freeze_panes = 'B6'
    return t2


# ─────────────────────────────── ③ 明细汇总（左边汇总、右边逐笔） ───────────────────────────────
def resolve(tpl, p):
    """公式模板里的 %X（资金日记账某列）/ %PER（时段条件）换成真引用"""
    v = tpl.replace('%PER', in_period(J('B'), p))
    for col in ('AG', 'B', 'E', 'F', 'G', 'H', 'I', 'Y', 'Z'):
        v = v.replace(f'%{col},', f'{J(col)},').replace(f'%{col})', f'{J(col)})')
    return v


def detail_sheet(ws, ttl, note_txt, month, mon, flt, kpis, blocks, right, right_col='F', det_rows=600,
                 kpi_cols=('A', 'C', 'E')):
    """flt：资金日记账第 {r} 行是否入选的条件表达式（不含时段）；
       kpis：[(标签, SUMIFS 值表达式)]；blocks：[(标题, 键列, 槽数, [(列名, 公式模板{n}{r})])]；
       right：[(表头, 取值模板({src}=资金日记账行号), 格式, 宽)]"""
    wipe(ws)
    ws.sheet_properties.tabColor = GREEN_TAB
    rc0 = CI(right_col)
    last = L(rc0 + len(right) - 1)
    title(ws, ttl, 'A', last)
    p = period(ws, 2, 'A', month=month, year=2026, mon=mon)
    note(ws, note_txt, 'A', last, 3, 30)
    s, e = p['start'], p['end']
    per = lambda r: f'{jr("B", r)}>={s},{jr("B", r)}<={e}'
    # 第 4 行 KPI
    vcells = [f'{L(CI(c) + 1)}4' for c in kpi_cols]
    for c, (lab, expr) in zip(kpi_cols, kpis):
        cell(ws, f'{c}4', lab, F_LBL, FL_LBL, AC)
        expr = resolve(expr, p).replace('@1', vcells[0]).replace('@2', vcells[1])
        cell(ws, f'{L(CI(c) + 1)}4', f'=ROUND({expr},2)', F_AUTOB, FL_AUTO, AR, MONEY)
    ws.row_dimensions[4].height = 26
    # 辅助列：从 right 区右边空两列开始
    hc = rc0 + len(right) + 2
    helpers = []
    def nxt():
        nonlocal hc
        c = L(hc); hc += 1
        return c
    # 左边分块
    row = 6
    for bt, keycol, slots, cols in blocks:
        k1, k2, k3 = nxt(), nxt(), nxt()
        keys, cnts, tot = pool(ws, k1, k2, k3, key_expr=f'IF(AND({flt("{r}")},{jr(keycol, "{r}")}<>"",{per("{r}")}),'
                                                        f'{jr(keycol, "{r}")},"")')
        header(ws, row, 'A', [bt] + [c[0] for c in cols])
        b0, b1 = row + 1, row + slots
        for r in range(b0, b1 + 1):
            cell(ws, f'A{r}', '=' + list_item(keys, cnts, tot, f'ROW()-{b0 - 1}'), F_AUTOB, FL_AUTO, AL)
            for i, (_, tpl) in enumerate(cols):
                c = L(2 + i)
                cell(ws, f'{c}{r}', f'=IF($A{r}="","",{resolve(tpl, p).format(n=f"$A{r}", key=J(keycol))})', al=AR,
                     fmt=INT if '笔' in _ else MONEY)
        t = b1 + 1
        cell(ws, f'A{t}', '合　计', F_TOT, FL_TOT, ACN)
        for i, (lab, _) in enumerate(cols):
            c = L(2 + i)
            cell(ws, f'{c}{t}', f'=ROUND(SUM({c}{b0}:{c}{b1}),2)', F_TOTN, FL_TOT, AR, INT if '笔' in lab else MONEY)
        row = t + 2
    # 右边逐笔明细
    cc = nxt()
    ws[f'{cc}3'] = '排序键'; ws[f'{cc}3'].font = F_HELP
    for r in range(JR0, JR1 + 1):
        ws[f'{cc}{r}'] = f'=IF(AND({flt(r)},{per(r)}),INT({jr("B", r)})*1000+{r},"")'
        ws[f'{cc}{r}'].font = F_HELP
    src = nxt()
    hide(ws, cc, src)
    header(ws, 6, right_col, [h for h, *_ in right])
    D0, D1 = 7, 7 + det_rows - 1
    for r in range(D0, D1 + 1):
        k = f'ROW()-{D0 - 1}'
        ws[f'{src}{r}'] = (f'=IFERROR(SMALL(${cc}${JR0}:${cc}${JR1},{k})-INT(SMALL(${cc}${JR0}:${cc}${JR1},{k})/1000)*1000,"")')
        ws[f'{src}{r}'].font = F_HELP
        for i, (h, tpl, fmt_, w) in enumerate(right):
            c = L(rc0 + i)
            cell(ws, f'{c}{r}', f'=IF(${src}{r}="","",{tpl.format(src=f"${src}{r}")})',
                 al=AR if fmt_ in (MONEY, INT) else (ACN if fmt_ else AL), fmt=fmt_)
    widths(ws, {'A': 16, 'B': 12, 'C': 13, 'D': 13, 'E': 10})
    for i, (h, tpl, fmt_, w) in enumerate(right):
        ws.column_dimensions[L(rc0 + i)].width = w
    ws.freeze_panes = 'A7'
    return p


def _idx(c):
    return f'INDEX({JD(c)},{{src}})'


def detail_sheets(wb):
    # 费用月度明细汇总（公司费用）
    ws = wb['费用月度明细汇总']
    flt = lambda r: f'N({jr("Z", r)})=1,N({jr("I", r)})>0'
    p_ = None
    right = [('日期', _idx('B'), DATE, 11), ('资金账户', _idx('C') + '&""', None, 12),
             ('摘要/用途', _idx('D') + '&""', None, 30), ('费用项目', _idx('Y') + '&""', None, 14),
             ('支出金额', f'N({_idx("I")})', MONEY, 13), ('备注/收款人补充', _idx('L') + '&""', None, 22),
             ('源行', '{src}', INT, 7), ('对应科目', _idx('K') + '&""', None, 12)]
    # KPI 需要时段格，先占位，建完再写
    p = detail_sheet(ws, '果 然 鲜 费 用 · 月 度 支 出 汇 总',
                     '按往来单位「果然鲜」提取外部资金支出（不含账户互转）。总支出不等于损益费用：「其他/待分类」是记到借款、采购、'
                     '往来等科目的支出，右边明细最后一列能看到对应科目。月份留空＝全年。',
                     True, 9, flt,
                     [('外部支出合计', 'SUMIFS(%I,%Z,1,%PER)'), ('已归费用支出', 'SUMIFS(%I,%AG,1,%PER)'),
                      ('其他/待分类', '@1-@2')],
                     [('费用项目', 'Y', 80, [('笔数', 'COUNTIFS({key},{n},%Z,1,%I,">0",%PER)'),
                                             ('支出金额', 'SUMIFS(%I,{key},{n},%Z,1,%PER)')])],
                     right, 'E')
    # 家用月度明细汇总
    ws = wb['家用月度明细汇总']
    flt = lambda r: f'{jr("F", r)}="家用",N({jr("H", r)})+N({jr("I", r)})<>0'
    right = [('日期', _idx('B'), DATE, 11), ('资金账户', _idx('C') + '&""', None, 11), ('往来单位', _idx('E') + '&""', None, 10),
             ('摘要/用途', _idx('D') + '&""', None, 26), ('费用项目', _idx('Y') + '&""', None, 12),
             ('收入', f'N({_idx("H")})', MONEY, 12), ('支出', f'N({_idx("I")})', MONEY, 12),
             ('备注', _idx('L') + '&""', None, 14), ('源行', '{src}', INT, 7)]
    cols = [('笔数', 'COUNTIFS({key},{n},%F,"家用",%PER)'), ('收入', 'SUMIFS(%H,{key},{n},%F,"家用",%PER)'),
            ('支出', 'SUMIFS(%I,{key},{n},%F,"家用",%PER)')]
    p = detail_sheet(ws, '家 用 费 用 · 月 度 收 支 明 细',
                     '口径＝【资金日记账】业务类型选「家用」的行。左上按往来单位（家小赫、家小满……）看收入、支出，左下按费用项目看，'
                     '右边逐笔明细。月份留空＝全年。',
                     True, 9, flt, [('家用收入', 'SUMIFS(%H,%F,"家用",%PER)'), ('家用支出', 'SUMIFS(%I,%F,"家用",%PER)'),
                                   ('净支出（支出－收入）', '@2-@1')],
                     [('往来单位', 'E', 20, cols), ('费用项目', 'Y', 60, cols)], right, 'F', kpi_cols=('A', 'C', 'F'))
    # 国外费用汇总（新表）
    idx = wb.sheetnames.index('家用月度明细汇总') + 1
    ws = wb.create_sheet('国外费用汇总', idx)
    flt = lambda r: f'{jr("F", r)}="国外",N({jr("H", r)})+N({jr("I", r)})<>0'
    right = [('日期', _idx('B'), DATE, 11), ('资金账户', _idx('C') + '&""', None, 11), ('往来单位', _idx('E') + '&""', None, 10),
             ('摘要/用途', _idx('D') + '&""', None, 26), ('费用项目', _idx('G') + '&""', None, 10),
             ('转出金额', f'N({_idx("I")})', MONEY, 13), ('收回/收入', f'N({_idx("H")})', MONEY, 12),
             ('备注', _idx('L') + '&""', None, 16), ('源行', '{src}', INT, 7)]
    cols = [('笔数', 'COUNTIFS({key},{n},%F,"国外",%PER)'), ('转出金额', 'SUMIFS(%I,{key},{n},%F,"国外",%PER)'),
            ('收回/收入', 'SUMIFS(%H,{key},{n},%F,"国外",%PER)')]
    p = detail_sheet(ws, '国 外 费 用 · 转 出 汇 总',
                     '口径＝【资金日记账】业务类型选「国外」的行，统计给国外那边一共转出多少。默认看全年，起始日/截止日可缩小范围；'
                     '左边按往来单位汇总，右边逐笔明细。',
                     False, None, flt, [('国外转出合计', 'SUMIFS(%I,%F,"国外",%PER)'), ('收回/收入', 'SUMIFS(%H,%F,"国外",%PER)'),
                                     ('净转出', '@1-@2')],
                     [('往来单位', 'E', 30, cols)], right, 'F', kpi_cols=('A', 'C', 'F'))


# ─────────────────────────────── 借款汇总 ───────────────────────────────
def loans(wb):
    idx = wb.sheetnames.index('国外费用汇总') + 1
    ws = wb.create_sheet('借款汇总', idx)
    ws.sheet_properties.tabColor = GREEN_TAB
    ws.sheet_view.showGridLines = False
    title(ws, '借 款 汇 总（我方借入 · 我方借出 · 余额）', 'A', 'O')
    p = period(ws, 2, 'A', month=False, year=2026, open_ended=True)
    s, e = p['start'], p['end']
    note(ws, '【资金日记账】业务类型或费用项目是「借款」「还款」的行：借款收到钱＝我方借入，付出钱＝我方借出；还款付出钱＝我方还别人，'
             '收到钱＝别人还我方。记账以前就有的借款填在下面第三段（只填一次）。年份/日期留空＝不限，余额算到截止日。', 'A', 'O', 3, 36)
    # 手工期初区
    M0, M1 = 80, 99
    # 辅助列（隐藏）：R 类别、S 往来单位、T 日期、U 金额；V/W/X 借入名单；Y/Z/AA 借出名单；AB 明细编号
    for r in range(JR0, JR1 + 1):
        F, G, H, I, E_ = (jr(c, r) for c in 'FGHIE')
        ws[f'R{r}'] = (f'=IF({jr("C", r)}="","",IF(OR({F}="借款",{G}="借款"),IF(N({H})>0,"借入",IF(N({I})>0,"借出","")),'
                       f'IF(OR({F}="还款",{G}="还款"),IF(N({I})>0,"我方还款",IF(N({H})>0,"对方还款","")),"")))')
        ws[f'S{r}'] = f'={E_}&""'
        ws[f'T{r}'] = f'=N({jr("B", r)})'
        ws[f'U{r}'] = f'=N({H})+N({I})'
    for j in range(M1 - M0 + 1):
        r, m = JR1 + 1 + j, M0 + j
        ws[f'R{r}'] = f'=IF(OR($A{m}="",N($C{m})=0),"",IF($B{m}="借出","借出","借入"))'
        ws[f'S{r}'] = f'=$A{m}&""'
        ws[f'T{r}'] = '=0'
        ws[f'U{r}'] = f'=N($C{m})'
    HR1 = JR1 + (M1 - M0 + 1)
    for c in 'RSTU':
        ws[f'{c}3'] = {'R': '类别', 'S': '往来单位', 'T': '日期', 'U': '金额'}[c]
        for r in range(3, HR1 + 1):
            ws[f'{c}{r}'].font = F_HELP
    rg = lambda c: f'${c}${JR0}:${c}${HR1}'
    kIN = pool(ws, 'V', 'W', 'X', f'IF(AND(OR($R{{r}}="借入",$R{{r}}="我方还款"),$S{{r}}<>"",$T{{r}}<={e}),$S{{r}},"")',
               r1=HR1)
    kOUT = pool(ws, 'Y', 'Z', 'AA', f'IF(AND(OR($R{{r}}="借出",$R{{r}}="对方还款"),$S{{r}}<>"",$T{{r}}<={e}),$S{{r}},"")',
                r1=HR1)
    ws['AB3'] = '排序键'
    for r in range(JR0, JR1 + 1):
        ws[f'AB{r}'] = f'=IF(AND($R{r}<>"",$T{r}>={s},$T{r}<={e}),INT($T{r})*1000+{r},"")'
        ws[f'AB{r}'].font = F_HELP
    hide(ws, 'R', 'S', 'T', 'U', 'AB', 'AC')

    def block(top, ttl_, who, t_add, t_sub, lab_add, lab_sub, lab_bal, keys):
        section(ws, ttl_, 'A', 'G', top)
        header(ws, top + 1, 'A', [who, '笔数', '期初余额', lab_add, lab_sub, lab_bal, '提示'])
        b0, b1 = top + 2, top + 31
        for r in range(b0, b1 + 1):
            n = f'$A{r}'
            sm = lambda t, cond: f'SUMIFS({rg("U")},{rg("S")},{n},{rg("R")},"{t}",{cond})'
            cell(ws, f'A{r}', '=' + list_item(keys[0], keys[1], keys[2], f'ROW()-{b0 - 1}'), F_AUTOB, FL_AUTO, AL)
            cell(ws, f'B{r}', f'=IF({n}="","",COUNTIFS({rg("S")},{n},{rg("R")},"{t_add}",{rg("T")},"<="&{e})'
                              f'+COUNTIFS({rg("S")},{n},{rg("R")},"{t_sub}",{rg("T")},"<="&{e}))', al=AR, fmt=INT)
            before = f'{rg("T")},"<"&{s}'
            within = f'{rg("T")},">="&{s},{rg("T")},"<="&{e}'
            cell(ws, f'C{r}', f'=IF({n}="","",ROUND({sm(t_add, before)}-{sm(t_sub, before)},2))', al=AR, fmt=MONEY)
            cell(ws, f'D{r}', f'=IF({n}="","",ROUND({sm(t_add, within)},2))', al=AR, fmt=MONEY)
            cell(ws, f'E{r}', f'=IF({n}="","",ROUND({sm(t_sub, within)},2))', al=AR, fmt=MONEY)
            cell(ws, f'F{r}', f'=IF({n}="","",ROUND(C{r}+D{r}-E{r},2))', F_AUTOB, FL_AUTO, AR, MONEY)
            cell(ws, f'G{r}', f'=IF({n}="","",IF(F{r}<0,"有还款、没有借款记录：请在第三段补记账前的借款",""))',
                 Font(name=SONG, size=9, color='C00000'), FL_AUTO, AL)
        t = b1 + 1
        cell(ws, f'A{t}', '合　计', F_TOT, FL_TOT, ACN)
        for c in 'BCDEF':
            cell(ws, f'{c}{t}', f'=ROUND(SUM({c}{b0}:{c}{b1}),2)', F_TOTN, FL_TOT, AR, INT if c == 'B' else MONEY)
        cell(ws, f'G{t}', None, F_TOTN, FL_TOT, AC)
        return t

    t1 = block(6, '一、我方借入（向别人借的钱）', '出借人', '借入', '我方还款', '本期借款', '本期还款', '期末余额（我方欠）', kIN)
    t2 = block(t1 + 2, '二、我方借出（借给别人的钱）', '借款人', '借出', '对方还款', '本期借出', '本期收回',
               '期末余额（对方欠我方）', kOUT)
    # 第 4 行 KPI
    kp = [('A', '总借款（本期借入）', f'=D{t1}'), ('C', '总还款（我方本期还）', f'=E{t1}'),
          ('E', '借入余额合计（我方欠）', f'=F{t1}'), ('G', '借出余额合计（别人欠我方）', f'=F{t2}'),
          ('I', '净借款结余（借入余额－借出余额）', f'=ROUND(F{t1}-F{t2},2)')]
    for c, lab, v in kp:
        cell(ws, f'{c}4', lab, F_LBL, FL_LBL, AC)
        cell(ws, f'{L(CI(c) + 1)}4', v, F_AUTOB, FL_AUTO, AR, MONEY)
    ws.row_dimensions[4].height = 30
    # 第三段：手工期初
    top = M0 - 2
    section(ws, '三、记账以前就有的借款（手工填，只填一次；方向选 借入/借出）', 'A', 'F', top)
    header(ws, top + 1, 'A', ['往来单位', '方向', '金额', '备注'])
    for r in range(M0, M1 + 1):
        for c in 'ABCD':
            cell(ws, f'{c}{r}', None, F_IN, FL_IN, AR if c == 'C' else AC, MONEY if c == 'C' else None)
    dv_list(ws, f'B{M0}:B{M1}', '"借入,借出"', err=True)
    dv = DataValidation(type='list', formula1='INDIRECT("科目表!$I$4:$I$1003")', allow_blank=True, showErrorMessage=False)
    dv.add(f'A{M0}:A{M1}'); ws.add_data_validation(dv)
    # 右边逐笔明细 H～O
    loc = lambda c: f'INDEX(${c}$1:${c}${HR1},{{src}})'
    right = [('日期', loc('T'), DATE, 11), ('资金账户', f'INDEX({JD("C")},{{src}})&""', None, 12),
             ('往来单位', loc('S'), None, 13), ('摘要', f'INDEX({JD("D")},{{src}})&""', None, 22),
             ('业务类型/费用项目', f'INDEX({JD("F")},{{src}})&IF(INDEX({JD("G")},{{src}})&""="","","/"&INDEX({JD("G")},{{src}}))', None, 16),
             ('类别', loc('R'), None, 9), ('金额', loc('U'), MONEY, 13), ('源行', '{src}', INT, 7)]
    header(ws, 7, 'H', [h for h, *_ in right])
    for r in range(8, 8 + 300):
        k = f'ROW()-7'
        ws[f'AC{r}'] = f'=IFERROR(SMALL($AB${JR0}:$AB${JR1},{k})-INT(SMALL($AB${JR0}:$AB${JR1},{k})/1000)*1000,"")'
        ws[f'AC{r}'].font = F_HELP
        for i, (h, tpl, fmt_, w) in enumerate(right):
            c = L(8 + i)
            cell(ws, f'{c}{r}', f'=IF($AC{r}="","",{tpl.format(src=f"$AC{r}")})', al=AR if fmt_ in (MONEY, INT) else (ACN if fmt_ else AL),
                 fmt=fmt_)
    for i, (h, tpl, fmt_, w) in enumerate(right):
        ws.column_dimensions[L(8 + i)].width = w
    widths(ws, {'A': 16, 'B': 11, 'C': 14, 'D': 14, 'E': 14, 'F': 16, 'G': 3})
    ws.column_dimensions['G'].width = 30
    ws.column_dimensions['I'].width = 12
    ws.freeze_panes = 'A7'


# ─────────────────────────────── ④ 删月度资金明细、清辅助列 ───────────────────────────────
def cleanup(wb):
    wb.remove(wb['月度资金明细'])
    # 记账凭证资金段的 凭证字号/票据号/制单人 还指着资金日记账 W/X/Y（那三列早就改成了辅助列），显示成「非往来」之类
    import re
    ws = wb['记账凭证']
    for row in ws.iter_rows(min_row=4, max_row=3609):
        for c in row:
            v = c.value
            if not (isinstance(v, str) and re.search(r'资金日记账!\$(W|X|Y)\$', v)):
                continue
            r = c.row
            if c.column_letter == 'C':
                c.value = f'=IF($E{r}="","","记-"&TEXT($B{r},"YYYYMM")&"-资金")'
            elif c.column_letter == 'L':
                c.value = None
            elif c.column_letter == 'N':
                c.value = f'=IF($E{r}="","","系统")'
    ws = wb['资金日记账']
    for c in ('AB', 'AC', 'AD', 'AE'):
        for r in range(3, JR1 + 1):
            ws[f'{c}{r}'].value = None


def apply(wb):
    fund_balance(wb)
    income_expense_matrices(wb)
    detail_sheets(wb)
    loans(wb)
    cleanup(wb)
