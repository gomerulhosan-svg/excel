# -*- coding: utf-8 -*-
"""报表：收支汇总表（照你的收支类型截图）、简易现金流量表（业务口径）、资金余额表、费用统计（按公司）、
   客户收入统计（新增/续费）、发票汇总。全部按【资金台帐】【发票导入】自动算，顶上选公司、年份。"""
import datetime as dt
from common import *

MCOLS = [CL(3 + i) for i in range(12)]     # C..N：1～12 月
Q = '"'                                    # 公式里的双引号（嵌在 f-string 里用）


def _yr(y, m=1, add=0):
    return f'DATE({y},{m}+{add},1)' if add else f'DATE({y},{m},1)'


def jsum(col, crit):
    return f'SUMIFS({jr(col)},{crit})'


def month_crit(y, m):
    return f'{jr(J_DATE)},">="&DATE({y},{m},1),{jr(J_DATE)},"<"&DATE({y},{m}+1,1)'


def year_crit(y):
    return f'{jr(J_DATE)},">="&DATE({y},1,1),{jr(J_DATE)},"<"&DATE({y}+1,1,1)'


def ym_crit(y, m):
    """某年某月：资金台帐的「年月」列（YYYYMM 一个数）等于它——比「日期 ≥ 月初、< 下月初」两个条件快"""
    return f'{jr(J_YM)},{y}*100+{m}'


def net(crit, sign=1):
    """净额（收 − 支）一次求和；sign=-1 是「支 − 收」（成本费用类）"""
    return f'{"" if sign > 0 else "-"}SUMIFS({jr(J_NET)},{crit})'


def _sel_block(ws, y_default=2026, co=True, yr=True, extra=None):
    if co:
        selector(ws, 'B3', '公司', 'C3', '全部', f'={AUX_CO_ALL}', prompt='选一家公司，或「全部」')
    if yr:
        selector(ws, 'E3', '年份', 'F3', y_default, YEARS)
    ws.row_dimensions[3].height = 24


# ─────────────────────────────── 收支汇总表 ───────────────────────────────
def build_sum(wb, ctx):
    ws = wb.create_sheet(SH_SUM)
    widths(ws, {'A': 11, 'B': 15, **{c: 11.5 for c in MCOLS}, 'O': 13, 'P': 8, 'Q': 4})
    title(ws, '收 支 汇 总 表（收支类型 × 1～12 月 · 按资金台帐实收实付）', 'P', C_RPT,
          '💡 选公司、年份。行就是「收支类型」：收入 → 变动成本 → 变动利润 → 固定成本 → 利润，下面是其他支出、实收资本、投资、往来。'
          '收入类＝收进来减退出去；成本费用类＝付出去减退回来。账户互转、内部公司之间划转不算收入支出（最下面单列）。')
    _sel_block(ws)
    cc, y = co_crit('$C$3'), '$F$3'
    ci = coi_crit('$C$3')
    header(ws, 4, [('A', '收支类型'), ('B', '项目')] + [(c, f'{i + 1}月') for i, c in enumerate(MCOLS)] +
           [('O', '全年合计'), ('P', '占收入')], C_RPT)
    ws['Q4'] = '项目名'
    ws['Q4'].font = F_HELP
    r = 5
    rows = {}

    def flow_cell(crit, sign):
        return net(crit, sign)

    def item_row(r, cls, item, sign, grp_fill):
        """一个项目一行：名字从【基础资料】③ 那一行取（改名跟着变），按「项目＋类别」求和"""
        k = IT_ROW[item]
        put(ws, f'B{r}', f'=IF({SH_BASE}!$W${k}="",{SH_BASE}!$U${k}&"",{SH_BASE}!$W${k}&"")', F_TXT, grp_fill, align=AC)
        ws[f'Q{r}'] = f'={SH_BASE}!$U${k}&""'
        ws[f'Q{r}'].font = F_HELP
        for i, c in enumerate(MCOLS):
            base = f'{jr(J_ITEM)},$Q{r},{jr(J_CLS)},"{cls}",{jr(J_COI)},{ci},{ym_crit(y, i + 1)}'
            put(ws, f'{c}{r}', f'=ROUND({flow_cell(base, sign)},2)', F_AUTO, grp_fill, MONEY, AR)
        put(ws, f'O{r}', f'=ROUND(SUM(C{r}:N{r}),2)', F_AUTOB, grp_fill, MONEY, AR)

    def rest_row(r, cls, sign, grp_fill, g0, item=None):
        """这一类的「其他」：整类合计 − 上面列出的项目（基础资料里后加的项目也落在这行，小计永远等于整类合计）"""
        if item:
            ws[f'Q{r}'] = f'={SH_BASE}!$U${IT_ROW[item]}&""'
            ws[f'Q{r}'].font = F_HELP
        n_orig = len(dict(ITEMS)[cls])
        put(ws, f'B{r}', f'=IF(COUNTIFS({IT_CLSS},"{cls}",{IT_NAMES},"?*")>{n_orig},"其他（含新加项目）","其他")', F_TXT, grp_fill, align=AC)
        for i, c in enumerate(MCOLS):
            base = f'{jr(J_CLS)},"{cls}",{jr(J_COI)},{ci},{ym_crit(y, i + 1)}'
            put(ws, f'{c}{r}', f'=ROUND({flow_cell(base, sign)}-SUM({c}{g0}:{c}{r - 1}),2)' if r > g0 else f'=ROUND({flow_cell(base, sign)},2)',
                F_AUTO, grp_fill, MONEY, AR)
        put(ws, f'O{r}', f'=ROUND(SUM(C{r}:N{r}),2)', F_AUTOB, grp_fill, MONEY, AR)

    def class_row(r, cls, label, sign, grp_fill):
        """整类一行（实收资本、投资、往来、内部划转、账户互转）"""
        put(ws, f'B{r}', label, F_TXT, grp_fill, align=AC)
        for i, c in enumerate(MCOLS):
            base = f'{jr(J_CLS)},"{cls}",{jr(J_COI)},{ci},{ym_crit(y, i + 1)}'
            put(ws, f'{c}{r}', f'=ROUND({flow_cell(base, sign)},2)', F_AUTO, grp_fill, MONEY, AR)
        put(ws, f'O{r}', f'=ROUND(SUM(C{r}:N{r}),2)', F_AUTOB, grp_fill, MONEY, AR)

    def total_row(r, label, formula_by_col, fl, bold_font=F_TXTB):
        put(ws, f'B{r}', label, bold_font, fl, align=AC)
        for c in MCOLS + ['O']:
            put(ws, f'{c}{r}', '=' + formula_by_col(c), F_AUTOB, fl, MONEY, AR)

    groups = [(CLS_IN, '收入', +1, C_IN, '小计'), (CLS_VAR, '变动成本', -1, C_VAR, '变动成本小计'),
              (CLS_FIX, '固定成本', -1, C_FIX, '固定成本小计'), (CLS_OTH, '其他支出', -1, C_MISC, '其他支出小计')]
    items_by_cls = dict(ITEMS)
    for cls, label, sign, col, sub in groups:
        g0 = r
        fl = fill(col)
        names = items_by_cls[cls]
        last_other = names[-1] if names[-1] in ('其他收入', '其他成本') else None
        for it in names:
            if it == last_other:
                continue
            item_row(r, cls, it, sign, fl)
            r += 1
        rest_row(r, cls, sign, fl, g0, last_other)       # 收入、变动成本：截图里的「其他」；固定成本、其他支出：补一行「其他」
        r += 1
        total_row(r, sub, lambda c, a=g0, b=r - 1: f'ROUND(SUM({c}{a}:{c}{b}),2)', fill('FFBDD7EE') if cls != CLS_IN else fill('FFA9D08E'))
        rows[cls] = r
        r += 1
        if cls == CLS_VAR:
            total_row(r, '变动利润', lambda c: f'ROUND({c}{rows[CLS_IN]}-{c}{rows[CLS_VAR]},2)', FILL_TOT, F_RED)
            rows['vp'] = r
            r += 1
        if cls == CLS_FIX:
            total_row(r, '利润', lambda c: f'ROUND({c}{rows["vp"]}-{c}{rows[CLS_FIX]},2)', FILL_TOT, F_RED)
            rows['p'] = r
            r += 1
        ws.merge_cells(f'A{g0}:A{r - 1}')
        put(ws, f'A{g0}', label, F_TXTB, fl, align=ACW)
        for rr in range(g0, r):
            ws[f'A{rr}'].border = BD
    # 收到为准 / 支出为准 / 往来
    for label, cls, name, sign in (('收到为准', CLS_CAP, '实收资本', +1), ('支出为准', CLS_INVT, '投资', -1),
                                   ('', CLS_WL, '往来（净收入）', +1)):
        class_row(r, cls, name, sign, fill(C_MISC))
        put(ws, f'A{r}', label, F_TXTB, fill(C_MISC), align=ACW)
        r += 1
    r += 1
    for cls, name in ((CLS_INTRA, '内部划转（净收入）'), (CLS_XFER, '账户互转（应为 0）')):
        put(ws, f'A{r}', '参考', F_NOTE, align=AC)
        class_row(r, cls, name, +1, FILL_AUTO)
        rows['intra' if cls == CLS_INTRA else 'xfer'] = r
        r += 1
    put(ws, f'A{r}', '参考', F_NOTE, align=AC)
    put(ws, f'B{r}', '未定收支项目', F_TXT, FILL_AUTO, align=AC)
    for i, c in enumerate(MCOLS):
        base = f'{jr(J_COI)},{ci},{ym_crit(y, i + 1)}'
        put(ws, f'{c}{r}', f'=ROUND(SUMIFS({jr(J_NET)},{base})-SUMIFS({jr(J_NET)},{jr(J_CLS)},"?*",{base})+SUMIFS({jr(J_NET)},{jr(J_CLS)},"？",{base}),2)',
            F_RED, FILL_AUTO, MONEY, AR)
    put(ws, f'O{r}', f'=ROUND(SUM(C{r}:N{r}),2)', F_RED, FILL_AUTO, MONEY, AR)
    rows['un'] = r
    last = r
    for rr in range(5, last + 1):
        if ws[f'O{rr}'].value is not None:
            put(ws, f'P{rr}', f'=IF(N($O${rows[CLS_IN]})=0,"",O{rr}/$O${rows[CLS_IN]})', F_NOTE, fmt=PCT, align=AR)
    put(ws, f'A{last + 2}', '说明：「未定收支项目」是资金台帐里还没分类的收付（净额，红字），分好类就会进到上面对应的行；'
                            '在【基础资料】③ 新加的项目算进同一类的「其他」行（名字后面会标「含新加项目」），小计、利润都是全的；'
                            '「占收入」＝全年合计 ÷ 收入小计。', F_NOTE, align=AL, border=False)
    hide(ws, 'Q')
    ws.freeze_panes = 'C5'
    ctx['sum_rows'] = rows
    return ws


# ─────────────────────────────── 简易现金流量表 ───────────────────────────────
def _open_bal(cc, ci, d):
    """某天开始时的资金余额（该公司所有账户）：建账日期在这天之前（或没填建账日期）的期初 ＋ 这天之前的收支"""
    return (f'(SUMIFS({AC_OPENS},{AC_COS},{cc})-SUMIFS({AC_OPENS},{AC_COS},{cc},{AC_ODATES},">="&{d})'
            f'+SUMIFS({jr(J_NET)},{jr(J_COI)},{ci},{jr(J_DATE)},"<"&{d}))')


def _new_acc(cc, d0, d1):
    """这段时间里新建账的账户的期初余额（建账日期在 [d0, d1) 之间）"""
    return f'SUMIFS({AC_OPENS},{AC_COS},{cc},{AC_ODATES},">="&{d0},{AC_ODATES},"<"&{d1})'


def build_cf(wb, ctx):
    ws = wb.create_sheet(SH_CF)
    widths(ws, {'A': 30, **{CL(2 + i): 11.5 for i in range(12)}, 'N': 13})
    title(ws, '简 易 现 金 流 量 表（业务口径 · 按资金台帐）', 'N', C_RPT,
          '💡 不是税务那张现金流量表，是看「钱从哪来、花到哪去」的业务表：期初资金 ＋ 经营（收入 − 成本费用）＋ 投资 ＋ 筹资往来 ＝ 期末资金。'
          '最下面一行跟账户余额核对。选公司看单家时，自家公司之间的内部划转算流入流出；选「全部」时内部划转两边抵掉。')
    _sel_block(ws)
    cc, y = co_crit('$C$3'), '$F$3'
    ci = coi_crit('$C$3')
    MC = [CL(2 + i) for i in range(12)]
    header(ws, 4, [('A', '项　　目')] + [(c, f'{i + 1}月') for i, c in enumerate(MC)] + [('N', '全年')], C_RPT)

    def flow(c, cls, sign, i=None):
        """sign=+1: 收−支（流入为正）；-1: 支−收（流出为正）；全年＝12 个月加起来"""
        if c == 'N':
            return f'ROUND(SUM({MC[0]}{{r}}:{MC[-1]}{{r}}),2)'
        return f'ROUND({net(f"{jr(J_CLS)},{Q}{cls}{Q},{jr(J_COI)},{ci},{ym_crit(y, i + 1)}", sign)},2)'

    lines = [
        ('open', '一、期初资金余额', 'open'),
        ('new', '　　加：本期新建账户的建账余额', 'new'),
        (None, '二、经营活动', None),
        ('in', '　　收入回款（收入类）', ('flow', CLS_IN, +1)),
        ('var', '　　支付变动成本', ('flow', CLS_VAR, -1)),
        ('fix', '　　支付固定成本', ('flow', CLS_FIX, -1)),
        ('oth', '　　其他支出（挂靠、贷款利息）', ('flow', CLS_OTH, -1)),
        ('op', '　经营活动净流量', ('net', [(+1, 'in'), (-1, 'var'), (-1, 'fix'), (-1, 'oth')])),
        (None, '三、投资活动', None),
        ('invt', '　　投资支出（支出为准）', ('flow', CLS_INVT, -1)),
        ('inv', '　投资活动净流量', ('net', [(-1, 'invt')])),
        (None, '四、筹资及往来', None),
        ('cap', '　　实收资本（收到为准）', ('flow', CLS_CAP, +1)),
        ('wl', '　　往来款净流入（借款、押金等）', ('flow', CLS_WL, +1)),
        ('intra', '　　内部公司划转净流入', ('flow', CLS_INTRA, +1)),
        ('xfer', '　　账户互转（同公司，应为 0）', ('flow', CLS_XFER, +1)),
        ('fin', '　筹资及往来净流量', ('net', [(+1, 'cap'), (+1, 'wl'), (+1, 'intra'), (+1, 'xfer')])),
        ('un', '五、还没分类的收支（资金台帐收支项目空着的）', 'un'),
        ('chg', '六、本期资金净增加', ('net', [(+1, 'op'), (+1, 'inv'), (+1, 'fin'), (+1, 'un')])),
        ('end', '七、期末资金余额', ('net', [(+1, 'open'), (+1, 'new'), (+1, 'chg')])),
        ('chk', '核对：各账户余额合计', 'chk'),
        ('ok', '', 'ok'),
    ]
    R0 = 5
    at = {key: R0 + k for k, (key, _, _) in enumerate(lines) if key}
    NACC = AC_R1 - AC_R0 + 1
    HR0 = R0 + len(lines) + 4       # 隐藏辅助行：每个账户在期末那天算不算期初
    for k, (key, label, kind) in enumerate(lines):
        r = R0 + k
        big = kind in ('open',) or (isinstance(kind, tuple) and kind[0] == 'net')
        put(ws, f'A{r}', label, F_TXTB if (kind is None or big) else F_TXT,
            fill('FFD9E1F2') if kind is None else (FILL_TOT if big else FILL_NONE), align=AL)
        for i, c in enumerate(MC + ['N']):
            d0 = f'DATE({y},1,1)' if c == 'N' else f'DATE({y},{i + 1},1)'
            d1 = f'DATE({y}+1,1,1)' if c == 'N' else f'DATE({y},{i + 2},1)'
            if kind is None:
                v = None
            elif kind == 'open':
                # 1 月（和全年）按日期算；2～12 月＝上个月的期末（少算好多遍）
                v = f'=ROUND({_open_bal(cc, ci, d0)},2)' if c in ('B', 'N') else f'={MC[i - 1]}{at["end"]}'
            elif kind in ('new', 'un') and c == 'N':
                v = f'=ROUND(SUM({MC[0]}{r}:{MC[-1]}{r}),2)'
            elif kind == 'new':
                v = f'=ROUND({_new_acc(cc, d0, d1)},2)'
            elif kind == 'un':
                crit = f'{jr(J_COI)},{ci},{ym_crit(y, i + 1)}'
                v = f'=ROUND(SUMIFS({jr(J_NET)},{crit})-SUMIFS({jr(J_NET)},{jr(J_CLS)},"?*",{crit})+SUMIFS({jr(J_NET)},{jr(J_CLS)},"？",{crit}),2)'
            elif kind == 'chk':
                v = f'=ROUND(SUM({c}{HR0}:{c}{HR0 + NACC - 1})+SUMIFS({jr(J_NET)},{jr(J_COI)},{ci},{jr(J_DATE)},"<"&{d1}),2)'
            elif kind == 'ok':
                v = (f'=IF(ABS({c}{at["end"]}-{c}{at["chk"]})<0.01,"✓","✗ 差 "&TEXT({c}{at["end"]}-{c}{at["chk"]},"#,##0.00"))')
            elif kind[0] == 'flow':
                v = '=' + flow(c, kind[1], kind[2], i).replace('{r}', str(r))
            else:
                v = '=ROUND(' + '+'.join(f'{"-" if s < 0 else ""}{c}{at[j]}' for s, j in kind[1]) + ',2)'
            if v is not None:
                put(ws, f'{c}{r}', v, F_AUTOB if big else F_AUTO, FILL_TOT if big else FILL_NONE,
                    MONEY if kind != 'ok' else None, AR if kind != 'ok' else AC)
            else:
                put(ws, f'{c}{r}', None, fill_=fill('FFD9E1F2'))
    put(ws, f'A{R0 + len(lines) + 1}', '「全年」列：期初＝1 月 1 日的余额，期末＝12 月 31 日的余额。第一年建账的账户，建账余额在「加：本期新建账户」这行。'
                                        '「核对」不按期初日期算：账户在期末前有流水就把期初余额算进去——对不上（✗）多半是【基础资料】里期初日期填晚了，或者有流水日期看不懂。',
        F_NOTE, align=AL, border=False)
    for k in range(NACC):
        r = HR0 + k
        b = AC_R0 + k
        acc, co, od = f'{SH_BASE}!$I${b}', f'{SH_BASE}!$J${b}', f'{SH_BASE}!$O${b}'
        ws[f'A{r}'] = f'={acc}&""'
        for i, c in enumerate(MC + ['N']):
            d1 = f'DATE({y}+1,1,1)' if c == 'N' else f'DATE({y},{i + 2},1)'
            ws[f'{c}{r}'] = (f'=IF(OR({acc}="",{co}="",AND($C$3<>"全部",{co}<>$C$3)),0,'
                             f'IF(OR(N({od})=0,N({od})<{d1},COUNTIFS({jr(J_ACC)},{acc},{jr(J_DATE)},"<"&{d1})>0),N({SH_BASE}!$N${b}),0))')
        ws.row_dimensions[r].hidden = True
        for c in ['A'] + MC + ['N']:
            ws[f'{c}{r}'].font = F_HELP
    ws.freeze_panes = 'B5'
    return ws


# ─────────────────────────────── 资金余额表 ───────────────────────────────
def build_bal(wb, ctx):
    ws = wb.create_sheet(SH_BAL)
    widths(ws, {'A': 6, 'B': 14, 'C': 12, 'D': 8, 'E': 22, 'F': 14, 'G': 14, 'H': 14, 'I': 14, 'J': 12, 'K': 16})
    title(ws, '资 金 余 额 表（每个账户 · 期初 ＋ 收 − 支 ＝ 期末）', 'K', C_RPT,
          '💡 选公司和起止日期（不填＝不限）。期初＝建账期初余额 ＋ 起始日之前的收支（起止日期中间才建账的账户，期初就是建账余额）。「银行核对」看资金台帐里带了银行余额的行：'
          '算出来的即时余额跟银行给的对不上的有几笔。下面按公司合计。')
    selector(ws, 'B3', '公司', 'C3', '全部', f'={AUX_CO_ALL}')
    selector(ws, 'E3', '起始日', 'F3', dt.datetime(2026, 1, 1), fmt=DATE)
    selector(ws, 'H3', '截止日', 'I3', None, fmt=DATE)
    ws['M3'] = '=IF(N($F$3)>0,$F$3,1)'
    ws['M4'] = '=IF(N($I$3)>0,$I$3+1,2958466)'
    ws['M3'].font = ws['M4'].font = F_HELP
    s, e = '$M$3', '$M$4'      # e 是截止日的下一天（不含）
    header(ws, 5, [('A', '序号'), ('B', '账户'), ('C', '所属公司'), ('D', '类型'), ('E', '账号'), ('F', '期初余额'),
                   ('G', '本期收入'), ('H', '本期支出'), ('I', '期末余额'), ('J', '最后一笔'), ('K', '银行核对')], C_RPT)
    R0 = 6
    n = AC_R1 - AC_R0 + 1
    for i in range(n):
        r = R0 + i
        b = AC_R0 + i
        acc = f'{SH_BASE}!$I${b}'
        show = f'AND({acc}<>"",OR($C$3="全部",{SH_BASE}!$J${b}=$C$3))'
        put(ws, f'A{r}', f'=IF({show},{i + 1},"")', F_AUTO, align=AC)
        put(ws, f'B{r}', f'=IF({show},{acc},"")', F_AUTOB, align=AC)
        put(ws, f'C{r}', f'=IF({show},{SH_BASE}!$J${b}&"","")', F_AUTO, align=AC)
        put(ws, f'D{r}', f'=IF({show},{SH_BASE}!$K${b}&"","")', F_AUTO, align=AC)
        put(ws, f'E{r}', f'=IF({show},{SH_BASE}!$L${b}&"","")', F_AUTO, align=AC)
        opening = (f'IF(N({SH_BASE}!$O${b})>={e},0,N({SH_BASE}!$N${b}))'
                   f'+SUMIFS({jr(J_NET)},{jr(J_ACC)},$B{r},{jr(J_DATE)},"<"&{s})')
        put(ws, f'F{r}', f'=IF($B{r}="","",ROUND({opening},2))', F_AUTO, fmt=MONEY, align=AR)
        per = f'{jr(J_ACC)},$B{r},{jr(J_DATE)},">="&{s},{jr(J_DATE)},"<"&{e}'
        put(ws, f'G{r}', f'=IF($B{r}="","",ROUND(SUMIFS({jr(J_INV)},{per}),2))', F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'H{r}', f'=IF($B{r}="","",ROUND(SUMIFS({jr(J_OUTV)},{per}),2))', F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'I{r}', f'=IF($B{r}="","",ROUND(F{r}+G{r}-H{r},2))', F_AUTOB, fmt=MONEY, align=AR)
        put(ws, f'J{r}', f'=IF($B{r}="","",{SH_BASE}!$Q${b})', F_AUTO, fmt=DATE, align=AC)
        put(ws, f'K{r}', f'=IF($B{r}="","",IF(COUNTIFS({jr(J_ACC)},$B{r},{jr(J_BCHK)},"差*")=0,'
                         f'IF(COUNTIFS({jr(J_ACC)},$B{r},{jr(J_BCHK)},"✓")=0,"（没带银行余额）","✓ 跟银行一致"),'
                         f'"✗ "&COUNTIFS({jr(J_ACC)},$B{r},{jr(J_BCHK)},"差*")&" 笔对不上"))', F_AUTO, align=AC)
    T = R0 + n
    put(ws, f'A{T}', '合计', F_TXTB, FILL_TOT, align=AC)
    ws.merge_cells(f'A{T}:E{T}')
    for c in 'FGHI':
        put(ws, f'{c}{T}', f'=ROUND(SUM({c}{R0}:{c}{T - 1}),2)', F_AUTOB, FILL_TOT, MONEY, AR)
    # 按公司
    S2 = T + 2
    section(ws, S2, 'A', 'I', '按公司合计', C_RPT)
    header(ws, S2 + 1, [('A', '序号'), ('B', '公司'), ('F', '期初余额'), ('G', '本期收入'), ('H', '本期支出'), ('I', '期末余额')], C_RPT)
    for i in range(CO_R1 - CO_R0 + 1):
        r = S2 + 2 + i
        co = f'{SH_BASE}!$B${CO_R0 + i}'
        put(ws, f'A{r}', i + 1, F_AUTO, align=AC)
        ws.merge_cells(f'B{r}:E{r}')
        put(ws, f'B{r}', f'=IF({co}="","",{co})', F_AUTOB, align=AC)
        for c in 'FGHI':
            put(ws, f'{c}{r}', f'=IF($B{r}="","",ROUND(SUMIFS({c}${R0}:{c}${T - 1},$C${R0}:$C${T - 1},$B{r}),2))',
                F_AUTO, fmt=MONEY, align=AR)
    ws.freeze_panes = f'C{R0}'
    return ws


# ─────────────────────────────── 费用统计（按公司对比） ───────────────────────────────
def build_exp(wb, ctx):
    ws = wb.create_sheet(SH_EXP)
    CC = [CL(3 + i) for i in range(8)]      # C..J：8 家公司
    widths(ws, {'A': 11, 'B': 15, **{c: 13 for c in CC}, 'K': 14, 'L': 4})
    title(ws, '费 用 统 计（按公司对比 · 收入、成本费用、利润）', 'K', C_RPT,
          '💡 选年份和月份范围（比如 1～12 是全年，9～9 是 9 月），每家公司一列，最右是几家合计。费用＝付出去减退回来。'
          '要看某一项费用的每一笔：到【资金台帐】「收支项目」那列点筛选。')
    selector(ws, 'B3', '年份', 'C3', 2026, YEARS)
    selector(ws, 'E3', '从几月', 'F3', 1, '"1,2,3,4,5,6,7,8,9,10,11,12"')
    selector(ws, 'H3', '到几月', 'I3', 12, '"1,2,3,4,5,6,7,8,9,10,11,12"')
    per = f'{jr(J_YM)},">="&($C$3*100+$F$3),{jr(J_YM)},"<="&($C$3*100+$I$3)'
    header(ws, 4, [('A', '类别'), ('B', '项目')] + [(c, f'=IF({SH_BASE}!$B${CO_R0 + i}="","",{SH_BASE}!$B${CO_R0 + i})')
                                                  for i, c in enumerate(CC)] + [('K', '合计')], C_RPT)
    ws['L4'] = '项目名'
    ws['L4'].font = F_HELP
    items_by_cls = dict(ITEMS)
    r = 5
    subs = {}
    for cls, sign, col in ((CLS_IN, +1, C_IN), (CLS_VAR, -1, C_VAR), (CLS_FIX, -1, C_FIX), (CLS_OTH, -1, C_MISC)):
        g0 = r
        fl = fill(col)
        names = items_by_cls[cls]
        last_other = names[-1] if names[-1] in ('其他收入', '其他成本') else None
        for it in [n for n in names if n != last_other] + ['__rest__']:
            rest = it == '__rest__'
            if rest and last_other:
                it = last_other
            if it != '__rest__':
                k = IT_ROW[it]
                put(ws, f'B{r}', f'=IF({SH_BASE}!$W${k}="",{SH_BASE}!$U${k}&"",{SH_BASE}!$W${k}&"")', F_TXT, fl, align=AC)
                ws[f'L{r}'] = f'={SH_BASE}!$U${k}&""'
            else:
                put(ws, f'B{r}', '其他', F_TXT, fl, align=AC)
            if rest:
                put(ws, f'B{r}', f'=IF(COUNTIFS({IT_CLSS},"{cls}",{IT_NAMES},"?*")>{len(names)},"其他（含新加项目）","其他")',
                    F_TXT, fl, align=AC)
            ws[f'L{r}'].font = F_HELP
            for k_, c in enumerate(CC):
                if rest:
                    base = f'{jr(J_CLS)},"{cls}",{jr(J_COI)},{k_ + 1},{per}'
                    v = f'{net(base, sign)}-SUM({c}{g0}:{c}{r - 1})'
                else:
                    base = f'{jr(J_ITEM)},$L{r},{jr(J_CLS)},"{cls}",{jr(J_COI)},{k_ + 1},{per}'
                    v = net(base, sign)
                put(ws, f'{c}{r}', f'=IF({c}$4="","",ROUND({v},2))', F_AUTO, fl, MONEY, AR)
            put(ws, f'K{r}', f'=ROUND(SUM(C{r}:J{r}),2)', F_AUTOB, fl, MONEY, AR)
            r += 1
        label = {'收入': '收入合计', '变动成本': '变动成本小计', '固定成本': '固定成本小计', '其他支出': '其他支出小计'}[cls]
        put(ws, f'B{r}', label, F_TXTB, FILL_SUB, align=AC)
        for c in CC + ['K']:
            put(ws, f'{c}{r}', f'=ROUND(SUM({c}{g0}:{c}{r - 1}),2)', F_AUTOB, FILL_SUB, MONEY, AR)
        subs[cls] = r
        r += 1
        ws.merge_cells(f'A{g0}:A{r - 1}')
        put(ws, f'A{g0}', cls, F_TXTB, fl, align=ACW)
        for rr in range(g0, r):
            ws[f'A{rr}'].border = BD
    for label, f in (('变动利润', lambda c: f'{c}{subs[CLS_IN]}-{c}{subs[CLS_VAR]}'),
                     ('利润', lambda c: f'{c}{subs[CLS_IN]}-{c}{subs[CLS_VAR]}-{c}{subs[CLS_FIX]}'),
                     ('利润（再扣其他支出）', lambda c: f'{c}{subs[CLS_IN]}-{c}{subs[CLS_VAR]}-{c}{subs[CLS_FIX]}-{c}{subs[CLS_OTH]}')):
        ws.merge_cells(f'A{r}:B{r}')
        put(ws, f'A{r}', label, F_RED, FILL_TOT, align=AC)
        ws[f'B{r}'].border = BD
        for c in CC + ['K']:
            put(ws, f'{c}{r}', f'=ROUND({f(c)},2)', F_AUTOB, FILL_TOT, MONEY, AR)
        r += 1
    hide(ws, 'L')
    ws.freeze_panes = 'C5'
    return ws


# ─────────────────────────────── 客户收入统计 ───────────────────────────────
CU_R0, CU_N = 14, 600


def build_cus(wb, ctx):
    ws = wb.create_sheet(SH_CUS)
    MC = [CL(2 + i) for i in range(12)]
    widths(ws, {'A': 6, 'B': 28, 'C': 12, 'D': 12, 'E': 12, 'F': 12, 'G': 12, 'H': 13, 'I': 13, 'J': 13, 'K': 13, 'L': 12,
                'M': 12, 'N': 13, 'O': 4, 'P': 4, 'Q': 4})
    title(ws, '客 户 收 入 统 计（新增 · 续费 · 每月新客户）', 'N', C_RPT,
          '💡 同一家公司第一次收到这个客户的钱＝「新增」，以后＝「续费」（在资金台帐自动认，个别不对的在资金台帐 L 列改）。'
          '上面是每个月新增了几个客户、新增和续费各收了多少；下面每个客户一行：首次来款、今年新增/续费/其他项目、累计收款、开票和应收余额。')
    _sel_block(ws)
    cc, y = co_crit('$C$3'), '$F$3'
    ci = coi_crit('$C$3')
    header(ws, 4, [('A', '')] + [(c, f'{i + 1}月') for i, c in enumerate(MC)] + [('N', '全年')], C_RPT)
    put(ws, 'A4', '项目', F_HDR, fill(C_RPT), align=AC)
    lines = [('新增客户数', 'cnt'), ('新增收入', '新增'), ('续费收入', '续费'), ('其他收入项目', 'oth'), ('收入合计', 'all')]
    for k, (lab, kind) in enumerate(lines):
        r = 5 + k
        put(ws, f'A{r}', lab, F_TXTB, FILL_SUB, align=AC)
        ws.column_dimensions['A'].width = 12
        for i, c in enumerate(MC + ['N']):
            base = f'{jr(J_COI)},{ci},{ym_crit(y, i + 1)}' if c != 'N' else None
            if c == 'N' and kind != 'oth':
                v = f'=ROUND(SUM({MC[0]}{r}:{MC[-1]}{r}),2)'
            elif kind == 'cnt':
                v = f'=COUNTIFS({jr(J_ITEM)},"新增",{jr(J_INV)},">0",{base})'
            elif kind in ('新增', '续费'):
                v = f'=ROUND({net(f"{jr(J_ITEM)},{Q}{kind}{Q},{base}")},2)'
            elif kind == 'all':
                v = f'=ROUND({net(f"{jr(J_CLS)},{Q}{CLS_IN}{Q},{base}")},2)'
            else:
                v = f'=ROUND({c}{5 + 4}-{c}{5 + 1}-{c}{5 + 2},2)'
            put(ws, f'{c}{r}', v, F_AUTOB if kind in ('cnt', 'all') else F_AUTO, FILL_NONE, INT if kind == 'cnt' else MONEY, AR)
    # 客户明细
    H = CU_R0 - 1
    section(ws, H - 1, 'A', 'N', '每个客户（【往来单位】里类型是「客户」的，按登记顺序）', C_RPT)
    header(ws, H, [('A', '序号'), ('B', '客户'), ('C', '首次来款'), ('D', '最近来款'), ('E', '本年新增'), ('F', '本年续费'),
                   ('G', '本年其他项目'), ('H', '本年收入合计'), ('I', '累计收款\n（到年底）'), ('J', '本年开票'),
                   ('K', '应收余额\n（到年底）'), ('L', '本年笔数'), ('M', '状态'), ('N', '备注')], C_RPT)
    # 辅助：往来单位里类型＝客户的，压紧排
    for i in range(PT_R1 - PT_R0 + 1):
        pr_ = PT_R0 + i
        ws[f'P{CU_R0 + i}'] = f'=IF({SH_PARTY}!$C${pr_}="客户",1,0)'
        ws[f'Q{CU_R0 + i}'] = f'=N(Q{CU_R0 + i - 1})+P{CU_R0 + i}' if i else f'=P{CU_R0 + i}'
        ws[f'P{CU_R0 + i}'].font = ws[f'Q{CU_R0 + i}'].font = F_HELP
    last = f'$Q${CU_R0 + PT_R1 - PT_R0}'
    yend = f'DATE({y}+1,1,1)'
    for i in range(CU_N):
        r = CU_R0 + i
        k = i + 1
        n = f'$B{r}'
        ne = f'$O{r}'           # 隐藏：名字转义后当条件用
        put(ws, f'A{r}', f'=IF({n}="","",{k})', F_AUTO, align=AC)
        put(ws, f'B{r}', f'=IF({k}>{last},"",INDEX({pr(PT_NAME)},MATCH({k},$Q${CU_R0}:{last},0)))', F_AUTOB, align=AL)
        ws[f'O{r}'] = f'={esc(n)}'
        ws[f'O{r}'].font = F_HELP
        pc = f'{jr(J_PARTY)},{ne},{jr(J_COI)},{ci}'
        upto = f'{jr(J_DATE)},"<"&DATE({y}+1,1,1)'        # 看的是所选那年年底为止
        # 没来过款时 MINIFS 得 0，格式里 0 不显示（少算一遍 COUNTIFS）
        put(ws, f'C{r}', f'=IF({n}="","",_xlfn.MINIFS({jr(J_DATE)},{pc},{jr(J_INV)},">0",{jr(J_CLS)},"{CLS_IN}",{upto}))',
            F_AUTO, fmt=DATE + ';;', align=AC)
        put(ws, f'D{r}', f'=IF(OR({n}="",N(C{r})=0),"",_xlfn.MAXIFS({jr(J_DATE)},{pc},{jr(J_INV)},">0",{jr(J_CLS)},"{CLS_IN}",{upto}))',
            F_AUTO, fmt=DATE, align=AC)
        yc = f'{pc},{year_crit(y)}'
        put(ws, f'E{r}', f'=IF({n}="","",ROUND({net(f"{yc},{jr(J_ITEM)},{Q}新增{Q}")},2))',
            F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'F{r}', f'=IF({n}="","",ROUND({net(f"{yc},{jr(J_ITEM)},{Q}续费{Q}")},2))',
            F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'H{r}', f'=IF({n}="","",ROUND({net(f"{yc},{jr(J_CLS)},{Q}{CLS_IN}{Q}")},2))',
            F_AUTOB, fmt=MONEY, align=AR)
        put(ws, f'G{r}', f'=IF({n}="","",ROUND(H{r}-E{r}-F{r},2))', F_AUTO, fmt=MONEY, align=AR)
        tc = f'{pc},{jr(J_DATE)},"<"&{yend}'
        put(ws, f'I{r}', f'=IF({n}="","",ROUND({net(f"{tc},{jr(J_CLS)},{Q}{CLS_IN}{Q}")},2))',
            F_AUTO, fmt=MONEY, align=AR)
        vc = f'{vr(V_PARTY)},{ne},{vr(V_CO)},{cc}'
        put(ws, f'J{r}', f'=IF({n}="","",ROUND(SUMIFS({vr(V_ARV)},{vc},{vr(V_DATE)},">="&DATE({y},1,1),{vr(V_DATE)},"<"&{yend}),2))',
            F_AUTO, fmt=MONEY, align=AR)
        op = f'SUMIFS({br(OP_AR, OP_R0, OP_R1)},{br(OP_PARTY, OP_R0, OP_R1)},{ne},{br(OP_CO, OP_R0, OP_R1)},{cc})'
        put(ws, f'K{r}', f'=IF({n}="","",ROUND({op}+SUMIFS({vr(V_ARV)},{vc},{vr(V_DATE)},"<"&{yend})-I{r},2))',
            F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'L{r}', f'=IF({n}="","",COUNTIFS({yc},{jr(J_INV)},">0",{jr(J_CLS)},"{CLS_IN}"))', F_AUTO, fmt=INT, align=AC)
        put(ws, f'M{r}', f'=IF({n}="","",IF(N(C{r})=0,"还没来过款",IF(E{r}>0,"本年新客户",'
                         f'IF(D{r}<DATE({y},1,1),"今年没来款",""))))', F_AUTO, align=AC)
        put(ws, f'N{r}', f'=IF({n}="","",IF(K{r}<0,"已收款未开票/预收",IF(K{r}>0,"开了票还没收齐","")))', F_NOTE, align=AL)
    hide(ws, 'O', 'P', 'Q')
    ws.freeze_panes = f'C{CU_R0}'
    return ws


# ─────────────────────────────── 发票汇总 ───────────────────────────────
def build_invs(wb, ctx):
    ws = wb.create_sheet(SH_INVS)
    widths(ws, {'A': 8, 'B': 9, 'C': 14, 'D': 12, 'E': 14, 'F': 9, 'G': 14, 'H': 14, 'I': 14, 'J': 14, 'K': 14, 'L': 14})
    title(ws, '发 票 汇 总（每月销项、进项 · 按发票导入）', 'L', C_INV,
          '💡 选公司、年份。销项＝开出去的票（应收），进项分「计应付」（登记过的供应商）和「报销票」。重复粘贴、作废的不算，红冲的正负抵掉。'
          '选单家公司时，自家公司之间开的票也算进这家的销项/进项（跟电子税务局一致）；选「全部」时内部的票两边抵掉，只在最右列显示。'
          'J、K 两列是同月资金台帐里实际收到的收入类款项、付出的成本费用类款项（变动成本＋固定成本＋其他支出），对照看开票和收款的差。')
    _sel_block(ws)
    cc, y = co_crit('$C$3'), '$F$3'
    ci = coi_crit('$C$3')
    header(ws, 4, [('A', '月份'), ('B', '销项张数'), ('C', '销项金额\n（不含税）'), ('D', '销项税额'), ('E', '销项价税合计'),
                   ('F', '进项张数'), ('G', '进项价税合计'), ('H', '其中计应付'), ('I', '其中报销票'),
                   ('J', '本月收入类收款'), ('K', '本月成本费用付款'), ('L', '自家公司之间的票\n（选单家时已含在\n销项/进项里）')], C_INV)
    one = '$C$3<>"全部"'
    for m in range(1, 14):
        r = 4 + m
        if m <= 12:
            vper = f'{vr(V_DATE)},">="&DATE({y},{m},1),{vr(V_DATE)},"<"&DATE({y},{m}+1,1)'
            jper = ym_crit(y, m)
            put(ws, f'A{r}', f'{m}月', F_TXTB, FILL_SUB, align=AC)
        else:
            vper = f'{vr(V_DATE)},">="&DATE({y},1,1),{vr(V_DATE)},"<"&DATE({y}+1,1,1)'
            jper = year_crit(y)
            put(ws, f'A{r}', '全年', F_TXTB, FILL_TOT, align=AC)
        s_ = f'{vr(V_DIR)},"销项",{vr(V_CO)},{cc},{vper}'
        p_ = f'{vr(V_DIR)},"进项",{vr(V_CO)},{cc},{vper}'
        si = f'{vr(V_DIR)},"内部",{vr(V_SCO)},$C$3,{vper}'        # 选单家：它开给自家另一家的票
        pi = f'{vr(V_DIR)},"内部",{vr(V_BCO)},$C$3,{vper}'        # 选单家：自家另一家开给它的票
        both = lambda f1, f2: f'{f1}+IF({one},{f2},0)'
        cost = '+'.join(net(f'{jr(J_CLS)},"{k}",{jr(J_COI)},{ci},{jper}') for k in CLS_COST)
        vals = {'B': both(f'COUNTIFS({s_},{vr(V_EFF)},"<>0")', f'COUNTIFS({si},{vr(V_EFF)},"<>0")'),
                'C': f'ROUND({both(f"SUMIFS({vr(V_EAMT)},{s_})", f"SUMIFS({vr(V_EAMT)},{si})")},2)',
                'D': f'ROUND({both(f"SUMIFS({vr(V_ETAX)},{s_})", f"SUMIFS({vr(V_ETAX)},{si})")},2)',
                'E': f'ROUND({both(f"SUMIFS({vr(V_EFF)},{s_})", f"SUMIFS({vr(V_EFF)},{si})")},2)',
                'F': both(f'COUNTIFS({p_},{vr(V_EFF)},"<>0")', f'COUNTIFS({pi},{vr(V_EFF)},"<>0")'),
                'G': f'ROUND({both(f"SUMIFS({vr(V_EFF)},{p_})", f"SUMIFS({vr(V_EFF)},{pi})")},2)',
                'H': f'ROUND(SUMIFS({vr(V_EFF)},{p_},{vr(V_USE)},"应付"),2)',
                'I': f'ROUND(SUMIFS({vr(V_EFF)},{p_},{vr(V_USE)},"报销票"),2)',
                'J': f'ROUND({net(f"{jr(J_CLS)},{Q}{CLS_IN}{Q},{jr(J_COI)},{ci},{jper}")},2)',
                'K': f'ROUND(-({cost}),2)',
                'L': f'ROUND(IF({one},SUMIFS({vr(V_EFF)},{si})+SUMIFS({vr(V_EFF)},{pi}),SUMIFS({vr(V_EFF)},{vr(V_DIR)},"内部",{vper})),2)'}
        for c, v in vals.items():
            put(ws, f'{c}{r}', '=' + v, F_AUTOB if m == 13 else F_AUTO, FILL_TOT if m == 13 else FILL_NONE,
                INT if c in 'BF' else MONEY, AR)
    ws.freeze_panes = 'B5'
    return ws
