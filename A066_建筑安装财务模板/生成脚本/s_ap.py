# -*- coding: utf-8 -*-
"""查看表：【应付账款】【应付对账】（全是公式，只有亮黄格可以选）。

【应付账款】截至日期、本期起、类型（只筛①）：
   汇总（按类型）→ ① 按供应商（只列有数的）→ ② 按项目（只列有数的工程项目＋没分项目）。
【应付对账】选供应商（可再选项目）＋起止：抬头、汇总（本期 / 截至止日 / 按税率）、④ 按项目小计，
   下面三张清单并排：① 应付明细 ② 付款明细 ③ 开票明细（按日期排，各最多 300 条，合计跟汇总核对）。

口径（报表口径.md §4）：
   应付＝期初（期初余额 类型「应付账款」、对象＝供应商）＋ 付_应付额（按 付_日期）；
   已付＝−收_净额（收_供应商＝它、收_冲应付＝1，含专户、个人户付的）；未付＝应付−已付（负数＝多付了）；
   已开票＝付_已开票（按 付_开票日期）；未开票＝应付登记−已开票（期初不算票）；进项税＝付_进项税（专票）；
   直接付款＝供应商在应付登记里没有行时付给他的钱（收_冲应付＝0，归类 材料款/分包款/机械费），直接进成本。
   类型：应付、开票按【应付登记】每行的类型（付_类型）；期初、已付按【供应商信息】这家的类型（供应商_类型）。"""
from layout import *
from common import *

ALL = '"<>@@全部@@"'                    # SUMIFS 条件「全部都算」（含空格子），不用通配符
RATES = [(0.13, '13%'), (0.09, '9%'), (0.06, '6%'), (0.03, '3%'), (0.01, '1%')]
RATE_FMT = '0%;-0%;"-"'
HR0 = 6                                  # 隐藏辅助列从第 6 行起（计数格在第 5 行）
TYPES = AP_TYPES                         # 材料 分包 机械 其他
COST3 = (('材料款', '材料'), ('分包款', '分包'), ('机械费', '机械'))
ZP = '付_发票类型,"专票"'
AT = lambda t: f'收_账户类型,"{t}"'
F_STMT = Font(name=YH, sz=15, bold=True, color='FF1F3864')
FILL_LBL = fill('FFD9E1F2')


# ───────────────────────── 小工具 ─────────────────────────
def _lbl(ws, coord, text, align=AC):
    put(ws, coord, text, F_KPI_L, FILL_LBL, align=align)


def _note(ws, coord, value, align=AL, font=F_NOTE):
    return put(ws, coord, value, font, align=align, border=False)


def _merge(ws, rng):
    ws.merge_cells(rng)


def _grid(ws, r0, r1, cols, fmts=None, aligns=None, font=F_TXT):
    fmts, aligns = fmts or {}, aligns or {}
    for r in range(r0, r1 + 1):
        for c in cols:
            cell = ws[f'{c}{r}']
            cell.font = font
            cell.border = BD
            cell.alignment = aligns.get(c, AR if c in fmts and fmts[c] in (MONEY, MONEY0) else AC)
            if c in fmts:
                cell.number_format = fmts[c]


def _tot(ws, r, cols, fmts=None):
    fmts = fmts or {}
    for c in cols:
        cell = ws[f'{c}{r}']
        cell.font = F_TXTB
        cell.fill = FILL_TOT
        cell.border = BD
        fm = fmts.get(c, MONEY)
        cell.alignment = AR if fm == MONEY else AC
        cell.number_format = fm or 'General'


def _skey(ws, col, ncol, r0, n, cond, datef):
    """按日期排序用的键（跟 common.skey 一样，只是「第几条」取本表常数列 ncol，整列公式一样、文件小）：
       满足条件的行＝日期×10000＋第几条，不满足＝空；返回放个数的格子"""
    for i in range(n):
        r = r0 + i
        ws[f'{col}{r}'] = f'=IF({cond(r)},{datef(r)}*10000+${ncol}{r},"")'
    ws[f'{col}{r0 - 1}'] = f'=COUNT({col}{r0}:{col}{r0 + n - 1})'
    return f'${col}${r0 - 1}'


def _ncol(ws, col, r0, n):
    """常数列：第 i 行＝i（1 起），给 INDEX(名称, 第几条) 用"""
    for i in range(n):
        ws[f'{col}{r0 + i}'] = i + 1


# ═════════════════════════════════════ 应付账款 ═════════════════════════════════════
def build_ap(ws, ctx):
    LAST = 'V'
    widths(ws, {'A': 5, 'B': 16, **{CL(i): 12.5 for i in range(3, 23)}})
    tip = ('💡 每家材料商、分包、机械：应付多少、付了多少、还欠多少、票到了多少（按税率分）。亮黄格可以改：截至日期（空＝首页截止日）、'
           '本期起（空＝首页年初）、类型（只筛①按供应商；空＝全部）。应付＝期初＋【应付登记】；已付＝【收支登记】里「供应商名称」选了这家的付款'
           '（专户、老板/员工微信付的都算）；未付是红色负数＝多付了。这家没登记过应付、直接付的钱算「直接付款」（直接进成本，不冲应付）。'
           '想看某一家每一笔、打对账单，去【应付对账】。')
    title(ws, '应付账款', LAST, C_VIEW, tip)
    # ── 选择格（第 3 行） ──
    E_, S_, T_, TC = '$D$3', '$I$3', '$N$3', '$Z$3'
    selector(ws, 'B3', '截至日期', 'C3', None, fmt=DATE)
    dv_date(ws, 'C3')
    put(ws, 'D3', '=IF(ISNUMBER(C3),INT(C3),P_截止)', F_AUTOB, FILL_AUTO, DATE, AC)
    _note(ws, 'E3', '（空＝首页截止日）')
    selector(ws, 'G3', '本期起', 'H3', None, fmt=DATE)
    dv_date(ws, 'H3')
    put(ws, 'I3', '=IF(ISNUMBER(H3),INT(H3),P_年初)', F_AUTOB, FILL_AUTO, DATE, AC)
    _note(ws, 'J3', '（空＝首页年初）')
    selector(ws, 'L3', '类型', 'M3', '全部', dv_formula='"全部,' + ','.join(TYPES) + '"', prompt='只筛下面①按供应商')
    put(ws, 'N3', '=IF(OR(TRIM(M3&"")="",TRIM(M3&"")="全部"),"全部",TRIM(M3&""))', F_AUTOB, FILL_AUTO, align=AC)
    _note(ws, 'O3', '（只筛①；空＝全部）')
    home_link(ws, 'V3')
    ws['Z3'] = f'=IF({T_}="全部",{ALL},"="&{T_})'                     # 付_类型 的条件
    ws.row_dimensions[3].height = 24

    # ── 隐藏辅助：供应商名单（第 i 家在 HR0+i−1 行）、收支逐条的「冲应付供应商的类型」 ──
    s0, s1 = HR0, HR0 + N_SUP - 1
    for i in range(N_SUP):
        r = s0 + i
        ws[f'BA{r}'] = f'=INDEX(供应商_名称,{i + 1})&""'
        ws[f'BB{r}'] = f'=IF(BA{r}="",0,SUMIFS(期初_金额,期初_类型,"应付账款",期初_对象,{esc(f"BA{r}")}))'
        ws[f'BC{r}'] = f'=IF(BA{r}="",0,IF(OR({T_}="全部",INDEX(供应商_类型,{i + 1})={T_}),1,0))'
    pay3 = lambda nm, d: '+'.join(f'COUNTIFS(收_供应商,{esc(nm)},收_归类,"{g}",收_日期,"<="&{d})' for g, _t in COST3)

    def sup_cond(i):
        r = s0 + i
        return (f'AND(BA{r}<>"",OR(AND(BC{r}=1,OR(BB{r}<>0,{pay3(f"BA{r}", E_)}>0)),'
                f'COUNTIFS(付_供应商,{esc(f"BA{r}")},付_有效,1,付_类型,{TC},付_日期,"<="&{E_})>0))')
    counter(ws, 'BD', s0, N_SUP, sup_cond)
    n_sup = cnt('BD', s0, N_SUP)
    OPEN_RNG = f'$BB${s0}:$BB${s1}'
    h0, h1 = HR0, HR0 + N_CASH - 1
    _ncol(ws, 'BE', h0, N_CASH)
    for i in range(N_CASH):
        r = h0 + i
        ws[f'BF{r}'] = (f'=IF(INDEX(收_冲应付,$BE{r})=1,IFERROR(INDEX(供应商_类型,MATCH(INDEX(收_供应商,$BE{r}),供应商_名称,0))&"",'
                        f'"其他"),"")')
    HT = f'$BF${h0}:$BF${h1}'                                             # 跟 收_xxx 一样 5000 行
    p0 = HR0

    def pj_cond(i):
        nm = f'INDEX(项目_名称,{i + 1})'
        return (f'AND({nm}<>"",OR(COUNTIFS(付_项目,{esc(nm)},付_有效,1,付_日期,"<="&{E_})>0,'
                f'COUNTIFS(收_项目,{esc(nm)},收_冲应付,1,收_日期,"<="&{E_})>0,'
                f'SUMIFS(期初_金额,期初_类型,"应付账款",期初_项目,{esc(nm)})<>0))')
    counter(ws, 'BH', p0, N_PJ, pj_cond)
    n_pj = cnt('BH', p0, N_PJ)

    # ── 汇总：按类型 ──
    R_S = 5
    section(ws, R_S, 'A', LAST, '汇总：按类型（截至日；包工包料的分包，材料那部分算材料）', C_VIEW)
    hdr = [('A', ''), ('B', '类型'), ('C', '期初应付'), ('D', '应付总额\n（期初＋登记）'), ('E', '已付'), ('F', '未付\n（负数＝多付了）'),
           ('G', '已开票'), ('H', '未开票\n（登记−已开票）'), ('I', '进项税\n（专票）')]
    header(ws, R_S + 1, hdr, C_VIEW)
    rt0 = R_S + 2
    for j, t in enumerate(TYPES):
        r = rt0 + j
        ws[f'A{r}'] = j + 1
        ws[f'B{r}'] = t
        ws[f'C{r}'] = f'=SUMIFS({OPEN_RNG},供应商_类型,B{r})'
        ws[f'D{r}'] = f'=C{r}+SUMIFS(付_应付额,付_类型,B{r},付_日期,"<="&{E_})'
        ws[f'E{r}'] = f'=-SUMIFS(收_净额,{HT},B{r},收_日期,"<="&{E_})'
        ws[f'F{r}'] = f'=D{r}-E{r}'
        ws[f'G{r}'] = f'=SUMIFS(付_已开票,付_类型,B{r},付_开票日期,"<="&{E_})'
        ws[f'H{r}'] = f'=D{r}-C{r}-G{r}'
        ws[f'I{r}'] = f'=SUMIFS(付_进项税,付_类型,B{r},付_开票日期,"<="&{E_})'
    rt1 = rt0 + len(TYPES) - 1
    RT = rt1 + 1
    _grid(ws, rt0, rt1, list('ABCDEFGHI'), {c: MONEY for c in 'CDEFGHI'})
    ws[f'B{RT}'] = '合计'
    for c in 'CDEFGHI':
        ws[f'{c}{RT}'] = f'=SUM({c}{rt0}:{c}{rt1})'
    _tot(ws, RT, list('ABCDEFGHI'), {'A': None, 'B': None})
    ws[f'B{RT}'].alignment = AC
    _note(ws, f'K{R_S + 1}', '说明：应付、已开票按【应付登记】每一行的「类型」分（分包商包工包料的材料行算材料）；期初、已付按【供应商信息】里这家的类型分。'
          '已付只算冲应付的付款；没登记应付直接付的钱不在这里（看①最后一列「直接付款」）。', align=ALW)
    _merge(ws, f'K{R_S + 1}:R{RT}')

    # ── ① 按供应商 ──
    R1 = RT + 3
    section(ws, R1, 'A', LAST, '① 按供应商（只列有数的；类型选了就只看这一类）', C_VIEW)
    rnote, rg, rh, R1T = R1 + 1, R1 + 2, R1 + 3, R1 + 4
    r10, NS = R1T + 1, 150
    r11 = r10 + NS - 1
    for c1, c2, t in (('J', 'O', '已开票按税率分（截至日，按开票日期）'), ('T', 'V', '本期（本期起～截至日）')):
        _merge(ws, f'{c1}{rg}:{c2}{rg}')
        put(ws, f'{c1}{rg}', t, F_HDR, fill(C_VIEW), align=AC)
        for k in range(CI(c1) + 1, CI(c2) + 1):
            ws.cell(row=rg, column=k).border = BD
    hdr = [('A', '序号'), ('B', '供应商'), ('C', '类型'), ('D', '期初应付'), ('E', '应付总额\n（期初＋登记）'), ('F', '已付'),
           ('G', '未付\n（负数＝多付了）'), ('H', '已开票'), ('I', '未开票\n（登记−已开票）')]
    hdr += [(CL(10 + j), lab) for j, (_v, lab) in enumerate(RATES)] + [('O', '其他税率')]
    hdr += [('P', '专票'), ('Q', '普票\n（不是专票的）'), ('R', '进项税\n（专票）'), ('S', '直接付款\n（没登记应付）'),
            ('T', '本期应付'), ('U', '本期已付'), ('V', '本期开票')]
    header(ws, rh, hdr, C_VIEW, height=40)
    for k in range(NS):
        r = r10 + k
        W, X, Y = f'$W{r}', f'$X{r}', f'$Y{r}'
        ws[f'W{r}'] = f'={kth(k + 1, "BD", s0, N_SUP)}'
        ws[f'X{r}'] = f'=IF({W}=0,0,INDEX($BC${s0}:$BC${s1},{W}))'
        ws[f'Y{r}'] = f'=IF({W}=0,"",{esc(f"$B{r}")})'
        g = lambda body: f'=IF({W}=0,"",{body})'
        fu = lambda *more: ','.join((f'付_供应商,{Y}', f'付_类型,{TC}') + more)
        fu_d, fu_s = f'付_日期,"<="&{E_}', f'付_日期,">="&{S_}'
        fk_d, fk_s = f'付_开票日期,"<="&{E_}', f'付_开票日期,">="&{S_}'
        ws[f'A{r}'] = g(k + 1)
        ws[f'B{r}'] = g(f'INDEX(供应商_名称,{W})')
        ws[f'C{r}'] = g(f'INDEX(供应商_类型,{W})')
        ws[f'D{r}'] = g(f'{X}*INDEX({OPEN_RNG},{W})')
        ws[f'E{r}'] = g(f'D{r}+SUMIFS(付_应付额,{fu(fu_d)})')
        ws[f'F{r}'] = g(f'-{X}*SUMIFS(收_净额,收_供应商,{Y},收_冲应付,1,收_日期,"<="&{E_})')
        ws[f'G{r}'] = g(f'E{r}-F{r}')
        ws[f'H{r}'] = g(f'SUMIFS(付_已开票,{fu(fk_d)})')
        ws[f'I{r}'] = g(f'E{r}-D{r}-H{r}')
        for j, (rv, _lab) in enumerate(RATES):
            ws[f'{CL(10 + j)}{r}'] = g(f'SUMIFS(付_已开票,{fu(f"付_税率,{rv}", fk_d)})')
        ws[f'O{r}'] = g(f'ROUND(H{r}-SUM(J{r}:N{r}),2)')
        ws[f'P{r}'] = g(f'SUMIFS(付_已开票,{fu(ZP, fk_d)})')
        ws[f'Q{r}'] = g(f'H{r}-P{r}')
        ws[f'R{r}'] = g(f'SUMIFS(付_进项税,{fu(fk_d)})')
        ws[f'S{r}'] = g(f'-{X}*(' + '+'.join(f'SUMIFS(收_净额,收_供应商,{Y},收_冲应付,0,收_归类,"{g_}",收_日期,"<="&{E_})'
                                             for g_, _t in COST3) + ')')
        ws[f'T{r}'] = g(f'SUMIFS(付_应付额,{fu(fu_s, fu_d)})')
        ws[f'U{r}'] = g(f'-{X}*SUMIFS(收_净额,收_供应商,{Y},收_冲应付,1,收_日期,">="&{S_},收_日期,"<="&{E_})')
        ws[f'V{r}'] = g(f'SUMIFS(付_已开票,{fu(fk_s, fk_d)})')
    money1 = [CL(i) for i in range(4, 23)]
    _grid(ws, r10, r11, ['A', 'B', 'C'] + money1, {c: MONEY for c in money1}, {'B': AL})
    ws[f'B{R1T}'] = '合计'
    for c in money1:
        ws[f'{c}{R1T}'] = f'=SUM({c}{r10}:{c}{r11})'
    _tot(ws, R1T, ['A', 'B', 'C'] + money1, {'A': None, 'B': None, 'C': None})
    ws[f'B{R1T}'].alignment = AC
    # 核对：类型＝全部时，①合计＝上面按类型的合计
    chk = (f'=IF({T_}<>"全部","（类型选了「"&{T_}&"」，①只列这一类，不跟上面核对）",'
           f'IF(AND(ROUND(E{R1T}-D{RT},2)=0,ROUND(F{R1T}-E{RT},2)=0,ROUND(H{R1T}-G{RT},2)=0),'
           f'"✓ 核对：①按供应商合计＝上面按类型合计",'
           f'"✗ ①合计跟按类型合计差 "&TEXT(E{R1T}-D{RT},"#,##0.00")&"（可能有供应商不在【供应商信息】里，或超过 {NS} 家没列全）"))')
    cnt1 = f'IF({n_sup}>{NS},"共 "&{n_sup}&" 家，只显示前 {NS} 家（请选类型）","共 "&{n_sup}&" 家")'
    put(ws, f'A{rnote}', f'={cnt1}&"　　"&{chk[1:]}', F_NOTE, align=AL, border=False)
    _merge(ws, f'A{rnote}:{LAST}{rnote}')

    # ── ② 按项目 ──
    R2 = r11 + 2
    section(ws, R2, 'A', LAST, '② 按项目（截至日；只列有数的工程项目）', C_VIEW)
    rnote2, rg2, rh2, R2T, R2N = R2 + 1, R2 + 2, R2 + 3, R2 + 4, R2 + 5
    r20, NP = R2N + 1, 40
    r21 = r20 + NP - 1
    groups = [(TYPES[j], CL(4 + 3 * j), CL(6 + 3 * j)) for j in range(4)] + [('合计', 'P', 'T')]
    for t, c1, c2 in groups:
        _merge(ws, f'{c1}{rg2}:{c2}{rg2}')
        put(ws, f'{c1}{rg2}', t, F_HDR, fill(C_VIEW), align=AC)
        for k in range(CI(c1) + 1, CI(c2) + 1):
            ws.cell(row=rg2, column=k).border = BD
    hdr = [('A', '序号'), ('B', '项目'), ('C', '期初应付')]
    for j in range(4):
        hdr += [(CL(4 + 3 * j), '应付'), (CL(5 + 3 * j), '已开票'), (CL(6 + 3 * j), '已付')]
    hdr += [('P', '应付总额\n（含期初）'), ('Q', '已付'), ('R', '未付\n（负数＝多付了）'), ('S', '已开票'), ('T', '未开票\n（登记−已开票）')]
    header(ws, rh2, hdr, C_VIEW, height=40)

    def pj_row(r, pc, oc, guard=None):
        """pc：付_项目/收_项目 的条件；oc：期初_项目 的条件"""
        g = (lambda body: f'=IF({guard},"",{body})') if guard else (lambda body: f'={body}')
        pf = f',付_项目,{pc}' if pc else ''
        sf = f',收_项目,{pc}' if pc else ''
        of = f',期初_项目,{oc}' if oc else ''
        ws[f'C{r}'] = g(f'SUMIFS(期初_金额,期初_类型,"应付账款"{of})')
        for j, t in enumerate(TYPES):
            a, b, c = CL(4 + 3 * j), CL(5 + 3 * j), CL(6 + 3 * j)
            ws[f'{a}{r}'] = g(f'SUMIFS(付_应付额,付_类型,"{t}"{pf},付_日期,"<="&{E_})')
            ws[f'{b}{r}'] = g(f'SUMIFS(付_已开票,付_类型,"{t}"{pf},付_开票日期,"<="&{E_})')
            ws[f'{c}{r}'] = g(f'-SUMIFS(收_净额,{HT},"{t}"{sf},收_日期,"<="&{E_})')
        ws[f'P{r}'] = g(f'C{r}+D{r}+G{r}+J{r}+M{r}')
        ws[f'Q{r}'] = g(f'F{r}+I{r}+L{r}+O{r}')
        ws[f'R{r}'] = g(f'P{r}-Q{r}')
        ws[f'S{r}'] = g(f'E{r}+H{r}+K{r}+N{r}')
        ws[f'T{r}'] = g(f'P{r}-C{r}-S{r}')

    money2 = [CL(i) for i in range(3, 21)]
    for k in range(NP):
        r = r20 + k
        W, Y = f'$W{r}', f'$Y{r}'
        ws[f'W{r}'] = f'={kth(k + 1, "BH", p0, N_PJ)}'
        ws[f'Y{r}'] = f'=IF({W}=0,"",{esc(f"$B{r}")})'
        ws[f'A{r}'] = f'=IF({W}=0,"",{k + 1})'
        ws[f'B{r}'] = f'=IF({W}=0,"",INDEX(项目_名称,{W}))'
        pj_row(r, Y, Y, guard=f'{W}=0')
    _grid(ws, r20, r21, ['A', 'B'] + money2, {c: MONEY for c in money2}, {'B': AL})
    ws[f'B{R2T}'] = '合计'
    pj_row(R2T, None, None)
    _tot(ws, R2T, ['A', 'B'] + money2, {'A': None, 'B': None})
    ws[f'B{R2T}'].alignment = AC
    ws[f'B{R2N}'] = '没分项目的'
    pj_row(R2N, '""', '""')
    _grid(ws, R2N, R2N, ['A', 'B'] + money2, {c: MONEY for c in money2}, {'B': AL}, font=F_TXTB)
    for c in ['A', 'B'] + money2:
        ws[f'{c}{R2N}'].fill = FILL_SUB
    ws[f'A{R2N}'] = '—'
    chk2 = (f'=IF(AND(ROUND(SUM(P{R2N}:P{r21})-P{R2T},2)=0,ROUND(SUM(Q{R2N}:Q{r21})-Q{R2T},2)=0,ROUND(SUM(S{R2N}:S{r21})-S{R2T},2)=0),'
            f'"✓ 各项目＋没分项目＝合计","✗ 各项目加起来跟合计差 "&TEXT(SUM(P{R2N}:P{r21})-P{R2T},"#,##0.00")&'
            f'"（项目超过 {NP} 个没列全，或期初余额的项目不是工程项目）")')
    cnt2 = f'IF({n_pj}>{NP},"共 "&{n_pj}&" 个项目，只列前 {NP} 个","共 "&{n_pj}&" 个项目")'
    put(ws, f'A{rnote2}', f'={cnt2}&"　　"&{chk2[1:]}', F_NOTE, align=AL, border=False)
    _merge(ws, f'A{rnote2}:{LAST}{rnote2}')
    # 跳转
    _note(ws, 'R3', '跳到：')
    for coord, txt, row in (('S3', '① 按供应商', R1), ('T3', '② 按项目', R2)):
        c = put(ws, coord, txt, Font(name=YH, sz=10, color='FF0563C1', underline='single'), align=AC, border=False)
        link(c, ws.title, f'A{row}')

    hide(ws, 'W', 'X', 'Y', 'Z', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BH')
    ws.freeze_panes = 'C4'
    print_setup(ws, '1:3', landscape=True)
    ws.print_area = f'A1:{LAST}{r21}'
    return dict(type_rows=(rt0, rt1), type_total=RT, sup_total=R1T, sup_rows=(r10, r11), pj_total=R2T, pj_none=R2N,
                pj_rows=(r20, r21))


# ═════════════════════════════════════ 应付对账 ═════════════════════════════════════
def build_apd(ws, ctx):
    LAST = 'X'
    widths(ws, {'A': 11, 'B': 14, 'C': 7, 'D': 22, 'E': 13, 'F': 13, 'G': 11, 'H': 7, 'I': 8, 'J': 2,
                'K': 14, 'L': 12, 'M': 13, 'N': 22, 'O': 13, 'P': 13, 'Q': 12, 'R': 2,
                'S': 11, 'T': 13, 'U': 13, 'V': 13, 'W': 9, 'X': 12})
    tip = ('💡 选一家供应商（亮黄格下拉；可以再选一个项目，空＝全部项目），出这家的对账单，可以直接打印给对方核对签字。'
           '起止空着＝从建账日到首页截止日。上面是汇总：期初未付、本期应付、本期已付（分银行/现金、专户、个人账户）、期末未付，'
           '截至止日的应付总额、已付、未付、已开票（按 13%/9%/6%/3%/1% 分）、未开票、进项税；右边按项目小计。'
           '下面三张清单按日期列每一笔：① 应付（送货/结算）② 付款 ③ 发票，各最多 300 条，合计跟汇总对得上会打 ✓。')
    title(ws, '应付对账', LAST, C_VIEW, tip)
    # ── 选择格 ──
    sup0 = ''
    aps = ctx.get('ap') or []
    if aps:
        from collections import Counter
        sup0 = Counter(x.get('供应商') for x in aps if x.get('供应商')).most_common(1)
        sup0 = sup0[0][0] if sup0 else ''
    selector(ws, 'A3', '供应商', 'B3', sup0 or None, dv_formula='=供应商列表', prompt='选一家供应商（空＝供应商信息第一家）')
    selector(ws, 'E3', '项目', 'F3', None, dv_formula='=项目列表', prompt='空＝全部项目')
    selector(ws, 'K3', '起', 'L3', None, fmt=DATE)
    selector(ws, 'O3', '止', 'P3', None, fmt=DATE)
    dv_date(ws, 'L3')
    dv_date(ws, 'P3')
    SUP, PJ, D0, D1, PJC, SE = '$Z$3', '$Z$4', '$M$3', '$Q$3', '$Z$7', '$Z$8'
    ws['Z3'] = '=IF(TRIM(B3&"")="",IFERROR(INDEX(供应商_名称,1)&"",""),TRIM(B3&""))'
    ws['Z4'] = '=TRIM(F3&"")'
    ws['Z7'] = f'=IF({PJ}="",{ALL},"="&{esc(PJ)})'
    ws['Z8'] = f'={esc(SUP)}'
    put(ws, 'C3', f'=IF(TRIM(B3&"")="","（空＝第一家「"&{SUP}&"」）",IF(ISNUMBER(MATCH({SUP},供应商_名称,0)),"✓","⚠ 不在【供应商信息】里"))',
        F_NOTE, align=AL, border=False)
    _merge(ws, 'C3:D3')
    put(ws, 'G3', f'=IF({PJ}="","（空＝全部项目）",IF(ISNUMBER(MATCH({PJ},项目_名称,0)),"✓ 只看这个项目","⚠ 不在【项目档案】里"))',
        F_NOTE, align=AL, border=False)
    _merge(ws, 'G3:I3')
    put(ws, 'M3', '=IF(ISNUMBER(L3),INT(L3),P_建账日)', F_AUTOB, FILL_AUTO, DATE, AC)
    _note(ws, 'N3', '（空＝建账日）')
    put(ws, 'Q3', '=IF(ISNUMBER(P3),INT(P3),P_截止)', F_AUTOB, FILL_AUTO, DATE, AC)
    _note(ws, 'S3', '（空＝首页截止日）')
    _merge(ws, 'S3:T3')
    home_link(ws, 'X3')
    ws.row_dimensions[3].height = 24

    # ── 常用公式片段 ──
    full = f'IFERROR(INDEX({inref(SH_SUP, SUP_COLS["开票单位"])},MATCH({SUP},供应商_名称,0)+1)&"","")'
    OPEN = f'SUMIFS(期初_金额,期初_类型,"应付账款",期初_对象,{SE},期初_项目,{PJC})'
    ap = lambda *c: f'SUMIFS(付_应付额,付_供应商,{SE},付_项目,{PJC}' + ''.join(',' + x for x in c) + ')'
    kp = lambda col, *c: f'SUMIFS({col},付_供应商,{SE},付_项目,{PJC}' + ''.join(',' + x for x in c) + ')'
    paid = lambda *c: f'-SUMIFS(收_净额,收_供应商,{SE},收_冲应付,1,收_项目,{PJC}' + ''.join(',' + x for x in c) + ')'
    le = lambda col, d: f'{col},"<="&{d}'
    ge = lambda col, d: f'{col},">="&{d}'
    PER_F = (ge('付_日期', D0), le('付_日期', D1))
    PER_K = (ge('付_开票日期', D0), le('付_开票日期', D1))
    PER_S = (ge('收_日期', D0), le('收_日期', D1))

    # ── 抬头 ──
    _merge(ws, 'A5:X5')
    put(ws, 'A5', f'=P_公司&"　应付对账单　（"&IF({full}<>"",{full},{SUP})&"）"', F_STMT, align=AC, border=False)
    ws.row_dimensions[5].height = 30
    info = [('A6', '供应商', 'B6', 'B6:C6', f'={SUP}'), ('D6', '开票单位', 'E6', 'E6:I6', f'=IF({full}="","（供应商信息里没填开票单位）",{full})'),
            ('K6', '项目', 'L6', 'L6:M6', f'=IF({PJ}="","全部项目",{PJ})'),
            ('N6', '期间', 'O6', 'O6:Q6', f'=TEXT({D0},"yyyy-mm-dd")&" 至 "&TEXT({D1},"yyyy-mm-dd")'),
            ('S6', '截至止日未付', 'U6', 'U6:V6', '=I12')]
    for lc, lt, vc, rng, f in info:
        _lbl(ws, lc, lt)
        _merge(ws, rng)
        put(ws, vc, f, F_TXTB, align=AL)
        a, b = rng.split(':')
        for k in range(CI(a.rstrip('0123456789')), CI(b.rstrip('0123456789')) + 1):
            ws.cell(row=6, column=k).border = BD
    _merge(ws, 'S6:T6')
    ws['U6'].number_format = MONEY
    ws['U6'].font = F_RED

    # ── 一、汇总 ──
    section(ws, 8, 'A', 'I', '一、汇总', C_VIEW)
    _merge(ws, 'A9:D9')
    put(ws, 'A9', '本期（起～止）', F_HDR, fill(C_VIEW), align=AC)
    _merge(ws, 'F9:I9')
    put(ws, 'F9', '截至止日（累计）', F_HDR, fill(C_VIEW), align=AC)
    left = [('期初未付（起日前一天）', f'={OPEN}+{ap(le("付_日期", D0 + "-1"))}-({paid(le("收_日期", D0 + "-1"))})'),
            ('本期应付', f'={ap(*PER_F)}'),
            ('本期已付（冲应付的）', f'={paid(*PER_S)}'),
            ('　其中：银行、现金', f'={paid(*PER_S, AT("银行"))}+({paid(*PER_S, AT("现金"))})'),
            ('　　　　专户（农民工专户等）', f'={paid(*PER_S, AT("专户"))}'),
            ('　　　　个人账户（微信等）', f'={paid(*PER_S, AT("个人户"))}'),
            ('期末未付（负数＝多付了）', '=D10+D11-D12'),
            ('本期开票', f'={kp("付_已开票", *PER_K)}'),
            ('直接付款（没登记应付，直接进成本）',
             '=-(' + '+'.join(f'SUMIFS(收_净额,收_供应商,{SE},收_冲应付,0,收_归类,"{g}",收_项目,{PJC},{PER_S[0]},{PER_S[1]})'
                              for g, _t in COST3) + ')')]
    right = [('应付总额（含期初）', f'={OPEN}+{ap(le("付_日期", D1))}'),
             ('已付总额', f'={paid(le("收_日期", D1))}'),
             ('未付（负数＝多付了）', '=I10-I11'),
             ('已开票', f'={kp("付_已开票", le("付_开票日期", D1))}'),
             ('未开票（应付登记−已开票）', f'=I10-{OPEN}-I13'),
             ('进项税（专票）', f'={kp("付_进项税", le("付_开票日期", D1))}'),
             ('核对', '=IF(ROUND(D16-I12,2)=0,"✓ 期末未付＝截至止日未付","✗ 期末未付跟未付差 "&TEXT(D16-I12,"#,##0.00"))')]
    for k, (t, f) in enumerate(left):
        r = 10 + k
        _merge(ws, f'A{r}:C{r}')
        put(ws, f'A{r}', t, F_TXTB if not t.startswith('　') else F_TXT, fill('FFF2F2F2'), align=AL)
        for c in 'BC':
            ws[f'{c}{r}'].border = BD
        put(ws, f'D{r}', f, F_AUTOB, None, MONEY, AR)
    for k, (t, f) in enumerate(right):
        r = 10 + k
        _merge(ws, f'F{r}:H{r}')
        put(ws, f'F{r}', t, F_TXTB, fill('FFF2F2F2'), align=AL)
        for c in 'GH':
            ws[f'{c}{r}'].border = BD
        put(ws, f'I{r}', f, F_AUTOB, None, MONEY, AR)
    ws['I16'].font = F_NOTE
    ws['I16'].alignment = AL
    for c in ('D16', 'I12'):
        ws[c].fill = FILL_TOT
    ws['I16'].number_format = 'General'
    _merge(ws, 'A21:I22')
    put(ws, 'A21', '以上往来核对无误。　　供应商（签章）：＿＿＿＿＿＿＿＿　　日期：＿＿＿＿＿＿　　　　经办：＿＿＿＿＿＿',
        F_TXT, align=ALW, border=False)

    # ── 已开票按税率（截至止日） ──
    section(ws, 8, 'S', 'W', '已开票按税率（截至止日）', C_VIEW)
    header(ws, 9, [('S', '税率'), ('T', '已开票'), ('U', '其中专票'), ('V', '普票\n（不是专票的）'), ('W', '进项税')], C_VIEW)
    kd = le('付_开票日期', D1)
    for j, (rv, lab) in enumerate(RATES + [(None, '其他')]):
        r = 10 + j
        ws[f'S{r}'] = lab
        if rv is not None:
            ws[f'T{r}'] = f'={kp("付_已开票", f"付_税率,{rv}", kd)}'
            ws[f'U{r}'] = f'={kp("付_已开票", f"付_税率,{rv}", ZP, kd)}'
            ws[f'W{r}'] = f'={kp("付_进项税", f"付_税率,{rv}", kd)}'
        else:
            ws[f'T{r}'] = '=ROUND(I13-SUM(T10:T14),2)'
            ws[f'U{r}'] = f'=ROUND({kp("付_已开票", ZP, kd)}-SUM(U10:U14),2)'
            ws[f'W{r}'] = '=ROUND(I15-SUM(W10:W14),2)'
        ws[f'V{r}'] = f'=T{r}-U{r}'
    _grid(ws, 10, 15, list('STUVW'), {c: MONEY for c in 'TUVW'})
    ws['S16'] = '合计'
    for c in 'TUVW':
        ws[f'{c}16'] = f'=SUM({c}10:{c}15)'
    _tot(ws, 16, list('STUVW'), {'S': None})
    ws['S16'].alignment = AC

    # ── 四、按项目小计（截至止日） ──
    section(ws, 8, 'K', 'Q', '四、按项目小计（截至止日）', C_VIEW)
    header(ws, 9, [('K', '项目'), ('L', '期初应付'), ('M', '应付登记'), ('N', '已付'), ('O', '未付'), ('P', '已开票'),
                   ('Q', '未开票')], C_VIEW)
    NPJ4 = 10
    q0, q1 = 10, 10 + NPJ4 - 1
    QN, QT = q1 + 1, q1 + 2
    p0 = HR0

    def pj4_cond(i):
        nm = f'INDEX(项目_名称,{i + 1})'
        return (f'AND({nm}<>"",{SUP}<>"",OR({PJ}="",{nm}={PJ}),OR('
                f'COUNTIFS(付_供应商,{SE},付_项目,{esc(nm)},付_有效,1,付_日期,"<="&{D1})>0,'
                f'COUNTIFS(收_供应商,{SE},收_项目,{esc(nm)},收_冲应付,1,收_日期,"<="&{D1})>0,'
                f'SUMIFS(期初_金额,期初_类型,"应付账款",期初_对象,{SE},期初_项目,{esc(nm)})<>0))')
    counter(ws, 'AH', p0, N_PJ, pj4_cond)
    n_pj4 = cnt('AH', p0, N_PJ)

    def pj4_row(r, crit, guard):
        g = lambda body: f'=IF({guard},"",{body})'
        ws[f'L{r}'] = g(f'SUMIFS(期初_金额,期初_类型,"应付账款",期初_对象,{SE},期初_项目,{crit})')
        ws[f'M{r}'] = g(f'SUMIFS(付_应付额,付_供应商,{SE},付_项目,{crit},付_日期,"<="&{D1})')
        ws[f'N{r}'] = g(f'-SUMIFS(收_净额,收_供应商,{SE},收_冲应付,1,收_项目,{crit},收_日期,"<="&{D1})')
        ws[f'O{r}'] = g(f'L{r}+M{r}-N{r}')
        ws[f'P{r}'] = g(f'SUMIFS(付_已开票,付_供应商,{SE},付_项目,{crit},付_开票日期,"<="&{D1})')
        ws[f'Q{r}'] = g(f'M{r}-P{r}')

    for k in range(NPJ4):
        r = q0 + k
        ws[f'AI{r}'] = f'={kth(k + 1, "AH", p0, N_PJ)}'
        ws[f'AJ{r}'] = f'=IF($AI{r}=0,"",{esc(f"$K{r}")})'
        ws[f'K{r}'] = f'=IF($AI{r}=0,"",INDEX(项目_名称,$AI{r}))'
        pj4_row(r, f'$AJ{r}', f'$AI{r}=0')
    ws[f'K{QN}'] = f'=IF({PJ}<>"","",IF({n_pj4}>{NPJ4},"其他项目＋没分项目","没分项目的"))'
    # 没分项目（项目空）＋超出 10 个的项目：按「合计−列出的」倒挤，保证合计＝汇总
    for c in 'LMNP':
        tot_ref = {'L': f'{OPEN}', 'M': None, 'N': 'I11', 'P': 'I13'}[c]
        if c == 'M':
            tot_ref = f'(I10-{OPEN})'
        ws[f'{c}{QN}'] = f'=IF({PJ}<>"","",ROUND({tot_ref}-SUM({c}{q0}:{c}{q1}),2))'
    ws[f'O{QN}'] = f'=IF({PJ}<>"","",L{QN}+M{QN}-N{QN})'
    ws[f'Q{QN}'] = f'=IF({PJ}<>"","",M{QN}-P{QN})'
    _grid(ws, q0, QN, list('KLMNOPQ'), {c: MONEY for c in 'LMNOPQ'}, {'K': AL})
    for c in 'KLMNOPQ':
        ws[f'{c}{QN}'].fill = FILL_SUB
    ws[f'K{QT}'] = '合计'
    for c in 'LMNOPQ':
        ws[f'{c}{QT}'] = f'=SUM({c}{q0}:{c}{QN})'
    _tot(ws, QT, list('KLMNOPQ'), {'K': None})
    ws[f'K{QT}'].alignment = AC
    # 核对：「没分项目」那行直接按项目空算一遍，跟倒挤的比（不等＝项目超过 10 个，或有项目不在项目档案）
    none_m = f'SUMIFS(付_应付额,付_供应商,{SE},付_项目,"",付_日期,"<="&{D1})'
    ok4 = '"✓ 合计＝左边汇总（截至止日）"'
    put(ws, f'K{QT + 1}', f'=IF({PJ}<>"",IF(AND(ROUND(L{QT}+M{QT}-I10,2)=0,ROUND(N{QT}-I11,2)=0,ROUND(P{QT}-I13,2)=0),{ok4},'
                          f'"✗ 合计跟左边汇总对不上"),IF(ROUND(M{QN}-{none_m},2)=0,{ok4},'
                          f'IF({n_pj4}>{NPJ4},"共 "&{n_pj4}&" 个项目，只列前 {NPJ4} 个，其余并在「其他项目」那行",'
                          f'"⚠ 「没分项目」那行里有不是工程项目的数（期初余额的项目？）")))', F_NOTE, align=AL, border=False)
    _merge(ws, f'K{QT + 1}:Q{QT + 1}')

    # ── 三张清单（并排） ──
    RS = 25
    rn, rh, RT_ = RS + 1, RS + 2, RS + 3
    l0 = RT_ + 1
    NL = 300
    l1 = l0 + NL - 1
    _ncol(ws, 'AD', HR0, max(N_AP, N_CASH))
    nA = _skey(ws, 'AA', 'AD', HR0, N_AP,
               lambda r: (f'AND({SUP}<>"",INDEX(付_有效,$AD{r})=1,INDEX(付_供应商,$AD{r})={SUP},'
                          f'OR({PJ}="",INDEX(付_项目,$AD{r})={PJ}),INDEX(付_日期,$AD{r})>={D0},INDEX(付_日期,$AD{r})<={D1})'),
               lambda r: f'INDEX(付_日期,$AD{r})')
    nB = _skey(ws, 'AB', 'AD', HR0, N_CASH,
               lambda r: (f'AND({SUP}<>"",INDEX(收_有效,$AD{r})=1,INDEX(收_供应商,$AD{r})={SUP},'
                          f'OR({PJ}="",INDEX(收_项目,$AD{r})={PJ}),INDEX(收_日期,$AD{r})>={D0},INDEX(收_日期,$AD{r})<={D1})'),
               lambda r: f'INDEX(收_日期,$AD{r})')
    nC = _skey(ws, 'AC', 'AD', HR0, N_AP,
               lambda r: (f'AND({SUP}<>"",INDEX(付_有效,$AD{r})=1,INDEX(付_供应商,$AD{r})={SUP},INDEX(付_已开票,$AD{r})<>0,'
                          f'OR({PJ}="",INDEX(付_项目,$AD{r})={PJ}),INDEX(付_开票日期,$AD{r})>={D0},INDEX(付_开票日期,$AD{r})<={D1})'),
               lambda r: f'INDEX(付_开票日期,$AD{r})')
    section(ws, RS, 'A', 'I', '① 应付明细（【应付登记】，按日期）', C_VIEW)
    section(ws, RS, 'K', 'Q', '② 付款明细（【收支登记】供应商＝这家，按日期）', C_VIEW)
    section(ws, RS, 'S', 'X', '③ 开票明细（按开票日期）', C_VIEW)
    cntmsg = lambda n: f'IF({n}>{NL},"共 "&{n}&" 条，只列前 {NL} 条（请缩短日期）","共 "&{n}&" 条")'
    header(ws, rh, [('A', '日期'), ('B', '项目'), ('C', '类型'), ('D', '摘要'), ('E', '应付额'), ('F', '已开票'), ('G', '开票日期'),
                    ('H', '税率'), ('I', '发票类型')], C_VIEW)
    header(ws, rh, [('K', '日期'), ('L', '账户'), ('M', '收支项目'), ('N', '摘要'), ('O', '付款金额'), ('P', '项目'),
                    ('Q', '冲应付？')], C_VIEW)
    header(ws, rh, [('S', '开票日期'), ('T', '项目'), ('U', '开票金额'), ('V', '税率'), ('W', '发票类型'), ('X', '进项税')], C_VIEW)
    for k in range(NL):
        r = l0 + k
        a, b, c = f'$AE{r}', f'$AF{r}', f'$AG{r}'
        ws[f'AE{r}'] = f'={ksorted(k + 1, "AA", HR0, N_AP, nA)}'
        ws[f'AF{r}'] = f'={ksorted(k + 1, "AB", HR0, N_CASH, nB)}'
        ws[f'AG{r}'] = f'={ksorted(k + 1, "AC", HR0, N_AP, nC)}'
        A = lambda nm, n=a: f'=IF({n}=0,"",INDEX({nm},{n}))'
        ws[f'A{r}'] = A('付_日期')
        ws[f'B{r}'] = f'=IF({a}=0,"",IF(INDEX(付_项目,{a})="","（没分项目）",INDEX(付_项目,{a})))'
        ws[f'C{r}'] = A('付_类型')
        ws[f'D{r}'] = A('付_摘要')
        ws[f'E{r}'] = A('付_应付额')
        ws[f'F{r}'] = A('付_已开票')
        for col, nm in (('G', '付_开票日期'), ('H', '付_税率'), ('I', '付_发票类型')):
            ws[f'{col}{r}'] = f'=IF({a}=0,"",IF(INDEX(付_已开票,{a})=0,"",INDEX({nm},{a})))'
        B = lambda nm: f'=IF({b}=0,"",INDEX({nm},{b}))'
        ws[f'K{r}'] = B('收_日期')
        ws[f'L{r}'] = B('收_账户')
        ws[f'M{r}'] = B('收_收支项目')
        ws[f'N{r}'] = B('收_摘要')
        ws[f'O{r}'] = f'=IF({b}=0,"",-INDEX(收_净额,{b}))'
        ws[f'P{r}'] = B('收_项目原')
        cat = f'INDEX(收_归类,{b})'
        ws[f'Q{r}'] = (f'=IF({b}=0,"",IF(INDEX(收_冲应付,{b})=1,"是",IF(OR({cat}="材料款",{cat}="分包款",{cat}="机械费"),"直接成本",'
                       f'"不冲（"&IF({cat}="","未分类",{cat})&"）")))')
        C = lambda nm: f'=IF({c}=0,"",INDEX({nm},{c}))'
        ws[f'S{r}'] = C('付_开票日期')
        ws[f'T{r}'] = f'=IF({c}=0,"",IF(INDEX(付_项目,{c})="","（没分项目）",INDEX(付_项目,{c})))'
        ws[f'U{r}'] = C('付_已开票')
        ws[f'V{r}'] = C('付_税率')
        ws[f'W{r}'] = C('付_发票类型')
        ws[f'X{r}'] = C('付_进项税')
    _grid(ws, l0, l1, list('ABCDEFGHI'), {'A': DATE, 'E': MONEY, 'F': MONEY, 'G': DATE, 'H': RATE_FMT}, {'D': AL, 'B': AL})
    _grid(ws, l0, l1, list('KLMNOPQ'), {'K': DATE, 'O': MONEY}, {'N': AL, 'P': AL})
    _grid(ws, l0, l1, list('STUVWX'), {'S': DATE, 'U': MONEY, 'V': RATE_FMT, 'X': MONEY}, {'T': AL})
    # 合计行（清单上面）＋核对
    for c0, cols, sums in (('A', 'ABCDEFGHI', 'EF'), ('K', 'KLMNOPQ', 'O'), ('S', 'STUVWX', 'UX')):
        ws[f'{c0}{RT_}'] = '合计'
        for c in sums:
            ws[f'{c}{RT_}'] = f'=SUM({c}{l0}:{c}{l1})'
        _tot(ws, RT_, list(cols), {x: (MONEY if x in sums else None) for x in cols})
        ws[f'{c0}{RT_}'].alignment = AC
    all_pay = f'-SUMIFS(收_净额,收_供应商,{SE},收_项目,{PJC},{PER_S[0]},{PER_S[1]})'
    msgs = [('A', 'I', f'={cntmsg(nA)}&IF(ROUND(E{RT_}-D11,2)=0,"　✓ 应付额合计＝汇总本期应付","　✗ 应付额合计跟汇总本期应付差 "&TEXT(E{RT_}-D11,"#,##0.00"))'),
            ('K', 'Q', f'={cntmsg(nB)}&IF(ROUND(O{RT_}-({all_pay}),2)=0,"　✓","　✗ 没列全")&"　其中冲应付 "&TEXT(D12,"#,##0.00")'
                       f'&"（＝汇总本期已付）、直接成本 "&TEXT(D18,"#,##0.00")&"、其他 "&TEXT(O{RT_}-D12-D18,"#,##0.00")'),
            ('S', 'X', f'={cntmsg(nC)}&IF(ROUND(U{RT_}-D17,2)=0,"　✓ 开票合计＝汇总本期开票","　✗ 开票合计跟汇总本期开票差 "&TEXT(U{RT_}-D17,"#,##0.00"))')]
    for c1, c2, f in msgs:
        _merge(ws, f'{c1}{rn}:{c2}{rn}')
        put(ws, f'{c1}{rn}', f, F_NOTE, align=AL, border=False)
    hide(ws, 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ')
    ws.freeze_panes = 'A4'
    print_setup(ws, None, landscape=True)
    ws.print_area = f'A1:{LAST}{l1}'
    return dict(list_rows=(l0, l1), list_total=RT_, pj_rows=(q0, q1), pj_none=QN, pj_total=QT)


def build(wb, ctx):
    build_ap(wb[SH_AP], ctx)
    build_apd(wb[SH_APD], ctx)
