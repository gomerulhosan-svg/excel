# -*- coding: utf-8 -*-
"""提速（opt）里要在 p1008 之后做的几处：凭证汇总、记账凭证按「凭证所在的那一块」查分录。

记账分录是按来源分块排的（流水 / 采购 / 领用 / 销售 / 工资 / 手工 / 折旧 / 调汇 / 成本 / 销售成本 / 兑换 / 结转损益 / 年末结转，
各块的起止行见 s_je.layout 和 p1008），一张凭证的分录只会在它自己那一块里、从「凭证ID 对应的那个槽位」往下：
  · 凭证ID＝来源代码×100000＋位置（结转损益 12、年末结转 13），位置＝这张凭证第一行在源表里是第几行（月末几种＝月份序号）；
  · 合并成一张的几笔（按天合并、合并号、单据号、手工凭证组号），第一笔就是凭证ID 指的那一笔，其余都在它下面。
所以原来「SUMIF(分录凭证ID,…)」「MATCH(查找键,分录键,0)」扫整张记账分录 1.8 万行的，改成只扫这一段（INDEX(…,起):INDEX(…,止)），
结果一样：凭证汇总 E/F 借贷合计、记账凭证每条分录在第几行、合计行。
凭证汇总 N 列（索引行）：超过本月凭证张数的行不再 MATCH（原来 450 行每行都在 6,898 行里找一遍，找不到才停）。
"""
from common import *
import p1008
from s_je import VL, VL_R1, VP


def windows(ctx):
    """[(相对起行, 每个位置几行, 块末相对行)] 按来源代码 1..13（相对行＝记账分录第 4 行算第 1 行）"""
    bl = ctx['blocks']
    out = {}
    for src, b in bl.items():
        o = b['je0'] - JE_R0 + 1
        out[SRC[src]] = (o, b['legs'], o + b['n'] * b['legs'] - 1)
    o = p1008.CLOSE0 - JE_R0 + 1
    out[12] = (o, p1008.MBLK, o + p1008.NM * p1008.MBLK - 1)
    o = p1008.YE0 - JE_R0 + 1
    out[13] = (o, 2, o + p1008.NY * 2 - 1)
    assert sorted(out) == list(range(1, 14)), sorted(out)
    assert out[13][2] == ctx['je_end'] - JE_R0 + 1
    return [out[i] for i in range(1, 14)]


def win_formulas(vid, w):
    s = f'INT({vid}/100000)'
    start = (f'CHOOSE({s},{",".join(str(x[0]) for x in w)})+(MOD({vid},100000)-1)*'
             f'CHOOSE({s},{",".join(str(x[1]) for x in w)})')
    end = f'CHOOSE({s},{",".join(str(x[2]) for x in w)})'
    return start, end


def seg(name, a, b):
    return f'INDEX({name},{a}):INDEX({name},{b})'


def apply(wb, ctx):
    log = []
    w = windows(ctx)
    # ── 凭证汇总 ──
    ws = wb[SH_VLIST]
    c = VL
    put(ws, 'N4', '=COUNTIFS(凭证月份,报表年度*100+报表月份,凭证有,1)', F_NOTE, border=False)
    put(ws, 'M4', '本月张数→', F_NOTE, border=False)
    mk = '(报表年度*100+报表月份)'
    for k in range(1, VL_N + 1):
        r = VL_R0 + k - 1
        X, V = f'${c["idx"]}{r}', f'${c["vid"]}{r}'
        old = ws[f'{c["idx"]}{r}'].value
        assert old == f'=IFERROR(MATCH({mk}*10000+{k},凭证查找键,0),"")', old
        ws[f'{c["idx"]}{r}'] = f'=IF({k}>$N$4,"",IFERROR(MATCH({mk}*10000+{k},凭证查找键,0),""))'
        st, en = win_formulas(V, w)
        ws[f'P{r}'] = f'=IF({V}="","",{st})'
        ws[f'Q{r}'] = f'=IF({V}="","",{en})'
        for col, side in ((c['dr'], '分录借方'), (c['cr'], '分录贷方')):
            assert ws[f'{col}{r}'].value == f'=IF({X}="","",SUMIF(分录凭证ID,{V},{side}))'
            ws[f'{col}{r}'] = (f'=IF({X}="","",SUMIF({seg("分录凭证ID", f"$P{r}", f"$Q{r}")},{V},'
                               f'{seg(side, f"$P{r}", f"$Q{r}")}))')
        for col in ('P', 'Q'):
            ws[f'{col}{r}'].font = F_NOTE
    for col, t in (('P', '分录起'), ('Q', '分录止')):
        put(ws, f'{col}{VL_HDR}', t, F_NOTE, FILL_AUTO, align=AC)
        ws.column_dimensions[col].hidden = True
    ws.column_dimensions['M'].hidden = True       # M、N 本来就是隐藏的辅助列（凭证ID、索引行）
    ws.column_dimensions['N'].hidden = True
    log.append(f'凭证汇总：索引行超过本月张数不再查；借贷合计只扫这张凭证所在的那一块（P、Q 隐藏列＝分录起止行）')

    # ── 记账凭证 ──
    ws = wb[SH_VPRT]
    v = VP
    n = 0
    for p in range(1, VP_PAGES + 1):
        R = (p - 1) * VP_ROWS + 1
        P0, VID = f'${v["page"]}${R}', f'${v["vid"]}${R}'
        st, en = win_formulas(VID, w)
        ws[f'{v["key"]}{R}'] = f'=IF({P0}="","",{st})'
        ws[f'{v["jrow"]}{R}'] = f'=IF({P0}="","",{en})'
        S, E = f'${v["key"]}${R}', f'${v["jrow"]}${R}'
        for c_ in (v['key'], v['jrow']):
            ws[f'{c_}{R}'].font = F_NOTE
        for i in range(1, VP_LINES + 1):
            r = R + 2 + i
            key = f'${v["key"]}{r}'
            old = ws[f'{v["jrow"]}{r}'].value
            assert old == f'=IF({P0}="","",IFERROR(MATCH({key},分录键,0),""))', old
            ws[f'{v["jrow"]}{r}'] = f'=IF({P0}="","",IFERROR(MATCH({key},{seg("分录键", S, E)},0)+{S}-1,""))'
            n += 1
        r = R + 3 + VP_LINES
        for col, side in ((v['dr'], '分录借方'), (v['cr'], '分录贷方')):
            old = ws[f'{col}{r}'].value
            assert old.startswith('=IF(IF(') and f'SUMIF(分录凭证ID,{VID},{side})' in old, old
            ws[f'{col}{r}'].value = old.replace(f'SUMIF(分录凭证ID,{VID},{side})',
                                                f'SUMIF({seg("分录凭证ID", S, E)},{VID},{seg(side, S, E)})')
    log.append(f'记账凭证：每条分录（{n} 格）、合计行只在这张凭证所在的那一块里找（P、Q 两列第一行＝分录起止行）')
    log += step_rows_check(wb)
    return log


ROWS_CHECK = [(SH_CASH, CASH_R0, CASH_R1), (SH_BUY, MOD_R0, BUY_R1), (SH_ISS, MOD_R0, ISS_R1), (SH_SALE, MOD_R0, SALE_R1),
              (SH_PAY, MOD_R0, PAY_R1), (SH_MAN, MOD_R0, MAN_R1)]


def step_rows_check(wb):
    """数据校验加第 41 项：录入表有没有被插入 / 删除整行。
    记账分录、凭证索引现在直接引用流水 / 业务表的第几行（原来 INDEX(整列,第几笔)），删了整行，后面的凭证号会对不上——
    规矩本来就是「不要插行删行」，这里加一道检查：表的行数（ROWS）一变就报 ✗。"""
    from copy import copy
    ck = wb[SH_CHK]
    assert ck['A44'].value == 40 and ck['A45'].value is None, (ck['A44'].value, ck['A45'].value)
    bad = '&'.join(f'IF(ROWS({q(sh)}!$A${r0}:$A${r1})<>{r1 - r0 + 1},"【{sh}】","")' for sh, r0, r1 in ROWS_CHECK)
    vals = (41, '录入表：没有插入 / 删除整行',
            f'=IF({bad}="","√ 没有","✗ "&{bad}&"插入或删除了整行")',
            '记账分录、凭证直接按位置取各录入表的第几行：插入或删除整行会让后面的凭证号、凭证汇总对不上。按 Ctrl+Z 撤销；'
            '要删一笔，选中那一行的内容按 Delete 清空，不要删整行；新的往下面空行填')
    ck.row_dimensions[45].height = ck.row_dimensions[44].height
    for col, v in zip('ABCD', vals):
        src, dst = ck[f'{col}44'], ck[f'{col}45']
        dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
        dst.alignment, dst.number_format = copy(src.alignment), src.number_format
        dst.value = v
    old = ck['C3'].value
    assert old.count('C5:C44') == 3, old
    ck['C3'].value = old.replace('C5:C44', 'C5:C45')
    assert ck.auto_filter.ref == 'A4:D44', ck.auto_filter.ref
    ck.auto_filter.ref = 'A4:D45'
    from openpyxl.formatting.formatting import ConditionalFormattingList
    cf_new = ConditionalFormattingList()
    for rg in ck.conditional_formatting:
        ref = str(rg.sqref).replace('C5:C44', 'C5:C45')
        for rule in rg.rules:
            cf_new.add(ref, rule)
    ck.conditional_formatting = cf_new
    hm = wb[SH_HOME]
    n = 0
    for row in hm.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('=') and '数据校验!$C$5:$C$44' in c.value:
                c.value = c.value.replace('数据校验!$C$5:$C$44', '数据校验!$C$5:$C$45')
                n += 1
    assert n >= 1, '首页没找到数据校验汇总格'
    return [f'数据校验加第 41 项「录入表没有插入 / 删除整行」（首页汇总 {n} 格范围跟着扩到第 45 行）']
