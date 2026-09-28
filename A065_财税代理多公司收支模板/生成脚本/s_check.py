# -*- coding: utf-8 -*-
"""【数据校验】把「会让报表算错」的情况一项项数出来，再列出流水、发票里还没登记的单位（复制去【往来单位】）；
   【首页】选公司、年份看关键数，表格导航，使用步骤。"""
from openpyxl.formatting.rule import FormulaRule
from common import *
from s_reports import year_crit
from s_ar import AR_R0, AR_N

TODO_N = 60      # 待登记清单各列 60 个


def build_check(wb, ctx):
    ws = wb.create_sheet(SH_CHK)
    widths(ws, {'A': 6, 'B': 44, 'C': 12, 'D': 8, 'E': 50, 'F': 22, 'G': 14})
    title(ws, '数 据 校 验（全自动 · 每项应为 √）', 'F', C_CHK,
          '💡 这里把「会让报表算错」的情况一项项数出来。✗ 要改；⚠ 是提醒，看一眼没问题就行。改完回来看，变成 √ 就好了。'
          '下面还列出流水、发票里出现了、但【往来单位】里还没登记的名字：复制名字，到【往来单位】B 列最下面「选择性粘贴 → 数值」、选类型'
          '（要绑收支项目的顺手选上），资金台帐和应收应付就会自动认出来。')
    header(ws, 4, [('A', '序号'), ('B', '检查项'), ('C', '数量'), ('D', '状态'), ('E', '怎么处理'), ('F', '去哪看')], C_CHK, height=28)
    PN = pr(PT_NAME)
    ITN = IT_NAMES
    i0, i1, x0, x1 = ctx['intra_chk']
    kw_i, kw_w = br(KW_ITEM, KW_R0, KW_R1), br(KW_WORD, KW_R0, KW_R1)
    op_p = br(OP_PARTY, OP_R0, OP_R1)
    checks = [
        ('资金台帐：认不出往来单位 / 收支项目的笔数', f'=COUNTIF({jr(J_CHK)},"✗ 认不出*")+COUNTIF({jr(J_CHK)},"✗ 请在 L 列*")', 'E',
         '到资金台帐筛选「校验」列：常来往的单位到【往来单位】登记（能绑项目的绑上）；社保挂靠、费用这类在 K/L 列手工选', SH_CASH),
        ('资金台帐：其他 ✗（没选账户、日期看不懂、金额不对…）',
         f'=COUNTIF({jr(J_CHK)},"✗*")-C5', 'E', '筛选「校验」列，按提示改', SH_CASH),
        ('资金台帐：算出的余额跟银行给的余额对不上', f'=COUNTIF({jr(J_CHK)},"⚠ 余额*")', 'W',
         '一般是漏粘了几行、重复粘了几行，或期初余额不对；从第一个对不上的行往上查', SH_CASH),
        ('资金台帐：疑似重复粘贴（跟上面某行一模一样）', f'=COUNTIF({jr(J_CHK)},"⚠ 跟上面*")', 'W',
         '真是重复的就把这一行删掉；同一天同金额两笔真实交易的，在 M 列备注写一下就不算重复了', SH_CASH),
        ('资金台帐：税务扣款（默认算税金）要确认是不是社保', f'=COUNTIF({jr(J_CHK)},"⚠ 税务扣款*")', 'W',
         '国库扣款里可能有社保、社保挂靠：是的话在 L 列改成人工-社保 / 社保挂靠，改了提醒就消失', SH_CASH),
        ('往来单位：没选类型', f'=SUMPRODUCT(({PN}<>"")*({pr(PT_TYPE)}=""))', 'E',
         '【往来单位】C 列选客户 / 供应商 / 个人……，不选的话客户来款认不出新增、续费', SH_PARTY),
        ('资金台帐：手工选的往来单位没在【往来单位】登记', f'=COUNTIF({jr(J_CHK)},"⚠ 往来单位没*")', 'W',
         '统计照算；想在应收应付、客户统计里单独一行，就去【往来单位】登记', SH_CASH),
        ('资金台帐：收入只选了项目、没选往来单位', f'=COUNTIF({jr(J_CHK)},"⚠ 收入没*")', 'W',
         '收支汇总照算，但客户收入统计、应收里算不到这笔；K 列把客户选上（名单见下面）', SH_CASH),
        ('发票导入：认不出是哪家公司的发票', f'=COUNTIF({vr(V_CHK)},"✗ 销方*")', 'E',
         '到【基础资料】① 公司，把 5 家公司的税号都填上（发票靠税号认公司）', SH_INV),
        ('发票导入：开票日期看不懂', f'=COUNTIF({vr(V_CHK)},"✗ 开票日期*")', 'E',
         '开票日期要像 2026-09-01 或 2026/9/1；看不懂的票不进应收应付和发票汇总', SH_INV),
        ('发票导入：重复粘贴的发票（只算一次）', f'=COUNTIF({vr(V_CHK)},"⚠ 重复*")', 'W', '不影响金额；想干净可以把重复的行删掉', SH_INV),
        ('发票导入：开了销项票、客户还没登记', f'=COUNTIF({vr(V_CHK)},"⚠ 客户没*")', 'W',
         '应收照算（在应收应付汇总最后一行「没登记的单位合计」）；登记后单独一行。名单见下面右边', SH_INV),
        ('基础资料：资金账户没选所属公司', f'=SUMPRODUCT(({AC_NAMES}<>"")*({AC_COS}=""))', 'E', '【基础资料】② 选所属公司', SH_BASE),
        ('基础资料：资金账户的所属公司不在公司清单里',
         f'=SUMPRODUCT(({AC_COS}<>"")*(COUNTIF({CO_NAMES},{AC_COS})=0))', 'E', '公司简称要跟 ① 里写的一模一样', SH_BASE),
        ('基础资料：资金账户名称重复', f'=SUMPRODUCT(({AC_NAMES}<>"")*(COUNTIF({AC_NAMES},{AC_NAMES})>1))', 'E',
         '每个账户起一个不一样的简称（比如 农行2896、微信-中智）', SH_BASE),
        ('基础资料：公司没填税号（这家的发票认不出来）', f'=SUMPRODUCT(({CO_NAMES}<>"")*({CO_TAXES}=""))', 'W',
         '【基础资料】① 补税号', SH_BASE),
        ('基础资料：收支项目没选类别', f'=SUMPRODUCT(({ITN}<>"")*({IT_CLSS}=""))', 'E', '类别决定落到报表哪一块，一定要选', SH_BASE),
        ('基础资料：收支项目名称重复', f'=SUMPRODUCT(({ITN}<>"")*(COUNTIF({ITN},{ITN})>1))', 'E',
         '两个项目不能同名（两个「其他」已经叫其他收入、其他成本）', SH_BASE),
        ('基础资料：摘要关键词对应的项目不在收支项目里',
         f'=SUMPRODUCT(({kw_w}<>"")*(COUNTIF({ITN},{kw_i})=0))', 'E', '④ 的收支项目从下拉选', SH_BASE),
        ('往来单位：名称重复', f'=SUMPRODUCT(({PN}<>"")*(COUNTIF({PN},{esc(PN)})>1))', 'E',
         '同一个单位只登记一行；别的写法填到别名 1、别名 2', SH_PARTY),
        ('往来单位：绑定的收支项目不在清单里',
         f'=SUMPRODUCT(({pr(PT_BIND)}<>"")*(COUNTIF({ITN},{pr(PT_BIND)})=0))', 'E', '从下拉选', SH_PARTY),
        ('期初往来：填了金额，没选所属公司或往来单位',
         f'=SUMPRODUCT(((({br(OP_AR, OP_R0, OP_R1)}<>0)+({br(OP_AP, OP_R0, OP_R1)}<>0)+({br(OP_OTH, OP_R0, OP_R1)}<>0))>0)'
         f'*((({br(OP_CO, OP_R0, OP_R1)}="")+(op_p=""))>0))'.replace('op_p', op_p), 'E', '这几行哪张报表都算不进去，补上公司和单位', SH_BASE),
        ('期初往来：单位没在【往来单位】登记',
         f'=SUMPRODUCT(({op_p}<>"")*(COUNTIF({PN},{op_p})=0)*(COUNTIF({CO_NAMES},{op_p})=0))', 'W',
         '期初照算进合计；登记后才会在应收应付里单独一行', SH_BASE),
        ('内部往来：两家公司各记的对不上（对数）', f'=COUNTIF({SH_INTRA}!$B${i0}:$I${i1},"差*")/2', 'W',
         '多半是另一家公司那个账户的流水还没导进来；都导了还对不上，看是不是有一笔没认成「内部划转」', SH_INTRA),
        ('账户互转：同一家公司转出、转入没抵平', f'=COUNTIF({SH_INTRA}!$C${x0}:$C${x1},"✗*")', 'W',
         '转入那个账户的流水还没导进来，或者只记了一边', SH_INTRA),
        ('资金台帐：流水日期比账户的期初日期还早',
         f'=SUMPRODUCT(({AC_NAMES}<>"")*ISNUMBER({AC_ODATES})*COUNTIFS({jr(J_ACC)},{AC_NAMES},{jr(J_DATE)},"<"&{AC_ODATES}))', 'E',
         '【基础资料】② 把期初日期改成这个账户第一笔流水那天（期初余额也要是那一笔之前的余额），不然报表的期初期末会少算', SH_BASE),
        ('往来单位：登记了自家公司（会被当成客户/供应商）',
         f'=SUMPRODUCT(({PN}<>"")*((COUNTIF({CO_NAMES},{esc(PN)})+COUNTIF({CO_FULLS},{esc(PN)}))>0))', 'E',
         '自家 5 家公司只在【基础资料】① 登记，从【往来单位】删掉，内部划转才认得出', SH_PARTY),
        ('资金台帐：已用行数（共 6000 行）', f'=COUNTA({jr(J_TIME)})', 'C',
         '跨年一直往下记。超过 5500 行时提醒：在倒数几行中间右键「插入」若干行，再把上面一行整行复制粘贴下来，各报表会自动算进去', SH_CASH),
        ('发票导入：已用行数（3000 行）', f'=COUNTA({vr(V_ENO)})+COUNTA({vr(V_NO)})-SUMPRODUCT(({vr(V_ENO)}<>"")*({vr(V_NO)}<>""))',
         'C', '超过 2700 行时提醒，加行方法同上', SH_INV),
    ]
    R0 = 5
    err_rows, warn_rows = [], []
    for k, (lab, f, lvl, how, where) in enumerate(checks):
        r = R0 + k
        put(ws, f'A{r}', k + 1, F_AUTO, align=AC)
        put(ws, f'B{r}', lab, F_TXT, align=ALW)
        put(ws, f'C{r}', f, F_AUTOB, fmt=INT if lvl != 'C' else '0', align=AC)
        if lvl == 'C':
            cap = 5500 if where == SH_CASH else 2700
            put(ws, f'D{r}', f'=IF(C{r}>{cap},"⚠","√")', F_AUTOB, align=AC)
            warn_rows.append(r)
        else:
            put(ws, f'D{r}', f'=IF(C{r}=0,"√","{"✗" if lvl == "E" else "⚠"}")', F_AUTOB, align=AC)
            (err_rows if lvl == 'E' else warn_rows).append(r)
        put(ws, f'E{r}', how, F_NOTE, align=ALW)
        c = put(ws, f'F{r}', where, Font(name=YH, sz=10, color='FF0563C1', underline='single'), align=AC)
        link(c, where)
        ws.row_dimensions[r].height = 30
    last = R0 + len(checks) - 1
    rng = f'D{R0}:D{last}'
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'D{R0}="✗"'], fill=FILL_WARN, font=Font(name=YH, sz=11, bold=True, color='FFC00000')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'D{R0}="⚠"'], fill=fill('FFFFEB9C')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'D{R0}="√"'], fill=FILL_OK, font=Font(name=YH, sz=11, bold=True, color='FF00B050')))
    s = last + 1
    put(ws, f'A{s}', '', F_TXTB, FILL_TOT)
    put(ws, f'B{s}', '要改的（✗）项数', F_TXTB, FILL_TOT, align=AC)
    put(ws, f'C{s}', f'=COUNTIF(D{R0}:D{last},"✗")', F_RED, FILL_TOT, fmt='0"项"', align=AC)
    put(ws, f'D{s}', f'=IF(C{s}=0,"√","✗")', F_RED, FILL_TOT, align=AC)
    put(ws, f'E{s}', f'="提醒（⚠）"&COUNTIF(D{R0}:D{last},"⚠")&" 项"', F_NOTE, FILL_TOT, align=AL)
    put(ws, f'F{s}', None, fill_=FILL_TOT)
    ctx['chk_err_cell'] = f'{SH_CHK}!$C${s}'

    # 待登记清单
    t = s + 2
    section(ws, t, 'A', 'C', f'资金台帐里认不出 / 没登记的对方户名（前 {TODO_N} 个）', C_CHK)
    section(ws, t, 'E', 'G', f'销项发票里没登记的客户（前 {TODO_N} 个）', C_CHK)
    header(ws, t + 1, [('A', '序号'), ('B', '对方户名（复制 → 【往来单位】B 列「选择性粘贴 → 数值」）'), ('C', '笔数'),
                       ('E', '客户名称（→ 往来单位 B 列，粘贴为数值）'), ('F', '纳税人识别号（→ H 列）'), ('G', '应收金额')], C_CHK, height=30)
    jt, jc = jr(J_TODO), jr(J_TODOC)
    vt, vc = vr(V_TODO), vr(V_TODOC)
    for k in range(1, TODO_N + 1):
        r = t + 1 + k
        put(ws, f'A{r}', f'=IF(B{r}="","",{k})', F_AUTO, align=AC)
        put(ws, f'B{r}', f'=IFERROR(INDEX({jt},MATCH({k},{jc},0)),"")', F_AUTOB, align=AL)
        put(ws, f'C{r}', f'=IF(B{r}="","",COUNTIF({jt},{esc(f"B{r}")}))', F_AUTO, fmt=INT, align=AC)
        put(ws, f'F{r}', f'=IFERROR(INDEX({vr(V_BTAX)},MATCH({k},{vc},0))&"","")', F_AUTO, fmt='@', align=AC)
        put(ws, f'E{r}', f'=IFERROR(INDEX({vt},MATCH({k},{vc},0)),"")',
            F_AUTOB, align=AL)
        put(ws, f'G{r}', f'=IFERROR(ROUND(SUMIFS({vr(V_ARV)},{vt},{esc(f"INDEX({vt},MATCH({k},{vc},0))")}),2),"")', F_AUTO, fmt=MONEY, align=AR)
    r = t + 2 + TODO_N
    put(ws, f'A{r}', f'="流水里共 "&MAX({jc})&" 个、发票里共 "&MAX({vc})&" 个没登记的名字"', F_NOTE, align=AL, border=False)
    ws.freeze_panes = 'A5'
    return ws


def build_home(wb, ctx):
    ws = wb.create_sheet(SH_HOME)
    widths(ws, {'A': 16, 'B': 16, 'C': 16, 'D': 16, 'E': 16, 'F': 16, 'G': 16, 'H': 16})
    ws.sheet_properties.tabColor = C_HOME[2:]
    ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:H1')
    put(ws, 'A1', '财税代理 · 多公司收支管理（资金台帐 ＋ 发票 → 收支 · 应收应付 · 现金流）',
        Font(name=YH, sz=18, bold=True, color='FFFFFFFF'), fill(C_HOME), align=AC, border=False)
    ws.row_dimensions[1].height = 42
    ws.merge_cells('A2:H2')
    put(ws, 'A2', '💡 5 家公司、所有银行卡/微信/支付宝都记在一本【资金台帐】里，【发票导入】装全部购销发票；'
                  '其余报表全自动，每张表顶上的黄格子选公司（或全部）、年份。跨年一直往下记，报表换年份就是新一年。'
                  '淡黄格子手填，淡蓝是粘贴区，灰格子是公式别动。点下面的表名直接跳过去。',
        F_TIP, FILL_TIP, align=ALW, border=False)
    ws.row_dimensions[2].height = 40
    selector(ws, 'A4', '公司', 'B4', '全部', f'={AUX_CO_ALL}')
    selector(ws, 'D4', '年份', 'E4', 2026, YEARS)
    ws.row_dimensions[4].height = 24
    cc, y = co_crit('$B$4'), '$E$4'
    yc = year_crit(y)

    def net(cls, sign=+1):
        a, b = (J_INV, J_OUTV) if sign > 0 else (J_OUTV, J_INV)
        return (f'SUMIFS({jr(a)},{jr(J_CLS)},"{cls}",{jr(J_CO)},{cc},{yc})'
                f'-SUMIFS({jr(b)},{jr(J_CLS)},"{cls}",{jr(J_CO)},{cc},{yc})')
    kpis = [
        ('本年收入', f'=ROUND({net(CLS_IN)},2)', MONEY),
        ('变动利润', f'=ROUND({net(CLS_IN)}-({net(CLS_VAR, -1)}),2)', MONEY),
        ('利润', f'=ROUND({net(CLS_IN)}-({net(CLS_VAR, -1)})-({net(CLS_FIX, -1)}),2)', MONEY),
        ('本年新增客户', f'=COUNTIFS({jr(J_ITEM)},"新增",{jr(J_INV)},">0",{jr(J_CO)},{cc},{yc})', '0"个"'),
        ('资金余额（现在）', f'=ROUND(SUMIFS({AC_OPENS},{AC_COS},{cc})+SUMIFS({jr(J_NET)},{jr(J_CO)},{cc}),2)', MONEY),
        ('客户欠款＊', f'=ROUND(SUMIF({SH_AR}!$G${AR_R0}:$G${AR_R0 + AR_N},">0"),2)', MONEY),
        ('欠供应商＊', f'=ROUND(SUMIF({SH_AR}!$K${AR_R0}:$K${AR_R0 + AR_N},">0"),2)', MONEY),
        ('待处理 / 校验✗', f'=COUNTIF({jr(J_CHK)},"✗*")&" 笔 / "&{ctx["chk_err_cell"]}&" 项"', None),
    ]
    for i, (lab, f, fmt) in enumerate(kpis):
        col = CL(i + 1)
        put(ws, f'{col}6', lab, Font(name=YH, sz=10, bold=True, color='FF595959'), fill('FFF2F2F2'), align=AC)
        put(ws, f'{col}7', f, Font(name=YH, sz=14, bold=True, color='FFC00000' if i < 7 else 'FF1F4E79'),
            fill('FFF2F2F2'), fmt, AC)
    ws.row_dimensions[7].height = 32
    ws.merge_cells('A8:H8')
    put(ws, 'A8', '收入、利润、新增客户按上面选的公司和年份；资金余额是所选公司所有账户到现在；'
                  '＊客户欠款、欠供应商按【应收应付汇总】里选的公司、年份（年底余额，只加正数）。', F_NOTE, align=AL, border=False)

    # 使用步骤
    put(ws, 'A10', '怎么用（第一次按 ①～④ 设好，以后每月做 ⑤⑥）', F_SEC, fill(C_HOME), align=AL)
    ws.merge_cells('A10:H10')
    steps = [
        ('① 基础资料', '填 5 家公司（简称、全称、税号）、15 个资金账户（选所属公司；期初余额＝第一笔流水之前的余额，期初日期＝第一笔流水那天）。收支项目已按你的截图建好，可以往下加。'),
        ('② 往来单位', '常来往的客户、供应商登记一下：流水上的户名写法不一样的填到别名；房东、刻章店、电信这类绑上收支项目，以后自动归类。老客户「建账前已合作」填 是。'),
        ('③ 摘要关键词', '【基础资料】④：摘要里出现「费用外收」就是银行手续费、「扣税」就是税金……已经填了常见的，可以再加。'),
        ('④ 期初往来', '建账前就有的应收、应付、借款余额，在【基础资料】⑤ 填一次。'),
        ('⑤ 每月导流水', '网银导出的明细整块粘到【资金台帐】C 列（农行格式直接粘；微信/支付宝/其他银行先过【流水格式转换】），B 列选账户。'
                        '现金、代付的手工记一行。校验列有 ✗ 的，在 K/L 列手工选往来单位、收支项目（社保挂靠、费用类在这里筛）。'),
        ('⑥ 每月导发票', '电子税务局 → 发票查询统计 → 导出（销项、进项都导），打开后「发票基础信息」页的明细整块粘到【发票导入】A 列。5 家公司的都粘一起，按税号自动分。'),
        ('⑦ 看报表', '收支汇总表、简易现金流量表、费用统计、客户收入统计、应收应付汇总、往来对账单、内部往来、资金余额表，顶上选公司/年份/日期。最后看一眼【数据校验】。'),
    ]
    for i, (a, b) in enumerate(steps):
        r = 11 + i
        put(ws, f'A{r}', a, F_TXTB, FILL_SUB, align=AC)
        ws.merge_cells(f'B{r}:H{r}')
        put(ws, f'B{r}', b, F_TXT, align=ALW)
        for c in 'CDEFGH':
            ws[f'{c}{r}'].border = BD
        ws.row_dimensions[r].height = 34
    # 导航
    n0 = 11 + len(steps) + 1
    put(ws, f'A{n0}', '表格导航（点表名直接跳过去）', F_SEC, fill(C_HOME), align=AL)
    ws.merge_cells(f'A{n0}:H{n0}')
    for c, t in (('A', '表名'), ('B', '干什么的'), ('F', '你要做什么')):
        put(ws, f'{c}{n0 + 1}', t, F_HDR, fill('FF595959'), align=AC)
    ws.merge_cells(f'B{n0 + 1}:E{n0 + 1}')
    ws.merge_cells(f'F{n0 + 1}:H{n0 + 1}')
    nav = [
        (SH_BASE, '公司、资金账户、收支项目（收支类型）、摘要关键词、期初往来', '★ 先设好，以后加账户/项目在这'),
        (SH_PARTY, '客户、供应商、个人、税务银行…… 绑定收支项目、流水户名别名', '★ 新客户/新供应商在这登记'),
        (SH_CASH, '所有账户的收支流水（粘贴 ＋ 手工），即时余额，自动认往来单位和收支项目', '★ 每月粘流水；✗ 的手工选'),
        (SH_CONV, '微信、支付宝、其他银行的导出格式 → 资金台帐的列顺序', '粘进来，复制右边，到资金台帐粘为数值'),
        (SH_INV, '电子税务局导出的购销发票（5 家公司的都粘一起）', '★ 每月粘发票'),
        (SH_SUM, '你那张「收支类型」× 1～12 月：收入、变动成本、变动利润、固定成本、利润', '选公司、年份看'),
        (SH_CF, '业务口径的简易现金流量表（经营 / 投资 / 筹资往来），跟账户余额核对', '选公司、年份看'),
        (SH_BAL, '每个账户期初、收、支、期末，跟银行余额核对', '选公司、起止日期看'),
        (SH_EXP, '收入、各项费用按公司对比', '选年份、月份范围看'),
        (SH_CUS, '每月新增客户、新增/续费收入；每个客户首次来款、续费、应收', '选公司、年份看'),
        (SH_AR, '每个往来单位的应收、应付、其他往来（借款押金）', '选公司、年份看'),
        (SH_STMT, '给某个客户/供应商的对账单：发票、收付款逐笔带余额', '选公司、单位、起止日期'),
        (SH_INTRA, '5 家公司之间划转的钱谁欠谁、两边对不对得上；自家之间开的票', '选截止日看'),
        (SH_INVS, '每月销项、进项发票金额，对照同月收付款', '选公司、年份看'),
        (SH_CHK, '所有数据问题一表列出；没登记的户名、客户名单', '有 ✗ 就按提示改'),
    ]
    for i, (sh, what, todo) in enumerate(nav):
        r = n0 + 2 + i
        c = put(ws, f'A{r}', sh, Font(name=YH, sz=11, bold=True, color='FF0563C1', underline='single'), align=AC)
        link(c, sh)
        ws.merge_cells(f'B{r}:E{r}')
        put(ws, f'B{r}', what, F_TXT, align=ALW)
        ws.merge_cells(f'F{r}:H{r}')
        star = todo.startswith('★')
        put(ws, f'F{r}', todo, F_IN if star else F_NOTE, FILL_IN if star else None, align=ALW)
        for cc_ in 'CDEGH':
            ws[f'{cc_}{r}'].border = BD
            if star and cc_ in 'GH':
                ws[f'{cc_}{r}'].fill = FILL_IN
        ws.row_dimensions[r].height = 30
    return ws
