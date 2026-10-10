# -*- coding: utf-8 -*-
"""查看表（绿）【利润表】＋ 往来表（橙）【发票跟进】。口径见 agent_common「口径速查」（利润表、发票、往来）。

利润表：黄格 B3 年份（空＝P_年度；填 2026、202610、或那年里的一个日期都认）。A 列项目，B～M 列 1～12 月，N 列全年（＝各月合计）。
   截止年月以后的月份空着（不是 0）；建账月以前的月份，只有资金台帐「所属月份」填到那个月的才显示（表头浅色），不然也空着。
   收入、采购按 销_年月 / 采_年月（≤截止），资金类按 资_所属年月（≤截止），折旧 P_月折旧（建账月～截止年月），
   库存变动（只算建账月以后）＝本月估了且前面有估值（前面估过的月，或者 P_期初库存）时：上一个估值－本月估值；
   上一个估值是建账以前的月份、又填了 P_期初库存 时，用 P_期初库存（建账日的估值）。
   备查最后一行：已付、所属月份在截止以后的费用，只在截止以后的月份列里显示（上面的数里都没有）。
   隐藏：AA/AB 标量（第 3 行起）；AD～AO 每月辅助（第 2+m 行）：月、年月、显示、估了、本月估值、上次估的年月、上次估值、库存变动、折旧、
   建账前（显示的建账前月份＝1）、以后（已付、算在这个月的费用：这个月在截止以后才显示）。

发票跟进：黄格 B3 年份（① 各月用；空＝P_年度）、D3 月份（④ 清单用：1～12 按 B3 的年份；空或认不出＝截止日期那个月（截止年月）；
   填 2026-9 / 202609 连年份一起认）。
   第 4、5 行 KPI（截至截止日，固定位置，首页可引用）：B5 客户没开票、D5 多开给客户的票、E5 供应商欠票、G5 供应商多给的票。
   ① 各月（固定位置）：开出/收到 张数、价税合计、税额、收到专票税额、开出税额－收到专票税额；合计行。
   ② 客户开票情况、③ 供应商欠票、④ 选的月的发票清单：从第 FLOW_R0 行起一块接一块往下排（前一块有几行就占几行，不留大段空行）。
   隐藏列：AD/AE 客户（未开票、排序键＝比它大的个数×1000＋家号），AF/AG 供应商（欠票、键），第 3 行起每家一行，SMALL 取第 k 个＝从大到小；
   AH 段（2/3/4）、AI 段内第几行、AJ 行类（1 标题 2 表头 3 合计 4 第二个合计 5 明细 6 共几条 0 空）、AK 第几条、AL 来源行号；
   ④ 按 票_排序键（日期×SORT_M＋n）：这个月 [月初, MIN(月末, 截止)] 的键是连着的一段，SMALL(票_排序键, 月前个数＋k)。"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from layout import *
from common import *

M = SORT_M
_thin = Side(style='thin', color='FFBFBFBF')
CF_BD = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
LBL = fill('FFD9E1F2')
F_MEMO = Font(name=YH, sz=9, color='FF808080')
F_MEMOL = Font(name=YH, sz=9, bold=True, color='FF808080')
F_KPI_B = Font(name=YH, sz=11, bold=True, color='FF1F3864')
C_MEMO = 'FFA6A6A6'


def cf_fill(rgb):
    return PatternFill('solid', fgColor=rgb, bgColor=rgb)


def cf_blocks(ws, r0, r1, rules):
    """条件格式：rules＝[(起列, 止列, 公式(用本块第一列、第 r0 行写), 样式)]，按优先顺序。
       列切成「规则集合一样」的连续块，每格只落在一个块里（LibreOffice / 有的 WPS 对重叠的块只认一个）。"""
    lo = min(CI(a) for a, _b, _c, _k in rules)
    hi = max(CI(b) for _a, b, _c, _k in rules)
    runs = []
    for i in range(lo, hi + 1):
        sig = tuple(j for j, (a, b, _c, _k) in enumerate(rules) if CI(a) <= i <= CI(b))
        if not sig:
            continue
        if runs and runs[-1][2] == sig and runs[-1][1] == i - 1:
            runs[-1][1] = i
        else:
            runs.append([i, i, sig])
    for i0, i1, sig in runs:
        for j in sig:
            _a, _b, cond, kw = rules[j]
            ws.conditional_formatting.add(f'{CL(i0)}{r0}:{CL(i1)}{r1}',
                                          FormulaRule(formula=[cond(CL(i0))], stopIfTrue=True, **kw))


def _scalars(ws, sl, sv, keys, formulas):
    rows = {k: 3 + i for i, k in enumerate(keys)}
    for k in keys:
        ws[f'{sl}{rows[k]}'] = k
        ws[f'{sv}{rows[k]}'] = formulas[k]
        ws[f'{sl}{rows[k]}'].font = ws[f'{sv}{rows[k]}'].font = F_HELP
    return lambda k: f'${sv}${rows[k]}'


def _year_parse(cell):
    """年份格：2026 / 202610 / 那年里的一个日期都认；空或认不出＝P_年度"""
    v = f'IFERROR(--TRIM({cell}&""),0)'
    return v, (lambda x: f'=IF(AND({x}>=2000,{x}<=2099,INT({x})={x}),{x},IF(AND({x}>=200001,{x}<=209912),INT({x}/100),'
                         f'IF(AND({x}>=36526,{x}<=73050),YEAR({x}),P_年度)))'), \
        (lambda x: f'=IF(TRIM({cell}&"")="",1,IF(OR(AND({x}>=2000,{x}<=2099,INT({x})={x}),AND({x}>=200001,{x}<=209912),'
                   f'AND({x}>=36526,{x}<=73050)),1,0))')


def _year_selector(ws, prompt):
    selector(ws, 'A3', '年份', 'B3', None)
    dv = DataValidation(type='whole', operator='between', formula1='2000', formula2='2099', allow_blank=True,
                        showErrorMessage=False, showInputMessage=True, promptTitle='提示', prompt=prompt)
    ws.add_data_validation(dv)
    dv.add('B3')


# ═══════════════════════════════ 利润表 ═══════════════════════════════
PL_LAST = 'N'
MCOL = [CL(2 + i) for i in range(12)]                  # B～M：1～12 月
PL_HDR, PL_R0 = 5, 6
PH = dict(月='AD', 年月='AE', 显示='AF', 估了='AG', 估值='AH', 上次='AI', 上次估值='AJ', 变动='AK', 折旧='AL', 建账前='AM',
          以后='AN')
SHOW_KINDS = EXPENSE_KINDS + ['其他收入', '买设备', '不算收支']   # 利润表上出现的资金归类（建账前的月份有这些才显示）

TIP_PL = ('💡 全自动不用填。黄格 B3 填年份（比如 2026；空着＝截止日期那年），出 1～12 月和全年。'
          '收入按发货日期（销售登记），料、外协、外购成品按到货日期（采购登记），工资、电费、费用、税金按资金台帐付钱的月份'
          '（填了所属月份按所属月份）。截止日期以后的月份空着。最下面灰字是备查，不进上面的数。金额显示到元。')

NOTES_PL = [
    '说明：',
    '① 金额按实际成交价（带票的是含税价），跟对账单、往来欠款一致，不是会计报表上的不含税数。',
    '② 生产这块不跟过程：料、外协、外购成品买进来就进成本；工资、电费、刀具辅料现买的等算生产成本；月底在【基础资料】⑤ 估一下库存值多少，'
    '差额调成本（库存变动＝上次估值－本次估值：库存比上次多，成本就少算）。',
    '③ 没估库存的月份，买进来的料可能还没用完，利润只供参考；看几个月合计更准。',
    '④ 收客户货款、付供应商货款是往来款，不进利润表（收入、成本按发货、到货算）；买设备不算当月费用（每月折旧在【基础资料】估一个数）；'
    '老板存取、借款、还款、保证金、内部转账不算收支。待定价的发货、到货收入成本里都没算，定了价补上单价就有了。',
    '⑤ 资金台帐里所属月份填到建账以前的（比如 7 月发的 6 月工资），算在那个月（表头浅色的列），跟【费用统计】【首页】一样：'
    '那个月没有收入，利润是负数，全年也含这笔，看全年时心里有数。所属月份填到截止日期以后的（提前付的），到那个月才算，先在最下面备查里列出来。',
    '⑥ 截止日期那个月只算到截止日（折旧按整月），跟别的整月比时注意。',
]


def _pl_items():
    sale = lambda t: (lambda y: f'ROUND(SUMIFS(销_金额,销_计入往来,1,销_年月,{y},销_日期,"<="&P_截止,销_业务类型,"{t}"),2)')
    pur = lambda t: (lambda y: f'ROUND(SUMIFS(采_金额,采_计入往来,1,采_年月,{y},采_日期,"<="&P_截止,采_采购类别,"{t}"),2)')
    kind = lambda k, col='资_支出额': (lambda y: f'ROUND(SUMIFS({col},资_归类,"{k}",资_所属年月,{y},资_日期,"<="&P_截止),2)')
    helper = lambda h: (lambda y, m: f'${PH[h]}${2 + m}')
    # (键, 项目名, 样式, 类型, 内容)；样式：''、key（小计）、big（三五六：带编号的数）、tot（利润）、pct、memo；
    # 类型：f（fn(y)）、h（每月辅助格）、sum（[(键, ±1)]）、pct（(分子键, 分母键)）、stk（库存估了没有）
    return [
        ('S', '一、销售收入（按发货日期）'),
        ('自产', '　　自产加工（接订单加工的）', '', 'f', sale('自产加工')),
        ('外购', '　　外购成品（买进来直接卖的）', '', 'f', sale('外购成品')),
        ('收入', '销售收入合计', 'key', 'sum', [('自产', 1), ('外购', 1)]),
        ('S', '二、销售成本'),
        ('原材料', '　　原材料', '', 'f', pur('原材料')),
        ('外协', '　　外协加工', '', 'f', pur('外协加工')),
        ('外购采', '　　外购成品（采购）', '', 'f', pur('外购成品')),
        ('刀具', '　　辅料刀具', '', 'f', pur('辅料刀具')),
        ('包装', '　　包装物', '', 'f', pur('包装物')),
        ('其他采', '　　其他采购', '', 'f', pur('其他')),
        ('生产', '　　生产工资电费等（资金台帐「生产成本」）', '', 'f', kind('生产成本')),
        ('折旧', '　　设备折旧（估）', '', 'h', helper('折旧')),
        ('库存', '　　库存变动（上次估值－本次估值）', '', 'h', helper('变动')),
        ('成本', '销售成本合计', 'key', 'sum', [(k, 1) for k in ('原材料', '外协', '外购采', '刀具', '包装', '其他采', '生产', '折旧', '库存')]),
        ('毛利', '三、毛利（＝销售收入－销售成本）', 'key', 'sum', [('收入', 1), ('成本', -1)]),
        ('毛利率', '　　毛利率（毛利÷销售收入）', 'pct', 'pct', ('毛利', '收入')),
        ('S', '四、期间费用'),
        ('销售费', '　　销售费用', '', 'f', kind('销售费用')),
        ('管理费', '　　管理费用', '', 'f', kind('管理费用')),
        ('财务费', '　　财务费用', '', 'f', kind('财务费用')),
        ('费用', '期间费用合计', 'key', 'sum', [('销售费', 1), ('管理费', 1), ('财务费', 1)]),
        ('税金', '五、税金', 'big', 'f', kind('税金')),
        ('S', '六、其他收入、其他支出'),
        ('其他收入', '　　其他收入（废料、利息等）', '', 'f', kind('其他收入', '资_净额')),
        ('其他支出', '　　其他支出（罚款、赞助等）', '', 'f', kind('其他支出')),
        ('利润', '七、利润（＝毛利－期间费用－税金＋其他收入－其他支出）', 'tot', 'sum',
         [('毛利', 1), ('费用', -1), ('税金', -1), ('其他收入', 1), ('其他支出', -1)]),
        ('利润率', '　　利润率（利润÷销售收入）', 'pct', 'pct', ('利润', '收入')),
        ('M', '备查（灰字，不进上面的数）'),
        ('估了', '库存估了没有（没估的月份利润只供参考）', 'memo', 'stk', None),
        ('待发', '待定价发货笔数（收入里没算）', 'memo', 'f',
         lambda y: f'COUNTIFS(销_待定价,1,销_计入往来,1,销_年月,{y},销_日期,"<="&P_截止)'),
        ('待到', '待定价到货笔数（成本里没算）', 'memo', 'f',
         lambda y: f'COUNTIFS(采_待定价,1,采_计入往来,1,采_年月,{y},采_日期,"<="&P_截止)'),
        ('设备', '买设备（不算当月费用）', 'memo', 'f', kind('买设备')),
        ('不算', '老板存取、借款等（不算收支；净额：＋进来 －出去）', 'memo', 'f', kind('不算收支', '资_净额')),
        ('以后', '已付、算在以后月份的费用（所属月份填到截止日期以后，到那个月才算）', 'memo', 'fut', helper('以后')),
    ]


def build_pl(wb):
    ws = wb[SH_PL]
    widths(ws, dict({'A': 40, PL_LAST: 13}, **{c: 11 for c in MCOL}))
    title(ws, '利 润 表（按月）', PL_LAST, C_VIEW, TIP_PL)

    # ── 第 3 行：年份 ──
    _year_selector(ws, '填年份，比如 2026；空着＝截止日期那年')
    SC = ['年份输入', '年份', '年份认出', '建账月']
    v, fy, fok = _year_parse('B3')
    S = _scalars(ws, 'AA', 'AB', SC, {
        '年份输入': f'={v}',
        '年份': fy('$AB$3'),
        '年份认出': fok('$AB$3'),
        '建账月': '=YEAR(P_建账日)*100+MONTH(P_建账日)',
    })
    ws.merge_cells('C3:F3')
    put(ws, 'C3', (f'="＝ "&{S("年份")}&" 年"&IF(TRIM(B3&"")="","（空着＝截止日期那年）","")'
                   f'&IF({S("年份认出")}=0,"（⚠ 没认出来，先按截止日期那年）","")'), F_NOTE, align=AL, border=False)
    ws.conditional_formatting.add('C3', FormulaRule(formula=[f'{S("年份认出")}=0'], font=F_RED))
    ws.merge_cells('G3:L3')
    put(ws, 'G3', '口径：金额按实际成交价（带票的是含税价），跟对账一致', F_NOTE, align=AL, border=False)
    home_link(ws, f'{PL_LAST}3')
    ws.row_dimensions[3].height = 26

    # ── 第 4 行：截止日、建账日 ──
    ws.merge_cells(f'A4:{PL_LAST}4')
    pre_n = f'SUM(${PH["建账前"]}$3:${PH["建账前"]}$14)'
    put(ws, 'A4', ('="截止日期 "&TEXT(P_截止,"yyyy-mm-dd")&"（【首页】改），建账日 "&TEXT(P_建账日,"yyyy-mm-dd")'
                   '&"：只算建账日～截止日期的发货、到货、付款；截止日期以后的月份空着"'
                   f'&IF(AND(YEAR(P_截止)={S("年份")},P_截止<DATE(YEAR(P_截止),MONTH(P_截止)+1,0)),'
                   '"；"&MONTH(P_截止)&" 月只算到 "&DAY(P_截止)&" 日（折旧按整月）","")'
                   f'&IF({pre_n}>0,"；表头浅色的是建账以前的月份，只有所属月份填到那个月的工资电费等","")&"。"'
                   f'&IF({S("年份")}<YEAR(P_建账日),"⚠ 选的年份在建账以前，基本没有数。","")'
                   f'&IF({S("年份")}>YEAR(P_截止),"⚠ 选的年份在截止日期以后，上面都是空的。","")'),
        F_NOTE, align=ALW, border=False)
    ws.row_dimensions[4].height = 30

    # ── 每月辅助（隐藏）：第 2+m 行 ──
    y0 = S('年份')
    for k, c in PH.items():
        ws[f'{c}2'] = k
        ws[f'{c}2'].font = F_HELP
    for m in range(1, 13):
        r = 2 + m
        h = lambda k: f'{PH[k]}{r}'
        bm = S('建账月')
        cash_n = '+'.join(f'COUNTIFS(资_所属年月,{h("年月")},资_资金有效,1,资_日期,"<="&P_截止,资_归类,"{k}")'
                          for k in SHOW_KINDS)
        f = {
            '月': m,
            '年月': f'={y0}*100+{m}',
            # 截止以后不显示；建账以前只有资金台帐所属月份填到这个月（利润表上出现的归类）才显示
            '显示': f'=IF({h("年月")}>P_截止年月,0,IF({h("年月")}>={bm},1,IF({cash_n}>0,1,0)))',
            '估了': f'=IF(COUNTIFS(库存_年月,{h("年月")},库存_有值,1)>0,1,0)',
            '估值': f'=SUMIFS(库存_估值,库存_年月,{h("年月")},库存_有值,1)',
            '上次': f'=IFERROR(LARGE(库存_键,COUNTIF(库存_键,">="&{h("年月")})+1),0)',
            # 上次估的是建账以前的月份、又填了期初库存（建账日的估值）→ 用期初库存
            '上次估值': (f'=IF(AND({h("上次")}>0,OR({h("上次")}>={bm},NOT(ISNUMBER(P_期初库存)))),'
                       f'SUMIFS(库存_估值,库存_年月,{h("上次")},库存_有值,1),IF(ISNUMBER(P_期初库存),P_期初库存,0))'),
            # 只算建账月以后（建账以前的月份不调成本）
            '变动': (f'=IF(AND({h("年月")}>={bm},{h("估了")}=1,OR({h("上次")}>0,ISNUMBER(P_期初库存))),'
                   f'ROUND({h("上次估值")}-{h("估值")},2),0)'),
            '折旧': f'=IF(AND({h("年月")}>={bm},{h("年月")}<=P_截止年月),P_月折旧,0)',
            '建账前': f'=IF(AND({h("显示")}=1,{h("年月")}<{bm}),1,0)',
            '以后': ('=ROUND(' + '+'.join(f'SUMIFS(资_支出额,资_归类,"{k}",资_所属年月,{h("年月")},资_日期,"<="&P_截止)'
                                          for k in EXPENSE_KINDS) + ',2)'),
        }
        for k, val in f.items():
            ws[f'{PH[k]}{r}'] = val
            ws[f'{PH[k]}{r}'].font = F_HELP

    # ── 第 5 行表头 ──
    header(ws, PL_HDR, [('A', f'="项目（"&{y0}&" 年）"')] + [(c, f'{m}月') for m, c in enumerate(MCOL, 1)]
           + [(PL_LAST, '全年')], C_VIEW, height=24)
    grey_h = cf_fill('FFA6A6A6')
    pre_h = cf_fill('FFA9D08E')
    for m, c in enumerate(MCOL, 1):
        ws.conditional_formatting.add(f'{c}{PL_HDR}', FormulaRule(formula=[f'${PH["显示"]}${2 + m}=0'], fill=grey_h, stopIfTrue=True))
        ws.conditional_formatting.add(f'{c}{PL_HDR}', FormulaRule(formula=[f'${PH["建账前"]}${2 + m}=1'], fill=pre_h,
                                                                  font=Font(bold=True, color='FF1F3864')))

    # ── 表身 ──
    items = _pl_items()
    R = {}
    r = PL_R0
    placed = []
    for it in items:
        placed.append((r, it))
        if it[0] not in ('S', 'M'):
            R[it[0]] = r
        r += 1
    last_item = r - 1
    for r, it in placed:
        if it[0] in ('S', 'M'):
            section(ws, r, 'A', PL_LAST, it[1], C_VIEW if it[0] == 'S' else C_MEMO)
            ws.row_dimensions[r].height = 20
            continue
        key, lbl, sty, typ, spec = it
        memo = sty == 'memo'
        fmt = PCT if sty == 'pct' else (INT if key in ('待发', '待到') else MONEY0)
        bold = sty in ('key', 'tot')
        fl = FILL_SUB if sty == 'key' else (FILL_TOT if sty == 'tot' else None)
        put(ws, f'A{r}', lbl, F_MEMOL if memo else (F_TXTB if (bold or sty == 'big') else F_TXT), fl, align=ALW)
        for m, c in enumerate(MCOL, 1):
            y = f'${PH["年月"]}${2 + m}'
            show = f'${PH["显示"]}${2 + m}'
            if typ == 'f':
                e = spec(y)
            elif typ == 'h':
                e = spec(y, m)
            elif typ == 'sum':
                e = 'ROUND(' + ''.join(f'{"+" if s > 0 else "-"}{c}{R[k]}' for k, s in spec).lstrip('+') + ',2)'
            elif typ == 'pct':
                a, b = spec
                e = f'IF({c}{R[b]}=0,"",{c}{R[a]}/{c}{R[b]})'
            elif typ == 'fut':   # 只在截止以后的月份显示（这些月份上面都是空的）
                e = spec(y, m)
            else:   # stk
                e = f'IF(${PH["估了"]}${2 + m}=1,"✓",IF({y}<{S("建账月")},"","没估"))'
            cf = (f'=IF(OR({y}<=P_截止年月,{e}=0),"",{e})' if typ == 'fut' else f'=IF({show}=0,"",{e})')
            put(ws, f'{c}{r}', cf, F_MEMO if memo else (F_AUTOB if bold else F_AUTO), fl, fmt,
                AC if typ == 'stk' else AR)
        rng = f'{MCOL[0]}{r}:{MCOL[-1]}{r}'
        if typ == 'pct':
            a, b = spec
            tot = f'=IF(N({PL_LAST}{R[b]})=0,"",{PL_LAST}{R[a]}/{PL_LAST}{R[b]})'
        elif typ == 'stk':
            tot = f'="估了 "&COUNTIF({rng},"✓")&" 个月"'
        else:
            tot = f'=ROUND(SUM({rng}),2)'
        put(ws, f'{PL_LAST}{r}', tot, F_MEMOL if memo else F_AUTOB, fl or (None if memo else FILL_AUTO), fmt,
            AC if typ == 'stk' else AR)
        ws.row_dimensions[r].height = 30 if len(lbl) > 19 else (20 if not memo else 17)
        if typ == 'fut':
            ws.row_dimensions[r].height = 30
    # 没估：橙色
    rs = R['估了']
    ws.conditional_formatting.add(f'B{rs}:M{rs}', FormulaRule(formula=[f'B{rs}="没估"'],
                                                               font=Font(name=YH, sz=9, bold=True, color='FFC65911')))
    # 说明
    r = last_item + 2
    for i, t in enumerate(NOTES_PL):
        ws.merge_cells(f'A{r}:{PL_LAST}{r}')
        put(ws, f'A{r}', t, F_NOTE if i else F_MEMOL, align=ALW, border=False)
        ws.row_dimensions[r].height = 30 if len(t) > 90 else 17
        r += 1
    last = r - 1

    hide(ws, 'AA', 'AB', *PH.values())
    ws.freeze_panes = f'B{PL_R0}'
    print_setup(ws, f'{PL_HDR}:{PL_HDR}', landscape=True)
    ws.print_area = f'A1:{PL_LAST}{last}'
    ws.oddFooter.center.text = '第 &P 页 共 &N 页'
    return R


# ═══════════════════════════════ 发票跟进 ═══════════════════════════════
IV_LAST = 'J'
IV_R1, IV_H1, IV_H2, IV_M0 = 6, 7, 8, 9              # ① 标题、两行表头、1 月所在行
IV_TOT = IV_M0 + 12                                  # ① 合计行
FLOW_R0 = IV_TOT + 2                                 # ②③④ 从这一行起往下排
N2, N3, N4 = 150, 150, 120                           # 最多列几家 / 几张
FLOW_N = (N2 + 4) + 1 + (N3 + 4) + 1 + (N4 + 5)
U0, U1 = 3, 3 + N_UNIT - 1                           # 每家一行的排序辅助：第 3 行起
KC, KS = 'AE', 'AG'                                  # 客户键、供应商键
KC_RNG, KS_RNG = f'${KC}${U0}:${KC}${U1}', f'${KS}${U0}:${KS}${U1}'
FH = dict(段='AH', 行='AI', 类='AJ', 第几='AK', 源='AL')

TIP_IV = ('💡 全自动（发票在【发票登记】登）。B3 年份管 ①，空着＝截止日期那年；D3 月份管 ④：填 1～12 按 B3 的年份，'
          '空着＝截止日期那个月，填 2026-9 连年份一起认。② ③ 是截至截止日期的数：未开票＝要开票的发货（含期初）－开出的发票，'
          '欠票＝要票的到货（含期初）－收到的发票。')


def build_inv(wb):
    ws = wb[SH_INVS]
    W = {'A': 24, 'B': 13.5, 'C': 14, 'D': 28, 'E': 17, 'F': 13.5, 'G': 13.5, 'H': 14, 'I': 16, 'J': 16}
    widths(ws, W)
    title(ws, '发 票 跟 进', IV_LAST, C_WL, TIP_IV)

    # ── 第 3 行：年份、月份 ──
    _year_selector(ws, '填年份，比如 2026；空着＝截止日期那年（管 ① 各月；D3 填 1～12 时 ④ 也用这个年份）')
    selector(ws, 'C3', '④ 月份', 'D3', None, '"1,2,3,4,5,6,7,8,9,10,11,12"',
             prompt='选 1～12（按 B3 的年份）；空着＝截止日期那个月；填 2026-9 或 202609 连年份一起认')
    v, fy, fok = _year_parse('B3')
    mv = '$AB$6'
    okm = f'AND({mv}>=200001,{mv}<=209912,INT({mv})={mv},MOD({mv},100)>=1,MOD({mv},100)<=12)'
    okd = f'AND({mv}>=36526,{mv}<=73050)'
    ok1 = f'AND({mv}>=1,{mv}<=12,INT({mv})={mv})'
    SC = ['年份输入', '年份', '年份认出', '月份输入', '清单年', '清单月', '月份认出', '月初', '月末', '清单止', '清单前', '清单数',
          '清单显示', '开出张数', '收到张数', '客户数', '客户显示', '供应商数', '供应商显示', 'L2', 's3', 'L3', 's4', 'L4', '尾',
          '末行', '客应开', '客已开', '客未开', '客欠款', '供应收', '供已收', '供欠票', '供欠他',
          '开价税', '开税额', '开不含税', '收价税', '收税额', '收不含税', '截止年月']
    rows = {k: 3 + i for i, k in enumerate(SC)}
    S = lambda k: f'$AB${rows[k]}'
    assert S('月份输入') == mv
    win = lambda d: f'{S("月初")},票_日期,"<="&{S("清单止")},票_方向,"{d}"'
    inv_sum = lambda col, d: f'=ROUND(SUMIFS({col},票_计入往来,1,票_日期,">="&{win(d)}),2)'
    unit_sum = lambda col, flag, key: f'=ROUND(SUMIFS({col},位_有效,1,{flag},1,{key},">0"),2)'
    f = {
        '年份输入': f'={v}',
        '年份': fy(S('年份输入')),
        '年份认出': fok(S('年份输入')),
        '月份输入': '=IFERROR(--TRIM(D3&""),0)',
        '清单年': f'=IF({okm},INT({mv}/100),IF({okd},YEAR({mv}),IF({ok1},{S("年份")},YEAR(P_截止))))',
        '清单月': f'=IF({ok1},{mv},IF({okm},MOD({mv},100),IF({okd},MONTH({mv}),MONTH(P_截止))))',
        '月份认出': f'=IF(TRIM(D3&"")="",1,IF(OR({ok1},{okm},{okd}),1,0))',
        '月初': f'=DATE({S("清单年")},{S("清单月")},1)',
        '月末': f'=DATE({S("清单年")},{S("清单月")}+1,0)',
        '清单止': f'=MIN({S("月末")},P_截止)',
        '清单前': f'=COUNTIF(票_排序键,"<"&{S("月初")}*{M})',
        '清单数': (f'=IF({S("清单止")}<{S("月初")},0,COUNTIFS(票_排序键,">="&{S("月初")}*{M},'
                 f'票_排序键,"<"&({S("清单止")}+1)*{M}))'),
        '清单显示': f'=MIN({S("清单数")},{N4})',
        '开出张数': f'=COUNTIFS(票_计入往来,1,票_日期,">="&{win("开出")})',
        '收到张数': f'=COUNTIFS(票_计入往来,1,票_日期,">="&{win("收到")})',
        '客户数': f'=COUNT({KC_RNG})',
        '客户显示': f'=MIN({S("客户数")},{N2})',
        '供应商数': f'=COUNT({KS_RNG})',
        '供应商显示': f'=MIN({S("供应商数")},{N3})',
        'L2': f'={S("客户显示")}+4',
        's3': f'={S("L2")}+1',
        'L3': f'={S("供应商显示")}+4',
        's4': f'={S("s3")}+{S("L3")}+1',
        'L4': f'={S("清单显示")}+5',
        '尾': f'={S("s4")}+{S("L4")}',
        '末行': f'={FLOW_R0}+{S("尾")}-1',
        '客应开': unit_sum('位_应开票', '位_是客户', KC_RNG),
        '客已开': unit_sum('位_已开票', '位_是客户', KC_RNG),
        '客未开': unit_sum('位_未开票', '位_是客户', KC_RNG),
        '客欠款': unit_sum('位_应收', '位_是客户', KC_RNG),
        '供应收': unit_sum('位_应收票', '位_是供应商', KS_RNG),
        '供已收': unit_sum('位_已收票', '位_是供应商', KS_RNG),
        '供欠票': unit_sum('位_欠票', '位_是供应商', KS_RNG),
        '供欠他': unit_sum('位_应付', '位_是供应商', KS_RNG),
        '开价税': inv_sum('票_价税合计', '开出'),
        '开税额': inv_sum('票_税额', '开出'),
        '开不含税': inv_sum('票_不含税', '开出'),
        '收价税': inv_sum('票_价税合计', '收到'),
        '收税额': inv_sum('票_税额', '收到'),
        '收不含税': inv_sum('票_不含税', '收到'),
        '截止年月': '=P_截止年月',
    }
    for k in SC:
        ws[f'AA{rows[k]}'] = k
        ws[f'AB{rows[k]}'] = f[k]
        ws[f'AA{rows[k]}'].font = ws[f'AB{rows[k]}'].font = F_HELP
    ws.merge_cells('E3:I3')
    put(ws, 'E3', (f'="实际用：① "&{S("年份")}&" 年　④ "&{S("清单年")}&" 年 "&{S("清单月")}&" 月"'
                   f'&IF({S("年份认出")}=0,"　⚠ 年份没认出来（先按截止日期那年）","")'
                   f'&IF({S("月份认出")}=0,"　⚠ 月份没认出来（先按截止日期那个月）","")'), F_NOTE, align=ALW, border=False)
    ws.conditional_formatting.add('E3', FormulaRule(formula=[f'OR({S("年份认出")}=0,{S("月份认出")}=0)'], font=F_RED))
    home_link(ws, f'{IV_LAST}3')
    ws.row_dimensions[3].height = 26

    # ── 第 4、5 行：KPI（截至截止日） ──
    cust = '位_有效,1,位_是客户,1,位_未开票'
    supp = '位_有效,1,位_是供应商,1,位_欠票'
    kpis = [('A', 'A', '截至日期', '=P_截止', DATE),
            ('B', 'C', f'="客户没开票（"&COUNTIFS({cust},">0.005")&" 家）"', f'=ROUND(SUMIFS(位_未开票,{cust},">0.005"),2)', MONEY),
            ('D', 'D', f'="多开给客户的票（"&COUNTIFS({cust},"<-0.005")&" 家）"', f'=ROUND(-SUMIFS(位_未开票,{cust},"<-0.005"),2)', MONEY),
            ('E', 'F', f'="供应商欠票（"&COUNTIFS({supp},">0.005")&" 家）"', f'=ROUND(SUMIFS(位_欠票,{supp},">0.005"),2)', MONEY),
            ('G', 'H', f'="供应商多给的票（"&COUNTIFS({supp},"<-0.005")&" 家）"', f'=ROUND(-SUMIFS(位_欠票,{supp},"<-0.005"),2)', MONEY)]
    for c1, c2, lbl, val, fmt in kpis:
        if c1 != c2:
            ws.merge_cells(f'{c1}4:{c2}4')
            ws.merge_cells(f'{c1}5:{c2}5')
        put(ws, f'{c1}4', lbl, F_KPI_L, LBL, align=ACW)
        put(ws, f'{c1}5', val, F_KPI_B, FILL_TOT, fmt, AC)
        if c1 != c2:
            put(ws, f'{c2}4', None, F_KPI_L, LBL)
            put(ws, f'{c2}5', None, F_KPI_B, FILL_TOT)
    ws.merge_cells('I4:J5')
    put(ws, 'I4', '多开、多给：票比货多，一般是给建账以前的货开的票', F_NOTE, align=ALW, border=False)
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 26

    # ── ① 各月 ──
    section(ws, IV_R1, 'A', 'I', f'="① "&{S("年份")}&" 年各月开出、收到的发票（按开票日期；只算建账日～截止日期）"', C_WL)
    ws.row_dimensions[IV_R1].height = 20
    ws.merge_cells(f'A{IV_H1}:A{IV_H2}')
    ws.merge_cells(f'B{IV_H1}:D{IV_H1}')
    ws.merge_cells(f'E{IV_H1}:H{IV_H1}')
    ws.merge_cells(f'I{IV_H1}:I{IV_H2}')
    header(ws, IV_H1, [('A', '月份'), ('B', '开出（开给客户的）'), ('C', None), ('D', None), ('E', '收到（供应商开来的）'),
                       ('F', None), ('G', None), ('H', None), ('I', '开出税额－收到专票税额\n（只供参考，不是报税数）')], C_WL, height=22)
    header(ws, IV_H2, [('A', None), ('B', '张数'), ('C', '价税合计'), ('D', '税额'), ('E', '张数'), ('F', '价税合计'), ('G', '税额'),
                       ('H', '其中专票税额\n（类型空着的算普票）'), ('I', None)], C_WL, height=42)
    base = '票_计入往来,1,票_年月,{y},票_日期,"<="&P_截止,票_方向,"{d}"'
    cols = {
        'B': lambda y: f'COUNTIFS({base.format(y=y, d="开出")})',
        'C': lambda y: f'ROUND(SUMIFS(票_价税合计,{base.format(y=y, d="开出")}),2)',
        'D': lambda y: f'ROUND(SUMIFS(票_税额,{base.format(y=y, d="开出")}),2)',
        'E': lambda y: f'COUNTIFS({base.format(y=y, d="收到")})',
        'F': lambda y: f'ROUND(SUMIFS(票_价税合计,{base.format(y=y, d="收到")}),2)',
        'G': lambda y: f'ROUND(SUMIFS(票_税额,{base.format(y=y, d="收到")}),2)',
        'H': lambda y: f'ROUND(SUMIFS(票_税额,{base.format(y=y, d="收到")},票_发票类型,"*专*"),2)',
    }
    for m in range(1, 13):
        r = IV_M0 + m - 1
        y = f'({S("年份")}*100+{m})'
        put(ws, f'A{r}', f'{m}月', F_TXT, align=AC)
        for c, fn in cols.items():
            put(ws, f'{c}{r}', f'=IF({y}>P_截止年月,"",{fn(y)})', F_AUTO, None, INT if c in 'BE' else MONEY, AR)
        put(ws, f'I{r}', f'=IF({y}>P_截止年月,"",ROUND(D{r}-H{r},2))', F_AUTO, FILL_AUTO, MONEY, AR)
        ws.conditional_formatting.add(f'A{r}', FormulaRule(formula=[f'{y}>{S("截止年月")}'], font=Font(color='FFA6A6A6')))
    put(ws, f'A{IV_TOT}', '合计', F_TXTB, FILL_TOT, align=AC)
    for c in 'BCDEFGHI':
        put(ws, f'{c}{IV_TOT}', f'=ROUND(SUM({c}{IV_M0}:{c}{IV_M0 + 11}),2)', F_AUTOB, FILL_TOT, INT if c in 'BE' else MONEY, AR)

    # ── 每家一行的排序辅助（隐藏）：客户按未开票、供应商按欠票，从大到小 ──
    for c, t in (('AD', '客户未开票'), ('AE', '客户键'), ('AF', '供应商欠票'), ('AG', '供应商键')):
        ws[f'{c}{U0 - 1}'] = t
        ws[f'{c}{U0 - 1}'].font = F_HELP
    for i in range(N_UNIT):
        r, n = U0 + i, i + 1
        for vc, kc, flag, val in (('AD', 'AE', '位_是客户', '位_未开票'), ('AF', 'AG', '位_是供应商', '位_欠票')):
            ws[f'{vc}{r}'] = f'=IF(AND(INDEX({flag},{n})=1,ABS(INDEX({val},{n}))>0.005),INDEX({val},{n}),"")'
            ws[f'{kc}{r}'] = f'=IF({vc}{r}="","",COUNTIF(${vc}${U0}:${vc}${U1},">"&{vc}{r})*1000+{n})'
            ws[f'{vc}{r}'].font = ws[f'{kc}{r}'].font = F_HELP

    # ── ②③④ 往下排 ──
    for k, c in FH.items():
        ws[f'{c}{FLOW_R0 - 1}'] = k
        ws[f'{c}{FLOW_R0 - 1}'].font = F_HELP
    flow_cols = 'ABCDEFGHIJ'
    # 同一列在 ②③ 和 ④ 里有的是金额、有的是文字：金额用带「* 」填充的格式（数字总是靠右），文字按列的对齐
    MONEY_R = '* #,##0.00;[Red]* -#,##0.00;* "-";@'
    fm = {'A': DATE, 'B': MONEY_R, 'C': MONEY_R, 'D': MONEY_R, 'E': MONEY_R, 'F': DATE, 'G': MONEY_R, 'H': MONEY_R, 'I': MONEY_R}
    al = {'A': 'left', 'B': 'center', 'C': 'center', 'D': 'center', 'E': 'center', 'F': 'center', 'G': 'center', 'H': 'center',
          'I': 'center', 'J': 'center'}
    U = lambda nm, ix: f'INDEX({nm},{ix})'
    neg = lambda val, word, extra='': f'IF({val}<-0.005,"{word} "&TEXT(-{val},"#,##0.00")&"{extra}",{val})'
    cutoff = 'TEXT(P_截止,"yyyy-mm-dd")'
    for j in range(FLOW_N):
        r = FLOW_R0 + j
        seg, o, kd, k, ix = (f'${FH[x]}{r}' for x in ('段', '行', '类', '第几', '源'))
        ws[f'{FH["段"]}{r}'] = f'=IF({j}<{S("s3")},2,IF({j}<{S("s4")},3,IF({j}<{S("尾")},4,0)))'
        ws[f'{FH["行"]}{r}'] = f'=IF({seg}=2,{j},IF({seg}=3,{j}-{S("s3")},IF({seg}=4,{j}-{S("s4")},0)))'
        n23 = f'IF({seg}=2,{S("客户显示")},{S("供应商显示")})'
        n4 = S('清单显示')
        ws[f'{FH["类"]}{r}'] = (f'=IF({seg}=0,0,IF({o}=0,1,IF({o}=1,2,IF({o}=2,3,IF({seg}=4,'
                               f'IF({o}=3,4,IF({o}<4+{n4},5,IF({o}=4+{n4},6,0))),'
                               f'IF({o}<3+{n23},5,IF({o}=3+{n23},6,0)))))))')
        ws[f'{FH["第几"]}{r}'] = f'=IF({kd}=5,IF({seg}=4,{o}-3,{o}-2),0)'
        ws[f'{FH["源"]}{r}'] = (f'=IF({k}=0,0,IF({seg}=2,MOD(SMALL({KC_RNG},{k}),1000),IF({seg}=3,MOD(SMALL({KS_RNG},{k}),1000),'
                               f'MOD(SMALL(票_排序键,{S("清单前")}+{k}),{M}))))')
        for c in FH.values():
            ws[f'{c}{r}'].font = F_HELP
        # 每列：{行类: (段2, 段3, 段4)}
        item = {
            'A': (U('位_简称', ix), U('位_简称', ix), U('票_日期', ix)),
            'B': (U('位_应开票', ix), U('位_应收票', ix), U('票_方向', ix)),
            'C': (U('位_已开票', ix), U('位_已收票', ix), U('票_单位', ix)),
            'D': (neg(U('位_未开票', ix), '多开', '（含建账前的货）'), neg(U('位_欠票', ix), '多给'), U('票_发票号码', ix)),
            'E': (neg(U('位_应收', ix), '预收'), neg(U('位_应付', ix), '预付'), U('票_发票类型', ix)),
            'F': (f'IF({U("位_最后发货", ix)}>0,{U("位_最后发货", ix)},"")', f'IF({U("位_最后到货", ix)}>0,{U("位_最后到货", ix)},"")',
                  f'IF(MOD(ROUND({U("票_税率", ix)}*1000,0),10)=0,TEXT({U("票_税率", ix)},"0%"),TEXT({U("票_税率", ix)},"0.0%"))'),
            'G': (U('位_联系人', ix), U('位_联系人', ix), U('票_价税合计', ix)),
            'H': (U('位_电话', ix), U('位_电话', ix), U('票_税额', ix)),
            'I': ('""', '""', U('票_不含税', ix)),
            'J': ('""', '""', U('票_对应单号', ix)),
        }
        title_ = {
            'A': ('"② 客户开票情况"', '"③ 供应商欠票"', f'"④ "&{S("清单年")}&"年"&{S("清单月")}&"月发票清单"'),
            'D': (f'"截至 "&{cutoff}', f'"截至 "&{cutoff}', '"按开票日期排"'),
            'E': ('"没开票多的在前"', '"欠票多的在前"', '""'),
        }
        head = {
            'A': ('"客户"', '"供应商"', '"日期"'), 'B': ('"应开票"', '"应收票"', '"开出/收到"'),
            'C': ('"已开票"', '"已收票"', '"往来单位"'), 'D': ('"未开票"', '"欠票"', '"发票号码"'),
            'E': ('"现在欠款"', '"现在欠他"', '"类型"'), 'F': ('"最后发货"', '"最后到货"', '"税率"'),
            'G': ('"联系人"', '"联系人"', '"价税合计"'), 'H': ('"电话"', '"电话"', '"税额"'),
            'I': ('""', '""', '"不含税"'), 'J': ('""', '""', '"对应单号"'),
        }
        tot = {
            'A': (f'"合计（全部 "&{S("客户数")}&" 家）"', f'"合计（全部 "&{S("供应商数")}&" 家）"',
                  f'"开出合计（全部 "&{S("开出张数")}&" 张）"'),
            'B': (S('客应开'), S('供应收'), '""'), 'C': (S('客已开'), S('供已收'), '""'),
            'D': (S('客未开'), S('供欠票'), '""'), 'E': (S('客欠款'), S('供欠他'), '""'),
            'G': ('""', '""', S('开价税')), 'H': ('""', '""', S('开税额')), 'I': ('""', '""', S('开不含税')),
        }
        tot2 = {'A': f'"收到合计（全部 "&{S("收到张数")}&" 张）"', 'G': S('收价税'), 'H': S('收税额'), 'I': S('收不含税')}
        trail = {
            'A': (f'IF({S("客户数")}=0,"客户的票都开齐了","共 "&{S("客户数")}&" 家")',
                  f'IF({S("供应商数")}=0,"供应商的票都收齐了","共 "&{S("供应商数")}&" 家")',
                  f'IF({S("清单数")}=0,IF({S("月初")}>P_截止,"这个月在截止日期以后","这个月没有发票"),"共 "&{S("清单数")}&" 张")'),
            'D': (f'IF({S("客户数")}>{N2},"只列了前 {N2} 家（合计是全部的）","")',
                  f'IF({S("供应商数")}>{N3},"只列了前 {N3} 家（合计是全部的）","")',
                  f'IF({S("清单数")}>{N4},"只列了前 {N4} 张（合计是全部的）","")'),
        }

        def by_seg(t):
            if t is None:
                return '""'
            if isinstance(t, str):
                return t
            a, b, c = t
            if a == b == c:
                return a
            if a == b:
                return f'IF({seg}=4,{c},{a})'
            return f'IF({seg}=2,{a},IF({seg}=3,{b},{c}))'

        for c in flow_cols:
            parts = [(5, by_seg(item.get(c))), (1, by_seg(title_.get(c))), (2, by_seg(head.get(c))), (3, by_seg(tot.get(c))),
                     (4, by_seg(tot2.get(c))), (6, by_seg(trail.get(c)))]
            parts = [(x, e) for x, e in parts if e != '""']
            out = '""'
            for x, e in reversed(parts):
                out = f'IF({kd}={x},{e},{out})'
            cell = ws[f'{c}{r}']
            cell.value = '=' + out
            cell.font = F_TXT
            cell.alignment = Alignment(horizontal=al[c], vertical='center', shrink_to_fit=(c in 'CDE'))
            if c in fm:
                cell.number_format = fm[c]
    r0, r1 = FLOW_R0, FLOW_R0 + FLOW_N - 1
    K = lambda: f'${FH["类"]}{r0}'
    SG = lambda: f'${FH["段"]}{r0}'
    st_title = dict(fill=cf_fill(C_WL), font=Font(bold=True, color='FFFFFFFF'))
    st_head = dict(fill=cf_fill('FFF8CBAD'), font=Font(bold=True, color='FF000000'), border=CF_BD)
    st_tot = dict(fill=cf_fill('FFFCE4D6'), font=Font(bold=True), border=CF_BD)
    seg4 = lambda cond: (lambda c: f'AND({cond},{SG()}=4)')
    cf_blocks(ws, r0, r1, [       # ②③ 只用到 H 列，I、J 只给 ④
        ('A', 'H', lambda c: f'{K()}=1', st_title),
        ('I', IV_LAST, seg4(f'{K()}=1'), st_title),
        ('A', 'H', lambda c: f'{K()}=2', st_head),
        ('I', IV_LAST, seg4(f'{K()}=2'), st_head),
        ('A', 'H', lambda c: f'OR({K()}=3,{K()}=4)', st_tot),
        ('I', IV_LAST, seg4(f'OR({K()}=3,{K()}=4)'), st_tot),
        ('A', IV_LAST, lambda c: f'{K()}=6', dict(font=Font(color='FF808080'))),
        ('D', 'D', lambda c: f'AND({K()}=5,{SG()}<4,LEFT($D{r0},1)="多")', dict(font=Font(bold=True, color='FFC65911'), border=CF_BD)),
        ('E', 'E', lambda c: f'AND({K()}=5,{SG()}<4,LEFT($E{r0},1)="预")', dict(font=Font(bold=True, color='FF2F75B5'), border=CF_BD)),
        ('B', 'B', lambda c: f'AND({K()}=5,{SG()}=4,$B{r0}="收到")', dict(font=Font(color='FFC65911'), border=CF_BD)),
        ('A', 'H', lambda c: f'{K()}=5', dict(border=CF_BD)),
        ('I', IV_LAST, seg4(f'{K()}=5'), dict(border=CF_BD)),
    ])

    hide(ws, 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', *FH.values())
    ws.freeze_panes = 'A4'
    print_setup(ws, None, landscape=True)
    ws.oddFooter.center.text = '第 &P 页 共 &N 页'
    q = f"'{ws.title}'"
    ws.defined_names['Print_Area'] = DefinedName(
        'Print_Area', attr_text=f'{q}!$A$1:INDEX({q}!${IV_LAST}$1:${IV_LAST}${r1},{q}!{S("末行")})')
    return S


def build(wb, ctx=None):
    build_pl(wb)
    build_inv(wb)
