# -*- coding: utf-8 -*-
"""录入表（蓝）：收支登记（登记汇总表）、应收应付登记；基础资料（灰）：基础资料、往来单位。
   录入表只放录入列，外加最右边两列只给人看的显示列（别的公式不引用它们，插进来的行没有这两格也不影响汇总）。"""
import datetime as dt
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from layout import *
from common import *


def _d(v):
    if v in (None, ''):
        return None
    if isinstance(v, (dt.date, dt.datetime)):
        return dt.datetime(v.year, v.month, v.day)
    try:
        return dt.datetime.strptime(str(v)[:10], '%Y-%m-%d')
    except ValueError:
        return v


def dv_dec(ws, sqref, what='金额'):
    dv = DataValidation(type='decimal', operator='between', formula1='-999999999', formula2='999999999', allow_blank=True,
                        showErrorMessage=True, errorStyle='stop', errorTitle=what, error=f'{what}要填数字（不要带文字、单位）')
    ws.add_data_validation(dv)
    dv.add(sqref)


def _frame(ws, last_col, hdr, heads, color, tip, W, r0, nstyle, fm, al, freeze):
    widths(ws, W)
    title(ws, ws.title, last_col, color, tip)
    header(ws, hdr, heads, color)
    style_rows(ws, r0, r0 + nstyle - 1, [c for c, _ in heads], fmts=fm, aligns=al)
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f'A{hdr}:{last_col}{r0 + nstyle - 1}'


def _write_rows(ws, rows, cols, r0, date_keys=('日期',), num_keys=()):
    for i, x in enumerate(rows):
        r = r0 + i
        for k, col in cols.items():
            v = x.get(k)
            if v in (None, ''):
                continue
            if k in date_keys:
                v = _d(v)
            elif v == 'QP':                      # 金额＝数量×单价（跟用户原表一样写公式）
                v = f'=ROUND({cols["数量"]}{r}*{cols["单价"]}{r},2)'
            ws[f'{col}{r}'] = v


# ───────────────────────── 收支登记（登记汇总表） ─────────────────────────
def build_cash(ws, ctx):
    C = CASH_COLS
    S = CASH_SHOW
    heads = [(C['日期'], '日期'), (C['板块'], '业务板块'), (C['收支项目'], '收支项目'), (C['摘要'], '摘要'),
             (C['往来单位'], '往来单位\n（客户/供应商）'), (C['数量'], '数量\n（KG）'), (C['单价'], '单价'),
             (C['收入'], '收入金额'), (C['支出'], '支出金额'), (C['账户'], '账户\n（空＝现金）'), (C['经手人'], '经手人'),
             (C['备注'], '备注'), (S['结余'], '本账户结余\n（自动）'), (S['校验'], '这一行的问题\n（自动）')]
    tip = ('💡 所有板块的钱都在这一张表登记，一笔一行：日期、业务板块、收支项目都从下拉选，各板块表、查询、资金台账自动拆分汇总。'
           '收入/支出填金额；粮食买卖、烘干按斤算的，也可以只填数量和单价，金额空着就按「数量×单价」算（收支项目是收入类算收入，支出类算支出）。'
           '收回以前的欠款选「收回欠款」、付掉欠人家的钱选「支付欠款」，往来单位一定要选，应收应付和对账单会自动冲掉。'
           '现金存银行、银行取现：记两行「内部转账」（一行支出、一行收入），不算收支。可以插行、删行、排序（整行一起动）。')
    W = {C['日期']: 11, C['板块']: 12, C['收支项目']: 12, C['摘要']: 26, C['往来单位']: 14, C['数量']: 10, C['单价']: 9,
         C['收入']: 13, C['支出']: 13, C['账户']: 10, C['经手人']: 8, C['备注']: 18, S['结余']: 14, S['校验']: 34}
    fm = {C['日期']: DATE, C['数量']: '#,##0.##', C['单价']: '#,##0.00##', C['收入']: MONEY, C['支出']: MONEY, S['结余']: MONEY}
    al = {C['摘要']: AL, C['备注']: AL, C['收入']: AR, C['支出']: AR, S['结余']: AR, S['校验']: AL, C['数量']: AR, C['单价']: AR}
    rows = ctx.get('cash', [])
    nshow = min(max(len(rows) + 1000, 1500), N_CASH)
    _frame(ws, CASH_LAST, CASH_HDR, heads, C_IN, tip, W, CASH_R0, nshow, fm, al, f'D{CASH_R0}')
    ws.row_dimensions[CASH_HDR].height = 40
    # 第 3 行：合计（筛选后只合计看得见的行）
    last = CASH_R0 + N_CASH + 2000
    put(ws, 'A3', '合计（筛选后只算看得见的行）', F_KPI_L, fill('FFD9E1F2'), align=AC)
    ws.merge_cells('A3:G3')
    for col in (C['收入'], C['支出']):
        put(ws, f'{col}3', f'=SUBTOTAL(9,{col}{CASH_R0}:{col}{last})', F_AUTOB, FILL_AUTO, MONEY, AR)
    put(ws, f'{C["账户"]}3', '收支相抵', F_KPI_L, fill('FFD9E1F2'), align=AC)
    put(ws, f'{C["经手人"]}3', f'={C["收入"]}3-{C["支出"]}3', F_AUTOB, FILL_AUTO, MONEY, AR)
    ws.merge_cells(f'{C["经手人"]}3:{C["备注"]}3')
    # 数据
    _write_rows(ws, rows, C, CASH_R0)
    # 显示列
    base = f'ROW()-ROW(${C["日期"]}${CASH_HDR})'
    for i in range(nshow):
        r = CASH_R0 + i
        n = base
        ws[f'{S["结余"]}{r}'] = (f'=IFERROR(IF(INDEX(收_有效,{n})<>1,"",ROUND(SUMIFS(账户_期初余额,账户_名称,INDEX(收_账户,{n}))'
                                f'+SUMIFS(收_净额,收_账户,INDEX(收_账户,{n}),收_排序键,"<="&INDEX(收_排序键,{n})),2)),"")')
        ws[f'{S["校验"]}{r}'] = f'=IFERROR(INDEX(收_校验,{n})&"","")'
        for col in (S['结余'], S['校验']):
            ws[f'{col}{r}'].fill = FILL_AUTO
            ws[f'{col}{r}'].font = F_AUTO
    ws.conditional_formatting.add(f'{S["校验"]}{CASH_R0}:{S["校验"]}{CASH_R0 + nshow - 1}',
                                  FormulaRule(formula=[f'LEFT({S["校验"]}{CASH_R0},1)="✗"'], font=F_RED))
    # 下拉、格式校验（覆盖到容量）
    rng = lambda col: f'{col}{CASH_R0}:{col}{CASH_R0 + N_CASH - 1}'
    dv_date(ws, rng(C['日期']))
    dv_list(ws, rng(C['板块']), '=板块列表', prompt='从【基础资料】的业务板块里选；账户之间倒钱可以不选')
    dv_list(ws, rng(C['收支项目']), '=收支项目列表', prompt='收回以前的欠款选「收回欠款」，付掉欠人家的钱选「支付欠款」')
    dv_list(ws, rng(C['往来单位']), '=往来单位列表', stop=False, prompt='客户、供应商从【往来单位】选；收回/支付欠款一定要选')
    dv_list(ws, rng(C['账户']), '=账户列表', prompt='空着＝基础资料里第一个账户（现金）')
    dv_list(ws, rng(C['经手人']), '=经手人列表', stop=False)
    for k in ('数量', '单价'):
        dv_dec(ws, rng(C[k]), k)
    for k in ('收入', '支出'):
        dv_dec(ws, rng(C[k]))
    print_setup(ws, f'{CASH_HDR}:{CASH_HDR}')


# ───────────────────────── 应收应付登记 ─────────────────────────
def build_wl(ws, ctx):
    C = WL_COLS
    S = WL_SHOW
    heads = [(C['日期'], '日期'), (C['板块'], '业务板块'), (C['往来单位'], '往来单位'), (C['类型'], '应收 / 应付'),
             (C['收支项目'], '业务内容\n（收支项目）'), (C['摘要'], '摘要'), (C['数量'], '数量\n（KG）'), (C['单价'], '单价'),
             (C['金额'], '金额\n（空＝数量×单价）'), (C['约定日期'], '约定收/付款日期\n（空＝按账期）'), (C['经手人'], '经手人'),
             (C['备注'], '备注'), (S['未结'], '这笔还没结的\n（自动）'), (S['校验'], '这一行的问题\n（自动）')]
    tip = ('💡 赊出去、赊进来（钱当时没收/没付）的业务在这里登记：卖给人家没收钱＝应收，买人家的没给钱＝应付。'
           '金额空着就按「数量×单价」算；折让、少收少付记负数。以后收到钱/付了钱，到【收支登记】记一笔「收回欠款」「支付欠款」并选同一个往来单位，'
           '这里「这笔还没结的」会按先欠先还自动减掉。建账以前的老欠款也在这里记一行（摘要写「期初余额」）。可以插行、删行。')
    W = {C['日期']: 11, C['板块']: 12, C['往来单位']: 14, C['类型']: 8, C['收支项目']: 12, C['摘要']: 24, C['数量']: 10,
         C['单价']: 9, C['金额']: 13, C['约定日期']: 13, C['经手人']: 8, C['备注']: 16, S['未结']: 13, S['校验']: 32}
    fm = {C['日期']: DATE, C['约定日期']: DATE, C['数量']: '#,##0.##', C['单价']: '#,##0.00##', C['金额']: MONEY, S['未结']: MONEY}
    al = {C['摘要']: AL, C['备注']: AL, C['金额']: AR, S['未结']: AR, S['校验']: AL, C['数量']: AR, C['单价']: AR}
    rows = ctx.get('wl', [])
    nshow = min(max(len(rows) + 500, 800), N_WL)
    _frame(ws, WL_LAST, WL_HDR, heads, C_IN, tip, W, WL_R0, nshow, fm, al, f'D{WL_R0}')
    ws.row_dimensions[WL_HDR].height = 40
    last = WL_R0 + N_WL + 2000
    put(ws, 'A3', '合计（筛选后只算看得见的行）', F_KPI_L, fill('FFD9E1F2'), align=AC)
    ws.merge_cells('A3:H3')
    put(ws, f'{C["金额"]}3', f'=SUBTOTAL(9,{C["金额"]}{WL_R0}:{C["金额"]}{last})', F_AUTOB, FILL_AUTO, MONEY, AR)
    _write_rows(ws, rows, C, WL_R0, date_keys=('日期', '约定日期'))
    base = f'ROW()-ROW(${C["日期"]}${WL_HDR})'
    for i in range(nshow):
        r = WL_R0 + i
        ws[f'{S["未结"]}{r}'] = f'=IFERROR(IF(INDEX(往_有效,{base})<>1,"",INDEX(往_未结,{base})),"")'
        ws[f'{S["校验"]}{r}'] = f'=IFERROR(INDEX(往_校验,{base})&"","")'
        for col in (S['未结'], S['校验']):
            ws[f'{col}{r}'].fill = FILL_AUTO
            ws[f'{col}{r}'].font = F_AUTO
    ws.conditional_formatting.add(f'{S["校验"]}{WL_R0}:{S["校验"]}{WL_R0 + nshow - 1}',
                                  FormulaRule(formula=[f'LEFT({S["校验"]}{WL_R0},1)="✗"'], font=F_RED))
    rng = lambda col: f'{col}{WL_R0}:{col}{WL_R0 + N_WL - 1}'
    dv_date(ws, rng(C['日期']))
    dv_date(ws, rng(C['约定日期']))
    dv_list(ws, rng(C['板块']), '=板块列表')
    dv_list(ws, rng(C['往来单位']), '=往来单位列表', prompt='从【往来单位】选；新单位先去那里加一行')
    dv_list(ws, rng(C['类型']), '"' + ','.join(WL_TYPES) + '"', prompt='卖出去没收钱＝应收；买进来没付钱＝应付')
    dv_list(ws, rng(C['收支项目']), '=收支项目列表', stop=False)
    dv_list(ws, rng(C['经手人']), '=经手人列表', stop=False)
    for k in ('数量', '单价'):
        dv_dec(ws, rng(C[k]), k)
    dv_dec(ws, rng(C['金额']))
    print_setup(ws, f'{WL_HDR}:{WL_HDR}')


# ───────────────────────── 基础资料 ─────────────────────────
def build_base(ws, ctx):
    W = {'A': 14, 'B': 13, 'C': 14, 'D': 13, 'E': 16, 'F': 2, 'G': 12, 'H': 13, 'I': 14, 'J': 2, 'K': 13, 'L': 7, 'M': 9,
         'N': 7, 'O': 22, 'P': 2, 'Q': 10, 'R': 14}
    widths(ws, W)
    title(ws, '基 础 资 料', 'R', C_MASTER,
          '💡 一次填好、偶尔改：单位名称、建账日期；业务板块（各板块表、查询按它分）；账户（现金、银行，填建账日的余额）；'
          '收支项目（收支登记的下拉，「方向」「用途」「库存」决定怎么算）；经手人。往来单位在【往来单位】那张表。'
          '每块都可以在中间插行加一个，名称不要重复，中间不要留空行。')
    pv = ctx.get('params', {})
    put(ws, 'A3', '① 参数', F_SEC, fill(C_MASTER), align=AL)
    ws.merge_cells('A3:C3')
    for k, lbl, default in PARAMS:
        r = PAR_ROW[k]
        put(ws, f'A{r}', lbl, F_KPI_L, fill('FFD9E1F2'), align=AL)
        ws.merge_cells(f'A{r}:B{r}')
        v = pv.get(k, default)
        if k == '建账日':
            v = _d(v)
        put(ws, f'C{r}', v, F_IN, FILL_IN, DATE if k == '建账日' else None, AC)
    dv_date(ws, f'C{PAR_ROW["建账日"]}')
    notes = ['收支项目的「方向」：收入类 / 支出类 / 双向（内部转账、板块调拨这种两头都有的）。',
             '「用途」：普通＝一般收支；冲应收＝收回欠款（减应收）；冲应付＝支付欠款（减应付）；固定资产＝买设备、房屋（查询里「固定资产」）；内部转账＝账户之间倒钱（不算收支）。',
             '「库存」：入库＝粮食购进（加粮食存量）；出库＝粮食销售（减粮食存量）。按 KG 记数量。',
             '默认账期：应收应付登记没填「约定收/付款日期」时，按登记日期＋账期算到期日；往来单位自己填了账期的按单位的。']
    for i, t in enumerate(notes):
        put(ws, f'E{4 + i}', t, F_NOTE, align=AL, border=False)
        ws.merge_cells(f'E{4 + i}:R{4 + i}')
    # 各块
    blocks = [
        ('② 业务板块', SEG_COLS, [('名称', '板块名称'), ('期初结余', '期初结余\n（建账日）'), ('期初存量', '期初粮食存量\n（KG）'),
                                ('期初固定资产', '期初固定资产\n（原值）'), ('备注', '备注')], ctx.get('segments', []), N_SEG),
        ('③ 账户', ACC_COLS, [('名称', '账户名称'), ('期初余额', '期初余额\n（建账日）'), ('备注', '备注')], ctx.get('accounts', []), N_ACC),
        ('④ 收支项目', ITEM_COLS, [('名称', '收支项目'), ('方向', '方向'), ('用途', '用途'), ('库存', '库存'), ('备注', '说明')],
         ctx.get('items', []), N_ITEM),
        ('⑤ 经手人', PER_COLS, [('姓名', '经手人'), ('备注', '备注')], ctx.get('persons', []), N_PER),
    ]
    for ttl, cols, hd, data, cap in blocks:
        c1, c2 = cols[hd[0][0]], cols[hd[-1][0]]
        put(ws, f'{c1}{BASE_HDR - 1}', ttl, F_SEC, fill(C_MASTER), align=AL)
        ws.merge_cells(f'{c1}{BASE_HDR - 1}:{c2}{BASE_HDR - 1}')
        header(ws, BASE_HDR, [(cols[k], t) for k, t in hd], C_MASTER, height=40)
        fm = {cols[k]: MONEY for k in ('期初结余', '期初固定资产', '期初余额') if k in cols}
        if '期初存量' in cols:
            fm[cols['期初存量']] = '#,##0.##'
        style_rows(ws, BASE_R0, BASE_R0 + max(len(data) + 5, 10) - 1, [cols[k] for k, _ in hd], fmts=fm,
                   aligns={cols[k]: AL for k in ('名称', '备注', '姓名') if k in cols})
        for i, x in enumerate(data):
            for k, _t in hd:
                v = x.get(k)
                if v not in (None, ''):
                    ws[f'{cols[k]}{BASE_R0 + i}'] = v
        for k in ('期初结余', '期初存量', '期初固定资产', '期初余额'):
            if k in cols:
                dv_dec(ws, f'{cols[k]}{BASE_R0}:{cols[k]}{BASE_R0 + cap - 1}')
    rngI = lambda k: f'{ITEM_COLS[k]}{BASE_R0}:{ITEM_COLS[k]}{BASE_R0 + N_ITEM - 1}'
    dv_list(ws, rngI('方向'), '"' + ','.join(ITEM_DIRS) + '"')
    dv_list(ws, rngI('用途'), '"' + ','.join(ITEM_USES) + '"', prompt='空着＝普通')
    dv_list(ws, rngI('库存'), '"' + ','.join(ITEM_STOCK) + '"', prompt='只有粮食购进（入库）、粮食销售（出库）要选')
    ws.freeze_panes = f'A{BASE_R0}'
    print_setup(ws, None)


# ───────────────────────── 往来单位 ─────────────────────────
def build_unit(ws, ctx):
    C = UNIT_COLS
    heads = [(C['名称'], '单位 / 个人名称'), (C['类型'], '类型'), (C['板块'], '主要业务板块'), (C['联系人'], '联系人'),
             (C['电话'], '电话'), (C['地址'], '地址'), (C['账期'], '账期（天）\n（空＝默认）'), (C['备注'], '备注')]
    tip = ('💡 客户、供应商（单位或个人）一个一行，名称不要重复。收支登记、应收应付登记的「往来单位」从这里下拉；'
           '应收应付跟进、对账单按这个名称汇总。账期＝赊账多少天内要结清，空着按【基础资料】的默认账期。可以插行、删行。')
    W = {C['名称']: 18, C['类型']: 12, C['板块']: 13, C['联系人']: 10, C['电话']: 14, C['地址']: 26, C['账期']: 11, C['备注']: 24}
    rows = ctx.get('units', [])
    _frame(ws, 'H', UNIT_HDR, heads, C_MASTER, tip, W, UNIT_R0, max(len(rows) + 50, 100), {},
           {C['名称']: AL, C['地址']: AL, C['备注']: AL}, f'B{UNIT_R0}')
    _write_rows(ws, rows, C, UNIT_R0, date_keys=())
    rng = lambda col: f'{col}{UNIT_R0}:{col}{UNIT_R0 + N_UNIT - 1}'
    dv_list(ws, rng(C['类型']), '"' + ','.join(UNIT_TYPES) + '"')
    dv_list(ws, rng(C['板块']), '=板块列表', stop=False)
    dv = DataValidation(type='whole', operator='between', formula1='0', formula2='3650', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng(C['账期']))
    print_setup(ws, f'{UNIT_HDR}:{UNIT_HDR}')


BUILDERS = {SH_CASH: build_cash, SH_WL: build_wl, SH_BASE: build_base, SH_UNIT: build_unit}


def build(wb, ctx):
    for name, fn in BUILDERS.items():
        fn(wb[name], ctx)
