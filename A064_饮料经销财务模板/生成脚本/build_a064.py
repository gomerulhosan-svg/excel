# -*- coding: utf-8 -*-
"""A064 饮料经销财务模板 · 生成脚本

跑法：python3 build_a064.py [输出路径]
输入：../参考/原表1_存条出库库存总览.xlsx、../参考/原表2_过期报损统计.xlsx
输出：../A064_饮料经销财务模板.xlsx

做法：以原表 1 为底（出库/存条/总览/商品库存/收付款明细/客户查询 的格式原样保留），
把原表 2 的四张报损表连格式搬进来，再加上新表。全部公式都是 Excel 2010 起就有的函数
（外加 MAXIFS，WPS 和 Excel 2019+ 都认），不用 FILTER/UNIQUE/XLOOKUP 这类动态数组。
"""
import json
import os
import sys
import datetime as dt

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from common import *
import fixes
import data_prep
import s_base, s_cash, s_orig, s_baosun, s_reports, s_stmt, s_check, s_fee

ROOT = os.path.dirname(HERE)
SRC1 = os.path.join(ROOT, '参考', '原表1_存条出库库存总览.xlsx')
SRC2 = os.path.join(ROOT, '参考', '原表2_过期报损统计.xlsx')
SRC3 = os.path.join(ROOT, '参考', '原表3_费用报销明细（补充）.xlsx')
OUT = os.path.join(ROOT, 'A064_饮料经销财务模板.xlsx')


def prepare(s1, s2):
    ctx = {'year': fixes.BOOK_YEAR, 'kpi_month': 9,
           'cus_merge': dict(fixes.CUSTOMER_MERGE), 'goods_merge': dict(fixes.GOODS_MERGE)}
    ctx['customers'] = data_prep.build_customer_master(s1)
    # 商品档案：fixes 里整理好的 + 明细里出现过但还没列进去的
    goods = [dict(g) for g in fixes.GOODS]
    names = {g['name'] for g in goods}
    extra = []
    for rec in s1.outbound() + s1.deposits():
        n = rec['goods']
        if n and n not in names:
            names.add(n)
            extra.append(n)
            goods.append(dict(name=n))
    ctx['goods'] = goods
    ctx['goods_extra'] = extra
    ctx['suppliers'] = fixes.SUPPLIERS
    ctx['accounts'] = fixes.ACCOUNTS
    ctx['categories'] = fixes.CATEGORIES
    ctx['expense_items'] = fixes.EXPENSE_ITEMS
    ctx['baosun_months'] = fixes.BAOSUN_MONTHS
    # 资金台帐：原收付款明细的 117 笔客户收款，按日期排好搬进来
    recs = s1.receipts()
    rows = []
    for k, x in enumerate(recs):
        d = x['date']
        rows.append(dict(_k=(d or dt.date(2100, 1, 1), k), _src=x['row'], _raw=x['raw_date'],
                         date=dt.datetime(d.year, d.month, d.day) if d else x['raw_date'],
                         account=x['method'], category='客户回款',
                         memo=x['note'] or '收货款', **{'in': x['amount']}, customer=x['customer']))
    rows.sort(key=lambda t: t['_k'])
    ctx['cash_rows'] = rows
    ctx['out_last'] = max(r['row'] for r in s1.outbound())
    ctx['cun_last'] = max(r['row'] for r in s1.deposits())
    ctx['default_客户列表'] = ctx['customers'][0][0]
    ctx['default_供应商列表'] = fixes.SUPPLIERS[0]['name']
    return ctx


# 按位置（INDEX(区域,k)）取数的辅助公式：用户在明细表/基础资料里删行后区域变短，
# 最后几个位置会越界成 #REF!——一个错值会把 SMALL() 整列带崩，所以统一包一层 IFERROR（给对类型的默认值）
SAFE_COLS = {
    SH_AUX: {'A': '""', 'D': '""', 'E': '""', 'F': '""', 'H': '""', 'J': '""', 'K': '""',
             'N': '0', 'O': '0', 'P': '0', 'Q': '0', 'R': '0', 'S': '""', 'AA': '""', 'AB': '1'},
    SH_COST: {'B': '""', 'C': '""', 'D': '""', 'E': '""', 'F': '0', 'G': '0'},
    SH_STK: {'C': '""'},
    SH_EXP: {'B': '""', 'C': '""'},
    SH_RPS: {'B': '""', 'C': '""', 'K': '""', 'L': '""'},
    SH_BSS: {'B': '""'},
}


def wrap_iferror(wb):
    n = 0
    for sh, cols in SAFE_COLS.items():
        ws = wb[sh]
        for col, dflt in cols.items():
            for (c,) in ws.iter_rows(min_col=CI(col), max_col=CI(col)):
                v = c.value
                if isinstance(v, str) and v.startswith('=') and 'INDEX(' in v and not v.startswith('=IFERROR(IF('):
                    c.value = f'=IFERROR({v[1:]},{dflt})'
                    n += 1
    return n


def main(out_path=OUT):
    s1 = data_prep.Src1(SRC1)
    s2 = data_prep.Src2(SRC2)
    ctx = prepare(s1, s2)

    wb = openpyxl.load_workbook(SRC1)
    del wb['Sheet1']                         # 原表只有一格草稿 =4200+1000+…，没被任何地方引用
    wb.create_sheet(SH_AUX)
    changes = []

    # 原表存盘时带着筛选（总览汇总只显示「李记副食」、存条明细只显示「梅窖建军批发」…），几乎所有行都被藏起来了——全部放开
    for ws in wb.worksheets:
        ws.auto_filter.filterColumn = []
        ws.auto_filter.sortState = None
        for rd in ws.row_dimensions.values():
            rd.hidden = False

    s_base.build(wb, ctx)
    s_cash.build_cash(wb, ctx)
    s_orig.build_buy(wb, ctx)
    src3 = openpyxl.load_workbook(SRC3)
    s_fee.build_fee(wb, ctx, src3, changes)
    s_fee.build_qr(wb, ctx, src3)
    s_orig.fix_out(wb, ctx, changes)
    s_orig.fix_cun(wb, ctx, changes)
    s_orig.fix_ov(wb, ctx)
    s_cash.build_rp(wb, ctx)
    s_orig.build_aux_combo(wb, ctx)
    s_orig.fix_inv(wb, ctx)
    s_orig.fix_q(wb, ctx)
    s_baosun.build(wb, ctx, s2)
    s_reports.build_cost(wb, ctx)
    s_reports.build_stock(wb, ctx)
    s_reports.build_exp(wb, ctx)
    s_reports.build_aux_fee(wb, ctx)
    s_reports.build_pl(wb, ctx)
    s_reports.build_bal(wb, ctx)
    s_reports.build_aux_balances(wb, ctx)
    s_stmt.build_aux_stmt(wb, ctx)
    s_stmt.build_rps(wb, ctx)
    s_stmt.build_cst(wb, ctx)
    s_stmt.build_sst(wb, ctx)
    s_check.build_aux_check(wb, ctx)
    s_check.build_check(wb, ctx)
    s_check.build_home(wb, ctx)
    n_safe = wrap_iferror(wb)

    # 表的顺序、隐藏辅助表、打开时落在首页
    order = [wb[n] for n in SHEET_ORDER]
    assert len(order) == len(wb.worksheets), set(ws.title for ws in wb.worksheets) - set(SHEET_ORDER)
    wb._sheets = order
    wb[SH_AUX].sheet_state = 'hidden'
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = False
    wb.active = 0
    wb[SH_HOME].sheet_view.tabSelected = True
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)

    log = {'changes': changes, 'goods_extra': ctx['goods_extra'],
           'customers': [(n, d, src) for n, d, src in ctx['customers']],
           'cash_rows': [(r['_src'], str(r['_raw']), str(r['date']), r['customer'], r['in']) for r in ctx['cash_rows']]}
    with open(os.path.join(HERE, '_build_log.json'), 'w', encoding='utf-8') as f:
        json.dump(log, f, ensure_ascii=False, indent=1, default=str)
    print(f'✓ 生成 {out_path}')
    print(f'  客户 {len(ctx["customers"])} 个，商品 {len(ctx["goods"])} 个（明细里补进来的：{ctx["goods_extra"]}），'
          f'资金台帐搬入 {len(ctx["cash_rows"])} 笔，名称规范化 {len(changes)} 处')


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else OUT)
