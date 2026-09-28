# -*- coding: utf-8 -*-
"""A065 财税代理多公司收支模板 · 生成脚本

跑法：python3 build.py [输出路径]
输入：../参考/ 里你发来的农行流水、销项发票、进项发票（只读，当演示数据装进去）
输出：../A065_财税代理多公司收支模板.xlsx

全部公式都是 Excel 2010 起就有的函数（MAXIFS/MINIFS 除外，Excel 2019 / WPS 都有），不用 FILTER/UNIQUE/XLOOKUP 这类动态数组。
"""
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from common import *
import data_prep
import s_base, s_cash, s_conv, s_inv, s_reports, s_ar, s_check

ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, 'A065_财税代理多公司收支模板.xlsx')


def main(out_path=OUT):
    acct, opening, rows = data_prep.bank_rows()
    sales = data_prep.invoice_rows(data_prep.F_SALE)
    buys = data_prep.invoice_rows(data_prep.F_BUY)
    ctx = {'my_co': data_prep.MY_CO, 'my_acc': data_prep.MY_ACC, 'bank': (acct, opening, rows),
           'sales': sales, 'buys': buys, 'parties': data_prep.parties(rows, sales, buys)}
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    s_base.build_base(wb, ctx)
    s_base.build_party(wb, ctx)
    s_cash.build_cash(wb, ctx)
    s_conv.build_conv(wb, ctx)
    s_inv.build_inv(wb, ctx)
    s_reports.build_sum(wb, ctx)
    s_reports.build_cf(wb, ctx)
    s_reports.build_bal(wb, ctx)
    s_reports.build_exp(wb, ctx)
    s_reports.build_cus(wb, ctx)
    s_ar.build_ar(wb, ctx)
    s_ar.build_stmt(wb, ctx)
    s_ar.build_intra(wb, ctx)
    s_reports.build_invs(wb, ctx)
    s_check.build_check(wb, ctx)
    s_check.build_home(wb, ctx)
    s_base.build_aux(wb, ctx)

    order = [wb[n] for n in SHEET_ORDER]
    assert len(order) == len(wb.worksheets), set(ws.title for ws in wb.worksheets) - set(SHEET_ORDER)
    wb._sheets = order
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = False
    wb.active = 0
    wb[SH_HOME].sheet_view.tabSelected = True
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    print(f'✓ 生成 {out_path}')
    print(f'  流水 {len(rows)} 笔（期初 {opening}），销项 {len(sales)} 张，进项 {len(buys)} 张，往来单位 {len(ctx["parties"])} 个')


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else OUT)
