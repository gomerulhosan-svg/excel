# -*- coding: utf-8 -*-
"""A065 财税代理多公司收支模板 · 生成脚本（两本联动版）

跑法：python3 build.py
输入：../参考/ 里的农行流水、销项发票、进项发票（只读，当演示数据装进去）
输出：../A065-1_流水发票导入.xlsx   录入：基础资料、往来单位、流水1～8、手工记账、资金台帐、发票导入、数据校验
      ../A065-2_汇总报表.xlsx       报表：收支汇总、现金流量、资金余额、费用、客户收入、应收应付、对账单、内部往来、发票汇总
                                    （隐藏的取数表用跨工作簿链接取 ① 的数）

做法：
  ① 先拼一本「合并算数本」（① 的全部表 ＋ ② 的报表，报表直接引用同名的表），LibreOffice 整本重算、数报错；
  ② 生成两本成品，一整列同样的公式改成共享公式（文件小、打开快），把合并本算出来的数写进去当缓存
     （手机、微信预览也能看到数；Excel/WPS 打开时照样重算）；
  ③ ② 里跨工作簿链接的缓存也按 ① 的数写好。成品不经 LibreOffice 存盘。
全部公式都是 Excel 2010 起就有的函数（MAXIFS/MINIFS 除外，Excel 2019 / WPS 都有）。
"""
import glob
import json
import os
import shutil
import subprocess
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from common import *
import data_prep
import s_base, s_src, s_cash, s_inv, s_reports, s_ar, s_check, s_mirror
from share import share
from inject_cache import inject
from refresh_link_cache import refresh

ROOT = os.path.dirname(HERE)
OUT1 = os.path.join(ROOT, WB1_FILE)
OUT2 = os.path.join(ROOT, WB2_FILE)
CALC = os.path.join(HERE, '_calc')


def make_ctx():
    acct, opening, rows = data_prep.bank_rows()
    sales = data_prep.invoice_rows(data_prep.F_SALE)
    buys = data_prep.invoice_rows(data_prep.F_BUY)
    samples = {j: data_prep.sample_raw(a[7]) for j, a in enumerate(data_prep.ACCOUNTS, 2) if a[7]}
    return {'my_co': data_prep.MY_CO, 'my_acc': data_prep.MY_ACC, 'bank': (acct, opening, rows),
            'companies': data_prep.COMPANIES, 'accounts': data_prep.ACCOUNTS, 'samples': samples,
            'bank_raw': data_prep.bank_raw(), 'sales': sales, 'buys': buys,
            'parties': data_prep.parties(rows, sales, buys)}


def reports(wb, ctx):
    s_reports.build_sum(wb, ctx)
    s_reports.build_cf(wb, ctx)
    s_reports.build_bal(wb, ctx)
    s_reports.build_exp(wb, ctx)
    s_reports.build_cus(wb, ctx)
    s_ar.build_ar(wb, ctx)
    s_ar.build_stmt(wb, ctx)
    s_ar.build_intra(wb, ctx)
    s_reports.build_invs(wb, ctx)


def finish(wb, order):
    assert set(order) == set(wb.sheetnames), set(order) ^ set(wb.sheetnames)
    wb._sheets = [wb[n] for n in order]
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = False
    wb.active = 0
    wb.worksheets[0].sheet_view.tabSelected = True
    # 打开时不强制整本重算：成品里已经写好了算出来的数（Excel/WPS 改了哪格只重算受影响的格子）
    wb.calculation.fullCalcOnLoad = False
    wb.calculation.calcId = 191029


def build_wb1(ctx, composite=False):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    s_base.build_base(wb, ctx)
    s_base.build_party(wb, ctx)
    s_src.build_all(wb, ctx)
    s_cash.build_cash(wb, ctx, with_stmt=composite)
    s_inv.build_inv(wb, ctx, with_stmt=composite)
    s_base.build_aux(wb, ctx, sources=True)
    s_check.build_check(wb, ctx)
    s_check.build_home1(wb, ctx)
    order = list(WB1_ORDER)
    if composite:
        reports(wb, ctx)
        s_check.build_home2(wb, ctx, name=SH_HOME2)
        order = order[:-1] + WB2_REPORTS + [SH_HOME2, SH_AUX]
    finish(wb, order)
    return wb


def build_wb2(ctx, wb1_names):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    s_mirror.build_mirrors(wb)
    s_base.build_aux(wb, ctx, sources=False, link='[1]')
    reports(wb, ctx)
    s_check.build_home2(wb, ctx)
    s_mirror.add_external_link(wb, wb1_names)
    finish(wb, [SH_HOME] + WB2_REPORTS + WB2_MIRRORS + [SH_AUX])
    return wb


def find_recalc():
    cands = [os.environ.get('RECALC_PY', ''), '/mnt/skills/public/xlsx/scripts/recalc.py']
    cands += glob.glob(os.path.expanduser('~/.claude/skills/**/xlsx/scripts/recalc.py'), recursive=True)
    for c in cands:
        if c and os.path.exists(c):
            return c


def lo_recalc(path, timeout=3000):
    rc = find_recalc()
    if not rc:
        print('✗ 没找到 recalc.py（LibreOffice 重算脚本），成品里没有缓存值——Excel/WPS 打开时会自己算')
        return None
    r = subprocess.run([sys.executable, rc, path, str(timeout), '--force'], capture_output=True, text=True,
                       timeout=timeout + 300)
    out = r.stdout
    i, j = out.find('{'), out.rfind('}')
    return json.loads(out[i:j + 1]) if i >= 0 else {'error': out or r.stderr}


def main():
    ctx = make_ctx()
    os.makedirs(CALC, exist_ok=True)
    comp = os.path.join(CALC, '合并算数本.xlsx')
    print('① 合并算数本')
    wb = build_wb1(ctx, composite=True)
    wb.save(comp)
    share(comp)
    st = lo_recalc(comp)
    if st is None:
        return 0
    if 'error' in st:
        print('✗ 重算没跑成：', str(st['error'])[:800])
        return 2
    print(f"   LibreOffice 重算：公式 {st['total_formulas']} 个，报错 {st['total_errors']} 处")
    for k, v in (st.get('error_summary') or {}).items():
        print(f'    {k} ×{v["count"]}  {v["locations"][:10]}')
    print('② 工作簿 1')
    wb1 = build_wb1(make_ctx())
    names1 = wb1.sheetnames
    wb1.save(OUT1)
    s1 = share(OUT1)
    n, _ = inject(OUT1, comp)
    print(f'   {WB1_FILE}：共享公式 {sum(s1.values())} 格，写回缓存 {n} 格')
    print('③ 工作簿 2')
    wb2 = build_wb2(make_ctx(), names1)
    wb2.save(OUT2)
    s2 = share(OUT2)
    n, _ = inject(OUT2, comp, names={SH_HOME: SH_HOME2})
    refresh(OUT2, {WB1_FILE: OUT1})
    print(f'   {WB2_FILE}：共享公式 {sum(s2.values())} 格，写回缓存 {n} 格')
    for p in (OUT1, OUT2):
        print(f'   {os.path.basename(p)} {os.path.getsize(p) / 1e6:.1f} MB')
    return 0 if st['total_errors'] == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
