# -*- coding: utf-8 -*-
"""A075 矿业美元财务账套 —— 生成脚本入口（重新生成用户 10/14 上传的那本 A075）。

    python3 -I build.py [输出.xlsx]          只生成（openpyxl 写盘，不压共享公式、不重算）
    python3 -I make_all.py [输出.xlsx]       生成 → 压共享公式 → LibreOffice 在副本上重算查错 → 把算好的值写回成品

怎么来的（详见 README.md）：
  A062 生成脚本（9/26，生成的就是仓库里 A062 第一版，逐格一样）
  ＋ 用户 10/05 在 A062 上的改动里属于「结构」的：流水删到 3,960 行（记账分录 / 凭证索引的流水槽位仍是 4,000）
  ＋ p1008.py（A062 10/08 修改脚本 m1008.py 搬到生成脚本的列位上：按天合并凭证、结转损益、A4 打印…）
  ＋ A075 的改动（common.py「A075 参数」：本位币美元、汇率「1 美元 = X 外币」、月度汇率 K:R、下拉名称不用 OFFSET…）
  ＋ 用户的东西从原件读（ref_a075.py：参数、账户、科目表、记账规则、物料、汇率；首页 / 使用说明整张照搬）
"""
import datetime as dt
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from openpyxl import Workbook
from openpyxl.workbook.properties import CalcProperties

from common import *
import ref_a075
import s_base, s_cash, s_mods, s_mend, s_je, s_ledger, s_reports, s_check
import p1008
import p_a075
import p_opt

OUT = os.path.join(HERE, '..', 'A075_矿业美元财务账套（提速版）.xlsx')


def prepare(ref):
    d = ref_a075.read(ref)
    p = d['params']
    ctx = dict(ledger=d['ledger'], params=p, fx=d['fx'], chart=d['chart'], overrides={}, counterparties=d['counterparties'],
               projects=d['projects'], rules=d['rules'], months=d['months'], accounts=d['accounts'], rates=d['rates'],
               mat_cats=d['fx'].MAT_CATS, home_year=None, home_month=None, n_fee=0, ref=ref,
               gl_default=ref[SH_GL]['B3'].value)
    return ctx


def build(out=OUT, ref_path=ref_a075.REF, log=print, tweak=None):
    """tweak(wb, ctx)：存盘前再改一下（做场景测试用，比如填几笔流水）"""
    t0 = time.time()
    ref = ref_a075.load(ref_path)
    log(f'  读原件 {time.time() - t0:.0f} 秒')
    ctx = prepare(ref)
    wb = Workbook()
    wb.remove(wb.active)
    s_base.build_base(wb, ctx)
    s_base.build_rules(wb, ctx)
    s_base.build_rates(wb, ctx)
    s_base.build_accounts(wb, ctx)
    s_cash.build_cash(wb, ctx)
    s_mods.build_buy(wb, ctx)
    s_mods.build_issue(wb, ctx)
    s_mods.build_prod(wb, ctx)
    s_mods.build_sale(wb, ctx)
    s_mods.build_pay(wb, ctx)
    s_mods.build_fa(wb, ctx)
    s_mods.build_man(wb, ctx)
    wb.create_sheet(SH_AUX)
    s_mend.build_inv_grid(wb, ctx)
    s_mend.build_ore(wb, ctx)
    s_mend.build_mend(wb, ctx)
    s_je.build_je(wb, ctx)
    s_je.build_vindex(wb, ctx)
    s_je.build_vlist(wb, ctx)
    s_je.build_vprint(wb, ctx)
    s_ledger.build_tb(wb, ctx)
    s_ledger.build_gl(wb, ctx)
    s_ledger.build_acct(wb, ctx)
    s_ledger.build_arap(wb, ctx)
    s_ledger.build_inv(wb, ctx)
    s_reports.build_bs(wb, ctx)
    s_reports.build_pl(wb, ctx)
    s_reports.build_cf(wb, ctx)
    s_reports.build_ops(wb, ctx)
    s_check.build_check(wb, ctx)
    wb.create_sheet(SH_HOME)            # 首页、使用说明：p_a075 从原件整张照搬
    wb.create_sheet(SH_HELP)
    s_base.define_names(wb)
    wb._sheets = [wb[n] for n in SHEET_ORDER]
    log(f'  A062 生成脚本部分 {time.time() - t0:.0f} 秒')
    for line in p1008.apply(wb, ctx):
        log('   · 10/08 ' + line)
    for line in p_a075.apply(wb, ctx):
        log('   · A075 ' + line)
    for line in p_opt.apply(wb, ctx):
        log('   · 提速 ' + line)
    wb[SH_AUX].sheet_state = 'hidden'
    wb[SH_VIDX].sheet_state = 'hidden'
    wb[SH_CLOSE].sheet_state = 'hidden'
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == SH_HOME
    if tweak:
        tweak(wb, ctx)
    wb.calculation = CalcProperties(calcId=191029, fullCalcOnLoad=False)     # 不写 fullCalcOnLoad（每次打开整本重算太慢；缓存值由 make_all 写回）
    wb.save(out)
    log(f'✓ 生成 {os.path.basename(out)}（{time.time() - t0:.0f} 秒）')
    return ctx


if __name__ == '__main__':
    build(sys.argv[1] if len(sys.argv) > 1 else OUT)
