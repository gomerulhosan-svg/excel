# -*- coding: utf-8 -*-
"""把各张表拼成一本：录入（蓝）→ 档案（灰）→ 查看（绿）→ 项目（橙）→ 老板报表（深红）→ 数据校验；隐藏的取数表放最后。"""
import openpyxl
from layout import *
import demo
import s_base, s_master, s_open, s_cash, s_regs, s_att, s_agg

TAB = {'录入': 'FF2F75B5', '档案': 'FF7F7F7F', '查看': 'FF548235', '项目': 'FFC65911', '报表': 'FF833C0C', '首页': 'FF1F3864'}
GROUPS = [
    ('首页', [SH_HOME]),
    ('录入', [SH_CASH, SH_AP, SH_ATT, SH_INV, SH_REV, SH_OFF]),
    ('档案', [SH_PROJ, SH_UNIT, SH_RATE, SH_OPEN, SH_BASE]),
    ('查看', [SH_ACC, '资金报表', SH_AR, SH_APS, SH_PER, SH_PAY, SH_PAYS, SH_LAB, SH_INVS]),
    ('项目', [SH_PL, SH_PPL, SH_ALLOC]),
    ('报表', [SH_IS, SH_BS, SH_BE, SH_CHK]),
]
HIDDEN = [SH_MS, SH_PS, SH_BALX, SH_AUX]


def build(scenario=None):
    ctx = demo.build_ctx()
    if scenario:
        scenario(ctx)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    s_cash.build_cash(wb, ctx)
    s_regs.build_ap(wb, ctx)
    s_att.build_att(wb, ctx)
    s_regs.build_inv(wb, ctx)
    s_regs.build_rev(wb, ctx)
    s_regs.build_off(wb, ctx)
    s_master.build_proj(wb, ctx)
    s_master.build_unit(wb, ctx)
    s_master.build_rate(wb, ctx)
    s_open.build_open(wb, ctx)
    s_base.build_base(wb, ctx)
    s_agg.build_ms(wb, ctx)
    s_agg.build_ps(wb, ctx)
    s_agg.build_balx(wb, ctx)
    for mod, fn in (('s_views', 'build_all'), ('s_views2', 'build_all'), ('s_proj', 'build_all'), ('s_fin', 'build_all'), ('s_check', 'build_all')):
        try:
            m = __import__(mod)
        except ImportError:
            continue
        getattr(m, fn)(wb, ctx)
    s_base.build_aux(wb, ctx)
    order = []
    for grp, names in GROUPS:
        for n in names:
            if n in wb.sheetnames:
                order.append(n)
                wb[n].sheet_properties.tabColor = TAB[grp][2:]
    order += [n for n in HIDDEN if n in wb.sheetnames]
    rest = [n for n in wb.sheetnames if n not in order]
    assert not rest, rest
    wb._sheets = [wb[n] for n in order]
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = False
    wb.active = 0
    wb.worksheets[0].sheet_view.tabSelected = True
    wb.calculation.fullCalcOnLoad = False
    wb.calculation.calcId = 191029
    return wb
