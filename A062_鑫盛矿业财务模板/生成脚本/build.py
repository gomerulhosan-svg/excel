# -*- coding: utf-8 -*-
"""A062 鑫盛矿业财务模板 —— 生成脚本入口。

跑法：python3 build.py            （只生成成品，不重算）
      python3 make_all.py         （生成 → LibreOffice 在副本上重算查错 → 把算好的值写回成品）
数据来源：../参考/原表2（现金流水、科目表、流水用的参数和月度汇率缓存）；修正清单在 fixes.py。"""
import collections
import datetime as dt
import glob
import os
import sys

from openpyxl import Workbook
from openpyxl.workbook.properties import CalcProperties

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from common import *
import fixes
from data_prep import Src2
import s_base, s_cash, s_mods, s_mend, s_je, s_ledger, s_reports, s_check

OUT = os.path.join(HERE, '..', 'A062_鑫盛矿业财务模板.xlsx')
REF = os.path.join(HERE, '..', '参考')


def _key(code):
    return int((code + '0000000000')[:10])


def prepare():
    f2 = glob.glob(os.path.join(REF, '原表2*.xlsx'))[0]
    s2 = Src2(f2)
    ledger = s2.ledger()
    params = s2.params()
    L = {x['row']: x for x in ledger}

    # 科目表：原表 287 个 + 补的明细
    chart = s2.accounts()
    have = {a['code'] for a in chart}
    for a in chart:
        if a['code'] in fixes.ACCOUNT_FIX:
            a.update(fixes.ACCOUNT_FIX[a['code']])
    for code, name, cls, d, aux, fx, note in fixes.ACCOUNT_ADD:
        assert code not in have, code
        chart.append(dict(code=code, name=name, cls=cls, dir=d, aux=aux, fx=fx, cashf='是' if code[:4] in ('1001', '1002') else '否',
                          stat='启用', para='否', added=True, note=note))
    chart.sort(key=lambda a: _key(a['code']))
    codes = {a['code'] for a in chart}
    for c in codes:                     # 每个明细的上级都要在
        for n in (4, 6, 8):
            if len(c) > n:
                assert c[:n] in codes, (c, c[:n])

    # 逐行修正（写进「指定对方科目」）
    ovr = {r: (code, fixes.OVERRIDE_WHY[r]) for r, code in fixes.ROW_OVERRIDE.items()}
    for b, c in fixes.NBU_PAIRS:
        cat = L[c]['category']
        code = '122105' if cat == '购汇/结汇' else '122104'
        ovr[b] = (code, f'银行「1900/4800 注册资本」这笔是把现金存进公户：同一天现金账户第 {c} 行（{L[c]["account"]}·{cat}）'
                        f'付出了同样的钱。股东的钱在人民币现金「投资款」里已经记过实收资本，这里改记'
                        f'{"外币兑换待对冲" if code == "122105" else "资金划转待对冲"}，不再重复记实收资本。原类别「{L[b]["category"]}」没动。')
    nfee = 0
    for r, x in L.items():
        if x['category'] == '内部转账' and fixes.FEE_IN_TRANSFER_WORD in (x['memo'] or '') and r not in ovr:
            ovr[r] = ('560302', '记在「内部转账」里的银行手续费：摘要写着手续费，也没有对面账户的一笔，改记财务费用-手续费')
            nfee += 1

    # 往来单位（流水里出现过的，按出现次数排）
    cnt = collections.Counter(x['cp'] for x in ledger if x['cp'])
    cps = []
    for name, _ in cnt.most_common():
        if name in fixes.SHAREHOLDERS:
            cps.append(dict(name=name, type='股东', cap=fixes.SHAREHOLDERS[name]))
        elif name in fixes.SUPPLIERS:
            cps.append(dict(name=name, type='供应商'))
        elif name in fixes.CUSTOMERS:
            cps.append(dict(name=name, type='客户', note='可能是甲方，待确认'))
        elif name in fixes.OTHER_CP:
            cps.append(dict(name=name, type=fixes.OTHER_CP[name]))
        else:
            cps.append(dict(name=name, type='其他'))
    for name, code in fixes.SHAREHOLDERS.items():     # 没在流水里出现的股东也列上
        if name not in cnt and not any(c['name'] == name for c in cps) and name.isascii() is False:
            cps.append(dict(name=name, type='股东', cap=code))
    # 项目
    pcnt = collections.Counter(x['project'] for x in ledger if x['project'])
    projects = [p for p, _ in pcnt.most_common()]
    for extra in ('矿山生产', '基建工程', '设备物资', '火工品'):
        if extra not in projects:
            projects.append(extra)
    # 收支类别规则：按原表 38 个类别的顺序，后面接新加的
    rmap = {r[0]: r for r in fixes.RULES}
    rules = []
    for name, attr in params['categories']:
        assert name in rmap, f'类别「{name}」没有记账规则'
        r = rmap[name]
        rules.append((name, attr) + tuple(r[2:]))
    for name, r in rmap.items():
        if name not in dict(params['categories']):
            rules.append(r)
    for x in ledger:
        assert x['category'] is None or x['category'] in rmap, x
    start = params['start']
    months = [dt.datetime(start.year + (start.month - 1 + i) // 12, (start.month - 1 + i) % 12 + 1, 1) for i in range(N_MONTHS)]
    # 报表默认：流水最后一个整月
    last = max(x['date'] for x in ledger if isinstance(x['date'], dt.datetime))
    ctx = dict(ledger=ledger, params=params, chart=chart, overrides=ovr, counterparties=cps, projects=projects,
               rules=rules, months=months, accounts=params['accounts'], mat_cats=fixes.MAT_CATS,
               home_year=last.year, home_month=last.month, gl_default='100207 NBU银行苏姆结算户-6688',
               n_fee=nfee)
    return ctx


def build(out=OUT):
    ctx = prepare()
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
    s_check.build_home(wb, ctx)
    s_check.build_help(wb, ctx)
    s_base.define_names(wb)
    wb._sheets = [wb[n] for n in SHEET_ORDER]
    wb[SH_AUX].sheet_state = 'hidden'
    wb[SH_VIDX].sheet_state = 'hidden'
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == SH_HOME
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(out)
    return ctx


def main(out=OUT):
    ctx = build(out)
    print(f'✓ 生成 {os.path.basename(out)}：流水 {len(ctx["ledger"])} 行，逐行修正 {len(ctx["overrides"])} 行（其中手续费 {ctx["n_fee"]}），'
          f'科目 {len(ctx["chart"])} 个，记账分录 {ctx["je_r1"] - JE_R0 + 1} 行，凭证槽位 {ctx["vx_r1"] - VX_R0 + 1}')
    return ctx


if __name__ == '__main__':
    main()
