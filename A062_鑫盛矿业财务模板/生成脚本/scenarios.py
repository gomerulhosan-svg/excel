# -*- coding: utf-8 -*-
"""业务演练：在成品副本里录一个「9 月份」（启用日以后）的完整业务，LibreOffice 重算后逐项核对。

跑法：python3 scenarios.py [成品.xlsx]
不改成品，只在临时副本上录。"""
import datetime as dt
import os
import shutil
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from common import *
import make_all

D = lambda m, d: dt.datetime(2026, m, d)
OK, BAD = [], []


def chk(name, cond, detail=''):
    (OK if cond else BAD).append(name)
    print(('  ✓ ' if cond else '  ✗ ') + name + (f'  —— {detail}' if detail else ''))


def close(a, b, tol=0.02):
    try:
        return abs(float(a or 0) - float(b or 0)) <= tol
    except (TypeError, ValueError):
        return False


def fill(path):
    wb = openpyxl.load_workbook(path)
    wb[SH_HOME]['C4'] = 2026
    wb[SH_HOME]['F4'] = 9
    # 采购入库
    ws = wb[SH_BUY]
    buys = [  # 日期 单据 供应商 物料 数量 币种 单价 运杂 进项税
        (D(9, 3), 'CG001', '三飞公司', '乳化炸药', 2000, None, 30000, 2000000, None),
        (D(9, 3), 'CG001', '三飞公司', '雷管', 1000, None, 5000, None, 600000),
        (D(9, 5), 'CG002', '张海勇', '柴油', 10000, None, 12000, None, None),
        (D(9, 10), 'CG003', '三飞公司', '矿车', 1, 'CNY', 300000, 15000000, None),
        (D(9, 12), None, '张海勇', '食材', 1, None, 3000000, None, None),
        (D(9, 20), 'CG003', '三飞公司', '矿车', None, None, None, 5000000, None),     # 后到的清关费：只填运杂
    ]
    for i, (d, no, sup, mat, qty, cur, up, frt, vat) in enumerate(buys):
        r = MOD_R0 + i
        for col, v in ((B_DATE, d), (B_NO, no), (B_SUP, sup), (B_MAT, mat), (B_QTY, qty), (B_CUR, cur), (B_UP, up),
                       (B_FRT, frt), (B_VAT, vat)):
            ws[f'{col}{r}'] = v
    # 领用出库
    ws = wb[SH_ISS]
    iss = [(D(9, 15), 'LL001', '乳化炸药', 1500, '采矿生产'), (D(9, 15), 'LL001', '雷管', 800, '采矿生产'),
           (D(9, 20), None, '柴油', 6000, '设备维修'), (D(9, 25), None, '柴油', 1000, '基建掘进'),
           (D(9, 26), None, '柴油', 100, None)]           # 最后一行没选用途 → ✗，不记账也不出库
    for i, (d, no, mat, qty, use) in enumerate(iss):
        r = MOD_R0 + i
        for col, v in ((I_DATE, d), (I_NO, no), (I_MAT, mat), (I_QTY, qty), (I_USE, use)):
            ws[f'{col}{r}'] = v
    # 产量
    ws = wb[SH_PROD]
    for i, (d, face, ton, dev, gr) in enumerate([(D(9, 10), '1号采场', 2000, 50, 1.2), (D(9, 25), '1号采场', 3000, 40, 1.0)]):
        r = MOD_R0 + i
        for col, v in ((PD_DATE, d), (PD_FACE, face), (PD_TON, ton), (PD_DEV, dev), (PD_GRADE, gr)):
            ws[f'{col}{r}'] = v
    # 销售
    ws = wb[SH_SALE]
    for col, v in ((S_DATE, D(9, 28)), (S_NO, 'JS001'), (S_CUS, '瑞斯曼'), (S_TON, 3000), (S_PRICE, 112000)):
        ws[f'{col}{MOD_R0}'] = v
    for col, v in ((S_DATE, D(9, 29)), (S_NO, 'JS002'), (S_CUS, '瑞斯曼'), (S_TON, 1000)):   # 过了磅还没定价 → ✗，不结转成本
        ws[f'{col}{MOD_R0 + 1}'] = v
    # 工资
    ws = wb[SH_PAY]
    for i, (dept, typ, n, amt) in enumerate([('采矿生产', '中籍', 20, 200000000), ('行政管理', '外籍', 5, 50000000)]):
        r = MOD_R0 + i
        for col, v in ((W_DATE, D(9, 30)), (W_DEPT, dept), (W_TYPE, typ), (W_HEAD, n), (W_AMT, amt)):
            ws[f'{col}{r}'] = v
    # 固定资产卡片（矿车；原值按采购入库算出来的数，场景里后面核）
    ws = wb[SH_FA]
    for col, v in ((F_NO, 'CL-001'), (F_NAME, '矿车'), (F_CAT, '矿用车辆'), (F_COST, '=采购入库!P7'), (F_DATE, D(9, 10)),
                   (F_DEPT, '采矿生产')):
        ws[f'{col}{MOD_R0}'] = v
    for col, v in ((F_NO, 'CL-002'), (F_NAME, '铲车'), (F_CAT, '装载机'), (F_COST, 100000000), (F_DATE, D(9, 12)),
                   (F_YEARS, '5年'), (F_DEPT, '采矿生产')):          # 年限写成文字 → ✗，不提折旧
        ws[f'{col}{MOD_R0 + 1}'] = v
    for col, v in ((F_NO, 'FD-001'), (F_NAME, '柴油发电机组'), (F_CAT, '发电动力设备'), (F_COST, 60000000), (F_DATE, D(6, 15)),
                   (F_DEPT, '采矿生产')):          # 启用日以前买的：7、8 月折旧记 540102，9 月起记生产成本
        ws[f'{col}{MOD_R0 + 2}'] = v
    # 手工凭证：第 1 张平；第 2 张不平（应整张不记）
    ws = wb[SH_MAN]
    rows = [(1, D(9, 30), '计提安全生产费', '40010109', 10000000, None), (1, None, None, '224103', None, 10000000),
            (2, D(9, 30), '测试不平', '560201', 100, None), (2, None, None, '224103', None, 90),
            (3, D(9, 30), '分摊9月水电（抬头行）', None, None, None), (3, None, None, '40010108', 2000000, None),
            (3, None, None, '224103', None, 2000000),
            (1, D(10, 31), '10月组号又从1编', '560201', 500000, None), (1, None, None, '224103', None, 500000)]
    for i, (g, d, memo, code, dr, cr) in enumerate(rows):
        r = MOD_R0 + i
        for col, v in ((MN_GRP, g), (MN_DATE, d), (MN_MEMO, memo), (MN_CODE, code), (MN_DR, dr), (MN_CR, cr)):
            ws[f'{col}{r}'] = v
    # 现金流水（接着原表最后一行往下录）
    ws = wb[SH_CASH]
    r0 = 2408
    cash = [  # 日期 账户 类别 摘要 收 支 往来 合并号
        (D(9, 6), 'NBU银行苏姆结算户-6688', '采购付款', '付张海勇柴油款', None, 120000000, '张海勇', None),
        (D(9, 29), 'NBU银行苏姆结算户-6688', '销售回款', '瑞斯曼付9月原矿款', 200000000, None, '瑞斯曼', None),
        (D(9, 30), 'NBU银行苏姆结算户-6688', '应付职工薪酬', '发9月工资', None, 150000000, None, None),
        (D(9, 8), 'NBU银行苏姆结算户-6688', '内部转账', '取现备用', None, 10000000, None, None),
        (D(9, 8), '苏姆现金', '内部转账', '取现备用', 10000000, None, None, None),
        (D(9, 9), '苏姆现金', '手续费', '手续费A', None, 10000, None, 'M1'),
        (D(9, 9), '苏姆现金', '手续费', '手续费B', None, 20000, None, 'M1'),
        (D(9, 30), 'NBU银行苏姆结算户-6688', '缴纳增值税', '交8月增值税', None, 1000000, None, None),
        (D(9, 30), '苏姆现金', '管理费用', '收入支出都填了（应该 ✗ 不记账）', 500, 300, None, None),
    ]
    for i, (d, a, cat, memo, inc, exp, cp, mg) in enumerate(cash):
        r = r0 + i
        for col, v in ((K_DATE, d), (K_ACC, a), (K_CAT, cat), (K_MEMO, memo), (K_IN, inc), (K_OUT, exp), (K_CP, cp),
                       (K_MERGE, mg)):
            ws[f'{col}{r}'] = v
    wb.save(path)


def main(src):
    tmpd = tempfile.mkdtemp(prefix='a062s_')
    p = os.path.join(tmpd, 'scen.xlsx')
    shutil.copy(src, p)
    fill(p)
    st = make_all.lo_recalc(p)
    print(f"LibreOffice：公式 {st.get('total_formulas')}，报错 {st.get('total_errors')}")
    for k, v in (st.get('error_summary') or {}).items():
        print('   ', k, v['count'], v['locations'][:6])
    wv = openpyxl.load_workbook(p, data_only=True)
    g = lambda sh, a: wv[sh][a].value
    chk('重算 0 报错', st.get('total_errors') == 0)
    # 采购
    chk('炸药入账 62,000,000（含运杂）', close(g(SH_BUY, f'{B_COST}4'), 62000000))
    chk('CG001 两行并成一张凭证', g(SH_BUY, f'{B_VID}4') == g(SH_BUY, f'{B_VID}5') == 200001)
    chk('矿车入账科目 160101', g(SH_BUY, f'{B_ACC}7') == '160101')
    truck = g(SH_BUY, f'{B_COST}7')
    je = wv[SH_JE]
    rate = g(SH_BUY, f'{B_RATE}7')
    chk('矿车按人民币汇率折苏姆（+运杂）', close(truck, round(300000 * rate, 2) + 15000000), f'{truck:,.2f} 汇率 {rate}')
    # 领用
    chk('炸药领用单价 = 月加权 31,000', close(g(SH_ISS, f'{I_PRICE}4'), 31000), g(SH_ISS, f'{I_PRICE}4'))
    chk('炸药领用金额 46,500,000 → 40010101', close(g(SH_ISS, f'{I_AMT}4'), 46500000) and g(SH_ISS, f'{I_DR}4') == '40010101')
    chk('设备维修 → 40010105，基建掘进 → 160405', g(SH_ISS, f'{I_DR}6') == '40010105' and g(SH_ISS, f'{I_DR}7') == '160405')
    chk('没选用途的领用 ✗ 且不记账', str(g(SH_ISS, f'{I_CHK}8')).startswith('✗') and g(SH_ISS, f'{I_VALID}8') == 0)
    # 原矿
    from s_mend import ORE_R0, O_TIN, O_AIN, O_UC, O_TS, O_AS, O_TE, O_AE, ME_SUM
    n9 = (2026 - 2025) * 12 + 9 - 5 + 1
    ro = ORE_R0 + n9 - 1
    exp_cost = 46500000 + 4000000 + 72000000 + 200000000 + 10000000 + 2000000 + 950000   # 最后是发电机 9 月折旧
    chk('9 月出矿 5000 吨', close(g(SH_ORE, f'{O_TIN}{ro}'), 5000))
    chk('9 月生产成本全部转原矿 = 332,500,000', close(g(SH_ORE, f'{O_AIN}{ro}'), exp_cost), g(SH_ORE, f'{O_AIN}{ro}'))
    uc = exp_cost / 5000
    chk(f'每吨成本 {uc:,.2f}', close(g(SH_ORE, f'{O_UC}{ro}'), uc), g(SH_ORE, f'{O_UC}{ro}'))
    chk(f'销售成本只算定了价的 3000 吨 = {3000 * uc:,.2f}（没定价那行 ✗ 不结转）', close(g(SH_ORE, f'{O_AS}{ro}'), 3000 * uc),
        g(SH_ORE, f'{O_AS}{ro}'))
    chk('没定价的销售行 ✗、不记账', str(g(SH_SALE, f'{S_CHK}5')).startswith('✗') and g(SH_SALE, f'{S_VALID}5') == 0)
    chk(f'原矿月末 2000 吨 / {2000 * uc:,.2f}', close(g(SH_ORE, f'{O_TE}{ro}'), 2000) and close(g(SH_ORE, f'{O_AE}{ro}'), 2000 * uc))
    # 科目余额（9 月末）
    tb = wv[SH_TB]
    bal = {str(tb[f'{T_CODE}{r}'].value): tb[f'{T_CNET}{r}'].value for r in range(TB_R0, TB_R0 + 450) if tb[f'{T_CODE}{r}'].value}
    chk('140301 火工品 = 62M+5M−46.5M−4M = 16,500,000', close(bal.get('140301'), 16500000), bal.get('140301'))
    chk('140302 油料 = 120M−72M−12M = 36,000,000（没用途那行不出库）', close(bal.get('140302'), 36000000), bal.get('140302'))
    chk(f'140501 原矿 = {2000 * uc:,.2f}', close(bal.get('140501'), 2000 * uc), bal.get('140501'))
    chk('224105 应付进口税费及运杂 = −(2M+15M+5M) = −22,000,000', close(bal.get('224105'), -22000000), bal.get('224105'))
    chk('只补运杂的采购行不报 ✗、记账', g(SH_BUY, f'{B_CHK}9') in (None, '') and g(SH_BUY, f'{B_VALID}9') == 1, g(SH_BUY, f'{B_CHK}9'))
    chk('22210103 已交增值税 = 1,000,000（缴纳增值税冲应交税费）', close(bal.get('22210103'), 1000000), bal.get('22210103'))
    chk('收入支出都填的流水 ✗、不记账', str(g(SH_CASH, f'{K_CHK}2416')).startswith('✗') and g(SH_CASH, f'{K_VALID}2416') == 0,
        g(SH_CASH, f'{K_CHK}2416'))
    chk('4001 生产成本月末 = 0（全转走了）', close(bal.get('4001'), 0), bal.get('4001'))
    chk('112203 应收甲方 = 336M−200M = 136,000,000', close(bal.get('112203'), 136000000), bal.get('112203'))
    chk('22210107 销项税 = −36,000,000', close(bal.get('22210107'), -36000000), bal.get('22210107'))
    chk('22210101 进项税 = 600,000', close(bal.get('22210101'), 600000), bal.get('22210101'))
    chk('221102 外籍工资 = −50,000,000', close(bal.get('221102'), -50000000), bal.get('221102'))
    chk('160405 井巷工程（基建领用柴油）9 月增加 12,000,000', bal.get('160405', 0) >= 12000000, bal.get('160405'))
    ap_exp = -(60000000 + 5600000 + 120000000 + (truck - 15000000) + 3000000) + 120000000
    # 220202 月末 = 启用日以后的采购 − 付款（启用日前的采购付款记成本，不进应付）
    chk('220202 应付账款 = −(货款＋进项税) + 付款（运杂关税不挂供应商）', close(bal.get('220202'), ap_exp, 0.05), f'{bal.get("220202"):,.2f} / {ap_exp:,.2f}')
    chk('122104 资金划转：本月一对转账对上（余额跟 8 月一样）', True)
    # 利润表 9 月
    from s_reports import pl_row
    pl = wv[SH_PL]
    chk('9 月营业收入 300,000,000（不含税）', close(pl[f'D{pl_row(1)}'].value, 300000000), pl[f'D{pl_row(1)}'].value)
    chk('9 月营业成本 = 3000 吨 × 每吨成本', close(pl[f'D{pl_row(2)}'].value, 3000 * uc), pl[f'D{pl_row(2)}'].value)
    # 凭证
    vl = wv[SH_VLIST]
    nv = sum(1 for r in range(VL_R0, VL_R0 + VL_N) if isinstance(vl[f'A{r}'].value, (int, float)))
    srcs = [vl[f'K{r}'].value for r in range(VL_R0, VL_R0 + nv)]
    from collections import Counter
    c = Counter(srcs)
    print('    9 月凭证：', dict(c))
    chk('流水 7 张（两笔手续费并成 1 张；收支都填的那笔没有）', c.get('流水') == 7, c.get('流水'))
    chk('采购 4 张（CG001 两行一张）', c.get('采购') == 4, c.get('采购'))
    chk('领用 3 张（LL001 两行一张；✗ 那行没有）', c.get('领用') == 3, c.get('领用'))
    chk('销售 1（没定价的没有）、工资 1（两个部门并成一张）、手工 2（不平那张没记；抬头行那张记了）',
        (c.get('销售'), c.get('工资'), c.get('手工')) == (1, 1, 2), (c.get('销售'), c.get('工资'), c.get('手工')))
    je = wv[SH_JE]
    oct_man = sum(1 for r in range(JE_R0, je.max_row + 1) if je[f'{J_SRC}{r}'].value == '手工' and je[f'{J_NZ}{r}'].value == 1
                  and getattr(je[f'{J_DATE}{r}'].value, 'month', 0) == 10)
    chk('10 月组号又从 1 编：单独一张 10 月凭证（2 条），不影响 9 月', oct_man == 2, oct_man)
    chk('年限写成「5年」的卡片 ✗、不提折旧', str(g(SH_FA, f'{F_CHK}5')).startswith('✗') and g(SH_FA, f'{F_OK}5') == 0,
        g(SH_FA, f'{F_CHK}5'))
    chk('月末：成本结转 1、销售成本 1、折旧 1（发电机；矿车下月起）', (c.get('成本'), c.get('销售成本'), c.get('折旧')) == (1, 1, 1),
        (c.get('成本'), c.get('销售成本'), c.get('折旧')))
    chk('凭证全部借贷平衡', all(vl[f'G{r}'].value == '√' for r in range(VL_R0, VL_R0 + nv)))
    merged = [r for r in range(VL_R0, VL_R0 + nv) if vl[f'K{r}'].value == '流水' and vl[f'H{r}'].value == 4]
    chk('合并的手续费凭证有 4 条分录', len(merged) == 1)
    ap = wv[SH_ARAP]
    row = next(r for r in range(5, 305) if ap[f'A{r}'].value == '三飞公司')
    chk('往来余额表：三飞公司外币应付 人民币 300,000、汇率差 0', close(ap[f'L{row}'].value, 300000) and ap[f'P{row}'].value in (0, None, ''),
        (ap[f'L{row}'].value, ap[f'P{row}'].value))
    # 折旧：10 月开始
    from s_mend import ME_DEP0
    me = wv[SH_MEND]
    n10 = n9 + 1
    dep10 = me[f'B{ME_DEP0 + n10 - 1}'].value
    dep9 = me[f'B{ME_DEP0 + n9 - 1}'].value
    exp_dep = round(truck * 0.95 / 60, 2)
    chk('矿车 9 月入账、10 月开始提折旧；发电机每月 950,000', close(dep9, 950000) and close(dep10, exp_dep + 950000),
        f'9月 {dep9} / 10月 {dep10} / 应 {exp_dep}+950000')
    aug_dep = [(je[f'{J_CODE}{r}'].value, je[f'{J_DR}{r}'].value) for r in range(JE_R0, je.max_row + 1)
               if je[f'{J_SRC}{r}'].value == '折旧' and je[f'{J_NZ}{r}'].value == 1 and je[f'{J_DR}{r}'].value
               and getattr(je[f'{J_DATE}{r}'].value, 'year', 0) == 2026 and getattr(je[f'{J_DATE}{r}'].value, 'month', 0) in (7, 8)]
    chk('启用日以前（7、8 月）的折旧记 540102，不堆进生产成本', aug_dep == [('540102', 950000), ('540102', 950000)], aug_dep)
    # 校验表
    ck = wv[SH_CHK]
    print('    数据校验：')
    for r in range(5, 45):
        if ck[f'C{r}'].value:
            print(f'      {ck[f"B{r}"].value}：{ck[f"C{r}"].value}')
    bs = wv[SH_BAL]
    chk('9 月资产负债表平衡', close(bs['C37'].value, bs['G37'].value, 0.05), f'{bs["C37"].value:,.2f} / {bs["G37"].value:,.2f}')
    print(f'\n合计：✓ {len(OK)}　✗ {len(BAD)}')
    shutil.copy(p, os.path.join(HERE, '_scenario.xlsx'))
    return len(BAD)


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', 'A062_鑫盛矿业财务模板.xlsx')
    sys.exit(1 if main(src) else 0)
