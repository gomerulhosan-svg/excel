# -*- coding: utf-8 -*-
"""把你发来的三个文件读进来：农行流水（→ 资金台帐演示数据）、销项/进项发票（→ 发票导入演示数据），
   并整理出往来单位初始名单。只读原件，不改原件。"""
import os
import xlrd
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
REF = os.path.join(HERE, '..', '参考')
F_BANK = os.path.join(REF, '农行流水_2896_202609.xls')
F_SALE = os.path.join(REF, '销项发票_中智优力_202607-08.xlsx')
F_BUY = os.path.join(REF, '进项发票_中智优力_202607-08.xlsx')

MY_CO = ('中智优力', '中智优力（宁波）财务咨询有限公司', '91330212MA2CKKG076')
MY_ACC = '农行2896'


def bank_rows():
    """农行导出：第 3 行起是明细，列顺序 交易时间|收入金额|支出金额|账户余额|对方账号|对方户名|对方开户行|摘要。
       原样保留（交易时间是文本，金额是数字），跟你从导出文件复制粘贴过来一样。"""
    sh = xlrd.open_workbook(F_BANK).sheet_by_index(0)
    head = [sh.cell_value(1, c) for c in range(sh.ncols)]
    acct = str(head[0]).split(':', 1)[-1].strip()
    rows = []
    for r in range(3, sh.nrows):
        v = [sh.cell_value(r, c) for c in range(8)]
        if not str(v[0]).startswith('20'):
            continue
        out = []
        for i, x in enumerate(v):
            if i in (1, 2, 3):
                out.append(float(x) if x not in ('', None) else None)
            else:
                out.append(str(x) if x not in ('', None) else None)
        rows.append(out)
    first_in = rows[0][1] or 0
    first_out = rows[0][2] or 0
    opening = round(rows[0][3] - first_in + first_out, 2)
    return acct, opening, rows


def bank_raw():
    """农行导出文件原样（标题、账号行、表头、明细、末尾几行都在），给【流水1】当演示：就像你把导出文件整份复制粘贴过来"""
    sh = xlrd.open_workbook(F_BANK).sheet_by_index(0)
    out = []
    for r in range(sh.nrows):
        row = []
        for c in range(sh.ncols):
            v = sh.cell_value(r, c)
            row.append(None if v in ('', None) else v)
        out.append(row)
    while out and all(v is None for v in out[-1]):
        out.pop()
    return out


def invoice_rows(path):
    """「发票基础信息」那一页，19 列原样；最后的「合计行」不要"""
    ws = openpyxl.load_workbook(path)['发票基础信息']
    out = []
    for row in ws.iter_rows(min_row=2, max_col=19, values_only=True):
        if row[0] in (None, '') or str(row[0]).startswith('合计'):
            continue
        out.append(list(row))
    return out


def parties(bank, sales, buys):
    """初始往来单位：销项的购买方＋流水里给我们打钱的单位 → 客户；看得出用途的几个付款对象 → 绑好收支项目。
       进项发票上的都是吃饭、加油、过路费这类报销票，不登记（不登记就不算应付）。"""
    cus = {}
    for r in sales:
        cus.setdefault(r[7].replace('(', '（').replace(')', '）'), r[6])
    for r in bank:
        if r[1] and r[5] and r[5] != '郭伟健':
            cus.setdefault(r[5], None)
    out = []
    for name, tax in cus.items():
        out.append(dict(名称=name, 类型='客户', 税号=tax))
    out.append(dict(名称='中国电信股份有限公司宁波分公司', 类型='供应商', 绑定='办公费用',
                    别名1='中国电信股份有限公司宁波分公司（电信）', 备注='公共缴费：电话宽带'))
    out.append(dict(名称='国库（税款）', 类型='税务银行', 绑定='税金',
                    别名1='待报解预算收入－国库信息处理系统扣税', 备注='银行扣税'))
    return out


if __name__ == '__main__':
    a, o, b = bank_rows()
    print(a, o, len(b), b[0])
    s, p = invoice_rows(F_SALE), invoice_rows(F_BUY)
    print(len(s), len(p))
    for x in parties(b, s, p):
        print(x)
