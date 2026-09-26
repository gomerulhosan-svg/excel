# -*- coding: utf-8 -*-
"""从两个原表里把业务数据读出来，套上 fixes.py 里逐条写明的修正，交给生成脚本。

原则：用户手工录的数（数量、单价、金额、日期、备注）一律原样搬；
只有 fixes.py 里写明了「为什么要改」的那几处才改，改动清单同时写进说明.md。
"""
import datetime as dt
import re
import openpyxl

import fixes


def _s(v):
    """单元格文字：去首尾空格；None → ''"""
    if v is None:
        return ''
    return str(v).strip() if isinstance(v, str) else v


def parse_date(v):
    """原表里的日期有真日期，也有 '2026.6.25'、'206-7-31' 这种文字，统一成 date"""
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 40000 < v < 60000:
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(v))).date()
    if isinstance(v, str):
        t = v.strip()
        if t in fixes.DATE_TEXT:
            return fixes.DATE_TEXT[t]
        m = re.fullmatch(r'(\d{4})[.\-/年](\d{1,2})[.\-/月](\d{1,2})日?', t)
        if m:
            return dt.date(int(m[1]), int(m[2]), int(m[3]))
    return None


def canon_customer(name):
    n = _s(name)
    return fixes.CUSTOMER_MERGE.get(n, n)


def canon_goods(name):
    n = _s(name)
    return fixes.GOODS_MERGE.get(n, n)


class Src1:
    """原表1：总览汇总 / 出库明细 / 收付款明细 / 存条明细 / 商品库存"""

    def __init__(self, path):
        self.path = path
        self.wb = openpyxl.load_workbook(path)                  # 公式
        self.wv = openpyxl.load_workbook(path, data_only=True)  # 缓存值

    def _rows(self, sheet, r0, ncol, key_cols):
        ws, wv = self.wb[sheet], self.wv[sheet]
        rows = []
        for r in range(r0, ws.max_row + 1):
            vals = [ws.cell(r, c).value for c in range(1, ncol + 1)]
            if all(_s(vals[c - 1]) in ('', None) for c in key_cols):
                continue
            cached = [wv.cell(r, c).value for c in range(1, ncol + 1)]
            rows.append({'row': r, 'v': vals, 'cached': cached})
        return rows

    def outbound(self):
        """出库明细 → [{row, date, raw_date, customer, goods, qty, price, amount_cell, note}]"""
        out = []
        for x in self._rows('出库明细', 4, 8, key_cols=(2, 3, 4, 5, 6)):
            v = x['v']
            out.append(dict(row=x['row'], raw_date=v[1], date=parse_date(v[1]),
                            customer_raw=_s(v[2]), customer=canon_customer(v[2]),
                            goods_raw=_s(v[3]), goods=canon_goods(v[3]),
                            qty=v[4], price=v[5], amount=v[6], amount_cached=x['cached'][6],
                            note=v[7]))
        return out

    def deposits(self):
        """存条明细 → 同上结构"""
        res = []
        for x in self._rows('存条明细', 4, 8, key_cols=(2, 3, 4, 5, 6)):
            v = x['v']
            res.append(dict(row=x['row'], raw_date=v[1], date=parse_date(v[1]),
                            customer_raw=_s(v[2]), customer=canon_customer(v[2]),
                            goods_raw=_s(v[3]), goods=canon_goods(v[3]),
                            qty=v[4], price=v[5], amount=v[6], amount_cached=x['cached'][6],
                            note=v[7]))
        return res

    def receipts(self):
        """收付款明细 → [{row, date, customer, amount, method, note}]"""
        res = []
        for x in self._rows('收付款明细', 4, 6, key_cols=(2, 3, 4)):
            v = x['v']
            res.append(dict(row=x['row'], raw_date=v[1], date=parse_date(v[1]),
                            customer_raw=_s(v[2]), customer=canon_customer(v[2]),
                            amount=v[3], method=_s(v[4]) or '收款码', note=_s(v[5])))
        return res

    def overview(self):
        """总览汇总 B/C → [(row, 客户, 初始存条金额)]"""
        ws = self.wb['总览汇总']
        res = []
        for r in range(5, 155):
            b = _s(ws.cell(r, 2).value)
            if b:
                res.append((r, b, ws.cell(r, 3).value))
        return res


class Src2:
    """原表2：报损"""

    def __init__(self, path):
        self.path = path
        self.wb = openpyxl.load_workbook(path)
        self.wv = openpyxl.load_workbook(path, data_only=True)

    def products(self):
        """所有报损月表里出现过的 (所属公司, 产品名称, 规格, 单价)，按首次出现排序"""
        seen, res = set(), []
        for sh in ['报损明细台账', '2026年7月报损', '2026年8月报损']:
            ws = self.wb[sh]
            for r in range(6, ws.max_row + 1):
                comp, name = _s(ws.cell(r, 2).value), _s(ws.cell(r, 3).value)
                if not name or name in seen:
                    continue
                seen.add(name)
                res.append(dict(company=comp, name=name, spec=_s(ws.cell(r, 4).value),
                                price=ws.cell(r, 7).value, unit=_s(ws.cell(r, 6).value)))
        return res


def build_customer_master(s1):
    """客户名单：先按总览汇总原顺序（重名只留第一次），再补上只在流水里出现过的客户（按首次日期）。
    返回 [(客户, 初始存条金额合计, 来源说明)]"""
    order, dep, src = [], {}, {}
    for r, name, c in s1.overview():
        n = canon_customer(name)
        if n not in dep:
            order.append(n)
            dep[n] = 0.0
            src[n] = []
        val = fixes.DEPOSIT_OVERRIDE.get((r, name), c)
        if isinstance(val, (int, float)) and (r, name) not in fixes.DEPOSIT_DROP:
            dep[n] += float(val)
        src[n].append(r)
    extra = []
    for rec in s1.deposits() + s1.outbound() + s1.receipts():
        n = rec['customer']
        if n and n not in dep and n not in [e[0] for e in extra]:
            extra.append((n, rec['date'] or dt.date(2100, 1, 1)))
    extra.sort(key=lambda t: t[1])
    res = [(n, (dep[n] if dep[n] else None), src[n]) for n in order]
    res += [(n, None, []) for n, _ in extra]
    return res
