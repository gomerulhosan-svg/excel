# -*- coding: utf-8 -*-
"""独立复算：不看表里的公式，直接从各录入表（资金日记账、送货单、外发、订单、工资、手工分录）和档案的「录入格」用 Python 按同样的口径重算，
   再跟 LibreOffice 算出来的表里的数逐项比（款式×月 的双数/收入/成本、成本分摊表、订单成本、利润表、资产负债表、账户余额、供应商欠款……）。

跑法：python3 verify.py <LibreOffice 算过的副本.xlsx>      （build.py --keep 留下的 _calc/…_算数.xlsx）
"""
import datetime as dt
import re
import sys
from collections import defaultdict

import openpyxl
from layout import *

TOL = 0.011


def d2(x):
    return round(x + 1e-9 if x >= 0 else x - 1e-9, 2)


def num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return None          # 文字


def isdate(v):
    return isinstance(v, (dt.date, dt.datetime))


def todate(v):
    return v.date() if isinstance(v, dt.datetime) else v


def norm_name(s):
    return str(s).strip().replace(' ', '').replace('(', '（').replace(')', '）')


class Model:
    def __init__(self, path):
        self.wb = openpyxl.load_workbook(path, data_only=True)
        wb = self.wb
        b = wb[SH_BASE]
        self.year = int(b[P['YEAR'].split('!')[1].replace('$', '')].value)
        self.open = todate(b[P['OPEN'].split('!')[1].replace('$', '')].value)
        self.lock = int(num(b[P['LOCK'].split('!')[1].replace('$', '')].value) or 0)
        self.basis = b[P['BASIS'].split('!')[1].replace('$', '')].value
        self.cust0 = str(b[P['CUST'].split('!')[1].replace('$', '')].value or '').strip()
        self.lagd = num(b[P['LAGD'].split('!')[1].replace('$', '')].value)
        self.open_m = self.open.month if self.open.year == self.year else 1
        # 账户
        self.acc = {}
        nb = npay = 0
        for r in range(AC_R0, AC_R1 + 1):
            nm = b[f'{AC_NAME}{r}'].value
            if not nm:
                continue
            tp = b[f'{AC_TYPE}{r}'].value
            if tp == '现金':
                code = '1001'
            elif tp == '银行':
                nb += 1
                code = f'1002{nb:02d}'
            else:
                npay += 1
                code = f'1012{npay:02d}'
            self.acc[str(nm).strip()] = dict(code=code, open=num(b[f'{AC_OPEN}{r}'].value) or 0)
        self.cats = {}
        for r in range(CT_R0, CT_R1 + 1):
            nm = b[f'{CT_NAME}{r}'].value
            if nm:
                self.cats[nm] = dict(opp=str(b[f'{CT_OPP}{r}'].value or ''), need=b[f'{CT_NEED}{r}'].value, cf=b[f'{CT_CF}{r}'].value)
        self.fees = {}
        for r in range(FE_R0, FE_R1 + 1):
            nm = b[f'{FE_NAME}{r}'].value
            if nm:
                cls = b[f'{FE_CLS}{r}'].value
                self.fees[nm] = dict(cls=cls, code=FE_CLASSES[cls][0])
        self.depts = {}
        for r in range(DP_R0, DP_R1 + 1):
            nm = b[f'{DP_NAME}{r}'].value
            if nm:
                self.depts[nm] = b[f'{DP_TYPE}{r}'].value
        # 固定资产月折旧
        self.dep = defaultdict(float)          # (m, use) → 折旧
        self.fa_before = [0.0, 0.0]
        for r in range(FA_R0, FA_R1 + 1):
            nm = b[f'{FA_NAME}{r}'].value
            if not nm:
                continue
            d0 = todate(b[f'{FA_DATE}{r}'].value)
            cost = num(b[f'{FA_COST}{r}'].value) or 0
            mon = num(b[f'{FA_MON}{r}'].value) or 0
            res = num(b[f'{FA_RES}{r}'].value) or 0
            dep0 = num(b[f'{FA_DEP0}{r}'].value) or 0
            use = b[f'{FA_USE}{r}'].value
            if isdate(d0) and d0 < self.open:
                self.fa_before[0] += cost
                self.fa_before[1] += dep0
            if not isdate(d0) or not cost or mon <= 0:
                continue
            dm = d2(cost * (1 - res) / mon)
            start = d0.year * 12 + d0.month + 1
            t_open = self.open.year * 12 + self.open.month
            for m in range(1, 13):
                t = self.year * 12 + m
                if t < t_open or t < start:
                    continue
                v = d2(max(0, min(dm, cost * (1 - res) - dep0 - dm * (t - max(start, t_open)))))
                self.dep[(m, use)] += v
        # 往来单位
        u = wb[SH_UNIT]
        self.units = {}
        for r in range(UN_R0, UN_R1 + 1):
            nm = u[f'{UN_NAME}{r}'].value
            if nm:
                self.units[str(nm).strip()] = dict(type=u[f'{UN_TYPE}{r}'].value, ap0=num(u[f'{UN_AP0}{r}'].value) or 0,
                                                   ar0=num(u[f'{UN_AR0}{r}'].value) or 0, or0=num(u[f'{UN_OR0}{r}'].value) or 0,
                                                   op0=num(u[f'{UN_OP0}{r}'].value) or 0, row=r)
        # 款式
        s = wb[SH_STY]
        self.sty = []           # 按档案行
        self.sty_idx = {}
        for r in range(ST_R0, ST_R1 + 1):
            code = s[f'{ST_CODE}{r}'].value
            key = str(code).strip() if code not in (None, '') else ''
            self.sty.append(dict(code=key, price=num(s[f'{ST_PRICE}{r}'].value) or 0,
                                 km=1 if s[f'{ST_KM}{r}'].value in (None, '') else num(s[f'{ST_KM}{r}'].value),
                                 kl=1 if s[f'{ST_KL}{r}'].value in (None, '') else num(s[f'{ST_KL}{r}'].value),
                                 sdm=num(s[f'{ST_SDM}{r}'].value) or 0, sdl=num(s[f'{ST_SDL}{r}'].value) or 0))
            if key and key.upper() not in self.sty_idx:
                self.sty_idx[key.upper()] = len(self.sty) - 1
        sdms = [x['sdm'] for x in self.sty if x['sdm'] > 0]
        sdls = [x['sdl'] for x in self.sty if x['sdl'] > 0]
        am, al_ = (sum(sdms) / len(sdms) if sdms else 1), (sum(sdls) / len(sdls) if sdls else 1)
        for x in self.sty:
            if not x['code']:
                x['wm'] = x['wl'] = 0
            elif self.basis == '标准成本':
                x['wm'] = x['sdm'] if x['sdm'] > 0 else x['km'] * am
                x['wl'] = x['sdl'] if x['sdl'] > 0 else x['kl'] * al_
            else:
                x['wm'], x['wl'] = x['km'], x['kl']
        self.je = []            # (月, 借, 贷, 金额, 单位, 来源, 日期)
        self.cost = []          # (类型 款/公, 款式行号 or None, 组件, 月, 金额)
        self._journal()
        self._dn()
        self._out()
        self._orders()
        self._wages()
        self._mj()
        self._engine()

    def month_of(self, d):
        if not isdate(d):
            return 0
        d = todate(d)
        return d.month if d.year == self.year else 0

    def sty_of(self, code):
        if code in (None, ''):
            return None
        return self.sty_idx.get(str(code).strip().upper())

    # ── 资金日记账
    def _journal(self):
        ws = self.wb[SH_CASH]
        self.jrows = []
        for r in range(J_R0, J_R1 + 1):
            v = {c: ws[f'{c}{r}'].value for c in (J_DATE, J_ACC, J_CAT, J_UNIT, J_FEE, J_MEMO, J_IN, J_OUT, J_TOACC, J_STY)}
            if all(v[c] in (None, '') for c in (J_DATE, J_ACC, J_IN, J_OUT)):
                continue
            i_, o_ = num(v[J_IN]), num(v[J_OUT])
            row = dict(r=r, ok=False, net=0.0, acc=str(v[J_ACC] or '').strip(), cat='', to='')
            self.jrows.append(row)
            if i_ is None or o_ is None or i_ < 0 or o_ < 0 or (i_ and o_):
                continue
            net = d2(i_ - o_)
            row['net'] = net
            unit = str(v[J_UNIT] or '').strip()
            utype = self.units.get(unit, {}).get('type', '')
            cat = v[J_CAT] or ''
            if not cat:
                if v[J_TOACC]:
                    cat = '内部转账'
                elif v[J_FEE]:
                    cat = '费用支出'
                elif utype == '材料供应商':
                    cat = '付材料款'
                elif utype == '外发加工厂':
                    cat = '付加工费'
                elif utype == '客户':
                    cat = '收货款'
            row['cat'] = cat
            if cat == '内部转账':
                row['to'] = str(v[J_TOACC] or '').strip()
            d = todate(v[J_DATE])
            m = self.month_of(d)
            if not row['acc'] in self.acc or not isdate(d) or m == 0 or d < self.open or net == 0 or cat not in self.cats:
                continue
            need = self.cats[cat]['need']
            if need == '单位' and (not unit or not utype):
                continue
            if need == '费用项目':
                if v[J_FEE] not in self.fees:
                    continue
                opp = self.fees[v[J_FEE]]['code']
            elif need == '对方账户':
                if row['to'] not in self.acc or row['to'] == row['acc']:
                    continue
                opp = self.acc[row['to']]['code']
            else:
                opp = self.cats[cat]['opp']
            if not opp:
                continue
            row['ok'] = True
            a = self.acc[row['acc']]['code']
            dr, cr = (a, opp) if net > 0 else (opp, a)
            self.je.append((m, dr, cr, abs(net), unit, '日记账', d))
            comp = {'400101': '材料', '400102': '外发', '400103': '人工', '4101': '制造'}.get(opp)
            if comp:
                si = self.sty_of(v[J_STY])
                self.cost.append(('款' if si is not None else '公', si, comp, m, -net))
            row['cf'] = None if cat == '内部转账' else (
                ('支付的各项税费' if self.fees[v[J_FEE]]['cls'] in ('税金及附加', '所得税') else '支付其他与经营有关的现金')
                if need == '费用项目' else self.cats[cat]['cf'])
            row['m'] = m

    # ── 送货单 / 外发
    def _reg(self, sh, c_date, c_rdate, c_sup, c_qty, c_price, c_amt, c_sty, comp, r0, r1, code):
        ws = self.wb[sh]
        out = []
        for r in range(r0, r1 + 1):
            d, rd, sup = ws[f'{c_date}{r}'].value, ws[f'{c_rdate}{r}'].value, ws[f'{c_sup}{r}'].value
            q_, p_, a_ = num(ws[f'{c_qty}{r}'].value), num(ws[f'{c_price}{r}'].value), ws[f'{c_amt}{r}'].value
            if d is None and sup is None and a_ is None and ws[f'{c_qty}{r}'].value is None:
                continue
            if q_ is None or p_ is None or (a_ not in (None, '') and num(a_) is None):
                continue
            amt = d2(num(a_)) if a_ not in (None, '') else d2(q_ * p_)
            if not isdate(d) or not sup or str(sup).strip() not in self.units or amt == 0:
                continue
            d = todate(d)
            if rd not in (None, '') and not isdate(rd):
                continue
            rd = todate(rd) if rd not in (None, '') else None
            if rd and rd < d:
                continue
            def mm(x):
                if x.year < self.year:
                    return 0
                if x.year > self.year:
                    return 13
                return x.month
            sm = mm(d)
            rm = mm(rd) if rd else sm
            if sm == 13:
                continue
            if d < self.open:
                if (rd or d) < self.open:
                    continue
                cm = max(rm, self.open_m)
            elif self.lock > 0 and sm <= self.lock and rm > self.lock:
                cm = rm
            else:
                cm = sm
            if cm < 1 or cm > 12:
                continue
            si = self.sty_of(ws[f'{c_sty}{r}'].value)
            self.cost.append(('款' if si is not None else '公', si, comp, cm, amt))
            self.je.append((cm, code, '2202', amt, str(sup).strip(), sh, d))
            out.append(dict(r=r, cm=cm, sm=sm, rm=rm, amt=amt, sup=str(sup).strip(), lag=(rd - d).days if rd else 0))
        return out

    def _dn(self):
        self.dn = self._reg(SH_DN, DN_DATE, DN_RDATE, DN_SUP, DN_QTY, DN_PRICE, DN_AMTIN, DN_STY, '材料', DN_R0, DN_R1, '400101')

    def _out(self):
        self.ot = self._reg(SH_OUT, OT_DATE, OT_RDATE, OT_SUP, OT_QTY, OT_PRICE, OT_AMTIN, OT_STY, '外发', OT_R0, OT_R1, '400102')

    # ── 订单
    def _orders(self):
        ws = self.wb[SH_ORD]
        self.od = []
        for r in range(OD_R0, OD_R1 + 1):
            v = {c: ws[f'{c}{r}'].value for c in (OD_NO, OD_DATE, OD_CUST, OD_STY, OD_SPEC, OD_QTY, OD_DDATE, OD_DQTY, OD_PRICE)}
            if all(v[c] in (None, '') for c in (OD_NO, OD_STY, OD_QTY, OD_DQTY)):
                continue
            si = self.sty_of(v[OD_STY])
            dd = v[OD_DDATE]
            dq = 0.0
            if isdate(dd):
                dq = (num(v[OD_QTY]) or 0) if v[OD_DQTY] in (None, '') else (num(v[OD_DQTY]) or 0)
            dm = self.month_of(dd) if (isdate(dd) and dq != 0 and todate(dd) >= self.open) else 0
            price = num(v[OD_PRICE]) if v[OD_PRICE] not in (None, '') else (self.sty[si]['price'] if si is not None else 0)
            amt = d2(dq * price) if dm else 0
            cust = str(v[OD_CUST]).strip() if v[OD_CUST] not in (None, '') else self.cust0
            ok = bool(dm and amt != 0 and v[OD_NO] not in (None, '') and v[OD_STY] not in (None, ''))
            self.od.append(dict(r=r, no=v[OD_NO], si=si, dq=dq, dm=dm, amt=amt if ok else 0, ok=ok, cust=cust,
                                oq=num(v[OD_QTY]) or 0))
            if ok:
                self.je.append((dm, '1122', '5001', amt, cust, '收入', None))

    # ── 工资
    def _wages(self):
        ws = self.wb[SH_WAGE]
        self.wg = []
        for r in range(WG_R0, WG_R1 + 1):
            m = ws[f'{WG_MON}{r}'].value
            dept = ws[f'{WG_DEPT}{r}'].value
            vals = [num(ws[f'{c}{r}'].value) for c in (WG_PQ, WG_PP, WG_BASE, WG_OT, WG_DED)]
            if m is None and dept is None and all(x == 0 for x in vals if x is not None):
                continue
            if any(x is None for x in vals) or not isinstance(m, (int, float)) or int(m) != m or not 1 <= m <= 12:
                continue
            m = int(m)
            if m < self.open_m or dept not in self.depts:
                continue
            pq, pp, base, ot, ded = vals
            amt = d2(d2(pq * pp) + base + ot - ded)
            if amt == 0:
                continue
            tp = self.depts[dept]
            comp, code, _ = DP_TYPES[tp]
            self.je.append((m, code, '2211', amt, '', '工资', None))
            if comp:
                si = self.sty_of(ws[f'{WG_STY}{r}'].value)
                self.cost.append(('款' if si is not None else '公', si, comp, m, amt))
            self.wg.append(dict(m=m, amt=amt, dept=dept, code=code))

    # ── 手工分录
    def _mj(self):
        ws = self.wb[SH_MJ]
        codes = {str(self.wb[SH_COA][f'{COA_CODE}{r}'].value).strip() for r in range(COA_R0, COA_R1 + 1)
                 if self.wb[SH_COA][f'{COA_CODE}{r}'].value}
        cmap = {'400101': '材料', '400102': '外发', '400103': '人工', '4101': '制造'}
        for r in range(MJ_R0, MJ_R1 + 1):
            d = ws[f'{MJ_DATE}{r}'].value
            drc, crc = str(ws[f'{MJ_DRIN}{r}'].value or '').strip(), str(ws[f'{MJ_CRIN}{r}'].value or '').strip()
            a = num(ws[f'{MJ_AMTIN}{r}'].value)
            if not isdate(d) or not drc or not crc or a in (None, 0) or drc == crc or drc not in codes or crc not in codes:
                continue
            m = self.month_of(d)
            if not m or todate(d) < self.open:
                continue
            if drc in ('400104',) or crc in ('400104',) or drc.startswith('5401') and len(drc) == 6 or crc.startswith('5401') and len(crc) == 6:
                continue
            if drc in cmap and crc in cmap:
                continue
            a = d2(a)
            self.je.append((m, drc, crc, a, str(ws[f'{MJ_UNIT}{r}'].value or '').strip(), '手工', todate(d)))
            if drc in cmap:
                self.cost.append(('公', None, cmap[drc], m, a))
            elif crc in cmap:
                self.cost.append(('公', None, cmap[crc], m, -a))

    # ── 成本引擎
    def _engine(self):
        n = len(self.sty)
        M = range(1, 13)
        self.pairs = [[0.0] * 13 for _ in range(n)]
        self.rev = [[0.0] * 13 for _ in range(n)]
        for o in self.od:
            if o['ok'] and o['si'] is not None:
                if o['dq'] > 0:
                    self.pairs[o['si']][o['dm']] += o['dq']
                self.rev[o['si']][o['dm']] += o['amt']
        direct = defaultdict(float)     # (si, comp, m)
        pub = defaultdict(float)        # (comp, m)
        for t, si, comp, m, a in self.cost:
            if t == '款':
                direct[(si, comp, m)] += a
            else:
                pub[(comp, m)] += a
        for m in M:
            pub[('制造', m)] += self.dep[(m, '车间')]
        self.pub = pub
        self.direct_m = defaultdict(float)
        for (si, comp, m), a in direct.items():
            self.direct_m[(comp, m)] += a
        # 直接转出
        T = defaultdict(float)
        for si in range(n):
            for c in COMPS:
                done = 0.0
                cum = 0.0
                for m in M:
                    cum += direct[(si, c, m)]
                    if self.pairs[si][m] > 0:
                        T[(si, c, m)] = d2(cum - done)
                        done += T[(si, c, m)]
        self.T = T
        wm = [x['wm'] for x in self.sty]
        wl = [x['wl'] for x in self.sty]
        self.base_m = {m: sum(self.pairs[i][m] * wm[i] for i in range(n)) for m in M}
        self.base_l = {m: sum(self.pairs[i][m] * wl[i] for i in range(n)) for m in M}
        self.wip_est = {}
        a = self.wb[SH_ALLOC]
        for m in M:
            self.wip_est[m] = num(a[f'{al_col(m)}{AL_ROWS["在制估计"]}'].value) or 0
        self.alloc = defaultdict(float)
        carry = defaultdict(float)
        self.carry = {}
        for m in M:
            avail = {c: carry[c] + pub[(c, m)] for c in COMPS}
            tot = sum(avail.values())
            for c in COMPS:
                keep = min(avail[c], self.wip_est[m] * avail[c] / tot) if tot else 0
                base = self.base_m[m] if c == '材料' else self.base_l[m]
                self.alloc[(c, m)] = avail[c] - keep if base > 0 else 0
                carry[c] = avail[c] - self.alloc[(c, m)]
                self.carry[(c, m)] = carry[c]
        self.C = defaultdict(float)
        for si in range(n):
            for m in M:
                p = self.pairs[si][m]
                for c in COMPS:
                    v = T[(si, c, m)]
                    base = self.base_m[m] if c == '材料' else self.base_l[m]
                    w = wm[si] if c == '材料' else wl[si]
                    if p > 0 and base > 0:
                        v += self.alloc[(c, m)] * p * w / base
                    self.C[(si, c, m)] = v
        self.out_c = {(c, m): sum(self.C[(si, c, m)] for si in range(n)) for c in COMPS for m in M}
        for m in M:
            self.je.append((m, '400104', '4101', d2(pub[('制造', m)] + self.direct_m[('制造', m)]), '', '制造转入', None))
            for c in COMPS:
                self.je.append((m, COGS_CODE[c], COMP_CODE[c], d2(self.out_c[(c, m)]), '', '成本结转', None))
            for use, code in FA_USES.items():
                if self.dep[(m, use)]:
                    self.je.append((m, code, '1602', d2(self.dep[(m, use)]), '', '折旧', None))

    # ── 账
    def mov(self, prefix, m_from=1, m_to=12, unit=None):
        dr = sum(a for m, d, c, a, u, *_ in self.je if m_from <= m <= m_to and d.startswith(prefix) and (unit is None or u == unit))
        cr = sum(a for m, d, c, a, u, *_ in self.je if m_from <= m <= m_to and c.startswith(prefix) and (unit is None or u == unit))
        return dr, cr

    def pl(self, m1, m2):
        def cr_dr(p):
            d, c = self.mov(p, m1, m2)
            return c - d
        def dr_cr(p):
            d, c = self.mov(p, m1, m2)
            return d - c
        rev = cr_dr('5001') + cr_dr('5051')
        cost = dr_cr('5401') + dr_cr('5402')
        tax, sell, adm, fin = dr_cr('5403'), dr_cr('5601'), dr_cr('5602'), dr_cr('5603')
        oi, oe, it = cr_dr('5301'), dr_cr('5711'), dr_cr('5801')
        op = rev - cost - tax - sell - adm - fin
        return dict(收入=rev, 成本=cost, 毛利=rev - cost, 营业利润=op, 利润总额=op + oi - oe, 净利润=op + oi - oe - it,
                    材料=dr_cr('540101'), 外发=dr_cr('540102'), 人工=dr_cr('540103'), 制造=dr_cr('540104'))

    def bal(self, prefix, m, natural_debit=True, open_=0.0, unit=None):
        d, c = self.mov(prefix, 1, m, unit)
        return open_ + (d - c if natural_debit else c - d)


# ─────────────────────────── 跟表里比 ───────────────────────────
def compare(path):
    md = Model(path)
    wb = md.wb
    bad, n_ok = [], 0

    def chk(what, got, exp, tol=TOL):
        nonlocal n_ok
        g = got if isinstance(got, (int, float)) else (0 if got in (None, '') else got)
        if not isinstance(g, (int, float)) or abs(g - exp) > tol:
            bad.append(f'{what}: 表里 {got!r}，复算 {exp:.2f}')
        else:
            n_ok += 1

    sm = wb[SH_SM]
    for i, x in enumerate(md.sty):
        r = SM_R0 + i
        if not x['code']:
            continue
        for m in range(1, 13):
            chk(f'_款式月 {x["code"]} {m}月 双数', sm[f'{sm_col("双数", m)}{r}'].value, md.pairs[i][m])
            chk(f'_款式月 {x["code"]} {m}月 收入', sm[f'{sm_col("收入", m)}{r}'].value, md.rev[i][m])
            for c in COMPS:
                chk(f'_款式月 {x["code"]} {m}月 转{c}', sm[f'{sm_col("转" + c, m)}{r}'].value, md.T[(i, c, m)])
                chk(f'_款式月 {x["code"]} {m}月 本{c}', sm[f'{sm_col("本" + c, m)}{r}'].value, md.C[(i, c, m)], 0.02)
    al = wb[SH_ALLOC]
    for m in range(1, 13):
        chk(f'成本分摊表 {m}月 双数', al[f'{al_col(m)}{AL_ROWS["双数"]}'].value, sum(md.pairs[i][m] for i in range(len(md.sty))))
        for c in COMPS:
            occ = md.pub[(c, m)] + md.direct_m[(c, m)]
            chk(f'成本分摊表 {c} {m}月 本月发生', al[f'{al_col(m)}{AL_CROWS[c]["本月发生"]}'].value, occ)
            chk(f'成本分摊表 {c} {m}月 公共本月分摊', al[f'{al_col(m)}{AL_CROWS[c]["公共本月分摊"]}'].value, md.alloc[(c, m)])
            chk(f'成本分摊表 {c} {m}月 本月转出合计', al[f'{al_col(m)}{AL_CROWS[c]["本月转出合计"]}'].value, md.out_c[(c, m)], 0.05)
    # 订单行成本合计 = 款式成本合计
    od = wb[SH_ORD]
    tot_line = sum((od[f'{OD_COSTV}{o["r"]}'].value or 0) for o in md.od if isinstance(od[f'{OD_COSTV}{o["r"]}'].value, (int, float)))
    chk('订单明细 成本列合计 vs 复算款式成本合计', tot_line, sum(md.C.values()), 0.5)
    for o in md.od:
        chk(f'订单明细 第{o["r"]}行 交货金额(隐藏AMT)', od[f'{OD_AMT}{o["r"]}'].value, o['amt'])
    # 账户余额
    b = wb[SH_BASE]
    for r in range(AC_R0, AC_R1 + 1):
        nm = b[f'{AC_NAME}{r}'].value
        if not nm:
            continue
        nm = str(nm).strip()
        exp = md.acc[nm]['open'] + sum(j['net'] for j in md.jrows if j['acc'] == nm) - sum(j['net'] for j in md.jrows if j.get('to') == nm)
        chk(f'基础资料 {nm} 当前余额', b[f'{AC_NOW}{r}'].value, exp)
    # 往来单位现在余额（供应商）
    u = wb[SH_UNIT]
    for nm, x in md.units.items():
        if x['type'] in ('材料供应商', '外发加工厂'):
            d, c = md.mov('2202', 1, 12, nm)
            chk(f'往来单位 {nm} 现在欠款', u[f'M{x["row"]}'].value, x['ap0'] + c - d)
        elif x['type'] == '客户':
            d, c = md.mov('1122', 1, 12, nm)
            chk(f'往来单位 {nm} 现在应收', u[f'M{x["row"]}'].value, x['ar0'] + d - c)
    # 利润表：按标签找行，1～12 月 + 全年
    is_ = wb[SH_IS]
    lab = {}
    for row in is_.iter_rows(min_row=1, max_row=80):
        for cc in row[:3]:
            if isinstance(cc.value, str):
                t = re.sub(r'[\s一二三四五六七八九十、：:（）()减加其中]', '', cc.value)
                lab.setdefault(t, cc.row)
    mcol = {}
    for row in is_.iter_rows(min_row=1, max_row=8):
        for cc in row:
            if isinstance(cc.value, str) and re.fullmatch(r'(1[0-2]|[1-9])月', cc.value.strip()):
                mcol[int(cc.value.strip()[:-1])] = cc.column_letter
    want = {'营业收入': '收入', '营业成本': '成本', '营业利润': '营业利润', '利润总额': '利润总额', '净利润': '净利润'}
    if len(mcol) == 12:
        for k, key in want.items():
            if k in lab:
                for m in range(1, 13):
                    chk(f'利润表 {k} {m}月', is_[f'{mcol[m]}{lab[k]}'].value, md.pl(m, m)[key], 0.02)
            else:
                bad.append(f'利润表：找不到「{k}」这一行')
    else:
        bad.append(f'利润表：找不到 1～12 月的表头（找到 {sorted(mcol)}）')
    # 资产负债表平衡（表里自己的核对格）
    bsck = wb[SH_BS][BS_CHECK.split('!')[1].replace('$', '')].value
    chk('资产负债表 BS_CHECK（应为 0）', bsck, 0)
    # 生产成本余额 = 累计发生 − 累计转出（12 月末）
    wip = sum(md.pub[(c, m)] + md.direct_m[(c, m)] for c in COMPS for m in range(1, 13)) - sum(md.out_c.values())
    d, c = md.mov('4001')
    chk('生产成本 4001 年末余额', d - c, wip, 0.05)
    print(f'核对 {n_ok + len(bad)} 项：对上 {n_ok}，不对 {len(bad)}')
    for x in bad[:60]:
        print('  ✗', x)
    return md, bad


if __name__ == '__main__':
    compare(sys.argv[1])
