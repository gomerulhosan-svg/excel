# -*- coding: utf-8 -*-
"""独立核对：不看新表的公式，直接从原表的【数据录入】【门店月度收支统计】用 Python 重算一遍，跟成品里的数逐格比；
再核对原文件的每个部件都原样保留（只有 workbook / rels / Content_Types / app / styles 五个文件加了东西）。

跑法：python3 verify.py [成品] [原表]
"""
import collections
from decimal import Decimal, ROUND_HALF_UP
import datetime as dt
import os
import re
import sys
import zipfile

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', '多帐户收支登记表_加门店余额核对.xlsx')
SRC = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, '..', '参考', '原表_数据录入_2026.9.27.xlsx')
SH, AUX = '门店余额核对', '门店余额核对辅助'
NONSTORE = ['公司', '配送', '预收', '收款', '其他收款', '股东借款', '内部转款', '房租及其他', '12月门店']
BAD = []
N_OK = [0]


def chk(name, got, want, tol=0.005):
    if isinstance(want, (int, float)) and not isinstance(want, bool):
        ok = isinstance(got, (int, float)) and abs(got - want) <= tol
    else:
        ok = (got if got is not None else '') == (want if want is not None else '')
    if ok:
        N_OK[0] += 1
    else:
        BAD.append((name, got, want))


def r2(x):
    return float(Decimal(repr(float(x))).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def serial(d):
    return (d - dt.datetime(1899, 12, 30)).days


# ---------------- 原表数据 ----------------
src = load_workbook(SRC, data_only=True, read_only=True)
de = [list(r) for r in src['数据录入'].iter_rows(min_row=1, max_row=30191, max_col=15, values_only=True)]
row = lambda r: de[r - 1] if r - 1 < len(de) else [None] * 15
accts = [row(r)[2] for r in range(5, 19)]
bank0 = sum(num(row(r)[7]) for r in range(5, 19))
book_y = row(5)[1].year
ms = [list(r) for r in src['门店月度收支统计'].iter_rows(min_row=1, max_row=1004, max_col=11, values_only=True)]
ms_year = ms[1][1]


def eff(r):
    b, L = r[1], r[11]
    if isinstance(b, dt.datetime):
        return serial(b)
    if isinstance(b, (int, float)) and not isinstance(b, bool) and 36526 <= b <= 2958465:
        return int(b)
    if isinstance(L, (int, float)) and L >= 190001 and 1 <= L % 100 <= 12:
        y, m = int(L // 100), int(L % 100)
        nxt = dt.datetime(y + (m == 12), m % 12 + 1, 1)
        return serial(nxt - dt.timedelta(days=1))
    return 0


flows = []
for r in range(19, 30192):
    x = row(r)
    flows.append(dict(r=r, eff=eff(x), bank=1 if (x[2] not in (None, '') and x[2] in accts) else 0,
                      store=0 if (x[8] in (None, '') or x[8] in NONSTORE) else 1,
                      F=num(x[5]), G=num(x[6]), C=x[2], E=x[4], I=x[8], L=x[11],
                      hasF=x[5] not in (None, ''), hasG=x[6] not in (None, '')))

# 门店名单（门店月度收支统计 B 列，不重复，去掉不算门店的名字）
names, seen = [], set()
for r in range(5, 1005):
    b = ms[r - 1][1] if r - 1 < len(ms) else None
    if b in (None, '') or b in NONSTORE or b in seen:
        continue
    seen.add(b)
    names.append(b)
msG, msD, msE = (collections.defaultdict(float) for _ in range(3))
for r in range(5, 1005):
    x = ms[r - 1] if r - 1 < len(ms) else [None] * 11
    if x[1]:
        msG[x[1]] += num(x[6]); msD[x[1]] += num(x[3]); msE[x[1]] += num(x[4])


def agg(cut, pred):
    f = g = 0.0
    for x in flows:
        if x['eff'] <= cut and pred(x):
            f += x['F']; g += x['G']
    return f, g


def store_tables(cut):
    out = []
    for n in names:
        K = sum(x['F'] - x['G'] for x in flows if x['I'] == n and isinstance(x['L'], (int, float))
                and ms_year * 100 + 1 <= x['L'] <= ms_year * 100 + 12)
        Lall = sum(x['F'] - x['G'] for x in flows if x['I'] == n)
        f, g = agg(cut, lambda x, n=n: x['I'] == n)
        adj = r2(msD[n] - msE[n] - K)
        out.append(dict(name=n, open=msG[n], adj=adj, F=f, G=g, end=msG[n] + adj + f - g, miss=r2(Lall - K)))
    return out


# ---------------- 成品 ----------------
wb = load_workbook(OUT, data_only=True)
ws, wa = wb[SH], wb[AUX]
v = lambda ref: ws[ref].value


def run(cut_label, cut):
    st = store_tables(cut)
    open_sum = sum(s['open'] for s in st)
    adj_sum = sum(s['adj'] for s in st)
    sf, sg = agg(cut, lambda x: x['store'] == 1)
    nf, ng = agg(cut, lambda x: x['store'] == 0)
    bf, bg = agg(cut, lambda x: x['bank'] == 1)
    auto = bank0 - open_sum - adj_sum
    store_bal = open_sum + adj_sum + sf - sg
    should = r2(store_bal + nf - ng + auto)
    bank = r2(bank0 + bf - bg)
    p = f'[{cut_label}] '
    chk(p + '① 门店余额', v('C9'), store_bal)
    chk(p + '① 收款', v('C10'), nf)
    chk(p + '① 付款', v('C11'), ng)
    chk(p + '① 公司年初结余', v('C12'), auto)
    chk(p + '① 应存', v('C13'), should)
    chk(p + '① 实存', v('C14'), bank)
    chk(p + '① 差额', v('C15'), r2(bank - should))
    # ③ 没走银行的记录
    tiao = agg(cut, lambda x: x['C'] in (None, '') and '调拨' in str(x['E'] or ''))
    blank = agg(cut, lambda x: x['C'] in (None, ''))
    zd = agg(cut, lambda x: x['C'] == '货款直抵费用')
    nbn = agg(cut, lambda x: x['bank'] == 0 and x['C'] not in (None, ''))
    exp = [tiao, (blank[0] - tiao[0], blank[1] - tiao[1]), zd, (nbn[0] - zd[0], nbn[1] - zd[1])]
    for i, (f, g) in enumerate(exp):
        chk(p + f'③ 第{i + 1}行 收入', v(f'D{35 + i}'), f)
        chk(p + f'③ 第{i + 1}行 支出', v(f'E{35 + i}'), g)
    nb_net = sum(f - g for f, g in exp)
    chk(p + '③ 小计净额', v('F39'), nb_net)
    chk(p + '③ 合计＝差额', v('F41'), r2(-nb_net))
    chk(p + '差额＝－没走银行的净额（恒等式）', r2(bank - should), r2(-nb_net))
    cnt = lambda pred: sum((x['hasF']) + (x['hasG']) for x in flows if x['eff'] <= cut and pred(x))
    chk(p + '③ 调拨笔数', v('C35'), cnt(lambda x: x['C'] in (None, '') and '调拨' in str(x['E'] or '')))
    chk(p + '③ 货款直抵费用笔数', v('C37'), cnt(lambda x: x['C'] == '货款直抵费用'))
    # ④
    tot_f = tot_g = 0
    for i, n in enumerate(NONSTORE):
        f, g = agg(cut, lambda x, n=n: x['I'] == n)
        bf_, bg_ = agg(cut, lambda x, n=n: x['I'] == n and x['bank'] == 1)
        chk(p + f'④ {n} 收款', v(f'C{48 + i}'), f); chk(p + f'④ {n} 付款', v(f'D{48 + i}'), g)
        chk(p + f'④ {n} 银行收', v(f'F{48 + i}'), bf_); chk(p + f'④ {n} 银行付', v(f'G{48 + i}'), bg_)
        tot_f += f; tot_g += g
    bl = agg(cut, lambda x: x['I'] in (None, ''))
    chk(p + '④ 没填门店 收款', v('C63'), bl[0]); chk(p + '④ 没填门店 付款', v('D63'), bl[1])
    chk(p + '④ 合计 收款', v('C64'), nf); chk(p + '④ 合计 付款', v('D64'), ng)
    # ⑤
    for i, s in enumerate(st):
        rr = 68 + i
        chk(p + f'⑤ {s["name"]} 名字', v(f'B{rr}'), s['name'])
        chk(p + f'⑤ {s["name"]} 年初', v(f'C{rr}'), s['open'])
        chk(p + f'⑤ {s["name"]} 手加', v(f'D{rr}'), s['adj'])
        chk(p + f'⑤ {s["name"]} 收入', v(f'E{rr}'), s['F'])
        chk(p + f'⑤ {s["name"]} 支出', v(f'F{rr}'), s['G'])
        chk(p + f'⑤ {s["name"]} 期末', v(f'G{rr}'), s['end'])
        chk(p + f'⑤ {s["name"]} 少算', v(f'H{rr}'), s['miss'])
    chk(p + '⑤ 名单后面是空行', v(f'B{68 + len(st)}'), '')
    unk = agg(cut, lambda x: x['store'] == 1 and x['I'] not in seen)
    chk(p + '⑤ 名字不在门店月度里的 收入', v('E168'), unk[0])
    chk(p + '⑤ 合计期末', v('G169'), store_bal)
    # ⑥
    for i, a in enumerate(accts):
        rr = 173 + i
        f, g = agg(cut, lambda x, a=a: x['C'] == a)
        o = num(row(5 + i)[7])
        chk(p + f'⑥ {a} 账户', v(f'B{rr}'), a)
        chk(p + f'⑥ {a} 账上余额', v(f'F{rr}'), r2(o + f - g))
    chk(p + '⑥ 合计＝实存', v('F193'), bank)
    chk(p + '③ 合计跟①差额一致的提示', v('G41'), '✓ 跟 ① 的差额一致')
    chk(p + '自查：④⑤⑥ 合计跟 ① 一致', v('G44'), '✓ 一致')
    return dict(store=store_bal, recv=nf, pay=ng, auto=auto, should=should, bank=bank)


res = run('全部', 2958465)
OPEN_ADJ = sum(s['open'] + s['adj'] for s in store_tables(2958465))
print('全部：', {k: round(x, 2) for k, x in res.items()}, '差额', round(res['bank'] - res['should'], 2))
# ② 按月
last = max(x['eff'] for x in flows)
for k in range(1, 13):
    rr = 19 + k
    me = serial(dt.datetime(book_y + (k == 12), k % 12 + 1, 1) - dt.timedelta(days=1))
    if serial(dt.datetime(book_y, k, 1)) > last:
        chk(f'② {k}月 空着', v(f'C{rr}'), '')
        continue
    sf, sg = agg(me, lambda x: x['store'] == 1)
    nf, ng = agg(me, lambda x: x['store'] == 0)
    bf, bg = agg(me, lambda x: x['bank'] == 1)
    store = OPEN_ADJ + sf - sg
    should = r2(r2(store) + r2(nf) - r2(ng) + r2(res['auto']))
    bank = r2(bank0 + bf - bg)
    chk(f'② {k}月 门店余额', v(f'C{rr}'), r2(store))
    chk(f'② {k}月 收款', v(f'D{rr}'), r2(nf))
    chk(f'② {k}月 付款', v(f'E{rr}'), r2(ng))
    chk(f'② {k}月 应存', v(f'G{rr}'), should)
    chk(f'② {k}月 实存', v(f'H{rr}'), bank)
    chk(f'② {k}月 差额', v(f'I{rr}'), r2(bank - should))
# 辅助表逐行
for x in flows:
    r = x['r']
    chk(f'辅助 A{r}', wa[f'A{r}'].value, x['eff'], 0)
    chk(f'辅助 B{r}', wa[f'B{r}'].value, x['bank'], 0)
    chk(f'辅助 C{r}', wa[f'C{r}'].value, x['store'], 0)
chk('辅助 J3 年初银行余额', wa['J3'].value, bank0)
chk('辅助 J2 最后一笔', wa['J2'].value, last, 0)

# ---------------- 原文件部件原样保留 ----------------
zs, zo = zipfile.ZipFile(SRC), zipfile.ZipFile(OUT)
changed = {'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml', 'docProps/app.xml', 'xl/styles.xml'}
for n in zs.namelist():
    if n in changed:
        continue
    chk(f'部件原样 {n}', zo.read(n) == zs.read(n), True)
new_parts = sorted(set(zo.namelist()) - set(zs.namelist()))
chk('只多了两张表', len(new_parts), 2, 0)
wo, wsrc = zo.read('xl/workbook.xml').decode(), zs.read('xl/workbook.xml').decode()
chk('workbook.xml 只在 </sheets> 前多了两张表', re.sub(r'<sheet name="门店余额核对[^>]*/>', '', wo), wsrc)
so, ssrc = zo.read('xl/styles.xml').decode(), zs.read('xl/styles.xml').decode()
chk('选中的工作表只有一张（WPS 不会进「工作组」）', sum(zo.read(n).count(b'tabSelected="1"') for n in zo.namelist()
                                           if n.startswith('xl/worksheets/sheet')), 1, 0)

print(f'核对 {N_OK[0] + len(BAD)} 项，一致 {N_OK[0]} 项，不一致 {len(BAD)} 项')
for b in BAD[:40]:
    print('  ✗', b)
