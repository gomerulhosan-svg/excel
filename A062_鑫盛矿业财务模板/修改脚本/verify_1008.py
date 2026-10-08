# -*- coding: utf-8 -*-
"""10/08 这一轮的独立核对：不看公式，只拿 LibreOffice 算出来的数（../_calc/A062_算数.xlsx）
   跟你 10/05 那本里 Excel/WPS 存的数（../参考/1008上传/）比，再用 Python 从流水原始数据把该有的凭证重算一遍。

   ① 6 笔溢价款 / 退回：对方科目 122106，现金流量 6 / 2
   ② 按天合并：每笔流水的凭证号＝同一天、同一账户、同一收付方向里第一笔；分录逐条（条号、科目、借、贷）跟 Python 重算的一样
   ③ 每张凭证借贷平衡；条号 1..N 连续；凭证索引的条数、页数对得上
   ④ 结转损益：每个月每个损益科目，结转分录＝本月发生额取反；损益科目每个月末余额 0；本年利润、未分配利润对得上
   ⑤ 利润表（报表月）本月 / 本年每一行跟改之前一样；资产负债表只有其他应收款、无形资产两行变了，变动额相等、总计不变
   ⑥ 凭证汇总＝本月凭证按日期排好；数据校验没有 ✗；打印设置 A4、一页两张
跑法：python3 verify_1008.py
"""
import os
import sys
from collections import defaultdict

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, '参考', '1008上传', 'A062_用户10.05版.xlsx')
CALC = os.path.join(ROOT, '_calc', 'A062_算数.xlsx')
OUT = os.path.join(ROOT, 'A062_鑫盛矿业财务模板.xlsx')

fails = []
n_ok = 0


def ok(cond, msg):
    global n_ok
    if cond:
        n_ok += 1
    else:
        fails.append(msg)


def r2(x):
    return round(float(x or 0), 2)


def main():
    print('读取……')
    new = openpyxl.load_workbook(CALC, data_only=True)
    old = openpyxl.load_workbook(SRC, data_only=True)
    cs = new['现金流水总表']

    # ---------- ① 保证金 ----------
    for r, sign in ((6, -1), (7, -1), (8, -1), (61, -1), (1631, 1), (2239, 1)):
        ok(str(cs[f'V{r}'].value) == '122106', f'V{r} 不是 122106')
        ok(str(cs[f'W{r}'].value).startswith('122106 其他应收款-保证金'), f'W{r}={cs[f"W{r}"].value}')
        ok(cs[f'AF{r}'].value == (6 if sign < 0 else 2), f'AF{r}={cs[f"AF{r}"].value}')
        ok(cs[f'AE{r}'].value == 1, f'AE{r} 无效')

    # ---------- ② 合并分组 ----------
    rows = {}
    for r in range(6, 3966):
        v = [cs.cell(r, c).value for c in range(1, 41)]
        if v[30] != 1:
            ok(v[34] in (None, ''), f'流水 {r} 无效行却有凭证ID')
            continue
        rows[r] = dict(A=v[0], B=v[1], O=float(v[14]), X=v[23], Z=str(v[25]), AB=str(v[27]),
                       AC=float(v[28]), AD=float(v[29]), AI=v[34])
    first = {}
    for r, x in rows.items():
        key = ('X', x['X']) if x['X'] not in (None, '') else (x['A'].toordinal(), x['B'], x['O'] > 0)
        x['key'] = key
        first.setdefault(key, r)
    groups = defaultdict(list)
    for r, x in rows.items():
        f = first[x['key']]
        ok(x['AI'] == 100000 + f - 5, f'流水 {r} 凭证ID {x["AI"]}，该是 {100000 + f - 5}')
        groups[f].append(r)
    exp = {}
    for f, rs in groups.items():
        rs.sort()
        vid = 100000 + f - 5
        xs = [rows[r] for r in rs]
        lines = []
        if xs[0]['X'] not in (None, ''):
            for g, x in enumerate(xs, 1):
                amt = x['AC'] + x['AD']
                d1 = x['Z'] if x['AC'] > 0 else x['AB']
                d2 = x['AB'] if x['AC'] > 0 else x['Z']
                lines += [(2 * g - 1, d1, r2(amt), 0.0), (2 * g, d2, 0.0, r2(amt))]
        elif xs[0]['O'] > 0:
            lines.append((1, xs[0]['Z'], r2(sum(x['AC'] for x in xs)), 0.0))
            lines += [(g + 1, x['AB'], 0.0, r2(x['AC'])) for g, x in enumerate(xs, 1)]
        else:
            lines += [(g, x['AB'], r2(x['AD']), 0.0) for g, x in enumerate(xs, 1)]
            lines.append((len(xs) + 1, xs[0]['Z'], 0.0, r2(sum(x['AD'] for x in xs))))
        exp[vid] = lines
    print(f'   流水 {len(rows)} 笔 → 凭证 {len(groups)} 张（改之前一笔一张）；最大一张 {max(len(v) for v in groups.values())} 笔')

    # ---------- 记账分录 ----------
    je = new['记账分录']
    ent = []
    for r in range(4, 18022):
        v = [je.cell(r, c).value for c in range(1, 18)]
        if v[15] == 1:
            ent.append(dict(row=r, A=v[0], B=v[1], C=v[2], D=str(v[3]), F=float(v[5] or 0), G=float(v[6] or 0),
                            J=v[9], L=v[11], M=v[12]))
    by = defaultdict(list)
    for e in ent:
        by[e['L']].append(e)
    got = {vid: sorted((e['M'], e['D'], r2(e['F']), r2(e['G'])) for e in by[vid]) for vid in by if vid < 200000}
    ok(set(got) == set(exp), f'流水凭证 ID 集合不一致：多 {sorted(set(got) - set(exp))[:5]} 少 {sorted(set(exp) - set(got))[:5]}')
    bad = [vid for vid in exp if sorted(exp[vid]) != got.get(vid)]
    ok(not bad, f'流水凭证分录不对 {len(bad)} 张，如 {bad[:3]}：该 {exp[bad[0]] if bad else ""} 有 {got.get(bad[0]) if bad else ""}')
    for vid, es in by.items():
        ok(abs(sum(e['F'] for e in es) - sum(e['G'] for e in es)) < 0.005, f'凭证 {vid} 借贷不平')
        ms = sorted(e['M'] for e in es)
        ok(ms == list(range(1, len(ms) + 1)), f'凭证 {vid} 条号不连续 {ms[:10]}')
    ix = new['_凭证索引']
    idx = {}
    for r in range(4, 6899):
        v = [ix.cell(r, c).value for c in range(1, 13)]
        if v[2] == 1:
            idx[v[0]] = dict(row=r, src=v[1], D=v[3], E=v[4], F=v[5], G=v[6], H=v[7], I=v[8], J=v[9], L=v[11])
    ok(set(idx) == set(by), f'索引里的凭证跟分录对不上：多 {sorted(set(idx) - set(by))[:5]} 少 {sorted(set(by) - set(idx))[:5]}')
    for vid, d in idx.items():
        ok(d['F'] == len(by.get(vid, [])), f'索引 {vid} 条数 {d["F"]} ≠ {len(by.get(vid, []))}')
        ok(d['G'] == max(1, -(-d['F'] // 8)), f'索引 {vid} 页数')
    # 各来源合计跟改之前一样（只多了结转损益、年末结转）
    oje = old['记账分录']
    tot_old, tot_new = defaultdict(float), defaultdict(float)
    for r in range(4, 16064):
        if oje.cell(r, 16).value == 1:
            tot_old[oje.cell(r, 10).value] += float(oje.cell(r, 6).value or 0)
    for e in ent:
        tot_new[e['J']] += e['F']
    for k in tot_old:
        if k == '流水':
            continue
        ok(abs(tot_old[k] - tot_new[k]) < 0.01, f'来源 {k} 借方合计变了 {tot_old[k]} → {tot_new[k]}')
    print('   各来源借方合计：', {k: round(v / 1e6, 1) for k, v in tot_new.items()}, '（百万苏姆）')
    # 流水：对方那一边每笔一行，本账户那一边并成一行——按科目汇总的借、贷不变
    def by_code(wb_rows):
        d = defaultdict(float)
        for code, f, g in wb_rows:
            d[code] += f - g
        return {k: round(v, 2) for k, v in d.items() if abs(v) > 0.004}
    o_flow = by_code((str(oje.cell(r, 4).value), float(oje.cell(r, 6).value or 0), float(oje.cell(r, 7).value or 0))
                     for r in range(4, 8004) if oje.cell(r, 16).value == 1)
    n_flow = by_code((e['D'], e['F'], e['G']) for e in ent if e['J'] == '流水')
    diff = {k: (o_flow.get(k, 0), n_flow.get(k, 0)) for k in set(o_flow) | set(n_flow) if abs(o_flow.get(k, 0) - n_flow.get(k, 0)) > 0.01}
    want = {'170101', '122106'}
    ok(set(diff) <= want, f'流水按科目的净额变了：{ {k: v for k, v in diff.items() if k not in want} }')
    print('   流水按科目净额的变化（只该是采矿权 → 保证金）：', diff)

    # ---------- ④ 结转损益 ----------
    pl_net = defaultdict(float)
    close = defaultdict(float)
    for e in ent:
        if not e['D'].startswith('5'):
            continue
        ym = (e['A'].year, e['A'].month)
        if e['J'] == '结转损益':
            close[(e['D'], ym)] += e['F'] - e['G']
        else:
            pl_net[(e['D'], ym)] += e['F'] - e['G']
    keys = {k for k, v in pl_net.items() if abs(v) > 0.004} | {k for k, v in close.items() if abs(v) > 0.004}
    badc = [k for k in keys if abs(pl_net[k] + close[k]) > 0.011]
    ok(not badc, f'结转损益不对 {len(badc)} 处，如 {[(k, pl_net[k], close[k]) for k in sorted(badc)[:3]]}')
    months = sorted({k[1] for k in keys})
    ytd = defaultdict(float)
    bal = defaultdict(float)
    by_month = defaultdict(list)
    for e in ent:
        by_month[(e['A'].year, e['A'].month)].append(e)
    for ym in sorted(by_month):
        for e in by_month[ym]:
            bal[e['D']] += e['F'] - e['G']
        left = {c: round(v, 2) for c, v in bal.items() if c.startswith('5') and abs(v) > 0.011}
        ok(not left, f'{ym} 月末损益科目没结平：{list(left.items())[:3]}')
        prof = -sum(pl_net[(c, y)] for (c, y) in list(pl_net) if y == ym)
        ytd[ym[0]] += prof
        if ym[1] == 12:
            ok(abs(bal.get('3103', 0)) < 0.011, f'{ym[0]} 年末本年利润没转完：{bal.get("3103")}')
        else:
            ok(abs(-bal.get('3103', 0) - ytd[ym[0]]) < 0.02, f'{ym} 本年利润 {-bal.get("3103", 0)} ≠ 本年累计利润 {ytd[ym[0]]}')
    done_years = sum(v for y, v in ytd.items() if (y, 12) in by_month)
    ok(abs(-bal.get('310415', 0) - done_years) < 0.02, f'未分配利润 {-bal.get("310415", 0)} ≠ 已结年度利润 {done_years}')
    print(f'   结转损益：{len(months)} 个月、{len(keys)} 个科目×月；各年利润 { {y: round(v / 1e6, 1) for y, v in ytd.items()} } 百万；'
          f'未分配利润（已结年度）{-bal.get("310415", 0) / 1e6:.1f} 百万')
    nclose = sum(1 for v in idx.values() if v['src'] == '结转损益')
    nye = sum(1 for v in idx.values() if v['src'] == '年末结转')
    ok(nclose == len(months), f'结转损益凭证 {nclose} 张，有损益的月份 {len(months)} 个')
    want_ye = sum(1 for y, v in ytd.items() if abs(v) > 0.01)
    ok(nye == want_ye, f'年末结转凭证 {nye} 张，有利润的年度 {want_ye} 个')

    # ---------- ⑤ 报表 ----------
    for sh, cols in (('利润表', 'CD'),):
        for r in range(5, 39):
            for c in cols:
                a, b = old[sh][f'{c}{r}'].value, new[sh][f'{c}{r}'].value
                if isinstance(a, (int, float)) or isinstance(b, (int, float)):
                    ok(abs(float(a or 0) - float(b or 0)) < 0.011, f'{sh}!{c}{r} {a} → {b}')
    ok(str(new['利润表']['A39'].value).startswith('√'), f'利润表核对：{new["利润表"]["A39"].value}')
    bs_diff = {}
    for r in range(6, 38):
        for c in 'CDGH':
            a, b = old['资产负债表'][f'{c}{r}'].value, new['资产负债表'][f'{c}{r}'].value
            if isinstance(a, (int, float)) or isinstance(b, (int, float)):
                if abs(float(a or 0) - float(b or 0)) > 0.011:
                    bs_diff[f'{c}{r} {new["资产负债表"][("A" if c in "CD" else "E") + str(r)].value}'] = (a, b)
    print('   资产负债表变化：', bs_diff)
    names = {k.split(' ', 1)[1] for k in bs_diff}
    ok(names <= {'其他应收款', '无形资产（采矿权等）', '流动资产合计', '非流动资产合计'}, f'资产负债表别的行也变了：{bs_diff}')
    for col in 'CD':                                  # 期末、年初：其他应收款增加的＝无形资产减少的，资产总计不变
        d = [float(b or 0) - float(a or 0) for k, (a, b) in bs_diff.items()
             if k.startswith(col) and ('其他应收款' in k or '无形资产' in k)]
        ok(abs(sum(d)) < 0.02, f'{col} 列其他应收款、无形资产变动不相抵：{d}')
    ok(str(new['资产负债表']['C39'].value).startswith('√'), f'资产负债表：{new["资产负债表"]["C39"].value}')
    for c in ('B8', 'D8', 'F8', 'H8'):
        a, b = old['首页'][c].value, new['首页'][c].value
        ok(abs(float(a or 0) - float(b or 0)) < 0.011, f'首页 {c} {a} → {b}')
    cf = {}
    for r in range(6, 31):
        for c in 'CD':
            a, b = old['现金流量表'][f'{c}{r}'].value, new['现金流量表'][f'{c}{r}'].value
            if isinstance(a, (int, float)) and abs(float(a) - float(b or 0)) > 0.011:
                cf[f'{c}{r} {old["现金流量表"][f"A{r}"].value}'] = (round(a), round(b))
    print('   现金流量表变化：', cf)
    for k in cf:
        ok(any(s in k for s in ('其他与经营', '经营活动产生', '处置固定资产', '购建固定资产', '投资活动产生')), f'现金流量表 {k} 不该变')
    ok(str(new['现金流量表']['A32'].value).startswith('√'), f'现金流量表：{new["现金流量表"]["A32"].value}')

    # ---------- ⑥ 凭证汇总、校验、打印 ----------
    sm = new['凭证汇总']
    ym = new['首页']['C4'].value * 100 + new['首页']['F4'].value
    mine = sorted((v for v in idx.values() if v['H'] == ym), key=lambda v: v['E'])
    lst = []
    for r in range(6, 456):
        if sm[f'A{r}'].value not in (None, ''):
            lst.append((sm[f'A{r}'].value, sm[f'M{r}'].value, sm[f'I{r}'].value, sm[f'J{r}'].value, sm[f'O{r}'].value, sm[f'K{r}'].value))
    ok([x[1] for x in lst] == [ix.cell(v['row'], 1).value for v in mine], '凭证汇总的顺序跟按日期排的不一样')
    ok([x[0] for x in lst] == list(range(1, len(lst) + 1)), '凭证号不连续')
    pg = 1
    for no, vid, pages, start, a4, src in lst:
        ok(start == pg and a4 == -(-start // 2), f'记-{no} 起始页 {start}/{a4}，该是 {pg}/{-(-pg // 2)}')
        pg += pages
    ok(sm['N3'].value == -(-(pg - 1) // 2), f'A4 张数 {sm["N3"].value}')
    print(f'   {ym}：凭证 {len(lst)} 张（改之前 {old["凭证汇总"]["D3"].value} 张），凭证页 {pg - 1}，A4 {sm["N3"].value} 张；'
          f'最后几张：{[(x[0], x[5]) for x in lst[-3:]]}')
    pz = new['记账凭证']
    for blk in range(min(6, pg - 1)):
        b0 = 14 * blk + 1
        no = pz[f'J{b0}'].value
        vid = pz[f'M{b0}'].value
        lp = pz[f'L{b0}'].value
        es = sorted(by[vid], key=lambda e: e['M'])[(lp - 1) * 8: lp * 8]
        for i, e in enumerate(es):
            rr = b0 + 3 + i
            ok(r2(pz[f'E{rr}'].value) == r2(e['F']) and r2(pz[f'F{rr}'].value) == r2(e['G']), f'打印第 {blk + 1} 页第 {i + 1} 行金额不对')
            ok(str(pz[f'D{rr}'].value).endswith(f'（{e["D"]}）'), f'打印第 {blk + 1} 页第 {i + 1} 行科目 {pz[f"D{rr}"].value}')
    ck = new['数据校验']
    xs = [(ck[f'B{r}'].value, ck[f'C{r}'].value) for r in range(5, 45) if str(ck[f'C{r}'].value).startswith('✗')]
    ok(not xs, f'数据校验有 ✗：{xs}')
    print('   数据校验：', new['数据校验']['C3'].value, '| 新加：', ck['C43'].value, '/', ck['C44'].value)

    out = openpyxl.load_workbook(OUT)
    ws = out['记账凭证']
    ps = ws.page_setup
    sc = int(ps.scale) / 100
    ok(int(ps.paperSize) == 9 and ps.orientation == 'portrait' and sc == 0.75, f'纸张 {ps.paperSize} {ps.orientation} {ps.scale}')
    brk = [b.id for b in ws.row_breaks.brk]
    ok(brk == [28 * i for i in range(1, 250)], f'分页 {brk[:5]}…')
    h = sum(ws.row_dimensions[r].height for r in range(1, 29))
    usable = 842 - (ws.page_margins.top + ws.page_margins.bottom) * 72
    ok(h * sc < usable - 15, f'两页凭证 {h * sc:.0f} 磅放不下 A4（{usable:.0f}）')
    px = sum(ws.column_dimensions[c].width * 8 + 5 for c in 'BCDEFG')      # 中文 Excel：宋体 11 每字宽 8 像素
    wide = 595 - (ws.page_margins.left + ws.page_margins.right) * 72
    ok(px * 0.75 * sc < wide - 20, f'宽 {px * 0.75 * sc:.0f} 磅放不下 A4（{wide:.0f}）')
    print(f'   打印：A4 纵向 {ps.scale}%，两页凭证共 {h * sc:.0f} 磅（纸上能用 {usable:.0f} 磅），宽 {px * 0.75 * sc:.0f} 磅（能用 {wide:.0f}），'
          f'中间空 {ws.row_dimensions[14].height * sc:.0f} 磅 ≈ {ws.row_dimensions[14].height * sc / 72 * 2.54:.1f} 厘米')

    print(f'\n核对 {n_ok + len(fails)} 项：{n_ok} 项对，{len(fails)} 项不对')
    for f in fails[:40]:
        print('  ✗', f)
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main())
