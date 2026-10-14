# -*- coding: utf-8 -*-
"""WPS 存盘不拆共享公式（wps）：生成完、存盘前做最后一遍；另有成品「WPS 体检」（make_all 调用）。

WPS 12.1 存 xlsx 的规矩（从 8 本用户用 WPS 存回的文件归纳，A074 评审 review_perf_gain §1.2）：
  同一列往下拖的公式（共享公式），只要直接写了「别的表!格子」或往下撑的区域（$A$5:A5、A$5:A5 这种一端绝对行、一端相对行），
  WPS 就把这一列拆成一格一条全文写进文件；只用名称（名称指向别的表也行）、只引用本表的照样共享。
  A075 原件就是这样被存成 85 MB XML 的（我们发过去的 956 组跨表共享公式全被拆开）。

这里做的（只动「会被压成共享公式」的那些列，也就是同一列连续 ≥4 格、R1C1 写法一样的；单个格子的公式 WPS 本来就全文写，不动，留着给人看）：
  1. 本表的格子写成了「本表名!A1」的，去掉表名（WPS 也当跨表）；
  2. 绝对的跨表引用（科目表!$A$5:$A$454、首页!$C$4…）换成名称（已有同一区域的名称就用它，没有就新建「区_表名_A5_A454」，隐藏）；
  3. 相对行的跨表单格（月末结转!C6 取 原矿产销存!D6 这种）换成 INDEX(名称,ROW()-常数)，名称＝这一串格子对着的那一段；
  4. 剩下的（跨表的相对区域、往下撑的区域）不认识就报错——这些要在各表的生成脚本里改写法（common.SEG 分段名称）。
SEG 分段名称（common.SEG_NAMES）也在这里写进工作簿（隐藏）。
"""
import re
import zipfile
from collections import defaultdict

from lxml import etree
from openpyxl.workbook.defined_name import DefinedName

import common
from share import normalize, sheet_parts, NS

MIN_RUN = 4              # 跟 share.share(min_run) 一样
SHEET_REF = re.compile(r"(?P<sh>'(?:[^']|'')+'|[A-Za-z_一-鿿][\w一-鿿.]*)!"
                       r"(?P<a>\$?[A-Z]{1,3}\$?\d+)(?::(?P<b>\$?[A-Z]{1,3}\$?\d+))?(?![\w(])")
CELL = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)$')
MIXED = re.compile(r'(?<![\w.$])\$?[A-Z]{1,3}(\$?)\d+:\$?[A-Z]{1,3}(\$?)\d+(?![\w(])')
# 白名单：这些格子 WPS 拆了也没多少（每段写法本来就不同、压不成共享的除外）
WHITE = set()


def _unq(sh):
    return sh[1:-1].replace("''", "'") if sh.startswith("'") else sh


def _parts(f):
    """公式按双引号切开：偶数段在字符串外"""
    return f.split('"')


def wps_split(f):
    """公式（文字）→ (直接跨表, 往下撑的区域)"""
    out = ''.join(_parts(f)[0::2])
    mixed = any(bool(m.group(1)) != bool(m.group(2)) for m in MIXED.finditer(out))
    return '!' in out, mixed


def xml_bytes(f):
    return len(f.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').encode())


def _norm_ref(text):
    """名称引用 → 比较用的标准写法：表名去引号，$ 去掉"""
    m = re.match(r"^(?P<sh>'(?:[^']|'')+'|[^!]+)!(?P<r>.+)$", text.strip())
    if not m:
        return None
    return _unq(m.group('sh')) + '!' + m.group('r').replace('$', '')


def _safe(s):
    return re.sub(r'[^\w一-鿿]', '', s)


def apply(wb):
    log = []
    # 0) 分段名称写进工作簿
    for nm, ref in common.SEG_NAMES.items():
        assert nm not in wb.defined_names, nm
        wb.defined_names[nm] = DefinedName(nm, attr_text=ref, hidden=True)
    log.append(f'分段名称 {len(common.SEG_NAMES)} 个（隐藏）')
    have = {}
    for nm, dn in wb.defined_names.items():
        k = _norm_ref(dn.attr_text or '')
        if k and re.fullmatch(r"[^!]+![A-Z]+\d+(:[A-Z]+\d+)?", k):
            have.setdefault(k, nm)
    new_names = {}
    titles = set(wb.sheetnames)

    def abs_name(sh, a, b):
        key = f'{sh}!{a.replace("$", "")}' + (f':{b.replace("$", "")}' if b else '')
        if key in have:
            return have[key]
        nm = f'区_{_safe(sh)}_{a.replace("$", "")}' + (f'_{b.replace("$", "")}' if b else '')
        ref = f"{common.q(sh)}!{_absref(a)}" + (f':{_absref(b)}' if b else '')
        if nm in new_names:
            assert new_names[nm] == ref
        else:
            assert nm not in wb.defined_names, nm
            new_names[nm] = ref
        have[key] = nm
        return nm

    rel_names = {}
    n_self = n_abs = n_rel = 0
    for ws in wb.worksheets:
        bycol = defaultdict(list)
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith('='):
                    bycol[c.column_letter].append((c.row, c))
        for col, cells in bycol.items():
            cells.sort(key=lambda x: x[0])
            i = 0
            while i < len(cells):
                r0, c0 = cells[i]
                key = normalize(c0.value, r0)
                j = i + 1
                while j < len(cells) and cells[j][0] == cells[j - 1][0] + 1 and normalize(cells[j][1].value, cells[j][0]) == key:
                    j += 1
                if j - i >= MIN_RUN:
                    run = cells[i:j]
                    xs, mx = wps_split(c0.value)
                    if xs or mx:
                        fixed = [_fix(ws.title, c.value, r, r0, titles, abs_name, rel_names, j - i) for r, c in run]
                        for (r, c), (f, a, b, s) in zip(run, fixed):
                            c.value = f
                            n_abs += a
                            n_rel += b
                            n_self += s
                        xs2, mx2 = wps_split(run[0][1].value)
                        if (xs2 or mx2) and (ws.title, col) not in WHITE:
                            raise AssertionError(f'WPS：{ws.title}!{col}{r0}:{col}{cells[j - 1][0]} 还会被拆（'
                                                 f'{"跨表" if xs2 else "往下撑的区域"}）：{run[0][1].value[:200]}')
                i = j
    for nm, ref in new_names.items():
        wb.defined_names[nm] = DefinedName(nm, attr_text=ref, hidden=True)
    log.append(f'成列的公式里：绝对跨表引用 {n_abs:,} 处换成名称（新建 {len(new_names)} 个），相对行跨表单格 {n_rel:,} 处换成 '
               f'INDEX(名称,ROW()-常数)（用到 {len(rel_names)} 个名称），本表名去掉 {n_self:,} 处')
    return log


def _absref(a):
    m = CELL.match(a)
    return f'${m.group(2)}${m.group(4)}'


def _fix(title, f, r, r0, titles, abs_name, rel_names, nrun):
    """改一条公式 → (新公式, 绝对换名处数, 相对换 INDEX 处数, 去本表名处数)"""
    parts = _parts(f)
    na = nr = ns = 0
    for p in range(0, len(parts), 2):
        s = parts[p]

        def sub(m):
            nonlocal na, nr, ns
            sh = _unq(m.group('sh'))
            a, b = m.group('a'), m.group('b')
            if sh not in titles:
                return m.group(0)
            if sh == title:                 # 本表：去掉表名
                ns += 1
                return a + (f':{b}' if b else '')
            ma, mb = CELL.match(a), CELL.match(b) if b else None
            row_abs = ma.group(3) == '$' and (mb is None or mb.group(3) == '$')
            if row_abs:
                na += 1
                return abs_name(sh, a, b)
            if b is None:                  # 相对行单格：这一串格子对着别的表的一段（行差固定）
                d = int(ma.group(4)) - r   # 目标行＝本行＋d
                t0, t1 = r0 + d, r0 + d + nrun - 1
                col = ma.group(2)
                nm = abs_name(sh, f'${col}${t0}', f'${col}${t1}')
                rel_names[nm] = 1
                nr += 1
                return f'INDEX({nm},ROW()-{r0 - 1})'
            raise AssertionError(f'WPS：{title}!第 {r} 行 跨表的相对区域 {m.group(0)} 没有自动改法（在生成脚本里改成段名称）：{f[:200]}')
        parts[p] = SHEET_REF.sub(sub, s)
    return '"'.join(parts), na, nr, ns


# ─────────────────────────── 成品体检（已压共享公式的 xlsx） ───────────────────────────
def estimate(path):
    """→ {表名: (XML 字节, WPS 存后估计字节, 会被拆的子格数, [(格, 原因, 公式)…])}。
    估法同 A074 评审 wps_est：主公式会被 WPS 拆的共享组，每个子格按主公式全文写（每格多 xml_bytes−20 字节）"""
    z = zipfile.ZipFile(path)
    out = {}
    for part, name in sheet_parts(z).items():
        size = z.getinfo(part).file_size
        master, kids = {}, defaultdict(int)
        for _ev, el in etree.iterparse(z.open(part), tag=f'{{{NS}}}f'):
            if el.get('t') == 'shared':
                if el.get('ref'):
                    master[el.get('si')] = (el.get('ref'), el.text or '')
                else:
                    kids[el.get('si')] += 1
            el.clear()
        add = n = 0
        bad = []
        for si, k in kids.items():
            ref, f = master.get(si, ('', ''))
            xs, mx = wps_split(f)
            if xs or mx:
                add += k * max(0, xml_bytes(f) - 20)
                n += k
                bad.append((ref, '跨表' if xs else '往下撑', f[:160]))
        out[name] = (size, size + add, n, bad)
    z.close()
    return out


def report(path, warn_mb=1.0, log=print):
    """打印「共享保住」和「WPS 存一次以后估计」两个 XML 数；WPS 后多出来超过 warn_mb 就算不过（返回 False）"""
    est = estimate(path)
    tot = sum(v[0] for v in est.values())
    totw = sum(v[1] for v in est.values())
    n = sum(v[2] for v in est.values())
    log(f'   WPS 体检：sheet XML 共享保住 {tot / 1e6:.1f} MB；WPS 存一次以后估计 {totw / 1e6:.1f} MB（会被拆的子格 {n:,} 个）')
    grow = sorted(((k, v[1] - v[0], v[3]) for k, v in est.items() if v[1] > v[0]), key=lambda x: -x[1])
    for k, d, bad in grow[:8]:
        log(f'     {k} +{d / 1e6:.2f} MB：' + '；'.join(f'{a}（{b}）{c[:80]}' for a, b, c in bad[:3]))
    ok = totw - tot <= warn_mb * 1e6
    if not ok:
        log(f'   ✗ WPS 存后会多 {(totw - tot) / 1e6:.1f} MB（超过 {warn_mb} MB）：上面这些列还有「表名!」或往下撑的区域')
    return ok, tot, totw
