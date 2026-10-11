# -*- coding: utf-8 -*-
"""静态环检测：Excel 把 INDEX / SUMIF / SUMIFS / MATCH / COUNTIF… 的区域参数整片都当成前驱，
LibreOffice 算得出来、0 报错，不代表 Excel 打开时不报「循环引用」。这里不算数，只看公式里的引用。

用法：cyc = find_cycles(wb)（openpyxl 的 Workbook，公式是字符串）→ 列表，空＝没环；fmt(cyc) 变成可打印的文字。
每个环：{'环': ['表!列起行:列止行', …], '例子': '表!格: =公式…', '种类': '环' 或 '自己引用自己'}

做法（正则一遍扫，几十万公式几秒）：
- 节点＝「表!列」里连续、写法一样（数字不算，字符串里的字不算）的一段公式：一整列同样的公式是一个节点，
  _参 这种一格一个写法的就是一格一个节点（所以 _参 里互相引用的不同参数不会被当成环）。
- 引用：A1、$A$1、A1:B9、整列 A:A、整行 4:4、表名!、带引号的 '表 名'!、定义名称（工作簿级、表级，展开成区域，名称套名称也展开）。
  ROW( ) / COLUMN( ) / ROWS( ) / COLUMNS( ) 括号里的引用不算前驱（Excel 不因它们报循环）。
- 边：被引用区域碰到的公式段 → 公式所在的段（同一段自己指自己的忽略）；一段里各行的引用取并集（相对引用按行滑动）。
- 强连通分量里节点数 ≥2 的就是环（报一条最短的回路＋一个例子公式）。
- 另外查「一格公式引用的区域里包含它自己」（Excel 一定报循环），按段报一次。
"""
import re
from bisect import bisect_left, bisect_right
from collections import defaultdict, deque

_DIG = re.compile(r'\d+')
_DIGC = re.compile(r'(\d+)')
_ROWFN = re.compile(r'(?<![\w.])(?:ROWS?|COLUMNS?)\(([^()]*)\)')
_REF = re.compile(
    r"(?<![\w.$'])"
    r"(?:(?:'(?P<qs>(?:[^']|'')+)'|(?P<us>[^\W\d][\w.]*))!)?"
    r"(?:(?P<c1>\$?[A-Z]{1,3})\$?(?P<r1>\d+)(?::(?P<c2>\$?[A-Z]{1,3})\$?(?P<r2>\d+))?"
    r"|(?P<k1>\$?[A-Z]{1,3}):(?P<k2>\$?[A-Z]{1,3})"
    r"|\$?(?P<w1>\d+):\$?(?P<w2>\d+)"
    r"|(?P<nm>[^\W\d][\w.]*))"
    r"(?![\w.(!])")
MAXR, MAXC = 1048576, 16384
_CI = {}


def _ci(s):
    s = s.lstrip('$')
    v = _CI.get(s)
    if v is None:
        v = 0
        for ch in s:
            v = v * 26 + ord(ch) - 64
        _CI[s] = v
    return v


def _cl(i):
    s = ''
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _mask(f):
    """字符串里的字换成空格（位置不变）；ROW()/COLUMN()/ROWS()/COLUMNS() 的参数也换成空格"""
    if '"' in f:
        p = f.split('"')
        for i in range(1, len(p), 2):
            if p[i]:
                p[i] = ' ' * len(p[i])
        f = '"'.join(p)
    if 'ROW' in f or 'COLUMN' in f:
        f = _ROWFN.sub(lambda mt: mt.group(0)[:mt.start(1) - mt.start(0)] + ' ' * len(mt.group(1)) + ')', f)
    return f


def _parse(m, f, sheet, resolve):
    """m：已 mask 的公式，f：原公式（位置一样；第几串数字按原公式数，跟同一段别的行对得上）。返回 [(表, c1, c2, r1, r2, 槽1, 槽2)]；槽＝行号是公式里第几串数字（None＝定死的，名称、整列）"""
    out = []
    ds = None
    for mt in _REF.finditer(m):
        g = mt.group
        nm = g('nm')
        qs = g('qs')
        sh = qs.replace("''", "'") if qs is not None else g('us')
        if nm is not None:
            out.extend(resolve(nm, sh or sheet, sh is not None))
            continue
        sh = sh or sheet
        if g('c1') is not None:
            if ds is None:
                ds = [x.start() for x in _DIG.finditer(f)]
            c1, r1, s1 = _ci(g('c1')), int(g('r1')), bisect_left(ds, mt.start('r1'))
            if g('c2') is not None:
                c2, r2, s2 = _ci(g('c2')), int(g('r2')), bisect_left(ds, mt.start('r2'))
            else:
                c2, r2, s2 = c1, r1, s1
            if c1 > c2:
                c1, c2 = c2, c1
            if r1 > r2:
                r1, r2, s1, s2 = r2, r1, s2, s1
            out.append((sh, c1, c2, r1, r2, s1, s2))
        elif g('k1') is not None:
            c1, c2 = sorted((_ci(g('k1')), _ci(g('k2'))))
            out.append((sh, c1, c2, 1, MAXR, None, None))
        else:
            r1, r2 = sorted((int(g('w1')), int(g('w2'))))
            out.append((sh, 1, MAXC, r1, r2, None, None))
    return out


def _name_table(wb):
    """定义名称 → 引用区域（名称套名称展开；OFFSET 之类只取写出来的格子）"""
    raw = {}
    for nm, dn in wb.defined_names.items():
        if dn.attr_text:
            raw[(None, nm.upper())] = dn.attr_text
    for ws in wb.worksheets:
        for nm, dn in getattr(ws, 'defined_names', {}).items():
            if dn.attr_text:
                raw[(ws.title, nm.upper())] = dn.attr_text
    cache = {}

    def expand(key, depth):
        if key in cache:
            return cache[key]
        cache[key] = []                                    # 名称自己套自己：当空
        txt = raw[key]
        refs = []
        if depth < 8:
            def res(nm, sh, explicit):
                return look(nm, sh if explicit else key[0], depth + 1)
            for r in _parse(_mask(txt), txt, key[0], res):
                if r[0] is not None:
                    refs.append((r[0], r[1], r[2], r[3], r[4], None, None))
        cache[key] = refs
        return refs

    def look(nm, sheet, depth=0):
        u = nm.upper()
        if (sheet, u) in raw:
            return expand((sheet, u), depth)
        if (None, u) in raw:
            return expand((None, u), depth)
        return ()

    return look


def find_cycles(wb, limit=200):
    look = _name_table(wb)
    rcache = {}

    def resolve(nm, sheet, explicit):
        k = (nm, sheet)
        v = rcache.get(k)
        if v is None:
            v = rcache[k] = look(nm, sheet)
        return v

    # ① 每列连续、写法一样的公式归成一段；每段算出引用区域的并集，顺便查「引用到自己」
    nodes = []          # (表, 列, 起行, 止行, 第一格公式)
    hulls = []          # 每段：[(表, c1, c2, r1, r2)]
    selfrefs = []
    colband = {}        # (表, 列) → (起行列表, 止行列表, 节点号列表)
    for ws in wb.worksheets:
        sheet = ws.title
        bycol = defaultdict(list)
        for (r, c), cell in ws._cells.items():
            if cell.data_type == 'f':
                v = cell._value
                if isinstance(v, str) and len(v) > 1 and v[0] == '=':
                    bycol[c].append((r, v))
        for c, lst in bycol.items():
            lst.sort()
            starts, ends, ids = [], [], []
            band = []

            def flush():
                if not band:
                    return
                nid = len(nodes)
                r0, f0, _d0 = band[0]
                rN = band[-1][0]
                nodes.append((sheet, c, r0, rN, f0))
                hull = []
                for (sh, c1, c2, r1, r2, s1, s2) in _parse(_mask(f0), f0, sheet, resolve):
                    if s1 is None or len(band) == 1:
                        hull.append((sh, c1, c2, r1, r2))
                        if sh == sheet and c1 <= c <= c2 and r1 <= rN and r0 <= r2:
                            hit = next((rr for rr, _f, _d in band if r1 <= rr <= r2), None)
                            if hit is not None:
                                selfrefs.append((nid, hit))
                        continue
                    a = [int(d[s1]) for _r, _f, d in band]          # 这一段每行这个引用的行号（相对引用按行滑动）
                    b = a if s2 == s1 else [int(d[s2]) for _r, _f, d in band]
                    lo, hi = min(min(a), min(b)), max(max(a), max(b))
                    hull.append((sh, c1, c2, lo, hi))
                    if sh == sheet and c1 <= c <= c2 and lo <= rN and r0 <= hi:
                        for i, (rr, _f, _d) in enumerate(band):
                            x1, x2 = (a[i], b[i]) if a[i] <= b[i] else (b[i], a[i])
                            if x1 <= rr <= x2:
                                selfrefs.append((nid, rr))
                                break
                hulls.append(hull)
                starts.append(r0)
                ends.append(rN)
                ids.append(nid)

            prev_r, prev_t = -9, None
            for r, f in lst:
                sp = _DIGC.split(f)                     # 非数字、数字、非数字…：写法＝非数字部分连起来，数字另存
                t = '\x00'.join(sp[::2])
                if not (r == prev_r + 1 and t == prev_t):
                    flush()
                    band = []
                band.append((r, f, sp[1::2]))
                prev_r, prev_t = r, t
            flush()
            colband[(sheet, c)] = (starts, ends, ids)
    fcols = defaultdict(list)
    for (sh, c) in colband:
        fcols[sh].append(c)
    for v in fcols.values():
        v.sort()

    # ② 边：被引用区域碰到的段 → 公式所在的段
    tcache = {}

    def targets(sh, c1, c2, r1, r2):
        k = (sh, c1, c2, r1, r2)
        v = tcache.get(k)
        if v is not None:
            return v
        out = []
        cols = fcols.get(sh)
        if cols:
            for c in cols[bisect_left(cols, c1):bisect_right(cols, c2)]:
                st, en, ids = colband[(sh, c)]
                out.extend(ids[bisect_left(en, r1):bisect_right(st, r2)])
        tcache[k] = out
        return out

    succ = [[] for _ in nodes]
    for b, hull in enumerate(hulls):
        seen = set()
        for ref in hull:
            for t in targets(*ref):
                if t != b and t not in seen:
                    seen.add(t)
                    succ[t].append(b)

    # ③ 强连通分量（Tarjan，非递归）
    n = len(nodes)
    index = [-1] * n
    low = [0] * n
    onst = [False] * n
    st = []
    comps = []
    idx = 0
    for root in range(n):
        if index[root] >= 0:
            continue
        work = [(root, 0)]
        index[root] = low[root] = idx
        idx += 1
        st.append(root)
        onst[root] = True
        while work:
            v, i = work[-1]
            sv = succ[v]
            if i < len(sv):
                work[-1] = (v, i + 1)
                w = sv[i]
                if index[w] < 0:
                    index[w] = low[w] = idx
                    idx += 1
                    st.append(w)
                    onst[w] = True
                    work.append((w, 0))
                elif onst[w] and index[w] < low[v]:
                    low[v] = index[w]
            else:
                work.pop()
                if work:
                    u = work[-1][0]
                    if low[v] < low[u]:
                        low[u] = low[v]
                if low[v] == index[v]:
                    comp = []
                    while True:
                        w = st.pop()
                        onst[w] = False
                        comp.append(w)
                        if w == v:
                            break
                    if len(comp) >= 2:
                        comps.append(comp)

    def label(i):
        sh, c, r0, r1, _f = nodes[i]
        cl = _cl(c)
        return f'{sh}!{cl}{r0}' if r0 == r1 else f'{sh}!{cl}{r0}:{cl}{r1}'

    def example(i):
        sh, c, r0, _r1, f = nodes[i]
        return f'{sh}!{_cl(c)}{r0}: {f[:300]}'

    out = []
    for comp in comps[:limit]:
        cs = set(comp)
        s = min(comp)                                   # 在分量里找一条从 s 出发回到 s 的最短回路
        prev = {s: None}
        dq = deque([s])
        back = None
        while dq and back is None:
            v = dq.popleft()
            for w in succ[v]:
                if w not in cs:
                    continue
                if w == s:
                    back = v
                    break
                if w not in prev:
                    prev[w] = v
                    dq.append(w)
        path = []
        v = back
        while v is not None:
            path.append(v)
            v = prev[v]
        path.reverse()                                  # s → … → back →（回到 s）
        out.append({'种类': '环', '环': [label(i) for i in path] + [label(s)], '节点数': len(comp),
                    '例子': example(path[1] if len(path) > 1 else s)})
    done = set()
    for nid, r in selfrefs:
        if nid in done:
            continue
        done.add(nid)
        sh, c, _r0, _r1, f = nodes[nid]
        out.append({'种类': '自己引用自己', '环': [f'{sh}!{_cl(c)}{r}'], '节点数': 1, '例子': example(nid)})
        if len(out) >= limit:
            break
    return out


def fmt(cycles):
    lines = []
    for k, c in enumerate(cycles, 1):
        lines.append(f'[{k}] {c["种类"]}（{c["节点数"]} 段）：' + ' → '.join(c['环']))
        lines.append(f'     例：{c["例子"]}')
    return '\n'.join(lines)


if __name__ == '__main__':
    import sys
    import time
    import openpyxl
    t = time.time()
    wb = openpyxl.load_workbook(sys.argv[1])
    t1 = time.time()
    cyc = find_cycles(wb)
    print(f'读 {t1 - t:.1f}s，查 {time.time() - t1:.1f}s，环 {len(cyc)} 个')
    print(fmt(cyc[:20]))
