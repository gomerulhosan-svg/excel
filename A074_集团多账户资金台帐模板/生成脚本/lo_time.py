# -*- coding: utf-8 -*-
"""在 LibreOffice 里量「改资金台帐一格要算多久」，报「占参照本的百分比」（v3.1，设计 §15.7 第 5 条；复核 compat-1/compat-2）。

跑法：python3 -I lo_time.py 成品.xlsx [另一本.xlsx …] [--ref 参照.xlsx | --no-ref] [--reps 3] [--no-entry] [--ref-entry] [--sandwich] [--threads]
      build.py 全容量时自动调（量的是**交付的成品**，不是 LibreOffice 另存过的 _calc 副本：那份公式全展开了，量出来快一倍，不是用户的样子）。

口径：
- 关 LibreOffice 多线程（UseThreadedCalculationForFormulaGroups=false）和 OpenCL，量 soffice.bin 进程的 CPU 秒（/proc/pid/stat），
  不量墙钟（这台机器别的任务也在跑，墙钟会抖一倍）。打开时不重算（OOXMLRecalcMode=1）。
- 先量参照本（默认 v2 成品＝git 4575e4b，第一次从 git 取出来放 _calc/ 下），紧接着量要量的本子（同一时段）；
  --sandwich：最后再量一遍参照本，分母取前后两遍的平均（机器忙闲有变化时用）。
  每本一个新的 LO 进程、一次只开一本（实测：两本同时开着交替改，两本每格的 CPU 秒都翻倍，v2 成品 12.4 → 25.8；
  同一个进程先开过参照本再开 v3.1 骨架，骨架每格 0.75 → 1.55，也翻倍。设计 §15.7 原来写的「同一个 LO 进程」不能照做）。
  只量 CPU 秒在这台 4 核（超线程）机器上照样会被别的任务拖慢一些，同一处 3 次取中、看百分比。
  每本都是先复制到临时目录再打开，只读不存（close 不保存），原文件不动。
- 每本：打开 → 关自动计算 → 整本算一遍（calculateAll）→ 改台帐（每次「填一格＋只算脏格 calculate()」一起计 CPU，3 次取中）：
    末尾：最后一笔下面那一行（先照上一行填好日期、账户，算一遍，不计时）改收入；
    中间：台帐中间一行改金额（收入栏有数改收入，否则改支出）；
    共用行：v3.1 两段共用的那一行（第 HDR＋SEG_ROWS×k 行，挑离中间最近的）改金额；
    摘要：中间那一行改摘要（文字）。
  要量的本子再开「录入模式」（基础资料①填「是」）量一遍同样的四处（末尾换成再下一行），最后清空录入模式、量「关掉那一下」整本要算多久。
  --ref-entry：参照本也量「报表暂停」时（对照评审的 68%）。
- 百分比＝本子的 CPU 中位数 ÷ 参照本同一处正常模式的 CPU 中位数。报警线 layout.TIME_WARN_NORMAL_PCT（30%）、TIME_WARN_ENTRY_PCT（12%），
  超了只报警不拦。别用 UNO setDataArray 把公式换成值做对照（评审实测会量错）。
- 只是 LibreOffice 的数：Excel 一般快几倍，WPS 没数据（交付前请用户实机量，设计 §15.7）。
输出：中文的一段结果；最后一行是 JSON（build.py 读）。"""
import json
import os
import shutil
import statistics
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
REF_COMMIT = '4575e4b'                                       # v2 成品（现在发给用户的那本）
REF_REL = 'A074_集团多账户资金台帐模板/A074_集团多账户资金台帐模板.xlsx'
SCAN_ROWS = 14000                                            # 找台帐最后一笔：从第 R0 行往下看这么多行（v2 12,000、v3.1 10,000）


def ref_path():
    """参照本：git 4575e4b 的成品，取出来放 _calc/参照_v2_4575e4b.xlsx（取一次，以后复用）"""
    p = os.path.join(HERE, '_calc', f'参照_v2_{REF_COMMIT}.xlsx')
    if not os.path.exists(p):
        os.makedirs(os.path.dirname(p), exist_ok=True)
        root = subprocess.run(['git', '-C', HERE, 'rev-parse', '--show-toplevel'], capture_output=True, text=True,
                              check=True).stdout.strip()
        blob = subprocess.run(['git', '-C', root, 'show', f'{REF_COMMIT}:{REF_REL}'], capture_output=True, check=True).stdout
        with open(p + '.tmp', 'wb') as fh:
            fh.write(blob)
        os.replace(p + '.tmp', p)
    return p


class Office:
    """一个 headless LibreOffice（临时配置目录），关多线程、OpenCL；cpu() 读 soffice.bin 的 CPU 秒"""

    def __init__(self, threads=False):
        import uno
        from com.sun.star.beans import PropertyValue
        self.PV = PropertyValue
        self.prof = tempfile.mkdtemp(prefix='a074_lo_')
        pipe = 'a074time%d' % os.getpid()
        self.proc = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', '--nodefault', '--nolockcheck',
                                      f'-env:UserInstallation=file://{self.prof}', f'--accept=pipe,name={pipe};urp;'],
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = None
        for _ in range(240):
            try:
                ctx = resolver.resolve(f'uno:pipe,name={pipe};urp;StarOffice.ComponentContext')
                break
            except Exception:
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError('连不上 LibreOffice')
        self.ctx = ctx
        self.desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
        self.cfg = {}
        self._set('/org.openoffice.Office.Calc/Formula/Calculation', UseThreadedCalculationForFormulaGroups=bool(threads))
        self._set('/org.openoffice.Office.Common/Misc', UseOpenCL=False)
        self._set('/org.openoffice.Office.Calc/Formula/Load', OOXMLRecalcMode=1)
        self.pid = None
        for _ in range(40):
            for d in os.listdir('/proc'):
                if d.isdigit():
                    try:
                        cl = open(f'/proc/{d}/cmdline', 'rb').read().decode('utf8', 'ignore')
                    except OSError:
                        continue
                    if 'soffice.bin' in cl and self.prof in cl:
                        self.pid = int(d)
            if self.pid:
                break
            time.sleep(0.25)
        if not self.pid:
            raise RuntimeError('找不到 soffice.bin 进程（量不了 CPU 秒）')
        self.tck = os.sysconf('SC_CLK_TCK')

    def _set(self, path, **props):
        cp = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', self.ctx)
        nd = self.PV()
        nd.Name, nd.Value = 'nodepath', path
        upd = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationUpdateAccess', (nd,))
        for k, v in props.items():
            try:
                upd.setPropertyValue(k, v)
            except Exception as e:                            # 版本没有这个设置：记下来，不拦
                self.cfg[k] = f'设不上：{e}'
        upd.commitChanges()
        rd = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (nd,))
        for k in props:
            if k not in self.cfg:
                self.cfg[k] = rd.getPropertyValue(k)

    def cpu(self):
        f = open(f'/proc/{self.pid}/stat').read().rsplit(')', 1)[1].split()
        return (int(f[11]) + int(f[12])) / self.tck

    def load(self, path):
        pv = self.PV()
        pv.Name, pv.Value = 'Hidden', True
        return self.desk.loadComponentFromURL('file://' + path, '_blank', 0, (pv,))

    def close(self):
        try:
            self.desk.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(20)
        except Exception:
            self.proc.kill()
        shutil.rmtree(self.prof, ignore_errors=True)


def _timed(lo, fn):
    c0, t0 = lo.cpu(), time.time()
    fn()
    return round(lo.cpu() - c0, 2), round(time.time() - t0, 2)


def _last_used(sh, r0):
    """台帐最后一笔（日期或账户有字）的行号；没有＝r0−1"""
    arr = sh.getCellRangeByName(f'A{r0}:B{r0 + SCAN_ROWS - 1}').getDataArray()
    last = r0 - 1
    for i, (a, b) in enumerate(arr):
        if (a not in ('', None)) or (b not in ('', None)):
            last = r0 + i
    return last


def _pick_rows(last, L):
    r0 = L.R0
    mid = max(r0, (r0 + last) // 2)
    seg = getattr(L, 'SEG_ROWS', 100)
    k = max(1, round((mid - L.HDR) / seg))
    shared = L.HDR + seg * k
    if shared == mid or shared > last:
        shared = L.HDR + seg * max(1, k - 1)
    if shared < r0 or shared > last:
        shared = None
    return mid, shared


def _fill_like(sh, row, src, cols):
    """新的一行：日期、账户照 src 那一行填（先算一遍，不计时）"""
    for col in cols:
        s = sh.getCellRangeByName(f'{col}{src}')
        d = sh.getCellRangeByName(f'{col}{row}')
        if s.getString() == '':
            continue
        if s.getValue() != 0:
            d.setValue(s.getValue())
        else:
            d.setString(s.getString())


def _amount_col(sh, row, L):
    f, g = L.LEDGER_COLS['收入'], L.LEDGER_COLS['支出']
    if sh.getCellRangeByName(f'{f}{row}').getString() == '' and sh.getCellRangeByName(f'{g}{row}').getString() != '':
        return g
    return f


def _param_row(doc, L):
    """基础资料①「录入模式」（v2 的参照本是「报表暂停」）在第几行 → (行, 名字)；没有＝(None, None)"""
    sh = doc.Sheets.getByName(L.SH_BASE)
    for r in range(L.PAR_HDR + 1, L.PAR_HDR + 15):
        t = sh.getCellRangeByName(f'A{r}').getString().strip()
        if t in ('录入模式', '报表暂停'):
            return r, t
    return None, None


POS = ('末尾', '中间', '共用行', '摘要')


class Book:
    """打开的一本（临时副本）：台帐最后一笔、要改的几行、结果"""

    def __init__(self, lo, path, L):
        self.lo, self.L, self.path = lo, L, path
        self.res = {'文件': path}
        self.tmpd = tempfile.mkdtemp(prefix='a074t_')
        tmp = os.path.join(self.tmpd, 'x.xlsx')
        shutil.copy(path, tmp)
        box = {}
        self.res['打开CPU'], self.res['打开秒'] = _timed(lo, lambda: box.__setitem__('d', lo.load(tmp)))
        self.doc = box['d']
        if self.doc is None:
            raise RuntimeError(f'打不开 {path}')
        self.doc.enableAutomaticCalculation(False)
        self.res['整本CPU'], self.res['整本秒'] = _timed(lo, self.doc.calculateAll)
        self.sh = self.doc.Sheets.getByName(L.SH_LEDGER)
        self.last = _last_used(self.sh, L.R0)
        self.mid, self.shared = _pick_rows(self.last, L)
        self.res['最后一笔'] = self.last
        self.prow, self.res['开关'] = _param_row(self.doc, L)
        self.new_row = None
        self.add_row()

    def add_row(self):
        """最后一笔下面再开一行（照最后一笔填日期、账户；算一遍，不计时）"""
        L = self.L
        self.new_row = (self.new_row or self.last) + 1
        _fill_like(self.sh, self.new_row, self.last, (L.LEDGER_COLS['日期'], L.LEDGER_COLS['账户']))
        self.res.setdefault('新行准备CPU', []).append(_timed(self.lo, self.doc.calculate)[0])

    def rows(self):
        return {'末尾': self.new_row, '中间': self.mid, '共用行': self.shared, '摘要': self.mid}

    def edit(self, pos, i):
        """改一处，「填＋calculate()」一起计 CPU → (CPU, 秒)；这一本没有这一处（如没有共用行）→ None"""
        L, sh = self.L, self.sh
        row = self.rows()[pos]
        if row is None:
            return None
        if pos == '摘要':
            c = sh.getCellRangeByName(f"{L.LEDGER_COLS['摘要']}{row}")
            fn = (lambda: c.setString(f'测试摘要租金{i}'))
        elif pos == '末尾':
            c = sh.getCellRangeByName(f"{L.LEDGER_COLS['收入']}{row}")
            fn = (lambda: c.setValue(1000 + 37 * i))
        else:
            c = sh.getCellRangeByName(f'{_amount_col(sh, row, L)}{row}')
            v = c.getValue()
            fn = (lambda: c.setValue(round((v or 100) + 1, 2)))
        return _timed(self.lo, lambda: (fn(), self.doc.calculate()))

    def mode(self, on):
        """开/关 录入模式（参照本是「报表暂停」）→ (CPU, 秒)"""
        cell = self.doc.Sheets.getByName(self.L.SH_BASE).getCellRangeByName(f'{self.L.BASE_PAR_VAL}{self.prow}')
        return _timed(self.lo, lambda: (cell.setString('是' if on else ''), self.doc.calculate()))

    def close(self):
        try:
            self.doc.close(True)
        except Exception:
            pass
        shutil.rmtree(self.tmpd, ignore_errors=True)


def _put(res, mode, pos, row, x):
    d = res.setdefault(mode, {}).setdefault(pos, {'行': row, 'CPU': [], '秒': []})
    d['CPU'].append(x[0])
    d['秒'].append(x[1])
    d['CPU中'] = round(statistics.median(d['CPU']), 2)
    d['秒中'] = round(statistics.median(d['秒']), 2)


def _phase(books, modes, reps, base=0):
    """交替量：每一处、每一次都按 参照 → 本子1 → 本子2… 的顺序各改一下（机器忙闲对各本一样）。modes：每本记到哪个模式名下"""
    for i in range(reps):
        for pos in POS:
            for b, m in zip(books, modes):
                x = b.edit(pos, base + i)
                if x is not None:
                    _put(b.res, m, pos, b.rows()[pos], x)


def pct(target, ref):
    """{模式: {位置: 百分比}}：本子（正常／录入）÷ 参照本正常（量了两遍＝--sandwich 时取两遍的平均）"""
    out = {}
    for mode in ('正常', '录入'):
        if mode not in target or '正常' not in ref:
            continue
        out[mode] = {}
        for pos, v in target[mode].items():
            rs = [ref[m][pos]['CPU中'] for m in ('正常', '正常2') if pos in ref.get(m, {})]
            den = sum(rs) / len(rs) if rs else 0
            if den > 0:
                out[mode][pos] = round(100 * v['CPU中'] / den, 1)
    return out


def main(argv):
    import layout as L
    args = [a for a in argv]

    def opt(k, d=None):
        if k in args:
            i = args.index(k)
            v = args[i + 1]
            del args[i:i + 2]
            return v
        return d

    def flag(k):
        if k in args:
            args.remove(k)
            return True
        return False
    reps = int(opt('--reps', 3))
    refp = opt('--ref')
    no_ref, no_entry, ref_entry, threads, sandwich = (flag('--no-ref'), flag('--no-entry'), flag('--ref-entry'), flag('--threads'),
                                                      flag('--sandwich'))
    targets = [os.path.abspath(a) for a in args]
    if not no_ref:
        refp = os.path.abspath(refp) if refp else ref_path()
    warn_n, warn_e = getattr(L, 'TIME_WARN_NORMAL_PCT', 30), getattr(L, 'TIME_WARN_ENTRY_PCT', 12)
    out = {'口径': f'LibreOffice CPU 秒（填一格＋calculate()），每处 {reps} 次取中；每本一个新 LO 进程，先参照本' + ('，最后再量一遍参照本' if sandwich else ''),
           '参照': refp, '本子': {}, '百分比': {}, '报警': []}
    out['设置'] = {}
    out['机器负载'] = os.getloadavg()[0]

    def run(path, entry, tag='正常'):
        """每本一个新的 LO 进程（同一个进程先开过 2 GB 的参照本再开别的本子，实测每改一格的 CPU 秒翻倍）"""
        lo = Office(threads=threads)
        out['设置'] = lo.cfg
        b = None
        try:
            b = Book(lo, path, L)
            _phase([b], [tag], reps)
            if entry and b.prow:                                  # 开录入模式（参照本是「报表暂停」）再量一遍
                b.res['开录入CPU'] = b.mode(True)[0]
                b.add_row()
                _phase([b], ['录入'], reps, base=10)
                b.res['关录入CPU'], b.res['关录入秒'] = b.mode(False)
        finally:
            if b is not None:
                b.close()
            lo.close()
        return b.res
    try:
        if refp:
            out['参照结果'] = run(refp, ref_entry)
        for t in targets:
            out['本子'][t] = run(t, not no_entry)
        if refp and sandwich:
            out['参照结果']['正常2'] = run(refp, False, '正常2')['正常2']
    except Exception as e:                                        # 量不成：记下来（LO 在 run 里已经关了）
        out['错误'] = repr(e)[:300]
    for t, r in out['本子'].items():
        if refp and '正常' in out.get('参照结果', {}):
            p = pct(r, out['参照结果'])
            out['百分比'][t] = p
            for pos, v in p.get('正常', {}).items():
                if v > warn_n:
                    out['报警'].append(f'{os.path.basename(t)} 正常模式 {pos} {v}% > {warn_n}%')
            for pos, v in p.get('录入', {}).items():
                if v > warn_e:
                    out['报警'].append(f'{os.path.basename(t)} 录入模式 {pos} {v}% > {warn_e}%')
    pn = [v for p in out['百分比'].values() for v in p.get('正常', {}).values()]
    pe = [v for p in out['百分比'].values() for v in p.get('录入', {}).values()]
    out['正常最高%'] = max(pn) if pn else None
    out['录入最高%'] = max(pe) if pe else None
    report(out)
    return out


def _fmt(d):
    return '、'.join(f"{k} {v['CPU中']}" for k, v in d.items())


def report(out):
    print(f"LibreOffice 计时（{out['口径']}；多线程 {out['设置'].get('UseThreadedCalculationForFormulaGroups')}，"
          f"OpenCL {out['设置'].get('UseOpenCL')}；开始时机器负载 {out['机器负载']:.1f}）" + (f"  ✗ {out['错误']}" if out.get('错误') else ''))
    rr = out.get('参照结果')
    if rr:
        print(f"  参照 {os.path.basename(rr['文件'])}：打开 {rr.get('打开秒')} 秒（CPU {rr.get('打开CPU')}），整本 {rr.get('整本CPU')}；"
              f"改一格 CPU 秒：{_fmt(rr.get('正常', {}))}" + (f"（第二轮 {_fmt(rr['正常2'])}）" if rr.get('正常2') else '')
              + (f"；{rr['开关']}时 {_fmt(rr['录入'])}" if rr.get('录入') else ''))
    for t, r in out['本子'].items():
        p = out['百分比'].get(t, {})
        print(f"  {os.path.basename(t)}：打开 {r.get('打开秒')} 秒（CPU {r.get('打开CPU')}），整本 {r.get('整本CPU')}")
        for mode in ('正常', '录入'):
            if mode in r:
                print(f"    {mode}：" + '、'.join(f"{k} {v['CPU中']} 秒" + (f"（{p[mode][k]}%）" if k in p.get(mode, {}) else '')
                                          for k, v in r[mode].items()))
        if '关录入CPU' in r:
            print(f"    清空录入模式那一下：整本 {r['关录入CPU']} CPU 秒")
    for w in out['报警']:
        print('  △ ' + w)


if __name__ == '__main__':
    res = main(sys.argv[1:])
    print(json.dumps(res, ensure_ascii=False, default=str))
