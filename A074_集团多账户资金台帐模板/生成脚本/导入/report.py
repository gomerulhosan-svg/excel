# -*- coding: utf-8 -*-
"""把清洗结果写成 json 和三份 md（整理记录、待确认、核对）。"""
import collections
import datetime
import json
import os
import re

from conf import (COMPANIES, ACCOUNTS, ACCOUNT_HOLDER, ITEMS, ITEM_SRC, NOT_IMPORT, NONCASH_OFFSET, ADD_ROWS,
                  MISSING_DAY, MANUAL_ITEM, MEMO_SPLIT, MEMO_FEE)
from rules_kw import RULES, PARTY_RULES

SHEET_ACC = {a[0]: a[1] for a in ACCOUNTS}
ACC_SHEET = {a[1]: a[0] for a in ACCOUNTS}
ACC_ORDER = [a[1] for a in ACCOUNTS]
ITEM_ATTR = {n: a for n, a, *_ in ITEMS}
ITEM_GROUP = {n: g for n, a, g, *_ in ITEMS}

# 古玩城农商银行 用户手算银行余额（R/S 列）各月末，见 journal_struct.md §5.1
GW_RS = [520288.77, 85757.53, 91114.08, 155258.06, 176581.58, 192347.62, 221452.50, 326729.94, 311559.15,
         410539.64]


def r2(x):
    return round(x + 0.0, 2)


def fmt(x):
    if x is None or x == '':
        return ''
    return '{:,.2f}'.format(x)


def md_table(head, rows):
    out = ['| ' + ' | '.join(head) + ' |', '|' + '|'.join('---' for _ in head) + '|']
    for r in rows:
        out.append('| ' + ' | '.join(str(c).replace('|', '／').replace('\n', ' ') for c in r) + ' |')
    return '\n'.join(out)


def dump(out, name, obj):
    with open(os.path.join(out, name), 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


# ============================================================ 主入口
def write_all(ctx):
    out = ctx['out']
    E = ctx['entries']
    CH = ctx['CH']
    meta = ctx['meta']
    openings = ctx['openings']
    im = ctx['im']
    dups = ctx['dups']
    rows = ctx['rows']

    by_acc = collections.defaultdict(list)
    for e in E:
        by_acc[e['账户']].append(e)

    # ---------------- accounts / companies
    accounts = []
    for i, (sheet, name, co, bank, kind, no, holder, status, note) in enumerate(ACCOUNTS, 1):
        es = by_acc[name]
        first = es[0]['日期'] if es else '2026-01-01'
        op = openings[sheet][0]
        start = '2026-01-01'
        if abs(op) < 0.005 and first > '2026-01-31':
            start = first
        accounts.append(dict(序号=i, 账户名=name, 公司简称=co, 公司全称=dict(COMPANIES)[co], 银行=bank, 户别=kind,
                             户号尾号=no, 持卡人=holder, 开账日期=start, 开账余额=op, 状态=status, 原表名=sheet,
                             子表名=name, 期末余额=es[-1]['余额'] if es else op, 流水行数=len(es),
                             最后一笔日期=es[-1]['日期'] if es else '', 说明=note))
        assert len(name) <= 20 and not re.search(r'[\[\]:*?/\\ ]', name), name
    dump(out, 'accounts.json', accounts)
    companies = []
    for i, (co, full) in enumerate(COMPANIES, 1):
        companies.append(dict(顺序=i, 简称=co, 全称=full, 账户数=sum(1 for a in ACCOUNTS if a[2] == co),
                              账户=[a[1] for a in ACCOUNTS if a[2] == co]))
    dump(out, 'companies.json', companies)

    # ---------------- items
    alias_rev = collections.defaultdict(list)
    for k, v in im['item_alias'].items():
        alias_rev[v['to']].append(k)
    ic = collections.Counter(e['收费项'] for e in E)
    iin = collections.Counter()
    iout = collections.Counter()
    for e in E:
        iin[e['收费项']] += e['收入']
        iout[e['收费项']] += e['支出']
    items = []
    for i, (n, attr, grp, hint, note) in enumerate(ITEMS, 1):
        items.append(dict(序号=i, 名称=n, 属性=attr, 大类=grp, 方向提示=hint, 说明=note, 原写法别名=alias_rev.get(n, []),
                          来源=ITEM_SRC[n], 本次笔数=ic[n], 收入合计=r2(iin[n]), 支出合计=r2(iout[n])))
    dump(out, 'items.json', items)

    # ---------------- rules
    rl = []
    k = 0
    for rid, d, co, where, kws, item, party, note in RULES:
        for kw in kws:
            k += 1
            rl.append(dict(序号=k, 规则号=rid, 关键词=kw, 查找位置=where, 适用方向=d, 适用公司=co, 收费项=item,
                           收付人=party, 说明=note))
    pl = []
    k = 0
    for rid, d, co, kws, party in PARTY_RULES:
        for kw in kws:
            k += 1
            pl.append(dict(序号=k, 规则号=rid, 关键词=kw, 查找位置='摘要', 适用方向=d, 适用公司=co, 收付人=party))
    dump(out, 'rules.json', dict(
        说明='按顺序从上往下找，命中第一条就停。关键词只要「包含」就算命中（不分大小写，不用正则）。'
           '适用方向：收=收入行、支=支出行、都=都适用；适用公司空=所有公司。查找位置：摘要或收付人。'
           '收费项规则命中「内部转账」时还要填对方账户；对方不在本册的改记「往来款」。'
           '收付人规则只在收付人空着时用。',
        收费项规则=rl, 收付人规则=pl))

    # ---------------- parties
    parties = build_parties(E, ctx['party_alias'])
    dump(out, 'parties.json', parties)

    # ---------------- entries
    keys = ['序号', '日期', '账户', '收费项', '属性', '收付人', '摘要', '收入', '支出', '余额', '对方账户', '归属公司', '备注',
            '原表位置', '来源标记', '账户内序号']
    dump(out, 'entries.json', [{k: e[k] for k in keys} for e in E])

    # ---------------- changes
    dump(out, 'changes.json', CH)
    write_changes_md(out, CH, E)

    # ---------------- check
    chk = write_check_md(out, ctx, accounts, by_acc)
    write_questions_md(out, ctx, by_acc, chk)
    print_summary(E, CH, chk)


# ============================================================ 收付人清单
BANKS = ('农商银行', '建设银行', '海峡银行', '招商银行', '中国银行', '民生银行', '兴业银行', '泉州银行', '工商银行', '农业银行')
GROUP_NAMES = set(ACCOUNT_HOLDER.values()) | {'福州市青翼创创业投资有限公司', '博烨篮球馆', '博烨羽毛球', '徐洁', '吴海萍',
                                               '王晓华', '谢依俤', '谢海亮', '福建力喆贸易有限公司'}
RENT_GROUPS = {'租赁收入', '物业收入', '押金'}


def build_parties(E, party_alias):
    st = collections.defaultdict(lambda: dict(收=0, 支=0, 租=0, 往=0, n=0, acc=collections.Counter()))
    for e in E:
        p = e['收付人']
        if not p:
            continue
        s = st[p]
        s['n'] += 1
        s['acc'][e['账户']] += 1
        if ITEM_ATTR.get(e['收费项']) == '往来':
            s['往'] += 1
        if e['收入']:
            s['收'] += 1
            if ITEM_GROUP.get(e['收费项']) in RENT_GROUPS or e['收费项'] in ('月卡收入',):
                s['租'] += 1
        else:
            s['支'] += 1
    alias = collections.defaultdict(set)
    for k, v in party_alias.items():
        kk = k.strip()
        if kk != v:
            alias[v].add(kk)
    out = []
    for p, s in sorted(st.items(), key=lambda kv: -kv[1]['n']):
        if p in GROUP_NAMES or re.search(r'力喆|元生万物|言吾言|艾泊烨|博烨体育|易佰兴信息|趋势创投|中策管理|青翼创', p):
            t = '内部'
        elif p in BANKS or p.endswith('银行') or re.search(r'银行.*(支行|营业部)$', p):
            t = '银行'
        elif re.search(r'税务|TIPS', p):
            t = '税务'
        elif re.search(r'收单|过渡户|清算|财付通|代发|支付科技|星驿|汇付|结算中心|抖音|微信|物联', p):
            t = '其他'
        elif '、' in p:
            t = '个人'
        elif re.search(r'中国移动|中国电信|中国联通|保险', p):
            t = '供应商'
        elif s['往'] * 2 >= s['n'] and s['n']:
            t = '个人' if len(p) <= 4 and not re.search(r'公司|店|厂|中心', p) else '其他'
        elif s['收'] >= s['支']:
            t = '租户' if s['租'] * 2 >= s['收'] and s['收'] else '客户'
            if p in ('地摊散户', '固定摊位商户', '临时停车散客', '月租车散户'):
                t = '客户'
        else:
            if re.search(r'公司|店|厂|中心|局|所|部|社|行|馆|院|集团|商行|工作室|站|会', p) or len(p) > 4:
                t = '供应商'
            else:
                t = '个人'
        if p == '员工（代发工资）':
            t = '个人'
        out.append(dict(名称=p, 类型=t, 别名=sorted(alias.get(p, [])), 笔数=s['n'],
                        常见账户=[a for a, _ in s['acc'].most_common(3)]))
    return out


# ============================================================ 整理记录
KIND_ORDER = ['期初', '不导', '补记', '日期', '金额方向', '拆行', '收费项-逐行改', '收费项-统一写法', '收费项-内部转账',
              '收费项-摘要规则', '收费项-待定', '收费项-分栏', '收付人', '收付人-补', '摘要', '收费项']
KIND_TITLE = {
    '期初': '开账余额', '不导': '不导入的行', '补记': '补记的行', '日期': '补日期', '金额方向': '金额和方向',
    '拆行': '一行拆成几行', '收费项-逐行改': '收费项：逐行改（填错的）', '收费项-统一写法': '收费项：统一写法',
    '收费项-内部转账': '收费项：内部转账（两头配上对的）', '收费项-摘要规则': '收费项：空的按摘要规则补',
    '收费项-待定': '收费项：定不了，标「待定」', '收费项-分栏': '收费项：分栏表按分栏列定',
    '收付人': '收付人：统一写法', '收付人-补': '收付人：空的补上', '摘要': '摘要：只去空格', '收费项': '收费项：其他'}


def write_changes_md(out, CH, E):
    c = collections.Counter(x['类别'] for x in CH)
    src = collections.Counter(e['来源标记'] for e in E)
    L = ['# 整理记录（A074 资金台帐导入）', '',
         '原表：2026年日记账（26 张账户表，不含「易佰兴基本户」）。每一条＝一处改动：原表位置、改了哪个字段、原值、新值、依据。'
         '「原表位置」写的是原表那一行的摘要格（易佰兴农商银行外是 E 列，其余是 F 列），一行拆成几行的在后面加 -1、-2。'
         '完整清单在 changes.json（%d 条），这里分类列出；分栏表按分栏列定收费项的 %d 条只列汇总。' % (len(CH), c['收费项-分栏']),
         '没有实际改动的（原值＝新值）不记。注意：原表的表名「博烨篮球馆 」末尾有一个空格，原表位置里写成「博烨篮球馆」，'
         '在原表里找这张表时按「博烨篮球馆 」找。', '',
         '## 数量统计', '']
    L.append(md_table(['类别', '条数'], [[KIND_TITLE.get(k, k), c[k]] for k in KIND_ORDER if c[k]] +
                      [['合计', len(CH)]]))
    L += ['', '总表 %d 行，收费项来源：' % len(E), '']
    L.append(md_table(['来源标记', '行数', '意思'], [
        [k, src[k], {'原有': '原表 D 列本来就填对了', '统一写法': '写法不统一，按对照表改', '逐行改': '填错了逐行改，或两头配上对改成内部转账',
                     '分栏': '分栏表按租金/物业/水电…列定', '摘要规则': '空的按摘要关键词规则补', '待定': '定不了，等用户确认',
                     '补记': '原表漏记补上的', '拆行': '从一行里拆出来的手续费、押金冲抵'}.get(k, '')]
        for k in ['原有', '统一写法', '逐行改', '分栏', '摘要规则', '待定', '补记', '拆行'] if src[k]]))
    L.append('')
    for k in KIND_ORDER:
        xs = [x for x in CH if x['类别'] == k]
        if not xs:
            continue
        L += ['## %s（%d 条）' % (KIND_TITLE.get(k, k), len(xs)), '']
        if k == '收费项-分栏':
            agg = collections.Counter()
            for x in xs:
                agg[(x['原表位置'].split('!')[0], x['新值'])] += 1
            L.append(md_table(['原表', '收费项', '条数'], [[a, b, n] for (a, b), n in sorted(agg.items())]))
            L.append('')
            continue
        if k == '收费项-摘要规则':
            agg = collections.Counter()
            for x in xs:
                m = re.match(r'规则 (\S+)：', x['依据'])
                agg[(m.group(1) if m else x['依据'][:12], x['新值'])] += 1
            L.append('按规则汇总：')
            L.append('')
            L.append(md_table(['规则号', '收费项', '条数'], [[a, b, n] for (a, b), n in sorted(agg.items())]))
            L.append('')
            L.append('逐条：')
            L.append('')
        L.append(md_table(['原表位置', '字段', '原值', '新值', '依据'],
                          [[x['原表位置'], x['字段'], short(x['原值']), short(x['新值']), x['依据']] for x in xs]))
        L.append('')
    with open(os.path.join(out, 'changes.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')


def short(s, n=60):
    s = '' if s is None else str(s)
    s = s.replace('\n', '⏎')
    return s if len(s) <= n else s[:n] + '…'


# ============================================================ 核对
def write_check_md(out, ctx, accounts, by_acc):
    E = ctx['entries']
    meta = ctx['meta']
    rows = ctx['rows']
    openings = ctx['openings']
    legs = ctx['legs']
    L = ['# 核对（A074 资金台帐导入）', '']
    res = dict(fail=[])

    # ---- 修正后口径（原表数按已定口径修正，仍是净额口径）
    corr = {}
    for a in ACCOUNTS:
        s = a[0]
        m = meta[s]
        corr[s] = dict(期初=m['期初'] or 0.0, 收入=m['收入合计'], 支出=m['支出合计'], 期末=m['期末_表上最后余额'], 说明=[])
    g = corr['古玩城农商银行']
    vir = r2(sum(r['inc'] for r in rows if r['sheet'] == '古玩城农商银行' and r['row'] in
                 {int(x['行']) for x in ctx['dups']['古玩城农商银行_同一行收支都记']['行']}))
    nc = r2(sum(r['inc'] for r in rows if (r['sheet'], r['row']) in NONCASH_OFFSET))
    g.update(期初=openings['古玩城农商银行'][0], 收入=r2(g['收入'] - vir - nc))
    g['期末'] = r2(g['期初'] + g['收入'] - g['支出'])
    g['说明'] += ['期初改银行口径 %s（原 P5 %s）' % (fmt(g['期初']), fmt(meta['古玩城农商银行']['期初'])),
                '收入去掉「吴海萍转入徐洁账户」虚收入 %s' % fmt(vir), '收入去掉以押抵租 1217–1219 行 %s（原表余额本来就不含）' % fmt(nc)]
    for (s, rw), why in NOT_IMPORT.items():
        r = [x for x in rows if x['sheet'] == s and x['row'] == rw][0]
        c = corr[s]
        c['收入'] = r2(c['收入'] - r['inc'])
        c['支出'] = r2(c['支出'] - r['exp'])
        c['期末'] = r2(c['期初'] + c['收入'] - c['支出'])
        c['说明'].append('去掉 %s!%d（%s %s，不过本账户）' % (s, rw, '支出' if r['exp'] else '收入', fmt(r['exp'] or r['inc'])))
    for ad in ADD_ROWS:
        c = corr[ad['sheet']]
        c['支出'] = r2(c['支出'] + ad['exp'])
        c['收入'] = r2(c['收入'] + ad['inc'])
        c['期末'] = r2(c['期初'] + c['收入'] - c['支出'])
        c['说明'].append('补记 9/20 转基本户 %s' % fmt(ad['exp']))
    t = corr['易佰兴趋势民生银行']
    if abs(t['期末'] - meta['易佰兴趋势民生银行']['期末_表上最后余额']) > 0.005:
        t['说明'].append('原表合计行把利息多加 0.01，表上期末 %s，逐笔算是 %s' % (
            fmt(meta['易佰兴趋势民生银行']['期末_表上最后余额']), fmt(t['期末'])))

    # ---- 新总表各账户
    fee = collections.Counter()
    off = collections.Counter()
    negc = collections.Counter()
    for e in E:
        s = e['_sheet']
        if e['_kind'] == 'fee':
            fee[s] += e['支出']
        if e['_kind'] == 'offset':
            off[s] += e['支出']
    neg_rows = [r for r in rows if r['inc'] < 0 or r['exp'] < 0]
    for r in neg_rows:
        negc[r['sheet']] += abs(r['inc']) + abs(r['exp'])
    # 按摘要拆（押金抵租）多出来的收、支（两边一样多）
    msx = collections.Counter()
    rowix = {(r['sheet'], r['row']): r for r in rows}
    for k in MEMO_SPLIT:
        es_ = [e for e in E if (e['_sheet'], e['_row']) == k]
        xi = r2(sum(e['收入'] for e in es_) - rowix[k]['inc'])
        xo = r2(sum(e['支出'] for e in es_) - rowix[k]['exp'])
        if abs(xi - xo) > 0.005:
            res['fail'].append('%s!%d 按摘要拆后收、支多出来的不一样：%s / %s' % (k[0], k[1], fmt(xi), fmt(xo)))
        msx[k[0]] += xi
    L += ['## 1. 每个账户：期初＋收入−支出＝期末', '',
          '三种口径：**原表**＝原表上的数（收入是扣手续费后的净额，不含本月合计行）；**修正后**＝按已定口径修正原表（古玩城用银行口径、'
          '去掉不过本账户的备忘行、补记漏记的转出），仍是净额；**新总表**＝本次导入的结果（手续费单独记支出，所以收入、支出都比净额大，余额不变）。', '']
    head = ['#', '账户', '原表名', '期初', '原表 收入', '原表 支出', '原表 期末', '修正后 收入', '修正后 支出', '修正后 期末',
            '新 收入', '新 支出', '新 期末', '新−修正后 期末', '原表笔数', '新行数']
    trs = []
    expl = []
    for i, a in enumerate(accounts, 1):
        s = a['原表名']
        es = by_acc[a['账户名']]
        m = meta[s]
        c = corr[s]
        ni = r2(sum(e['收入'] for e in es))
        no = r2(sum(e['支出'] for e in es))
        end = es[-1]['余额'] if es else a['开账余额']
        calc = r2(a['开账余额'] + ni - no)
        if abs(calc - end) > 0.005:
            res['fail'].append('%s 期初+收−支=%s，但最后一行余额 %s' % (a['账户名'], fmt(calc), fmt(end)))
        if abs(end - c['期末']) > 0.005:
            res['fail'].append('%s 新期末 %s ≠ 修正后期末 %s' % (a['账户名'], fmt(end), fmt(c['期末'])))
        trs.append([i, a['账户名'], s.strip(), fmt(a['开账余额']), fmt(m['收入合计']), fmt(m['支出合计']),
                    fmt(m['期末_表上最后余额']), fmt(c['收入']), fmt(c['支出']), fmt(c['期末']), fmt(ni), fmt(no), fmt(end),
                    fmt(r2(end - c['期末'])), m['行数']['流水'], len(es)])
        # 差异解释
        d_in = r2(ni - c['收入'])
        d_out = r2(no - c['支出'])
        parts = []
        if fee[s]:
            parts.append('手续费单独记支出 %s（收入按扣费前记，收、支各多 %s）' % (fmt(fee[s]), fmt(fee[s])))
        if off[s]:
            parts.append('以押抵租收入 %s 照记、另记押金冲抵支出 %s' % (fmt(off[s]), fmt(off[s])))
        if negc[s]:
            parts.append('原表负数收入 %s 改记支出（收、支各多 %s）' % (fmt(negc[s]), fmt(negc[s])))
        if abs(msx[s]) > 0.005:
            parts.append('押金抵租金/物业/水电按摘要拆开，抵掉的部分照记收入、另记退押金（收、支各多 %s）' % fmt(msx[s]))
        known = r2(fee[s] + off[s] * 2 * 0 + negc[s])
        # 收入多 fee+off+neg，支出多 fee+off+neg
        exp_d = r2(fee[s] + off[s] + negc[s] + msx[s])
        if abs(d_in - exp_d) > 0.005 or abs(d_out - exp_d) > 0.005:
            res['fail'].append('%s 新−修正后：收入差 %s、支出差 %s，能解释的 %s' % (a['账户名'], fmt(d_in), fmt(d_out),
                                                                       fmt(exp_d)))
        extra_rows = len(es) - m['行数']['流水']
        if c['说明'] or parts or extra_rows:
            expl.append((a['账户名'], c['说明'], parts, m['行数']['流水'], len(es)))
    L.append(md_table(head, trs))
    tot_o = r2(sum(a['开账余额'] for a in accounts))
    tot_e = r2(sum((by_acc[a['账户名']][-1]['余额'] if by_acc[a['账户名']] else a['开账余额']) for a in accounts))
    L += ['', '26 个账户开账余额合计 %s，期末（各账户最后一笔后）合计 %s。' % (fmt(tot_o), fmt(tot_e)), '']
    L += ['### 差异逐条解释', '']
    for nm, cs, parts, n0, n1 in expl:
        L.append('- **%s**' % nm)
        for x in cs:
            L.append('  - 原表→修正后：' + x)
        for x in parts:
            L.append('  - 修正后→新总表：' + x + '，余额不变')
        if n1 != n0:
            L.append('  - 行数 %d → %d：%s' % (n0, n1, row_delta_note(nm, E, ctx)))
    L.append('')

    # ---- 月末余额
    L += ['## 2. 各账户月末余额和原表逐月对', '',
          '原表月末＝按原表余额公式逐行重算的月末数（journal_sheets.json「月末余额_重算」）；古玩城农商银行用用户 R/S 列手算的银行余额。'
          '只列有差异的，差异都应该能用上面第 1 节解释。', '']
    mrows = []
    nmatch = 0
    for a in accounts:
        s = a['原表名']
        es = by_acc[a['账户名']]
        me = {}
        for e in es:
            me[int(e['日期'][5:7])] = e['余额']
        ref = {int(k): v for k, v in meta[s]['月末余额_重算'].items()}
        if s == '古玩城农商银行':
            ref = {i + 1: v for i, v in enumerate(GW_RS)}
        for mo, v in sorted(ref.items()):
            mine = me.get(mo)
            if mine is None:
                # 当月没有流水：取上月
                prev = [me[k] for k in me if k < mo]
                mine = prev[-1] if prev else a['开账余额']
            if abs(mine - v) > 0.005:
                mrows.append([a['账户名'], mo, fmt(v), fmt(mine), fmt(r2(mine - v))])
            else:
                nmatch += 1
    L.append('一致 %d 个月末；不一致 %d 个：' % (nmatch, len(mrows)))
    L.append('')
    if mrows:
        L.append(md_table(['账户', '月', '原表月末', '新总表月末', '差'], mrows))
    L += ['', '说明：博烨羽毛球 9 月起差 −7,478.06 是补记的转基本户；易佰兴趋势民生银行差 −0.01 是原表合计行重复加利息。', '']

    # ---- 守恒
    L += ['## 3. 拆行前后金额守恒', '']
    eff = collections.Counter()
    for e in E:
        if e['_kind'] != 'add':
            eff[(e['_sheet'], e['_row'])] += e['收入'] - e['支出']
    vir_rows = {(x['表'], int(x['行'])) for x in ctx['dups']['古玩城农商银行_同一行收支都记']['行']}
    bad = []
    nrow = 0
    for r in rows:
        k = (r['sheet'], r['row'])
        if k in NOT_IMPORT:
            continue
        nrow += 1
        want = r['inc'] - r['exp']
        if k in vir_rows:
            want = -r['exp']
        if k in NONCASH_OFFSET:
            want = 0.0
        if abs(eff[k] - want) > 0.005:
            bad.append([k[0], k[1], fmt(want), fmt(eff[k])])
    split_rows = sum(1 for k, v in collections.Counter((e['_sheet'], e['_row']) for e in E).items() if v > 1)
    L.append('原表 %d 行（不含不导的 %d 行）逐行比：新总表里同一原表行拆出来的各行「收入−支出」之和＝原表这一行对余额的影响'
             '（吴海萍转入徐洁的 22 行只算支出；以押抵租 1217–1219 行原表余额不含，按 0）。拆成多行的原表行 %d 行。不相等的：%d 行。'
             % (nrow, len(NOT_IMPORT), split_rows, len(bad)))
    if bad:
        L.append('')
        L.append(md_table(['原表', '行', '原表对余额影响', '新总表合计'], bad))
        res['fail'].append('拆行不守恒 %d 行' % len(bad))
    L.append('')

    # ---- 完整性
    L += ['## 4. 每一行都齐', '']
    item_names = {n for n, *_ in ITEMS}
    acc_names = set(ACC_ORDER)
    c1 = [e for e in E if not e['账户'] or e['账户'] not in acc_names]
    c2 = [e for e in E if not re.match(r'^2026-\d\d-\d\d$', e['日期'])]
    c3 = [e for e in E if e['收费项'] not in item_names]
    c4 = [e for e in E if e['收费项'] == '内部转账' and (not e['对方账户'] or e['对方账户'] not in acc_names)]
    c5 = [e for e in E if e['对方账户'] and e['对方账户'] == e['账户']]
    c6 = [e for e in E if (e['收入'] and e['支出']) or (not e['收入'] and not e['支出']) or e['收入'] < 0 or e['支出'] < 0]
    c7 = [e for e in E if e['对方账户'] and e['收费项'] != '内部转账']
    c8 = [e for e in E if not e['收付人']]
    checks = [('账户为空或不在账户清单', c1), ('日期不是 2026 年的真日期', c2), ('收费项不在清单（含「待定」）', c3),
              ('内部转账没填对方账户或对方不在账户清单', c4), ('对方账户就是本账户', c5), ('收入、支出同时有/都没有/负数', c6),
              ('填了对方账户但不是内部转账', c7)]
    L.append(md_table(['检查', '不合格行数'], [[n, len(x)] for n, x in checks] + [['（参考）收付人为空', len(c8)]]))
    for n, x in checks:
        if x:
            res['fail'].append('%s %d 行' % (n, len(x)))
    if c8:
        L += ['', '收付人补不出来、留空的 %d 行：' % len(c8), '']
        L.append(md_table(['原表位置', '收入', '支出', '摘要'], [[e['原表位置'], fmt(e['收入']), fmt(e['支出']), short(e['摘要'], 40)]
                                                       for e in c8]))
    neg = [e for e in E if e['余额'] < -0.004]
    L += ['', '余额出现负数的行：%d 行。' % len(neg)]
    if neg:
        L.append('')
        L.append(md_table(['原表位置', '日期', '账户', '支出', '余额', '原因'],
                          [[e['原表位置'], e['日期'], e['账户'], fmt(e['支出']), fmt(e['余额']),
                            '原表同一天先记支出、后记收入（原表这一行余额也是负数）' if e['原表位置'] == '易佰兴建行!F10' else '待查']
                           for e in neg]))
    L.append('')

    # ---- 按属性、大类、收费项
    L += ['## 5. 全年合计（2026-01-01 至 %s）' % max(e['日期'] for e in E), '', '### 5.1 按属性', '']
    attr = collections.OrderedDict((k, [0.0, 0.0, 0]) for k in ['收入', '支出', '往来', '内部转账', '待定'])
    for e in E:
        k = '待定' if e['收费项'] == '待定' else e['属性']
        attr[k][0] += e['收入']
        attr[k][1] += e['支出']
        attr[k][2] += 1
    L.append(md_table(['属性', '收入列合计', '支出列合计', '净额（收−支）', '行数'],
                      [[k, fmt(v[0]), fmt(v[1]), fmt(r2(v[0] - v[1])), v[2]] for k, v in attr.items()] +
                      [['合计', fmt(sum(v[0] for v in attr.values())), fmt(sum(v[1] for v in attr.values())),
                        fmt(r2(sum(v[0] - v[1] for v in attr.values()))), sum(v[2] for v in attr.values())]]))
    res['attr'] = {k: (r2(v[0]), r2(v[1]), v[2]) for k, v in attr.items()}
    op_in = attr['收入'][0] - attr['收入'][1]
    op_out = attr['支出'][1] - attr['支出'][0]
    L += ['', '经营口径：收入类净额 %s，支出类净额 %s，经营净额 %s（不含往来、内部转账、待定）。'
          '全部账户净增加＝期末合计−开账合计＝%s。' % (fmt(op_in), fmt(op_out), fmt(op_in - op_out), fmt(tot_e - tot_o)), '']
    # 冲减
    cj = [e for e in E if (e['属性'] == '收入' and e['支出']) or (e['属性'] == '支出' and e['收入'])]
    cj_in = r2(sum(e['支出'] for e in cj if e['属性'] == '收入'))
    cj_out = r2(sum(e['收入'] for e in cj if e['属性'] == '支出'))
    L += ['### 5.1.1 冲减行（收入类记在支出列、支出类记在收入列）', '',
          '这些是退回、冲回，收费项跟原来那笔一样，金额记在反方向。报表按收费项、属性、大类汇总时，**收入类要用 收入列−支出列，'
          '支出类要用 支出列−收入列**；只取一列的话，收入类会多算 %s、支出类会多算 %s。' % (fmt(cj_in), fmt(cj_out)), '']
    L.append(md_table(['原表位置', '日期', '账户', '收费项', '属性', '收入', '支出', '摘要'],
                      [[e['原表位置'], e['日期'], e['账户'], e['收费项'], e['属性'], fmt(e['收入']), fmt(e['支出']),
                        short(e['摘要'], 40)] for e in cj]))
    res['cj'] = (len(cj), cj_in, cj_out)
    L += ['', '### 5.1.2 给生成台帐的提示', '',
          '- items.json 里「待定」的属性写的是「按方向」：录入表如果用公式从收费项清单查属性，查到「按方向」时要按收支方向判断'
          '（收入列有数＝收入，支出列有数＝支出），否则这 %d 行进不了收入、支出合计。entries.json 里每行的「属性」已经按方向填好。'
          % sum(1 for e in E if e['收费项'] == '待定'),
          '- entries.json 的「日期」是文字 "2026-01-01"，写进 Excel 时要转成真日期，不能当文字写，否则按月 SUMIFS、筛选都会失效。',
          '- 冲减行见上面 5.1.1。', '']
    # 大类
    L += ['### 5.2 按大类', '']
    grp = collections.OrderedDict()
    for n, a_, gname, *_ in ITEMS:
        grp.setdefault((a_, gname), [0.0, 0.0, 0])
    for e in E:
        k = (ITEM_ATTR[e['收费项']], ITEM_GROUP[e['收费项']])
        grp[k][0] += e['收入']
        grp[k][1] += e['支出']
        grp[k][2] += 1
    L.append(md_table(['属性', '大类', '收入列', '支出列', '净额', '行数'],
                      [[k[0], k[1], fmt(v[0]), fmt(v[1]), fmt(r2(v[0] - v[1])), v[2]] for k, v in grp.items() if v[2]]))
    L += ['', '### 5.3 按收费项', '']
    it = collections.OrderedDict((n, [0.0, 0.0, 0]) for n, *_ in ITEMS)
    for e in E:
        it[e['收费项']][0] += e['收入']
        it[e['收费项']][1] += e['支出']
        it[e['收费项']][2] += 1
    L.append(md_table(['收费项', '属性', '大类', '收入列', '支出列', '净额', '行数'],
                      [[n, ITEM_ATTR[n], ITEM_GROUP[n], fmt(v[0]), fmt(v[1]), fmt(r2(v[0] - v[1])), v[2]]
                       for n, v in it.items() if v[2]]))
    unused = [n for n, v in it.items() if not v[2]]
    L += ['', '清单里今年没用到的收费项：' + '、'.join(unused), '']
    # 按月×属性
    L += ['### 5.4 按月 × 属性（收入列 / 支出列）', '']
    mon = collections.defaultdict(lambda: collections.defaultdict(lambda: [0.0, 0.0]))
    for e in E:
        k = '待定' if e['收费项'] == '待定' else e['属性']
        mon[int(e['日期'][5:7])][k][0] += e['收入']
        mon[int(e['日期'][5:7])][k][1] += e['支出']
    ks = list(attr.keys())
    L.append(md_table(['月'] + ['%s 收/支' % k for k in ks],
                      [[m] + ['%s / %s' % (fmt(mon[m][k][0]), fmt(mon[m][k][1])) for k in ks] for m in sorted(mon)]))
    L.append('')

    # ---- 内部转账
    L += ['## 6. 内部转账两边', '']
    ie = [e for e in E if e['收费项'] == '内部转账']
    tin = r2(sum(e['收入'] for e in ie))
    tout = r2(sum(e['支出'] for e in ie))
    L.append('内部转账 %d 行：转入合计 %s，转出合计 %s，差 %s。' % (len(ie), fmt(tin), fmt(tout), fmt(r2(tin - tout))))
    res['internal'] = (len(ie), tin, tout)
    L.append('')
    pair = collections.defaultdict(lambda: [0.0, 0.0])
    for e in ie:
        a, b = e['账户'], e['对方账户']
        if e['支出']:
            pair[(a, b)][0] += e['支出']
        else:
            pair[(b, a)][1] += e['收入']
    prow = []
    for (a, b), (o, i_) in sorted(pair.items(), key=lambda kv: (ACC_ORDER.index(kv[0][0]), ACC_ORDER.index(kv[0][1]))):
        prow.append([a, b, fmt(o), fmt(i_), fmt(r2(o - i_))])
    L.append('按「转出账户 → 转入账户」：转出方记的支出合计 vs 转入方记的收入合计（对方账户填的是谁就算谁）。')
    L.append('')
    L.append(md_table(['转出账户', '转入账户', '转出方记的', '转入方记的', '差'], prow))
    L.append('')
    pos_ix = {e['原表位置']: e for e in E}
    diffs = [r for r in prow if r[4] not in ('0.00', '-0.00')]
    L.append('有差的 %d 对，逐条：' % len(diffs))
    L.append('')
    dl = []
    for e in ie:
        if '金额差' in e['备注']:
            dl.append([e['原表位置'], e['日期'], e['账户'], e['对方账户'], fmt(e['收入'] or -e['支出']), e['备注']])
    for g_ in ctx['dups']['内部转账_一对多_净额_缺日期']:
        if g_['类型'] != '净额':
            continue
        legs_ = [g_['单笔']] + g_['对方多笔']
        for x in legs_:
            e = pos_ix.get('%s!F%d' % (x['表'].strip(), x['行']))
            if e:
                dl.append([e['原表位置'], e['日期'], e['账户'], e['对方账户'], fmt(e['收入'] or -e['支出']),
                           '净额记账：%s 记了一支一收，%s 只记净额 %s；两个方向的差 %s 正好抵消，净额为 0' % (
                               SHEET_ACC[g_['对方多笔'][0]['表']], SHEET_ACC[g_['单笔']['表']], fmt(g_['单笔']['金额']),
                               fmt(min(x2['金额'] for x2 in g_['对方多笔'])))])
    if dl:
        L.append(md_table(['原表位置', '日期', '账户', '对方账户', '金额', '说明'], dl))
    L += ['', '日期不同的配对（两边记账日期差 1~3 天，各按原表日期记，按日看内部转账净额会有临时差）：', '']
    dd = []
    for p in ctx['dups']['内部转账_一对一']:
        if p['日差']:
            dd.append(['%s!F%d' % (p['转出']['表'].strip(), p['转出']['行']), p['转出']['日期'],
                       '%s!F%d' % (p['转入']['表'].strip(), p['转入']['行']), p['转入']['日期'], fmt(p['金额'])])
    for g_ in ctx['dups']['内部转账_一对多_净额_缺日期']:
        dts = {g_['单笔']['日期']} | {x['日期'] for x in g_['对方多笔']}
        if len(dts) > 1 and '?' not in ''.join(dts):
            dd.append(['%s!F%d' % (g_['单笔']['表'].strip(), g_['单笔']['行']), g_['单笔']['日期'],
                       '、'.join('%s!F%d' % (x['表'].strip(), x['行']) for x in g_['对方多笔']),
                       '、'.join(sorted({x['日期'] for x in g_['对方多笔']})), fmt(g_['金额'])])
    L.append(md_table(['一边', '日期', '另一边', '日期', '金额'], dd))
    # 互相指向检查
    L += ['', '两头互指检查：每一行内部转账的对方账户里，都要有方向相反、对方账户填回本账户的内部转账。', '']
    cnt = collections.Counter()
    for e in ie:
        cnt[(e['账户'], e['对方账户'], '收' if e['收入'] else '支')] += 1
    lonely = [e for e in ie if cnt[(e['对方账户'], e['账户'], '支' if e['收入'] else '收')] == 0]
    L.append('找不到反向那一头的：%d 行。' % len(lonely))
    if lonely:
        res['fail'].append('内部转账只有一头 %d 行' % len(lonely))
        L.append(md_table(['原表位置', '账户', '对方账户', '金额'], [[e['原表位置'], e['账户'], e['对方账户'],
                                                              fmt(e['收入'] or -e['支出'])] for e in lonely]))
    L.append('')
    # 单边（不算内部转账）
    L += ['### 6.1 对方不在本册的转账（不算内部转账，按往来类记）', '']
    sl = []
    for x in ctx['dups']['单边转账']:
        p = '%s!F%d' % (x['表'].strip(), x['行'])
        e = pos_ix.get(p)
        if e is None:
            continue
        sl.append([p, e['日期'], e['收费项'], e['属性'], fmt(e['收入'] or -e['支出']), x['归类'], short(e['摘要'], 30)])
    L.append(md_table(['原表位置', '日期', '收费项', '属性', '金额', '对方', '摘要'], sl))
    L.append('')
    L += ['## 7. 自查结论', '']
    if res['fail']:
        L += ['**有 %d 项没过：**' % len(res['fail']), ''] + ['- ' + x for x in res['fail']]
    else:
        L.append('全部通过：每个账户 期初＋Σ收−Σ支＝期末＝修正后原表期末；拆行守恒；每行都有账户、日期；收费项都在清单里；'
                 '内部转账都有对方账户且两头互指。')
    with open(os.path.join(out, 'check.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    res['tot_open'] = tot_o
    res['tot_end'] = tot_e
    return res


def row_delta_note(nm, E, ctx):
    es = [e for e in E if e['账户'] == nm]
    s = ACC_SHEET[nm]
    c = collections.Counter(e['_kind'] for e in es)
    extra = sum(n - 1 for k, n in collections.Counter((e['_sheet'], e['_row']) for e in es if e['_kind'] != 'add').items()
                if n > 1)
    parts = []
    if extra:
        parts.append('拆行多出 %d 行（分栏、手续费、收支同行、按对方账户拆）' % extra)
    ni = sum(1 for k in NOT_IMPORT if k[0] == s)
    if ni:
        parts.append('不导 %d 行（不过本账户的备忘行）' % ni)
    n_amt = sum(1 for r in ctx['rows'] if r['sheet'] == s)
    blank = ctx['meta'][s]['行数']['流水'] - n_amt
    if blank:
        parts.append('原表 %d 行没有金额（空行或只写了文字），不导' % blank)
    if c['add']:
        parts.append('补记 %d 行' % c['add'])
    return '；'.join(parts)


# ============================================================ 待确认
def write_questions_md(out, ctx, by_acc, chk):
    E = ctx['entries']
    meta = ctx['meta']
    rows = ctx['rows']
    pos_ix = {e['原表位置']: e for e in E}

    def amt(p):
        e = pos_ix.get(p)
        return (e['收入'] or e['支出']) if e else 0

    Q = []

    def q(title, money, where, body):
        Q.append((title, money, where, body))

    tbd = [e for e in E if e['收费项'] == '待定']
    yf = [e for e in tbd if '悦丰' in e['收付人']]
    q('收费项「待定」的 %d 笔：性质要你说' % len(tbd), sum(e['收入'] + e['支出'] for e in tbd),
      '、'.join(e['原表位置'] for e in tbd),
      '- 艾泊烨农商银行!F17：4/8 支出「预付账款」1,600,000.00，付给厦门市鑫创达科技有限公司；同一天收到 160 万贷款（!F16）。'
      '是哪个项目的预付款？（可能记「丝绸厂车库机械设备」或别的项目）\n'
      '- 易佰兴农商银行外 福州悦丰投资有限公司转入 %d 笔，合计 %s（%s），其中 !E63 写「2-3月贷款利息款」。悦丰和易佰兴是什么关系？'
      '是借款（记借款转入）、投资（投资款），还是代付贷款利息？\n'
      '- 易佰兴农商银行外!E10：1/7「陈兴登转入」100,000.00（陈兴登是易佰兴法人）。是股东借款（借款转入）还是投资款？'
      % (len(yf), fmt(sum(e['收入'] for e in yf)), '、'.join(e['原表位置'].split('!')[1] for e in yf)))
    q('「易佰兴农商银行外」和「易佰兴基本户」只收了「外」那张', 840006.21,
      '易佰兴农商银行外（102 笔）、易佰兴基本户（51 笔）',
      '- 两张是同一个农商银行基本户：基本户 51 笔全部在「外」表里能找到，「外」表多 1 月份和悦丰转入、贷款利息等 52 笔。'
      '总表只导「外」表（期初 46,768.81，8/21 余额 8,127.51），基本户那张不导；原表汇总、Sheet1 用的是基本户（余额 −2,950.64），对不上是正常的。\n'
      '- 唯一不一样的一笔：基本户!F55 记 6/24「物业费」2,609.52，「外」表!E103 记 6/29「租金」2,609.52，总表按「外」表。请确认哪个对。\n'
      '- 这个账户 8/21 以后没再记，后面的流水要补录。')
    gw = [e for e in E if e['账户'] == '古玩城-农商银行(吴海萍私户)']
    q('古玩城农商银行（吴海萍卡）按银行口径导', 836000 + 8659,
      '古玩城农商银行!P5/R5；22 行「吴海萍转入徐洁账户」；1217–1219 行',
      '- 期初用 262,273.10（R5 / Sheet1!B2 的银行余额），不用 P5 账面 7,552,829.50；导完 10/9 余额 410,539.64，和你 R/S 列一致。\n'
      '- 22 行「吴海萍转入徐洁账户」原表同一行收、支都记（836,000.00），只导了支出一头，记内部转账、对方是徐洁招商卡。\n'
      '- 1217–1219 行刘瑞琴 A120 以押抵租（8,659.00）：租金、物业、水电收入照记，另记一笔「退押金」8,659.00，备注「以押抵租，不过银行」，余额不变。')
    sal_in = [e for e in E if e['收费项'] == '内部转账' and e['备注'].startswith('原收费项：薪资、社医保')]
    sal_out = [e for e in E if e['收费项'] == '内部转账' and e['账户'] == SHEET_ACC['易佰兴私户']
               and re.match(r'原收费项：(薪资|五险一金)', e['备注'])]
    q('易佰兴的工资社保经王晓华卡转给艾泊烨代发：按内部转账记了（艾泊烨收 %d 笔 %s，易佰兴私户付 %d 笔 %s）' % (
        len(sal_in), fmt(sum(e['收入'] for e in sal_in)), len(sal_out), fmt(sum(e['支出'] for e in sal_out))),
      sum(e['收入'] for e in sal_in) + sum(e['支出'] for e in sal_out),
      '易佰兴私户 16 笔（原记薪资/五险一金）、艾泊烨私户银行 8 笔和艾泊烨海峡银行 5 笔（原记「薪资、社医保」收入）',
      '- 钱的路线：易佰兴私户（谢依俤卡）→ 艾泊烨私户（王晓华卡）→ 艾泊烨海峡基本户 → 艾泊烨代发工资。原表在易佰兴记了薪资、五险一金，'
      '在艾泊烨又记了「薪资、社医保」收入，艾泊烨代发时再记一次薪资，集团合起来薪资算了两遍、收入也虚增。\n'
      '- 现在三段都记「内部转账」（原收费项写在备注），工资只在艾泊烨代发那一笔算一次。这样集团合计是对的，'
      '但看「易佰兴公司」自己的利润时就看不到这部分工资。你要按哪种看？')
    single = ctx['dups']['单边转账']
    sing_amt = sum(x['金额'] for x in single if x['归类'] != '外部单位往来(不是内部转账)' and '漏记' not in x['归类'])
    q('对方账户不在本册的转账（%d 笔）没算内部转账' % len([x for x in single if x['归类'] != '外部单位往来(不是内部转账)'
                                               and '漏记' not in x['归类']]), sing_amt,
      '见 check.md 6.1',
      '- 对方是王晓华建行卡、谢海亮另一张卡、徐洁工行 2426 / 海峡 1796、吴海萍招行、青翼创（账户表已删）、周月平个人。'
      '这些卡不在本册，钱确实进出了集团账户，所以按原收费项记（借款转入、还款、贷款转入、往来款、徐总费用…），属性是「往来」或支出。\n'
      '- 要不要把这几张卡也建成账户？建了之后这些就能改成内部转账。另外有 3 组是两头都在本册、只是经过这些卡中转'
      '（徐洁招商!F12→艾泊烨海峡!F16 130,000；力喆传媒!F7、!F9→言吾言!F17、!F20 各 5,000），已经按内部转账记，备注写了「经…中转」。\n'
      '- 徐洁招商!F254 10/13「吴海萍转入徐洁账户」5,000：古玩城农商银行只记到 10/9，那边补录这笔转出后，两头都改成内部转账。')
    q('中策-海峡银行、四季快安-兴业银行 标「暂停」', 119940.52 + 122.53, '原表两张隐藏表',
      '- 中策 3/21 以后、四季快安 4/16 以后没再记账，余额还在（119,940.52、122.53），已收进总表、状态「暂停」，老板看总资金时会算进去。'
      '这两个户还用不用？要不要接着记？\n'
      '- 中策!F7、!F11、!F16「收徐洁服务费」3 × 4,000 记了其他收入（徐洁招商卡里没有对应支出，应是徐洁另一张卡）；'
      '中策!F12「提取备用金」48,000 记了借款支出（现金去向没有表）。')
    q('博烨羽毛球补记一笔转出', 7478.06, '补记（对应 博烨体育农商银行!F133）',
      '- 博烨基本户 9/20 收到「博烨一般户（羽毛球馆）」7,478.06，羽毛球表没有这笔转出。按基本户补记了，羽毛球账户期末从 7,478.41 变成 0.35，请核对银行流水。')
    q('不导的两笔备忘行', 50000 + 13193.55, '艾泊烨私户银行!F143、易佰兴趋势民生银行!F9',
      '- 艾泊烨私户!F143 7/21「谢海亮垫付王庄篮球馆地板第二笔进度款」50,000：钱没过王晓华卡（原表没算进余额），没导。'
      '这笔要不要在博烨或艾泊烨记一笔「谢海亮垫付」的往来？\n'
      '- 易佰兴趋势民生!F9 2/13「易佰兴2026年1月工资」13,193.55（没有收付人）：没过本账户，没导。是从哪个账户发的？')
    q('内部转账两边金额或日期对不上', 50 + 50000 + 21000,
      '艾泊烨私户银行!F169 / 艾泊烨海峡银行!F269；易佰兴私户!F350 / 易佰兴建行!F82；古玩城农商银行!F1176 / 徐洁招商银行!F244',
      '- 艾泊烨私户!F169 9/15 转出 41,994.93，海峡!F269 收 41,944.93，差 50.00，哪边录错了？\n'
      '- 易佰兴私户!F350 记 4/30 转出 50,000，建行!F82 记 4/27 收到，差 3 天。\n'
      '- 古玩城!F1176 记 9/28 转出 21,000，徐洁招商!F244 记 9/30 收到。\n'
      '- 博烨私户!F53 摘要写「基本户转入私户」，实际是私户转回基本户 4,950.67（对方博烨基本户!F146）；'
      '艾泊烨农商银行!F83 摘要写「归还博烨体育借款」，钱进的是艾泊烨自己的海峡基本户（!F283）。摘要原样保留，请看要不要改。')
    q('账户归属、户名待确认', 0, 'accounts.json',
      '- 「力喆实业中国银行」表头写「力喆产业银行日记账」，「力喆产业农商银行」别的表转账时多写「福建力喆产业投资集团有限公司」（14 笔）、'
      '也有写力喆实业的（4 笔）。现在：中国银行归力喆实业、农商一般户归力喆产业，对吗？\n'
      '- 徐洁招商卡归在力喆实业（原汇总「力喆私户」），J 列写的归属公司放进了「归属公司」列（力喆实业 97 笔、古玩城 23 笔）。\n'
      '- 博烨私户持卡人按摘要推断是谢海亮；匹克球 847 户开在哪家银行（收单是泉州银行）？\n'
      '- 只有一个对公户的公司（力喆贸易、力喆传媒、元生万物、言吾言、趋势创投、中策、四季快安），户别先写「对公基本户」。\n'
      '- 易佰兴私户原表期初是 0，对吗？')
    q('收费项口径（新增和改动）', 0, 'items.json',
      '- 新增 4 项：场馆营业收入（博烨各球馆收款）、投资款（集团外投资，如万象会 200,000、林志 10）、内部转账、待定。\n'
      '- 属性分四种：收入 / 支出 / 往来 / 内部转账。往来＝押金、退押金、借款转入、借款支出、还款、收回借款、贷款转入、还贷款本金、'
      '投资款、股权投资款、股东分红、往来款，不算经营收支。股东分红、退押金放在往来，你同意吗？\n'
      '- 「往来款」原来 307 行大多是集团内互转，现在两头都在本册的改成「内部转账」；剩下的（对方不在本册）还叫往来款，属性往来。\n'
      '- 逐行改里把握低的：易佰兴建行!F150、!F152 电梯检测/维保 办公费→物业修缮费；徐洁招商!F86 利达大厦物业费 徐总费用→物业费。\n'
      '- 退回的钱用原来那笔的收费项冲减（记在收入列），不记其他收入：博烨体育农商!F106 退回公众责任保险 5,000 保留原表的办公费'
      '（和买保险的办公费 5,000 冲平）；中策!F13 退回旅游费用 800 记福利费（冲 !F9 旅游费 21,320）。\n'
      '- 原表负数收入改记支出、收费项不变（冲减）：艾泊烨私户!F106–F109（−5,991.77，冲回 2025 年电费和分成）、易佰兴私户!F480（−554.12，退预充值电费）。\n'
      '- 艾泊烨海峡!F90 原数退回 !F89 收的 225,103 水电费，记成「水电收入」冲减；博烨基本户!F140 退润楼体育订场费 900，记「场地租金收入」冲减。\n'
      '- 易佰兴趋势民生!F12 2/26「工资及年终奖退回」25,230.67 记「薪资」放收入列（冲减），不记其他收入：2/13 发了 36,324.22（!F10）'
      '和 2,100（!F11），退回后实发 36,324.22＋2,100−25,230.67＝13,193.55，正好是备忘行 !F9 的「1 月工资」数；'
      '记成其他收入会让其他收入、薪资各虚增 25,230.67。\n'
      '- 冲减行（收入类记在支出列、支出类记在收入列）共 %d 行，报表按「收入列−支出列」算净额，清单见 check.md 5.1.1。\n'
      '- 艾泊烨私户!F76、F77、F116、F142、F168 月卡收入原表记的是扣了微信提现手续费的净额，按摘要改记毛额'
      '（14,500 / 12,750 / 22,750 / 24,650 / 20,000.12），手续费单独一行「银行手续费」（共 84.60），余额不变。' % len(
          [e for e in E if (e['属性'] == '收入' and e['支出']) or (e['属性'] == '支出' and e['收入'])]))
    q('收付人写法待确认', 0, 'parties.json',
      '- 古玩城农商银行!F1161 写「林素琴」：B126 一直是柯秀强交租，是代付人还是「林淑琴」笔误？\n'
      '- 力喆贸易两种全称（福州力喆贸易有限公司 1 行、福建力喆贸易有限公司 9 行），统一成「福建力喆贸易有限公司」。\n'
      '- 拓远（福建）勘测规划有限公司 和 福建拓远智航科技发展有限公司 在同一铺位 A301，没合并。\n'
      '- 陈杜娟 / 陈杜鹃（同一车牌闽A818C6）统一成「陈杜娟」。\n'
      '- 古玩城 B129 4/17 换了租户：原表写「刘祖锣（宋少波）」「宋少波」的 %d 行（%s）收付人记「宋少波」，没并到刘祖锣；'
      '刘祖锣留 B130 和 4/17 以前的 B129。\n'
      '- 收付人空着的按铺位号（同一铺位最近一笔的租户）、停车品牌、地摊等补了，补不出来的留空（见 check.md 第 4 节）。\n'
      '- 易佰兴停车费只写了车牌的：同一车牌在别的行只有一个收付人的，按车牌补（%d 行）；摘要写临时、临停、停车券的记「临时停车散客」；'
      '查不到车主的月租车记「月租车散户」（见「月租车…查不到车主」那一条）。\n'
      '- 收单、提现手续费那一行的收付人：有收款通道的填通道（汇付、速停车、辰域、星驿、厦门路桥、财付通），没有的填开户行。\n'
      '- 内部转账的收付人统一填公司全称（原表写「博烨一般户（篮球馆）」「博烨一般户（羽毛球馆）」的 %d 行改成福州市博烨体育发展有限公司），'
      '是哪个账户看「对方账户」。' % (
          len([e for e in E if e['收付人'] == '宋少波']), fmt(sum(e['收入'] for e in E if e['收付人'] == '宋少波')),
          len([e for e in E if e.get('_p07', '').startswith('规则 P07a')]),
          len([c for c in ctx['CH'] if c['字段'] == '收付人' and c['原值'] in ('博烨一般户（篮球馆）', '博烨一般户（羽毛球馆）')])))
    yz = [e for e in E if e['收付人'] == '月租车散户']
    q('月租车 %d 笔查不到车主，收付人先写「月租车散户」' % len(yz), sum(e['收入'] for e in yz),
      '、'.join(e['原表位置'] for e in yz),
      '- 这些停车费摘要只有车牌和期间（月租、包月、年度缴费），车牌在别的行也没出现过收付人。原来按规则补成「临时停车散客」不对，'
      '改成「月租车散户」。请告诉车主是谁：\n' +
      '\n'.join('  - %s %s %s：%s' % (e['原表位置'], e['日期'][5:], fmt(e['收入']), e['摘要']) for e in yz) +
      '\n- 易佰兴私户!F84「闽ADF8706 2.5-3.4」和 !F167「闽ADF7806 3.5-4.4」（顺丰的车）期间正好接上、金额一样，'
      '可能是同一辆车、车牌两位数字写反了；没有确认前先不按顺丰记。')
    q('易佰兴建行!F136「B219-220 押金」8,800 原表记在租金栏，改记押金', 8800, '易佰兴建行!F136、!F137',
      '- 摘要只写「B219-220 押金」，金额却记在租金栏。同日 !F137 写明「B314-315押金7954.64元抵扣B219-220 6月22日-9月30日的部分租金…'
      '补足剩余租金4453.36元、物业2112元」：B219-220 的租金由旧押金抵和补交，这 8,800 应是 B219-220 的新押金。'
      '总表记「押金」（往来），房租收入少 8,800。如果确实是租金，告诉我们改回来。')
    L12 = ('- 博烨私户!F17「代取现金（王庄开业红包）」5,000 原记薪资：按收费项说明，红包属于公共关系费用（或福利费）。\n'
           '- 博烨体育农商银行!F40、!F49「vi设计费用」共 20,000 按规则记了办公费，也可以记营销费用。\n'
           '- 博烨体育农商银行!F127「采购场地长条椅子」、!F65「王庄篮球馆采购二手电脑」记了办公费，也可以记低值易耗品或王庄篮球馆。\n'
           '- 力喆产业农商银行!F15、!F16 是租金、物业、水电、税费合在一笔的收款，原记房租收入。摘要逐项写了金额'
           '（!F15：租金 4,100、1月物业及12月水费 500、电费 119.91、税费 47.20；!F16：租金 4,663.80、物业 886.40、水费 200、'
           '电费 120.75、税费 58.71），要不要按摘要拆成房租、物业、水电收入？「物业及水费 500」没分开，税费那部分记什么？')
    q('原有收费项，请顺便看一眼（不算错，先按原样/规则记）', 0, '博烨私户!F17、博烨体育农商银行!F40/F49/F65/F127、力喆产业农商银行!F15/F16',
      L12)
    q('现金表的钱去哪了', 73393 + 1200, '古玩城现金、易佰兴现金',
      '- 两张现金表只有收入（古玩城 73,393.00、易佰兴 1,200.00），没有存银行或交给谁的记录，所以现金余额一直涨。以后交现金、存银行请录一笔内部转账。\n'
      '- 古玩城现金!F25 7/3 郑用辉「C133 C134 5月租金及4月水电费」1,000，古玩城农商!F676 5/31 已经收过「C133 C134 5月租金」1,000，是不是重收了？')
    q('只有文字没金额的行没导', 9024, '易佰兴建行!F155、易佰兴私户!F775',
      '- 易佰兴建行!F155 6/30 魏龙 C215-216「以押抵租原押金9024元，抵扣…租金7308元、物业1492元…」只写了摘要没填金额。'
      '要不要像古玩城那样补「租金/物业收入 + 退押金」两行（余额不变）？\n'
      '- 易佰兴私户!F775 9/21「利息收入」没填金额。\n'
      '- 易佰兴农商银行外!E12 1/7「退垃圾清运费及装修保证金」：租金栏填了 3,000，收入列却是空的、支出 3,000；原表余额没算这 3,000 收入，'
      '总表也只记了支出 3,000（退押金）。租金栏那个数是不是填错了？')
    mixed = [r for r in rows if r['inc'] and r['exp'] and (r['sheet'], r['row']) not in
             {(x['表'], int(x['行'])) for x in ctx['dups']['古玩城农商银行_同一行收支都记']['行']}]
    msk = [k for k in MEMO_SPLIT if k not in {(r['sheet'], r['row']) for r in mixed} and
           any(e['收费项'] == '退押金' for e in E if (e['_sheet'], e['_row']) == k)]
    q('押金抵租金、物业、水电（以押抵租）%d 行：抵掉的部分照记收入，另记退押金' % (len(mixed) + len(msk)),
      sum(r['exp'] for r in mixed) + sum(e['支出'] for e in E if (e['_sheet'], e['_row']) in set(msk) and e['收费项'] == '退押金'),
      '、'.join(['%s!F%d' % (r['sheet'].strip(), r['row']) for r in mixed] + ['%s!F%d' % (k[0].strip(), k[1]) for k in msk]),
      '- 原表同一行里记了租金等收入和押金支出的 %d 行（钱没动，是押金抵租金），原表余额两边都算了。总表拆成收入（按分栏定收费项）'
      '和「退押金」两头，金额不变、余额不变。其中古玩城农商!F785 原表分栏空着、收入列整笔 8,268，按摘要拆成房租 7,158、物业 300、水电 810。\n'
      '- 原表只记了实际收付、押金抵掉的收入没记的 %d 行，按摘要补拆（余额不变）：古玩城农商!F96（退 B104 押金 7,800 扣 1 月物业 150、'
      '12 月水电 204，实退 7,446）；易佰兴建行!F137（B314-315 原押金 7,954.64 抵 B219-220 租金，补交的 6,565.36 分成租金 4,453.36、'
      '物业 2,112）；易佰兴建行!F146（原押金抵 6/1–6/23 租金 948.77，本账户实退 5,151.23）。\n'
      '- 押金冲抵的那部分不是真付款，看「退押金」合计时要知道里面含这些冲抵。' % (len(mixed), len(msk)))
    q('缺日的 4 行按前后定了日期', 1000 + 5370 + 280 + 1000,
      '古玩城农商银行!F97、!F98，古玩城现金!F51，易佰兴私户!F814',
      '\n'.join('- %s!F%d → %d/%d：%s' % (s.strip(), r, ctx_m(rows, s, r), d, why) for (s, r), (d, why) in MISSING_DAY.items()))

    Q.sort(key=lambda x: -x[1])
    L = ['# 待确认（A074 资金台帐导入）', '',
         '按涉及金额从大到小排。每条都已经按「先这样处理」导进总表了，你确认或告诉我们改法后再调整。', '']
    for i, (t, m, w, b) in enumerate(Q, 1):
        L += ['## %d. %s' % (i, t), '', '涉及金额：%s　　涉及：%s' % (fmt(m) if m else '—', w), '', b, '']
    with open(os.path.join(out, 'questions.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')


def ctx_m(rows, s, r):
    for x in rows:
        if x['sheet'] == s and x['row'] == r:
            return x['m']
    return 0


def print_summary(E, CH, chk):
    src = collections.Counter(e['来源标记'] for e in E)
    print('rows', len(E), dict(src))
    print('changes', len(CH), dict(collections.Counter(x['类别'] for x in CH)))
    print('attr', chk['attr'])
    print('internal', chk['internal'])
    print('open/end', chk['tot_open'], chk['tot_end'])
    print('FAIL' if chk['fail'] else 'ALL PASS', chk['fail'])
