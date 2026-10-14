# -*- coding: utf-8 -*-
"""A074 集团多账户资金台帐：把 2026年日记账 26 张账户表清洗成一张总表的流水。

用法（-I 下也能跑）：
    python3 -I clean.py [--src 日记账.xlsx] [--und 读懂报告目录] [--out 输出目录]
读：日记账（只取缓存值）、读懂报告的 journal_sheets.json / items_map.json / dups.json
写（out/）：accounts.json companies.json items.json parties.json rules.json entries.json
           changes.json changes.md questions.md check.md
"""
import argparse
import collections
import datetime
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from conf import (COMPANIES, ACCOUNTS, ACCOUNT_HOLDER, LAYOUT, CAT_ITEM, ITEMS, ITEM_SRC, NOT_IMPORT,  # noqa: E402
                  MISSING_DAY, NONCASH_OFFSET, ADD_ROWS, MANUAL_ITEM, PARTY_ALIAS_FIX, MEMO_SPLIT, MEMO_FEE,
                  MEMO_FEE_RE, STRAY_COLS, ITEM_FIX_BASIS)
from rules_kw import RULES, PARTY_RULES  # noqa: E402
from extract import read_journal, read_guwan_bank_opening, read_stray  # noqa: E402
import report  # noqa: E402

_SCR = '/tmp/claude-0/-home-user-excel/6401330f-a6e0-50ea-b764-91473ebf0b63/scratchpad/a074'
DEF_SRC = _SCR + '/src/journal.xlsx'
DEF_UND = _SCR + '/understand'
DEF_OUT = os.path.join(HERE, 'out')
YEAR = 2026

SHEET_ACC = {a[0]: a[1] for a in ACCOUNTS}
SHEET_CO = {a[0]: a[2] for a in ACCOUNTS}
SHEET_BANK = {a[0]: a[3] for a in ACCOUNTS}
ACC_IDX = {a[0]: i for i, a in enumerate(ACCOUNTS)}
ITEM_ATTR = {n: a for n, a, *_ in ITEMS}
ITEM_HINT = {n: h for n, a, g, h, s in ITEMS}

CH = []          # 整理记录
PARTY_ALIAS_TXT = '统一写法（旧写法记在【基础资料】⑤收付人的「别名」里）'
PARTY_SAME = '去掉首尾空格/换行'


def r2(x):
    return round(x + 0.0, 2)


def fmt(x):
    return '{:,.2f}'.format(x)


def pos(sheet, row, sub=None, col=None):
    col = col or (LAYOUT[sheet]['memo'])
    s = '%s!%s%d' % (sheet.strip(), col, row)
    return s + ('-%d' % sub if sub else '')


def log(p, field, old, new, basis, kind):
    if field != '整行' and str(old) == str(new):
        return          # 没有实际改动的不记
    CH.append(dict(原表位置=p, 字段=field, 原值=old, 新值=new, 依据=basis, 类别=kind))


def clean_text(s):
    s = (s or '').replace('&quot;', '"').replace('&amp;', '&')
    s = s.replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
    s = re.sub(r' {2,}', ' ', s)
    return s.strip()


# ============================================================ 内部转账配对
def build_internal(dups):
    """(原表, 行, 方向) → [(对方原表, 对方行, 金额, 说明)]"""
    M = collections.defaultdict(list)
    D = {'OUT': '支', 'IN': '收'}

    def k(t):
        return (t['表'], int(t['行']), D[t['方向']])

    for p in dups['内部转账_一对一']:
        a, b = k(p['转出']), k(p['转入'])
        M[a].append((b[0], b[1], p['金额'], ''))
        M[b].append((a[0], a[1], p['金额'], ''))
    for g in dups['内部转账_一对多_净额_缺日期']:
        s = k(g['单笔'])
        for x in g['对方多笔']:
            t = k(x)
            M[s].append((t[0], t[1], x['金额'], g['类型']))          # 对方那一笔的金额
            M[t].append((s[0], s[1], g['单笔']['金额'], g['类型']))  # 对方（单笔）的金额
    for p in dups['内部转账_金额不一致']:
        a, b = k(p['转出']), k(p['转入'])
        M[a].append((b[0], b[1], p['转入']['金额'], '金额不一致'))
        M[b].append((a[0], a[1], p['转出']['金额'], '金额不一致'))
    for p in dups['内部转账_经未建表个人账户中转']:
        a, b = k(p['转出']), k(p['转入'])
        via = '王晓华建行' if '王晓华' in (p['转出']['摘要'] + p['转出']['收付人']) else '吴海萍个人账户'
        M[a].append((b[0], b[1], p['金额'], '经' + via + '中转'))
        M[b].append((a[0], a[1], p['金额'], '经' + via + '中转'))
    return M


# ============================================================ 规则
def match_rule(direction, company, memo, party, account=''):
    """适用范围（comp）空＝全部；公司简称或账户名（v3.1，跟基础资料⑥一样）"""
    mu, pu = (memo or '').upper(), (party or '').upper()
    for rid, rdir, comp, where, kws, item, rparty, note in RULES:
        if rdir != '都' and rdir != direction:
            continue
        if comp and comp not in (company, account):
            continue
        text = mu if where == '摘要' else pu
        if not text:
            continue
        for kw in kws:
            if kw.upper() in text:
                return dict(rid=rid, kw=kw, item=item, party=rparty, where=where)
    return None


def match_party_rule(direction, company, memo, account=''):
    mu = (memo or '').upper()
    for rid, rdir, comp, kws, party in PARTY_RULES:
        if rdir != '都' and rdir != direction:
            continue
        if comp and comp not in (company, account):
            continue
        for kw in kws:
            if kw.upper() in mu:
                return rid, kw, party
    return None


CHANNEL_RE = r'收单|过渡户|清算|财付通|支付|备付金|结算中心|银行|微信|抖音|星驿|汇付|速停车|科拓|路桥|物联'
# 铺位号 A101、B 310、C-215：前面不能紧挨字母数字或车牌省份简称，后面不能紧跟字母数字（排除「闽A203H3」这种车牌）
PROV = '京津沪渝冀晋蒙辽吉黑苏浙皖闽赣鲁豫鄂湘粤桂琼川贵云藏陕甘青宁新'
SHOP_RE = re.compile(r'(?<![A-Za-z0-9%s])([A-Da-d])\s?-?(\d{3})(?![\dA-Za-z])' % PROV)
PLATE_RE = re.compile(r'[%s][A-Z][0-9A-Z]{4,6}' % PROV)
TEMP_PARK_RE = r'临时车|临时停车|临停|停车券'     # 跟 rules_kw.PARTY_RULES 的 P07 一样
MONTH_PARK_RE = r'包月|月租|期间|年度缴费|年缴|\d{1,2}\.\d{1,2}\s*[-－至]\s*\d{1,2}\.\d{1,2}'


# ============================================================ 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--src', default=DEF_SRC)
    ap.add_argument('--und', default=DEF_UND)
    ap.add_argument('--out', default=DEF_OUT)
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)

    meta_all = json.load(open(os.path.join(a.und, 'journal_sheets.json'), encoding='utf-8'))
    meta = meta_all['账户表']
    im = json.load(open(os.path.join(a.und, 'items_map.json'), encoding='utf-8'))
    dups = json.load(open(os.path.join(a.und, 'dups.json'), encoding='utf-8'))

    rows, openings, skipped = read_journal(a.src, meta)
    gw = read_guwan_bank_opening(a.src)
    old = openings['古玩城农商银行']
    openings['古玩城农商银行'] = (gw, '古玩城农商银行!R5（用户手算的银行余额列）')
    log('古玩城农商银行!P5', '开账余额', fmt(old[0]), fmt(gw),
        'P 列期初含约 7,290,556.40 非银行数；用银行口径 R5（=Sheet1!B2）。按此口径重算的各月末和用户 R/S 列逐月一致',
        '期初')

    # ---------------- 不导的行
    stray_cols = collections.OrderedDict()
    for sh, cell, v, kind in read_stray(a.src, meta):
        col = re.match(r'[A-Z]+', cell).group(0)
        if kind == '表外' and (sh, col) in STRAY_COLS:
            stray_cols.setdefault((sh, col), []).append((cell, v))
            continue
        log('%s!%s' % (sh.strip(), cell), '单元格', v, '（不导）',
            '原表数据区下面的残留（同一片还有 #REF! 公式残留），不在任何流水行里，原表余额也没算它' if kind == '残留' else
            '表格外面的孤立数（在本表表头以外的列，像是草稿），不是流水，原表余额也没算它', '不导')
    for (sh, col), cells in stray_cols.items():
        nn = sum(1 for _c, v in cells if isinstance(v, (int, float)))
        log('%s!%s%s:%s%s' % (sh.strip(), col, cells[0][0][len(col):], col, cells[-1][0][len(col):]), '整列',
            '%d 格（数字 %d、文字 %d）' % (len(cells), nn, len(cells) - nn), '（不导）', STRAY_COLS[(sh, col)], '不导')
    for r in skipped:
        log(pos(r['sheet'], r['row']), '整行', clean_text(r['memo_raw']) or clean_text(r['party_raw']), '（不导）',
            '原表这一行没有金额，只有文字', '不导')

    internal = build_internal(dups)
    virtual_rows = {(x['表'], int(x['行'])) for x in dups['古玩城农商银行_同一行收支都记']['行']}
    fix = {(x['sheet'], int(x['row'])): x for x in im['item_row_fix']}
    alias = {k: v['to'] for k, v in im['item_alias'].items()}
    party_alias = im['party_alias']

    # ---------------- 生成每一笔（leg）
    legs = []
    for r in rows:
        key = (r['sheet'], r['row'])
        if key in NOT_IMPORT:
            amt = r['inc'] or r['exp']
            log(pos(*key), '整行', '%s %s' % (clean_text(r['memo_raw']), fmt(amt)), '（不导）', NOT_IMPORT[key], '不导')
            continue
        lay = LAYOUT[r['sheet']]
        L = []          # (方向, 金额, kind, 分栏名)
        if key in MEMO_SPLIT:
            assert key not in virtual_rows and key not in NONCASH_OFFSET, key
            ms = MEMO_SPLIT[key]
            want = r2(r['inc'] - r['exp'])
            got = r2(sum(v if d_ == '收' else -v for d_, _it, v, _n in ms['legs']))
            assert abs(want - got) < 0.005, (key, want, got)
            for d_, it_, v, nt in ms['legs']:
                L.append((d_, v, 'msplit', (it_, nt)))
        elif key in MEMO_FEE:
            mf = re.search(MEMO_FEE_RE, clean_text(r['memo_raw']))
            gross, fee_ = float(mf.group(1).replace(',', '')), float(mf.group(2))
            assert abs(r2(gross - fee_) - r['inc']) < 0.005 and not r['exp'], key
            L.append(('收', gross, 'gross', None))
            L.append(('支', fee_, 'fee', None))
            log(pos(*key), '收入', fmt(r['inc']), '毛额 %s ／ 手续费 %s' % (fmt(gross), fmt(fee_)),
                '摘要写明「%s」，原表收入列记的是扣掉微信提现手续费后的净额；按口径（手续费单列）收入记毛额，手续费单独一行' %
                mf.group(0), '金额方向')
        elif lay['kind'] == 'cat':
            cats = r['cats']
            if cats:
                s = r2(sum(cats.values()) - r['fee'])
                if abs(s - r['inc']) < 0.005:
                    for nm, v in cats.items():
                        L.append(('收', v, 'cat', nm))
                    if r['fee']:
                        L.append(('支', r['fee'], 'fee', None))
                else:
                    log(pos(*key), '收入', '分栏合计 %s，收入列 %s' % (fmt(sum(cats.values())), fmt(r['inc'])),
                        '按收入列 %s' % fmt(r['inc']),
                        '分栏数和收入列对不上，原表余额按收入列算，分栏数不导（%s）' %
                        '、'.join('%s%s=%s' % (r['catcols'][k], r['row'], fmt(v)) for k, v in cats.items()),
                        '金额方向')
                    if r['inc']:
                        L.append(('收', r['inc'], 'main', None))
            elif r['inc']:
                L.append(('收', r['inc'], 'main', None))
            if r['exp']:
                L.append(('支', r['exp'], 'main', None))
        elif lay['kind'] == 'gross':
            if r['gross'] and r['fee'] and abs(r2(r['gross'] - r['fee']) - r['inc']) < 0.005:
                L.append(('收', r['gross'], 'gross', None))
                L.append(('支', r['fee'], 'fee', None))
            elif r['inc']:
                if r['gross'] and abs(r['gross'] - r['inc']) >= 0.005:
                    log(pos(*key), '收入', '毛 %s 手续费 %s 净 %s' % (fmt(r['gross']), fmt(r['fee']), fmt(r['inc'])),
                        '按净收入 %s' % fmt(r['inc']), '毛收入、手续费、净收入对不上，按原表余额用的净收入', '金额方向')
                L.append(('收', r['inc'], 'main', None))
            if r['exp']:
                L.append(('支', r['exp'], 'main', None))
        else:
            if r['inc']:
                L.append(('收', r['inc'], 'main', None))
            if r['exp']:
                L.append(('支', r['exp'], 'main', None))

        # 古玩城「吴海萍转入徐洁账户」收支同记：收入一头是虚的
        if key in virtual_rows:
            L = [x for x in L if x[0] != '收']
            log(pos(*key), '收入', fmt(r['inc']), '（不导）',
                '同一行收支都记的转账，收入一头是虚的（钱是转出去的），只导支出；和对方徐洁卡的收入配对', '不导')
        # 以押抵租不过银行：收入照导，另加一行押金冲抵
        if key in NONCASH_OFFSET:
            tot = r2(sum(x[1] for x in L if x[0] == '收'))
            L.append(('支', tot, 'offset', None))
        # 负数：收入负数改记支出，支出负数改记收入（收费项不变，相当于冲减）
        L2 = []
        for d_, v, kind, cat in L:
            if v < 0:
                nd = '支' if d_ == '收' else '收'
                log(pos(*key), '金额方向', '%s %s' % (d_ + '入' if d_ == '收' else d_ + '出', fmt(v)),
                    '%s %s' % ('支出' if nd == '支' else '收入', fmt(-v)),
                    '新模板金额不能填负数；改记反方向，收费项不变，汇总时照样冲减该项', '金额方向')
                L2.append((nd, -v, kind, cat, True))
            else:
                L2.append((d_, v, kind, cat, False))
        L = L2

        # 内部转账：对方跨两个账户的，按对方金额拆开
        L3 = []
        for d_, v, kind, cat, neg in L:
            peers = internal.get((r['sheet'], r['row'], d_), []) if kind in ('main', 'gross') else []
            accs = collections.OrderedDict()
            for ps, pr, pa, note in peers:
                accs.setdefault(ps, []).append((pr, pa, note))
            if len(accs) > 1:
                tot = r2(sum(pa for v2 in accs.values() for _, pa, _ in v2))
                if abs(tot - v) < 0.005:
                    for ps, lst in accs.items():
                        L3.append(dict(dir=d_, amt=r2(sum(x[1] for x in lst)), kind=kind, cat=cat, neg=neg,
                                       peers=[(ps,) + x for x in lst], splitpeer=True))
                    continue
            L3.append(dict(dir=d_, amt=v, kind=kind, cat=cat, neg=neg,
                           peers=[(ps,) + x for ps, lst in accs.items() for x in lst], splitpeer=False))
        n = len(L3)
        for i, x in enumerate(L3, 1):
            x.update(sheet=r['sheet'], row=r['row'], sub=(i if n > 1 else None), r=r)
            legs.append(x)

    # 补记
    for ad in ADD_ROWS:
        legs.append(dict(dir='支' if ad['exp'] else '收', amt=ad['exp'] or ad['inc'], kind='add', cat=None, neg=False,
                         peers=[(ad['peer_sheet'], ad['peer_row'], ad['exp'] or ad['inc'], '补记')], splitpeer=False,
                         sheet=ad['sheet'], row=10 ** 6, sub=None,
                         r=dict(sheet=ad['sheet'], row=10 ** 6, m=ad['m'], d=ad['d'], item='', party_raw=ad['party'],
                                memo_raw=ad['memo'], note='', owner='', extra={}, memocol='F'), add=ad))
        # 对方那一笔也要标成内部转账
        internal[(ad['peer_sheet'], ad['peer_row'], '收' if ad['exp'] else '支')].append(
            (ad['sheet'], 10 ** 6, ad['exp'] or ad['inc'], '补记'))
        log('补记（%s 9/20）' % SHEET_ACC[ad['sheet']], '整行', '（原表没有）',
            '支出 %s 内部转账，对方 %s' % (fmt(ad['exp']), SHEET_ACC[ad['peer_sheet']]),
            '对方 %s 收到「%s」%s（%s!F%d），本账户漏记转出；补记后期末 %s' % (
                SHEET_ACC[ad['peer_sheet']], '博烨一般户（羽毛球馆）', fmt(ad['exp']), ad['peer_sheet'], ad['peer_row'],
                '0.35'), '补记')
        # 补到原来给 peer 那一笔
    # peer 那一笔的 peers 已经在生成时取过了，这里补上
    for x in legs:
        if x['kind'] in ('main', 'gross') and not x['peers']:
            pk = (x['sheet'], x['row'], x['dir'])
            if pk in internal and any(p[3] == '补记' for p in internal[pk]):
                x['peers'] = [(p[0], p[1], p[2], p[3]) for p in internal[pk]]

    # ---------------- 收付人 统一写法 + 补
    def norm_party(raw, sheet):
        p0 = raw or ''
        p = clean_text(p0)
        p2 = p.replace('(', '（').replace(')', '）')
        for k in (p0, p, p2):
            if k in PARTY_ALIAS_FIX:
                v = PARTY_ALIAS_FIX[k][0]
                return v, (None if v == p0 else PARTY_SAME if v == p else '统一写法（更正：%s）' % PARTY_ALIAS_FIX[k][1])
        for k in (p0, p, p2):
            if k in party_alias:
                v = party_alias[k]
                return v, (None if v == p0 else PARTY_SAME if v == p else PARTY_ALIAS_TXT)
        if p == '银行':
            return SHEET_BANK[sheet], '「银行」按本账户开户行'
        if p != p0:
            return p, '去掉首尾空格/换行'
        return p, None

    # 铺位 → 收付人（按日期）
    shop_hist = collections.defaultdict(list)     # (场所, 铺位) → [(日期序, 收付人)]

    def place(sheet):
        return '古玩城' if sheet.startswith('古玩城') else ('易佰兴' if sheet.startswith('易佰兴') else None)

    def dkey(r):
        m = r['m'] or 0
        d = r['d'] or MISSING_DAY.get((r['sheet'], r['row']), (0,))[0]
        return m * 100 + d

    for x in legs:
        r = x['r']
        pl = place(r['sheet'])
        if not pl:
            continue
        pn, _ = norm_party(r['party_raw'], r['sheet'])
        if not pn:
            continue
        for m_ in SHOP_RE.finditer(clean_text(r['memo_raw'])):
            shop_hist[(pl, m_.group(1).upper() + m_.group(2))].append((dkey(r), r['row'], pn))

    def shop_party(r):
        pl = place(r['sheet'])
        if not pl:
            return None
        memo = clean_text(r['memo_raw'])
        for m_ in SHOP_RE.finditer(memo):
            shop = m_.group(1).upper() + m_.group(2)
            h = shop_hist.get((pl, shop))
            if not h:
                continue
            me = dkey(r)
            before = [t for t in h if t[0] <= me]
            if before:
                t = max(before)
                return shop, t[2], '同场所 %s 铺位最近一笔（%d/%d）的收付人' % (shop, t[0] // 100, t[0] % 100)
            t = min(h)
            return shop, t[2], '同场所 %s 铺位之后最近一笔（%d/%d）的收付人' % (shop, t[0] // 100, t[0] % 100)
        return None

    # ---------------- 每一笔定收费项、收付人
    row_dirs = collections.defaultdict(set)
    for x in legs:
        row_dirs[(x['sheet'], x['row'])].add(x['dir'])
    entries = []
    stat = collections.Counter()
    for x in legs:
        r = x['r']
        sheet, row, sub = x['sheet'], x['row'], x['sub']
        co = SHEET_CO[sheet]
        P = pos(sheet, row, sub, r.get('memocol')) if x['kind'] != 'add' else '补记（对应 %s!F%d）' % (
            x['add']['peer_sheet'], x['add']['peer_row'])
        P0 = pos(sheet, row, None, r.get('memocol')) if x['kind'] != 'add' else P
        memo = clean_text(r['memo_raw'])
        if x['kind'] != 'add' and (not sub or sub == 1) and memo != (r['memo_raw'] or ''):
            stat['摘要去空格'] += 1
            log(P0, '摘要', r['memo_raw'], memo, '去掉首尾空格、换行和连续空格' + (
                '；原表摘要里的引号是网页代码的写法（原来那一栏能看到），换回了普通引号' if re.search(r'&quot;|&amp;', r['memo_raw'] or '') else ''), '摘要')
        orig_item = r['item']
        notes = []
        if (sheet, row) in NONCASH_OFFSET:
            notes.append('以押抵租，不过银行')
        if x['kind'] == 'msplit' and x['cat'][1]:
            notes.append(x['cat'][1])
        if r.get('note') and not any(clean_text(r['note']) in n for n in notes):
            notes.append(clean_text(r['note']))
        for k, v in r.get('extra', {}).items():
            notes.append('%s：%s' % (k, v))
        # ---- 日期
        m, d = r['m'], r['d']
        if d is None:
            d, why = MISSING_DAY[(sheet, row)]
            if not sub or sub == 1:
                log(P0, '日期', '%d月（日空）' % m, '%d-%02d-%02d' % (YEAR, m, d), why, '日期')
        date = datetime.date(YEAR, m, d).isoformat()
        # ---- 收付人
        party, pwhy = norm_party(r['party_raw'], sheet)
        if pwhy and (not sub or sub == 1):
            log(P0, '收付人', r['party_raw'], party, pwhy, '收付人')
        # ---- D 列能不能用在这一笔上
        use_orig = ''
        if orig_item and x['kind'] in ('main', 'gross'):
            h = ITEM_HINT.get(alias.get(orig_item, orig_item))
            dirs_in_row = row_dirs[(sheet, row)]
            if len(dirs_in_row) > 1 and h in ('收', '支') and h != x['dir']:
                use_orig = ''   # 收支同行：D 列只用在方向对得上的那一笔
            else:
                use_orig = orig_item
        elif orig_item and x['kind'] == 'cat':
            ci = CAT_ITEM[x['cat']]
            if orig_item != ci and len(r['cats']) == 1 and ITEM_HINT.get(orig_item) == '收':
                log(P, '收费项', orig_item, ci, '按分栏「%s」列' % x['cat'], '收费项')

        # ---- 收费项
        item = src = basis = None
        peer_acc = ''
        if x['peers'] and x['kind'] in ('main', 'gross', 'add'):
            item, src = '内部转账', ('补记' if x['kind'] == 'add' else '逐行改')
            paccs = []
            for p in x['peers']:
                if SHEET_ACC[p[0]] not in paccs:
                    paccs.append(SHEET_ACC[p[0]])
            peer_acc = paccs[0]
            if len(paccs) > 1:
                notes.append('对方另有：' + '、'.join(paccs[1:]))
            pnotes = sorted({p[3] for p in x['peers'] if p[3] and p[3] not in ('一对多', '补记')})
            tags = []
            for p in x['peers']:
                tags.append('%s!F%s%s' % (p[0].strip(), p[1] if p[1] < 10 ** 6 else '(补记)',
                                          '' if abs(p[2] - x['amt']) < 0.005 else ' %s' % fmt(p[2])))
            basis = '配对：' + '、'.join(tags) + ('（%s）' % '；'.join(pnotes) if pnotes else '')
            for pn in pnotes:
                if pn.startswith('经'):
                    notes.append(pn)
                if pn == '金额不一致':
                    notes.append('和对方金额差 %s，待核对' % fmt(abs(x['amt'] - x['peers'][0][2])))
            if x['kind'] != 'add':
                if orig_item and orig_item != '内部转账':
                    notes.insert(0, '原收费项：' + orig_item)
                log(P, '收费项', orig_item or '（空）', '内部转账', basis, '收费项-内部转账')
            if x['kind'] == 'add':
                notes.append(x['add']['note'])
        elif x['kind'] == 'fee':
            item, src, basis = '银行手续费', '拆行', '原表手续费列，从收入里扣的那部分单独记一笔支出'
        elif x['kind'] == 'offset':
            item, src, basis = '退押金', '拆行', '以押抵租，押金冲抵租金'
        elif x['kind'] == 'msplit':
            item, src = x['cat'][0], '逐行改'
            basis = '按摘要逐项拆：' + MEMO_SPLIT[(sheet, row)]['basis']
            log(P, '收费项', orig_item or '（空）', item, basis, '收费项-逐行改')
        elif x['kind'] == 'cat':
            item, src, basis = CAT_ITEM[x['cat']], '分栏', '分栏「%s」列（%s%d）' % (x['cat'], r['catcols'][x['cat']], row)
            log(P, '收费项', orig_item or '（空）', item, basis, '收费项-分栏')
        elif (sheet, row, x['dir']) in MANUAL_ITEM and x['kind'] in ('main', 'gross'):
            item, basis = MANUAL_ITEM[(sheet, row, x['dir'])]
            src = '原有' if item == orig_item else ('待定' if item == '待定' else '逐行改')
            log(P, '收费项', orig_item or '（空）', item, '人工判断：' + basis,
                '收费项-待定' if item == '待定' else '收费项-逐行改')
        elif use_orig:
            key = (sheet, row)
            if key in fix and fix[key]['old'] == use_orig:
                item, src = fix[key]['new'], '逐行改'
                basis = ITEM_FIX_BASIS.get(key, fix[key]['basis']) + '（把握：%s）' % fix[key]['confidence']
                log(P, '收费项', use_orig, item, basis, '收费项-逐行改')
            elif use_orig in alias:
                item, src = alias[use_orig], '统一写法'
                basis = '统一成清单里的写法（旧写法记在【基础资料】④收费项的「别名」里）：' + im['item_alias'][use_orig]['basis']
                log(P, '收费项', use_orig, item, basis, '收费项-统一写法')
            else:
                item, src, basis = use_orig, '原有', ''
                h = ITEM_HINT.get(item)
                if h is None:
                    raise SystemExit('收费项不在清单：%s %s' % (P, item))
                if h in ('收', '支') and h != x['dir'] and not x['neg']:
                    raise SystemExit('收费项方向不对：%s %s %s' % (P, item, x['dir']))
        mk = (sheet, row, x['dir'])
        if item is None and mk in MANUAL_ITEM:
            item, basis = MANUAL_ITEM[mk]
            src = '待定' if item == '待定' else '逐行改'
            log(P, '收费项', orig_item or '（空）', item, '人工判断：' + basis,
                '收费项-待定' if item == '待定' else '收费项-逐行改')
        if item is None:
            # 分栏表里没分栏的收入、毛收入式里的毛收入、其他空项 → 摘要规则
            mr = match_rule(x['dir'], co, memo, party, SHEET_ACC[sheet])
            if mr:
                item, src = mr['item'], '摘要规则'
                basis = '规则 %s：%s含「%s」' % (mr['rid'], mr['where'], mr['kw'])
                if item == '内部转账':
                    item = '往来款'
                    basis += '，像是账户互转，但本册里找不到对方那一笔（对方不在本册或漏记），按往来款'
            elif x['kind'] == 'gross' and sheet == '艾泊烨农商银行':
                item, src, basis = '停车费收入', '分栏', '原表 G 列「停车收入」'
            else:
                item, src = '待定', '待定'
                basis = '规则判断不了'
            log(P, '收费项', orig_item or '（空）', item, basis,
                '收费项-摘要规则' if src in ('摘要规则', '分栏') and x['kind'] != 'cat' else '收费项-待定')
        stat['来源:' + src] += 1

        # ---- 收付人补
        defer = None    # 月租车/只有车牌：等所有行定完收付人后再按车牌查（见下面「车牌查收付人」）
        if not party:
            fill = None
            if item == '内部转账' and peer_acc:
                ps = x['peers'][0][0]
                fill = (ACCOUNT_HOLDER[ps], '内部转账，按对方账户的户名')
            if not fill:
                mm = match_party_rule(x['dir'], co, memo, SHEET_ACC[sheet])
                if mm and mm[0] == 'P07c':
                    defer = mm
                elif mm:
                    fill = (mm[2], '规则 %s：摘要含「%s」' % (mm[0], mm[1]))
            if not fill and re.search(r'利息', memo):
                fill = ('财付通（微信）' if '微信' in memo else ('招商银行' if '招行' in memo else SHEET_BANK[sheet]),
                        '利息收入，收付人填开户行')
            if not fill and re.search(r'手续费|短信(通)?服务费|回单|网银', memo) and '中国移动' not in memo:
                fill = (SHEET_BANK[sheet], '银行收费，收付人填开户行')
            if not fill:
                sp = shop_party(r)
                if sp:
                    fill = (sp[1], sp[2])
                    defer = None
            if not fill and not defer and sheet.startswith('易佰兴') and x['dir'] == '收' and PLATE_RE.search(memo):
                defer = ('P07', '车牌', '')
            if not fill and not defer and sheet.startswith('古玩城') and x['dir'] == '支':
                fill = ('林淑琴', '古玩城支出的经办人，沿用原表写法')
            if fill:
                party = fill[0]
                if not sub or sub == 1 or item == '内部转账':
                    log(P if sub else P0, '收付人', '（空）', party, fill[1], '收付人-补')
        if x['kind'] == 'fee':
            # 手续费是收款通道/银行扣的，收付人填通道或开户行，不填交款的租户
            if party and re.search(CHANNEL_RE, party):
                fp = party
            elif '微信' in (r.get('note') or '') or '微信' in memo:
                fp = '财付通（微信）'
            else:
                fp = SHEET_BANK[sheet]
            if fp != party:
                log(P, '收付人', party or '（空）', fp,
                    '手续费那一笔：收付人填收款通道或开户行（%s），不填交款人' % ('原表备注写微信' if fp == '财付通（微信）' else '开户行'),
                    '收付人-补')
                party = fp

        owner = r.get('owner', '')
        e = dict(日期=date, 账户=SHEET_ACC[sheet], 收费项=item, 属性=(ITEM_ATTR[item] if item != '待定' else
                                                                ('收入' if x['dir'] == '收' else '支出')),
                 收付人=party, 摘要=memo, 收入=r2(x['amt']) if x['dir'] == '收' else 0.0,
                 支出=r2(x['amt']) if x['dir'] == '支' else 0.0, 对方账户=peer_acc, 归属公司=owner,
                 备注='；'.join(n for n in notes if n), 原表位置=P, 来源标记=src,
                 _sheet=sheet, _row=row, _sub=sub or 0, _kind=x['kind'], _basis=basis or '', _defer=defer)
        if x['kind'] == 'add':
            e['摘要'] = x['add']['memo']
        entries.append(e)

    # ---------------- 车牌查收付人（规则 P07）：月租车、只有车牌的停车费
    # 先看同一车牌在别的行（原表本来就有收付人的，或按品牌、铺位补出来的）是不是只有一个收付人，是就用它；
    # 查不到的：摘要写临时/临停/停车券＝临时停车散客，写月租/包月/某段期间＝月租车散户
    plate_party = collections.defaultdict(lambda: collections.OrderedDict())
    for e in entries:
        if (e['_sheet'].startswith('易佰兴') and e['收付人'] and not e['_defer'] and e['_kind'] != 'fee'
                and e['收付人'] not in ('临时停车散客', '月租车散户')):
            for pl in dict.fromkeys(PLATE_RE.findall(e['摘要'])):
                plate_party[pl].setdefault(e['收付人'], []).append(e['原表位置'])
    for e in entries:
        if not e['_defer']:
            continue
        memo = e['摘要']
        found, unknown = collections.OrderedDict(), []
        for pl in dict.fromkeys(PLATE_RE.findall(memo)):
            d = plate_party.get(pl)
            if d and len(d) == 1:
                nm, ps = next(iter(d.items()))
                found.setdefault(nm, []).append('%s 见 %s' % (pl, '、'.join(ps[:3]) + (' 等 %d 行' % len(ps) if len(ps) > 3 else '')))
            else:
                unknown.append(pl + ('（有 %d 个收付人）' % len(d) if d else ''))
        if len(found) == 1:
            nm, why = next(iter(found.items()))
            party, why = nm, '同一车牌在别的行只有这一个收付人（%s）' % '；'.join(why) + (
                '；另 %s 查不到' % '、'.join(unknown) if unknown else '')
        elif re.search(TEMP_PARK_RE, memo):
            party, why = '临时停车散客', '规则 P07：摘要写临时车/临时停车/临停/停车券，车牌查不到车主'
        elif re.search(MONTH_PARK_RE, memo):
            party, why = '月租车散户', '规则 P07c：月租/包月/按期间交的停车费，车牌%s查不到车主（不叫「临时」）' % (
                '、'.join(unknown) if unknown else '')
        else:
            party, why = '临时停车散客', '摘要只有车牌，车牌查不到车主，按临时停车记'
        e['收付人'] = party
        e['_p07'] = why
        if e['_sub'] <= 1:
            log(e['原表位置'] if e['_sub'] else e['原表位置'], '收付人', '（空）', party, why, '收付人-补')

    # ---------------- 一行拆成几行的，每个原表行记一条
    grp = collections.OrderedDict()
    for e in entries:
        if e['_kind'] != 'add':
            grp.setdefault((e['_sheet'], e['_row']), []).append(e)
    rowmap = {(r['sheet'], r['row']): r for r in rows}
    for (sh, rw), es in grp.items():
        if len(es) < 2:
            continue
        r = rowmap[(sh, rw)]
        kinds = {e['_kind'] for e in es}
        why = []
        if 'cat' in kinds:
            why.append('分栏表一行有几个分栏，每个分栏一行')
        if 'fee' in kinds:
            why.append('从收入里扣的手续费单独一行支出「银行手续费」，收入按扣费前的数记')
        if 'offset' in kinds:
            why.append('以押抵租，钱没过银行（原表余额公式不含这笔收入）：租金等收入照记，押金冲抵另记一行退押金，余额不变')
        if 'msplit' in kinds:
            why.append(MEMO_SPLIT[(sh, rw)].get('why') or
                       '按摘要逐项金额拆（押金抵租金、物业、水电的，抵掉的部分照记收入，押金冲抵另记退押金），余额不变')
        if (len({('收' if e['收入'] else '支') for e in es}) > 1 and 'offset' not in kinds and 'fee' not in kinds
                and 'msplit' not in kinds):
            why.append('同一行既有收入又有支出（以押抵租/押金冲抵），一行只能记一个方向')
        if not why:
            why.append('对方是两个账户各转一笔（%s），按对方拆开，每笔都能对上对方账户' % '、'.join(
                '%s!F%s' % (e['对方账户'], '') for e in es).replace('!F', ''))
        net_new = r2(sum(e['收入'] - e['支出'] for e in es))
        log(pos(sh, rw, None, r.get('memocol')), '整行',
            '收入 %s / 支出 %s%s' % (fmt(r['inc']), fmt(r['exp']), ('（手续费 %s）' % fmt(r['fee'])) if r['fee'] else ''),
            '拆成 %d 行：%s' % (len(es), '；'.join('%s %s %s' % (e['原表位置'].rsplit('-', 1)[-1], e['收费项'],
                                                              fmt(e['收入'] or -e['支出'])) for e in es)),
            '；'.join(why) + '。拆后净额 %s' % fmt(net_new), '拆行')

    # ---------------- 排序、余额
    entries.sort(key=lambda e: (e['日期'], ACC_IDX[e['_sheet']], e['_row'], e['_sub']))
    bal = {s: openings[s][0] for s in SHEET_ACC}
    seq = collections.Counter()
    for i, e in enumerate(entries, 1):
        s = e['_sheet']
        bal[s] = r2(bal[s] + e['收入'] - e['支出'])
        seq[s] += 1
        e['序号'] = i
        e['账户内序号'] = seq[s]
        e['余额'] = bal[s]

    eff_alias = dict(party_alias)
    eff_alias.update({k: v[0] for k, v in PARTY_ALIAS_FIX.items()})
    ctx = dict(meta=meta, meta_all=meta_all, im=im, dups=dups, openings=openings, entries=entries, CH=CH,
               legs=legs, rows=rows, stat=stat, out=a.out, src=a.src, party_alias=eff_alias)
    report.write_all(ctx)


if __name__ == '__main__':
    main()
