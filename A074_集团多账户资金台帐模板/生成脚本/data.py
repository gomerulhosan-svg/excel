# -*- coding: utf-8 -*-
"""A074 集团多账户资金台帐 · 预置的基础资料和导入的流水。

读 生成脚本/导入/out/ 下清洗那边产出的文件（字段按清洗实际产出写；也认 设计.md §9 列的别名）：
  accounts.json   [{"序号","账户名","公司简称","公司全称","银行","户别","户号尾号","持卡人","开账日期","开账余额","状态","原表名","子表名",
                    "期末余额","流水行数","最后一笔日期","说明"}]     顺序＝子表顺序（跟 layout.SUB_SHEETS 一致），≤ 40
  companies.json  [{"顺序","简称","全称","账户数","账户"}]                                        ≤ 30
  items.json      [{"序号","名称","属性","大类","方向提示","说明","原写法别名":[…],"来源","本次笔数",…}]
                  属性：收入/支出/往来/内部转账/按方向（只给「待定」用）                           ≤ 150
  parties.json    [{"名称","类型","别名":[…],"笔数","常见账户"}]                                 ≤ 500
  rules.json      {"说明": …, "收费项规则": [{"序号","规则号","关键词","查找位置","适用方向","适用公司","收费项","收付人","说明"}],
                   "收付人规则": [{…同上，没有收费项…}]}                  两段接起来 ≤ 600 条，从上往下先命中的算
  entries.json    [{"序号","日期":"2026-01-05","账户","收费项","属性","收付人","摘要","收入","支出","余额","对方账户","归属公司","备注",
                    "原表位置","来源标记","账户内序号"}]   已按 日期 → 账户顺序 → 原表行序 排好；收入、支出只填一个（0＝空）
  changes.json    [{"原表位置","字段","原值","新值","依据","类别"}]                                 数据整理记录·修改明细
  questions.md    「## 1. 标题」＋「涉及金额：…　　涉及：…」＋「- 说明」                          数据整理记录·待确认（有 pending.json 就用它）
  check.md        第一张表：每个账户 原表/修正后/新总表 的期初、收、支、期末                        数据整理记录·账户对照（有 acct_check.json 就用它）
  params.json     {"集团名称","管理年度","开账日"}（可没有；开账日默认取 accounts 的开账日期，再没有＝2026-01-01）
读不到（清洗还没产出）时用下面的示例：26 个账户＋二十几个收费项＋三十几笔示例流水（备注「示例」）。
模板这边对清洗产出做的三处调整（设计 v2，见 设计.md §1.3、§4、评审意见处理）：
  ① 收费项属性：大类「项目建设」的项 → 「投资建设」；「待定」→「待分类」（ITEM_ATTR_OVERRIDE；总负责人不同意就把这两条删掉）；
  ② 收费项顺序：收入 → 内部转账 → 支出 → 投资建设 → 往来 → 待分类，同一属性里按清洗统计的笔数从多到少（下拉常用的在上面）；
     没有「账户转入」就补一个（往来，大类「账户并入」：年中新开的账户期初填 0，开户余额录一笔它）；
  ③ 收付人按笔数从多到少；暂停/销户的账户「停用日」＝最后一笔流水的日期（之前的流水不提示「暂停户还在记」）。
entries 里的「属性」（_属性）是清洗按旧口径给的，测试对数时按 ITEM_ATTR_OVERRIDE 换算（data.attr_v2()）。
名字只去空格、换行，不改括号（公式里比较时两边都用 layout.norm 规范，半角全角括号都认）。
测试小容量（A074_SMALL=1）时，流水每个账户只取前 SMALL_PER_ACC 笔。
"""
import os
import re
import json
import datetime as dt
from layout import (SUB_SHEETS, SUB_NAMES, SMALL, N_LEDGER, N_CO, N_ACC, N_ITEM, N_PARTY, N_RULE, ITEM_ATTRS, ACC_KINDS,
                    ACC_STATUS, ITEM_PENDING, ITEM_INTERNAL, ITEM_OPEN)

HERE = os.path.dirname(os.path.abspath(__file__))
IMPORT_DIR = os.path.join(HERE, '导入', 'out')
SMALL_PER_ACC = 15
EX = '示例'


# ───────────────────────── 小工具 ─────────────────────────
def nm(s):
    """名字：去空格、换行（括号不动）"""
    if s is None:
        return ''
    s = str(s)
    for a in (' ', '　', ' ', '\n', '\r', '\t'):
        s = s.replace(a, '')
    return s


def key(s):
    """比较用：再把半角括号改全角（跟公式里 layout.norm 一样）"""
    return nm(s).replace('(', '（').replace(')', '）')


def txt(s):
    """一般文字：去首尾空白、换行改空格；list 用「、」连"""
    if s is None:
        return ''
    if isinstance(s, (list, tuple)):
        return '、'.join(txt(x) for x in s if txt(x))
    return ' '.join(str(s).replace('\r', '\n').split('\n')).strip()


def money(x):
    if x is None or x == '':
        return None
    try:
        v = round(float(str(x).replace(',', '').replace('¥', '').replace('￥', '')), 2)
    except ValueError:
        return None
    return None if v == 0 else v


def day(x):
    """'2026-01-05' / '2026/1/5' / [2026,1,5] / Excel 序号 → date；看不懂＝None"""
    if x is None or x == '':
        return None
    if isinstance(x, dt.datetime):
        return x.date()
    if isinstance(x, dt.date):
        return x
    if isinstance(x, (list, tuple)) and len(x) == 3:
        return dt.date(int(x[0]), int(x[1]), int(x[2]))
    if isinstance(x, (int, float)) and 20000 < x < 80000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(x))
    s = str(x).strip().replace('/', '-').replace('.', '-').replace('年', '-').replace('月', '-').replace('日', '')
    try:
        y, m, d = (int(t) for t in s.split('-')[:3])
        return dt.date(y, m, d)
    except (ValueError, TypeError):
        return None


def _get(rec, *keys, default=None):
    for k in keys:
        if k in rec and rec[k] not in (None, '', []):
            return rec[k]
    return default


def _read(fn):
    p = os.path.join(IMPORT_DIR, fn)
    if not os.path.exists(p):
        return None
    with open(p, encoding='utf-8') as f:
        return f.read() if fn.endswith('.md') else json.load(f)


# ───────────────────────── 示例（导入数据没到时用） ─────────────────────────
DEFAULT_COMPANIES = [  # 公司简称, 公司全称（顺序同清洗 conf.COMPANIES）
    ('力喆实业', '福建力喆实业发展有限公司'), ('力喆产业', '福建力喆产业投资集团有限公司'), ('力喆贸易', '福建力喆贸易有限公司'),
    ('力喆传媒', ''), ('元生万物', '福州市元生万物投资有限公司'), ('言吾言', '福建言吾言物业管理有限公司'), ('古玩城', ''),
    ('艾泊烨', '福州艾泊烨停车科技有限公司'), ('博烨', '福州市博烨体育发展有限公司'), ('易佰兴', '福州易佰兴信息科技有限公司'),
    ('趋势创投', '福州趋势创投有限公司'), ('中策', '福州晋安中策管理咨询有限公司'), ('四季快安', ''),
]
# 账户（顺序＝layout.SUB_SHEETS）：公司, 户别, 开户行, 户号尾号, 持卡人, 开账余额（读懂报告 journal_struct §5；古玩城用银行口径）, 状态
_ACC_INFO = [
    ('力喆实业', '对公基本户', '中国银行', '', '', 20187.05, '在用'), ('力喆实业', '私户', '招商银行', '', '徐洁', 227902.83, '在用'),
    ('力喆产业', '对公一般户', '农商银行', '', '', 20301.19, '在用'), ('力喆贸易', '对公基本户', '建设银行', '', '', 14300.72, '在用'),
    ('力喆传媒', '对公基本户', '海峡银行', '', '', 10337.35, '在用'), ('元生万物', '对公基本户', '建设银行', '', '', 36305.74, '在用'),
    ('言吾言', '对公基本户', '中国银行', '', '', 3045.39, '在用'), ('古玩城', '私户', '农商银行', '', '吴海萍', 262273.10, '在用'),
    ('古玩城', '现金', '', '', '', 0, '在用'), ('艾泊烨', '对公基本户', '海峡银行', '', '', 4588.64, '在用'),
    ('艾泊烨', '对公一般户', '招商银行', '', '', 30632.96, '在用'), ('艾泊烨', '对公一般户', '农商银行', '', '', 3219.26, '在用'),
    ('艾泊烨', '私户', '农商银行', '', '王晓华', 13823.47, '在用'), ('博烨', '对公基本户', '农商银行', '', '', 1286.50, '在用'),
    ('博烨', '收单户', '农商银行', '146', '', 0, '在用'), ('博烨', '收单户', '农商银行', '468', '', 0, '在用'),
    ('博烨', '收单户', '民生银行', '751', '', 0, '在用'), ('博烨', '收单户', '', '847', '', 0, '在用'),
    ('博烨', '私户', '农商银行', '', '谢海亮', 0, '在用'), ('易佰兴', '对公基本户', '农商银行', '', '', 46768.81, '在用'),
    ('易佰兴', '对公一般户', '建设银行', '', '', 0, '在用'), ('易佰兴', '私户', '农商银行', '', '谢依俤', 0, '在用'),
    ('易佰兴', '现金', '', '', '', 0, '在用'), ('趋势创投', '对公基本户', '民生银行', '', '', 69.64, '在用'),
    ('中策', '对公基本户', '海峡银行', '', '', 176530.77, '暂停'), ('四季快安', '对公基本户', '兴业银行', '', '', 125.51, '暂停'),
]
DEFAULT_ITEMS = [  # 收费项, 属性, 大类（示例；正式的由清洗那边按付费项 77 项＋新增项给）
    ('房租收入', '收入', '租赁收入'), ('物业收入', '收入', '物业收入'), ('水电收入', '收入', '物业收入'), ('摊位收入', '收入', '租赁收入'),
    ('停车费收入', '收入', '停车收入'), ('场馆营业收入', '收入', '场馆收入'), ('利息收入', '收入', '财务费用'), ('其他收入', '收入', '其他收入'),
    ('薪资', '支出', '人工'), ('五险一金', '支出', '人工'), ('税费', '支出', '税费'), ('银行手续费', '支出', '财务费用'),
    ('水电费', '支出', '租赁及物业成本'), ('办公费', '支出', '行政办公'), ('徐总费用', '支出', '老板及个人'),
    ('押金', '往来', '押金'), ('退押金', '往来', '押金'), ('借款转入', '往来', '借款'), ('还款', '往来', '借款'),
    ('投资款', '往来', '投资'), ('往来款', '往来', '往来款'), ('内部转账', '内部转账', '内部转账'), ('王庄篮球馆', '支出', '项目建设'),
    (ITEM_PENDING, '按方向', '待定'),
]
ITEM_ATTR_OVERRIDE = [  # (条件, 新属性)：条件＝('大类', 值) 或 ('收费项', 值)。总负责人定的是四种属性，这两条是设计 v2 的建议（§12）
    (('大类', '项目建设'), '投资建设'),
    (('收费项', ITEM_PENDING), '待分类'),
]
ITEM_ORDER = ['收入', '内部转账', '支出', '按方向', '投资建设', '往来', '待分类']     # ④ 收费项的排列（同属性按笔数从多到少）


def attr_v2(item, attr, groups):
    """清洗给的属性 → 模板 v2 的属性（测试对数用；groups＝{收费项: 大类}，用 item_group_map(ctx['items'])）"""
    for (k, v), new in ITEM_ATTR_OVERRIDE:
        if (k == '收费项' and item == v) or (k == '大类' and groups.get(item) == v):
            return new
    return attr
DEFAULT_PARTIES = [('陈文', '租户'), ('汪亮', '租户'), ('江婉露', '租户'), ('福州顺丰速运有限公司', '租户'),
                   ('国网福建省电力有限公司', '供应商'), ('税务局（TIPS代扣）', '税务'), ('农商银行', '银行'), ('海峡银行', '银行'),
                   ('中国银行', '银行'), ('员工（代发工资）', '个人'), ('王晓华', '个人'), ('谢海亮', '个人'), ('地摊散户', '租户')]
DEFAULT_RULES = [  # 关键词, 查找位置, 方向, 适用范围, 收费项, 收付人
    ('手续费', '摘要', '支', '', '银行手续费', ''), ('短信', '摘要', '支', '', '银行手续费', ''), ('利息', '摘要', '收', '', '利息收入', ''),
    ('押金', '摘要', '收', '', '押金', ''), ('押金', '摘要', '支', '', '退押金', ''), ('社保', '摘要', '都', '', '五险一金', ''),
    ('工资', '摘要', '支', '', '薪资', ''), ('税务局', '收付人', '支', '', '税费', ''), ('电费', '摘要', '支', '', '水电费', ''),
    ('租金', '摘要', '收', '', '房租收入', ''), ('停车场收入', '摘要', '收', '艾泊烨', '停车费收入', ''),
    ('抖音', '摘要', '收', '博烨', '场馆营业收入', ''), ('地摊', '摘要', '收', '古玩城', '', '地摊散户'),
]


def _default_accounts():
    out = []
    for (sub, orig, _cap), (co, kind, bank, no, holder, op, st) in zip(SUB_SHEETS, _ACC_INFO):
        out.append({'账户名称': sub, '子表名': sub, '公司': co, '户别': kind, '开户行': bank, '户号尾号': no, '持卡人': holder,
                    '期初余额': op, '状态': st, '停用日': dt.date(2026, 3, 31) if st != '在用' else None,
                    '别名': orig.strip(), '备注': '', '原表名': orig.strip()})
    return out


def _default_entries():
    A = SUB_NAMES
    E = lambda d, a, item, party, memo, inc=None, exp=None, to='', co='', note=EX, pos='': dict(
        日期=d, 账户=A[a], 收费项=item, 收付人=party, 摘要=memo, 收入=inc, 支出=exp, 对方账户=(A[to] if to != '' else ''),
        归属公司=co, 备注=note, 原表位置=pos)
    return [
        E('2026-01-01', 7, '房租收入', '陈文', 'A105 1月租金', inc=7247),
        E('2026-01-04', 7, '', '地摊散户', '广场92号 固定摊位费', inc=600, note=EX + '：收费项空着，按摘要规则认'),
        E('2026-01-04', 7, '银行手续费', '农商银行', '广场92号 固定摊位费 手续费', exp=17.23),
        E('2026-01-05', 24, '银行手续费', '海峡银行', '支付短信通服务费', exp=10),
        E('2026-01-05', 9, '内部转账', '博烨', '博烨体育投资款（原收费项：股权投资款）', exp=200000, to=13, note=EX + '：内部转账转出'),
        E('2026-01-05', 13, '内部转账', '艾泊烨', '收艾泊烨投资款（原收费项：股权投资款）', inc=200000, to=9, note=EX + '：内部转账转入'),
        E('2026-01-08', 2, '薪资', '员工（代发工资）', '代发1月工资', exp=35600),
        E('2026-01-10', 1, '徐总费用', '', '还信用卡', exp=8800),
        E('2026-01-12', 1, '房租收入', '汪亮', '雅道巷63-1租金', inc=12832.11, co='力喆产业', note=EX + '：私户代收，归属公司填力喆产业'),
        E('2026-01-12', 1, '内部转账', '', '雅道巷租金转力喆', exp=12832.11, to=2),
        E('2026-01-12', 2, '内部转账', '', '徐洁转入雅道巷租金', inc=12832.11, to=1),
        E('2026-01-15', 21, '', '天虹', '天虹 闽A31J6H 月租停车费', inc=240, note=EX + '：收费项空，规则认不出'),
        E('2026-01-15', 20, '税费', '税务局（TIPS代扣）', '代扣增值税', exp=1520.5),
        E('2026-01-20', 12, '借款转入', '王晓华', '王晓华借款', inc=50000),
        E('2026-01-31', 7, '水电收入', '陈文', 'A105 1月水电', inc=812.4),
        E('2026-02-02', 7, '退押金', '江婉露', 'C103 押金抵2月租金', exp=3000, note='以押抵租，不过银行'),
        E('2026-02-02', 7, '房租收入', '江婉露', 'C103 2月租金', inc=3000, note='以押抵租，不过银行'),
        E('2026-02-05', 7, '内部转账', '徐洁', '吴海萍转入徐洁账户', exp=96000, to=1),
        E('2026-02-05', 1, '内部转账', '吴海萍', '吴海萍转入徐洁账户', inc=96000, to=7),
        E('2026-02-14', 14, '场馆营业收入', '', '统一收单清算', inc=455.6),
        E('2026-03-06', 0, '银行手续费', '中国银行', '回单箱年费', exp=400),
        E('2026-03-21', 24, '办公费', '', '提取备用金', exp=48000, note=EX + '：暂停户'),
        E('2026-05-11', 8, '摊位收入', '', '5.9-5.10 地摊费', inc=740, note=EX + '：收付人空着，按规则补'),
        E('2026-06-30', 17, ITEM_PENDING, '', '预付账款', exp=1600, note=EX + '：待定，请用户确认'),
        E('2026-08-04', 13, '内部转账', '谢海亮', '基本户转入私户', exp=32800, to=18),
        E('2026-08-04', 18, '内部转账', '谢海亮', '基本户转入私户', inc=32800, to=13),
        E('2026-09-20', 15, '内部转账', '', '一般户（羽毛球馆）转入基本户', exp=7478.06, to=13,
          note='原表漏记，按博烨基本户收款补记，请核对'),
        E('2026-09-20', 13, '内部转账', '', '收博烨一般户（羽毛球馆）', inc=7478.06, to=15),
        E('2026-10-08', 11, '', '', '停车场收入', inc=1234.5, note=EX + '：收费项空'),
        E('2026-10-09', 7, '房租收入', '陈文', 'A105 10月租金', inc=7247),
        E('2026-10-13', 1, '徐总费用', '', '世茂物业费', exp=2100),
    ]


CLEAN_RULES = [  # 数据整理记录 · 一、整理口径（给用户看的话）
    ('总表收哪些账户', '原日记账 27 张账户表收 26 张；「易佰兴基本户」是「易佰兴农商银行外」的不完整重复，不收。中策、四季快安两张隐藏表照收，状态「暂停」。'
                 'Sheet1、汇总、目录、付费项不是流水，不收（付费项清单搬进【基础资料】④收费项）。'),
    ('账户名', '统一成「公司简称-银行[户号](户别/持卡人)」，比如「古玩城-农商银行(吴海萍私户)」。原表名记在【基础资料】③账户的「别名」里。'),
    ('古玩城农商银行', '期初用银行口径 262,273.10（原表 R 列手算的数）。「吴海萍转入徐洁账户」只导支出一边，记内部转账、对方账户＝徐洁招商；'
                  '「以押抵租」拆成两行（收入 租金、支出 押金退还，金额相同，备注「以押抵租，不过银行」），余额不变、租金收入不丢。'),
    ('不过本账户的备忘行', '艾泊烨私户第 143 行（他人垫付 5 万）、易佰兴趋势民生第 9 行（薪资）是记着看的，不进余额，不导，写在修改明细里。'),
    ('漏记补记', '博烨羽毛球漏记 9/20 转基本户 7,478.06：按博烨基本户的收款补一行，备注「原表漏记，按博烨基本户收款补记，请核对」。'),
    ('分栏拆行', '租金/物业/水电/押金/地摊/停车/广告/利息/营业分成这些分栏，一行拆成几行（每个分栏一行），手续费单独一行支出「银行手续费」；'
             '「原表位置」写同一格加 -1、-2；拆完余额跟原来一致。毛收入/手续费/净收入式的表同理。'),
    ('内部转账', '两头都在本册、配上对的转账，两边收费项都改成「内部转账」，对方账户填对方的新账户名；原来的收费项和摘要留在摘要/备注里。'
             '对方不在本册的单边转账不算内部转账，按性质记借款转入、还款、往来款、押金等「往来」类收费项。'),
    ('收费项', '以原「付费项」77 项为底，加场馆营业收入、投资款、内部转账、银行手续费等；每项定属性（收入/支出/往来/内部转账）和大类。'
            '写法不统一的按对照表改；填错的逐行改；空的按分栏、摘要关键词补；补不出来的写「待定」，列在待确认里。'),
    ('收付人', '同一个人/单位的不同写法统一；空的按规则补（铺位号、停车品牌、地摊、银行、税务…）；补不出来的留空。'),
    ('日期和顺序', '原表只有月、日，年份一律 2026；缺日的按对方账户或前后行定。总表按日期排，同一天按账户顺序，同一账户按原表行序；'
               '每个账户的余额按这个顺序重算，期末跟原表（修正后的银行口径）一致。'),
]


# ───────────────────────── 读 ─────────────────────────
def _load_accounts():
    d = _read('accounts.json')
    if d is None:
        return _default_accounts()
    out = []
    for r in d[:N_ACC]:
        name = nm(_get(r, '账户名', '账户名称', '账户', 'name'))
        alias = [nm(_get(r, '原表名'))] + [nm(x) for x in (r.get('别名') or [])] if isinstance(r.get('别名'), list) \
            else [nm(_get(r, '原表名')), nm(_get(r, '别名'))]
        out.append({'账户名称': name, '子表名': nm(_get(r, '子表名', default=name)), '公司': nm(_get(r, '公司简称', '公司')),
                    '户别': txt(_get(r, '户别', '类型', default='其他')), '开户行': txt(_get(r, '银行', '开户行')),
                    '户号尾号': txt(_get(r, '户号尾号', '户号')), '持卡人': txt(_get(r, '持卡人')),
                    '期初余额': money(_get(r, '开账余额', '期初余额', '期初')) or 0,
                    '状态': txt(_get(r, '状态', default='在用')), '别名': '、'.join(a for a in dict.fromkeys(alias) if a),
                    '备注': txt(_get(r, '说明', '备注')), '原表名': nm(_get(r, '原表名')),
                    '开账日期': day(_get(r, '开账日期')), '期末余额': money(_get(r, '期末余额')),
                    '流水行数': _get(r, '流水行数'), '最后一笔日期': day(_get(r, '最后一笔日期'))})
        out[-1]['停用日'] = out[-1]['最后一笔日期'] if out[-1]['状态'] in ('暂停', '销户') else None
    return out


def _load_companies():
    d = _read('companies.json')
    if d is None:
        return [{'公司简称': a, '公司全称': b, '备注': ''} for a, b in DEFAULT_COMPANIES]
    d = sorted(d, key=lambda r: _get(r, '顺序', '序号', default=0) or 0)
    return [{'公司简称': nm(_get(r, '简称', '公司简称', '公司', 'name')), '公司全称': txt(_get(r, '全称', '公司全称')),
             '备注': txt(_get(r, '备注', '说明'))} for r in d][:N_CO]


def _item_fix(items):
    """属性调整（ITEM_ATTR_OVERRIDE）、补「账户转入」、按 ITEM_ORDER＋笔数排序"""
    for it in items:
        for (k, v), attr in ITEM_ATTR_OVERRIDE:
            if (k == '大类' and it['大类'] == v) or (k == '收费项' and it['收费项'] == v):
                if it['属性'] != attr:
                    it['说明'] = (it['说明'] + '；' if it['说明'] else '') + f'属性原为「{it["属性"]}」，模板改为「{attr}」'
                    it['属性'] = attr
    if not any(it['收费项'] == ITEM_OPEN for it in items):
        items.append({'收费项': ITEM_OPEN, '属性': '往来', '大类': '账户并入', '别名': '', '说明': '年中新开的账户：期初填 0，开户那天的余额录一笔这个',
                      '_笔数': 0})
    order = {a: i for i, a in enumerate(ITEM_ORDER)}
    idx = {id(it): i for i, it in enumerate(items)}
    items.sort(key=lambda it: (order.get(it['属性'], len(order)), -(it.get('_笔数') or 0), idx[id(it)]))
    return items[:N_ITEM]


def item_group_map(items):
    return {it['收费项']: it['大类'] for it in items}


def _load_items():
    d = _read('items.json')
    if d is None:
        return _item_fix([{'收费项': a, '属性': b, '大类': c, '别名': '', '说明': '', '_笔数': 0} for a, b, c in DEFAULT_ITEMS])
    return _item_fix([{'收费项': nm(_get(r, '名称', '收费项', 'name')), '属性': txt(_get(r, '属性', 'attr')),
                       '大类': txt(_get(r, '大类', 'group')), '别名': txt(_get(r, '原写法别名', '别名')),
                       '说明': txt(_get(r, '说明', 'note')), '_笔数': _get(r, '本次笔数', default=0) or 0} for r in d])


def _load_parties():
    d = _read('parties.json')
    if d is None:
        return [{'收付人': a, '类型': b, '别名': '', '备注': ''} for a, b in DEFAULT_PARTIES]
    out = [{'收付人': nm(_get(r, '名称', '收付人', 'name')), '类型': txt(_get(r, '类型')), '别名': txt(_get(r, '别名')),
            '备注': txt(_get(r, '备注', '说明')), '_笔数': _get(r, '笔数', default=0) or 0} for r in d]
    out.sort(key=lambda r: -(r['_笔数'] if isinstance(r['_笔数'], (int, float)) else 0))     # 常用的在上面（sort 是稳定的）
    return out[:N_PARTY]


def _load_rules():
    d = _read('rules.json')
    if d is None:
        return [{'关键词': a, '查找位置': b, '方向': c, '适用范围': e, '收费项': f, '收付人': g, '规则号': '', '说明': ''}
                for a, b, c, e, f, g in DEFAULT_RULES]
    rows = (d.get('收费项规则', []) + d.get('收付人规则', [])) if isinstance(d, dict) else d
    out = [{'关键词': txt(_get(r, '关键词', 'kw')), '查找位置': txt(_get(r, '查找位置')), '方向': txt(_get(r, '适用方向', '方向', 'dir')),
            '适用范围': nm(_get(r, '适用公司', '适用范围', 'scope')), '收费项': nm(_get(r, '收费项', 'item')),
            '收付人': nm(_get(r, '收付人', 'party')), '规则号': txt(_get(r, '规则号')), '说明': txt(_get(r, '说明', 'note'))}
           for r in rows]
    if len(out) > N_RULE:
        print(f'   ⚠ data：规则 {len(out)} 条超过容量 {N_RULE}，后面的不写')
    return out[:N_RULE]


def _load_entries():
    d = _read('entries.json')
    src = d is not None
    if d is None:
        d = _default_entries()
    out = []
    for r in d:
        out.append({'日期': day(_get(r, '日期', 'date')), '账户': nm(_get(r, '账户')), '收费项': nm(_get(r, '收费项')),
                    '收付人': nm(_get(r, '收付人')), '摘要': txt(_get(r, '摘要')), '收入': money(_get(r, '收入')),
                    '支出': money(_get(r, '支出')), '对方账户': nm(_get(r, '对方账户')), '归属公司': nm(_get(r, '归属公司')),
                    '备注': txt(_get(r, '备注')), '原表位置': txt(_get(r, '原表位置')),
                    '_属性': txt(_get(r, '属性')), '_余额': _get(r, '余额')})        # 下划线开头的不写进表，测试对数用
    return out, src


def _load_params(accounts):
    p = {'集团名称': '', '管理年度': None, '开账日': None}
    d = _read('params.json') or {}
    for k in p:
        if k in d:
            p[k] = d[k]
    ds = [a.get('开账日期') for a in accounts if a.get('开账日期')]
    p['开账日'] = day(p['开账日']) or (min(ds) if ds else None) or dt.date(2026, 1, 1)
    return p


def _load_cleanlog():
    d = _read('cleanlog.json') or _read('changes.json') or []
    return [{'类别': txt(_get(r, '类别')), '原表位置': txt(_get(r, '原表位置')), '字段': txt(_get(r, '字段')),
             '原值': txt(_get(r, '原值')), '新值': txt(_get(r, '新值')), '依据': txt(_get(r, '依据'))} for r in d]


def _parse_questions(md):
    """questions.md →[{问题, 涉及金额, 涉及, 说明}]"""
    out, cur = [], None
    for line in (md or '').splitlines():
        m = re.match(r'^##\s*\d+[.、．]\s*(.+)$', line.strip())
        if m:
            cur = {'问题': m.group(1).strip(), '涉及金额': '', '涉及': '', '说明': ''}
            out.append(cur)
            continue
        if cur is None:
            continue
        s = line.strip()
        m = re.match(r'^涉及金额[：:]\s*([^\s　]+)\s*[　 ]*(?:涉及[：:]\s*(.*))?$', s)
        if m:
            cur['涉及金额'] = m.group(1)
            cur['涉及'] = (m.group(2) or '').strip()
        elif s.startswith('- ') or s.startswith('* '):
            cur['说明'] += ('\n' if cur['说明'] else '') + s[2:].strip()
        elif s and not s.startswith('#'):
            cur['说明'] += ('\n' if cur['说明'] else '') + s
    return out


def _parse_first_table(md):
    """check.md 第一张 markdown 表 → [dict]"""
    rows, hdr = [], None
    for line in (md or '').splitlines():
        s = line.strip()
        if not s.startswith('|'):
            if hdr and rows:
                break
            continue
        cells = [c.strip().replace('**', '') for c in s.strip('|').split('|')]
        if hdr is None:
            hdr = cells
        elif set(''.join(cells)) <= set('-: '):
            continue
        else:
            rows.append(dict(zip(hdr, cells)))
    return rows


def _small(entries):
    seen, out = {}, []
    for e in entries:
        k = e['账户']
        if seen.get(k, 0) < SMALL_PER_ACC:
            seen[k] = seen.get(k, 0) + 1
            out.append(e)
    return out[:N_LEDGER - 20]


def _warn(ctx):
    """导入数据的基本检查：只打印，不拦"""
    accs = {key(a['账户名称']) for a in ctx['accounts']}
    cos = {key(c['公司简称']) for c in ctx['companies']}
    items = {key(i['收费项']) for i in ctx['items']}
    w = []
    if [key(a['账户名称']) for a in ctx['accounts'][:len(SUB_NAMES)]] != [key(s) for s in SUB_NAMES]:
        w.append('accounts 前 26 个账户名/顺序跟 layout.SUB_SHEETS 不一样（子表按位置对应：要么改 layout，要么改清洗产出）')
    for a in ctx['accounts']:
        if key(a['公司']) not in cos:
            w.append(f'账户「{a["账户名称"]}」的公司「{a["公司"]}」不在 companies 里')
        if a['户别'] not in ACC_KINDS or a['状态'] not in ACC_STATUS:
            w.append(f'账户「{a["账户名称"]}」户别/状态不在清单：{a["户别"]}/{a["状态"]}')
    for i in ctx['items']:
        if i['属性'] not in ITEM_ATTRS + ['']:
            w.append(f'收费项「{i["收费项"]}」属性不对：{i["属性"]}')
    es = ctx['entries']
    n = dict(账户不在清单=sum(1 for e in es if key(e['账户']) not in accs), 日期看不懂=sum(1 for e in es if e['日期'] is None),
             收费项不在清单=sum(1 for e in es if e['收费项'] and key(e['收费项']) not in items),
             收入支出都填了=sum(1 for e in es if e['收入'] and e['支出']), 没有金额=sum(1 for e in es if not e['收入'] and not e['支出']),
             对方账户不在清单=sum(1 for e in es if e['对方账户'] and key(e['对方账户']) not in accs))
    w += [f'流水 {v} 笔{k}' for k, v in n.items() if v]
    if len(es) > N_LEDGER:
        w.append(f'流水 {len(es)} 笔超过资金台帐容量 {N_LEDGER}（多的不写）')
    for t in w[:20]:
        print('   ⚠ data：' + t)
    return w


def build_ctx():
    entries, from_json = _load_entries()
    if SMALL:
        entries = _small(entries)
    accounts = _load_accounts()
    pend = _read('pending.json')
    chk = _read('acct_check.json')
    ctx = dict(params=_load_params(accounts), companies=_load_companies(), accounts=accounts, items=_load_items(),
               parties=_load_parties(), rules=_load_rules(), entries=entries[:N_LEDGER], cleanlog=_load_cleanlog(),
               pending=pend if pend is not None else _parse_questions(_read('questions.md')),
               acct_check=chk if chk is not None else _parse_first_table(_read('check.md')),
               clean_rules=CLEAN_RULES, source='导入/out' if from_json else '示例')
    if not from_json:
        print('   （导入/out/entries.json 还没有，用示例流水）')
    ctx['warnings'] = _warn(ctx)
    return ctx


if __name__ == '__main__':
    c = build_ctx()
    print({k: (len(v) if isinstance(v, list) else v) for k, v in c.items() if k not in ('params', 'warnings')}, c['params'])
    print('待确认', [p['问题'][:20] for p in c['pending']][:12])
    print('账户对照列', list(c['acct_check'][0].keys()) if c['acct_check'] else None)
