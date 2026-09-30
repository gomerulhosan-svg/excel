# -*- coding: utf-8 -*-
"""演示数据：把 data_prep 读出来的原表数整理成各张表要的样子（ctx）。

原则：你原表里的数原样用；汇总表里的「累计数」（应收、应付、分包、已付、已开票）当作建账前的期初，
      但要减掉 2026 年流水里已经付/收的（不然今年的流水会再算一遍）。拿不准的都在备注里写明，校验里也会提醒。"""
import datetime as dt
import openpyxl
import data_prep as d
import cats

D = dt.datetime
OPEN_DATE = D(2026, 3, 1)      # 跟你的用工成本表一样：25 年～26 年 2 月合计是期初，3 月起逐月
ACC_MAP = {'聂辉微信': '聂辉个人户', '王伟微信': '王伟个人户', '张滨微信': '张滨个人户'}


def workers():
    """4 月工资表里的人（日工资、开户行从原表缓存里读）＋ 2~3 月截图 ＋ 工资汇总、流水里出现的"""
    wb = openpyxl.load_workbook(d.F1, data_only=True)
    ws = wb['工资表月分表']
    out = {}
    for r in range(3, 35):
        n = ws.cell(r, 2).value
        if n:
            out[n] = dict(rate=ws.cell(r, 22).value, bank=ws.cell(r, 6).value)
    for n, rt in d.RATES.items():
        out.setdefault(n, dict(rate=rt, bank=None))
    return out


def build_ctx():
    ctx = {'company': ('竣辉', '（请填营业执照上的全称）', '（请填税号）'), 'open_date': OPEN_DATE}
    wk = workers()
    managers = {'聂辉': '老板', '王伟': '老板', '张滨': '项目部负责人', '张淑平': '办公室', '姚俊强': '办公室', '朱博文': '办公室'}
    # ── 往来单位及人员 ──
    units = []
    add = lambda **k: units.append(k)
    for n in ['九江浔辉建筑工程有限公司', '德安县市政工程公司', '德安县水利水电建筑工程公司', '江西揽之建设工程有限公司',
              '德安县博河产业控股集团有限公司']:
        add(name=n, type='甲方/总包', full=n)
    add(name='德安农商银行', type='甲方/总包', full='农村商业银行股份有限公司德安县总部', al1='农商银行')
    add(name='江西赣粤高速', type='甲方/总包', al1='赣粤高速公路', note='流水 5/21 付过工程款，全称请补')
    mat_alias = {'筑强': ('筑强建材', ''), '九江安洋钢材熊总': ('安洋', ''), '鑫城五金庄总': ('鑫城五金', ''), '浔阳楼电缆': ('浔阳楼', ''),
                 '启航五金': ('启航', ''), '杭州行消物资贸易': ('杭州行消', '行消'), '南昌盈厚建材': ('盈厚', ''),
                 '浙缆元通电缆': ('浙缆', '元通'), '江西允诺消防': ('允诺', ''), '江西沐丰水箱': ('沐丰', ''),
                 '上海诚果泵阀': ('上海诚果', '诚果'), '佩力通风': ('佩力', ''), '浙江寅盾智能': ('寅盾', '')}
    for m in d.MAT_SUP:
        a1, a2 = mat_alias.get(m[0], ('', ''))
        add(name=m[0], type='材料供应商', full=m[1] or None, al1=a1 or None, al2=a2 or None)
    for n, a1, note in (('华翔桥架', '华翔', '流水 3/14 付过德养材料款'), ('拓齐', '', '流水 3/11「支付拓齐-老叶-德养材料款」'),
                        ('四川久远智能消防', '久远', '流水 6/25、6/30 付设备款'), ('九江恒讯消防', '恒讯', '流水 7/31 付设备款')):
        add(name=n, type='材料供应商', al1=a1 or None, note=note)
    sub_alias = {'恩旗老叶': '老叶', '九江朱总': '朱总', '共安门业': '共安', '图揽成善': '图揽', '张小朋（分包）': '张小朋分包',
                 '程章赐（分包）': '程章赐分包'}
    subs = [s[0] if s[0] != '张小朋' else '张小朋（分包）' for s in d.SUB]
    for n in subs:
        add(name=n, type='分包', al1=sub_alias.get(n))
    for n, full, a1 in (('杨工', '', ''), ('徐工', '九江翌年行建筑工程劳务有限公司', ''), ('德安县恒建工程部', '', '恒建'),
                        ('航帆增辉', '', ''), ('章德吉货运', '', '章德吉')):
        add(name=n, type='机械运输', full=full or None, al1=a1 or None)
    for n in wk:
        if n in managers:
            add(name=n, type='管理人员', note=managers[n] + '（月薪）')
        else:
            al = {'王兴蛟': '王兴姣'}.get(n)
            add(name=n, type='工人', al1=al, bank=wk[n].get('bank'))
    for n in ('朱博文',):
        add(name=n, type='管理人员', note='办公室（流水里有 4、5 月工资和社保）')
    for n in ('陈德银', '邓付贞', '桂木生', '雷友才', '占善庆'):
        if n not in wk:
            add(name=n, type='工人', note='流水/工资汇总里有，日工资请补' if n not in d.RATES else None)
    add(name='张洪林', type='临时工', note='7/31 邹桥屋面保温棉搬运（已扣 80 个税）')
    for n in ('九江财小保', ):
        add(name=n, type='其他', note='代账公司')
    ctx['units'] = units
    # ── 工资标准 ──
    rates = []
    for n, v in wk.items():
        if n in managers:
            rates.append((n, D(2025, 1, 1), None, v['rate'], '月薪（4 月工资表）'))
        else:
            rates.append((n, D(2025, 1, 1), v['rate'], None, '4 月工资表的日工资' if n not in d.RATES or n in
                          [x for x in wk if wk[x].get('bank')] else '工资汇总表'))
    for n, rt in (('陈德银', 260), ('邓付贞', 400), ('桂木生', 300)):
        if n not in wk:
            rates.append((n, D(2025, 1, 1), rt, None, '工资汇总表'))
    ctx['rates'] = rates
    # ── 项目 ──
    projects = []
    for p in d.PROJECTS:
        projects.append(dict(name=p[0], full=p[1], cus=p[2] or None, amt=p[3], tax=p[6], kw=[k for k in p[7].split('、') if k],
                             note=p[8] or None, stat='在建'))
    ctx['projects'] = projects
    # ── 资金账户 ──
    ctx['accounts'] = [
        ('九江银行', '银行', None, None, 0, '对公基本户。1~4 月流水原表没有，期初余额请按银行对账单填'),
        ('农业银行', '银行', None, None, 0, '贷款户。期初余额请按对账单填'),
        ('工商银行', '银行', None, None, 0, '贷款户'),
        ('现金', '现金', None, None, 0, ''),
        ('聂辉个人户', '个人户', '聂辉', None, -237668,
         '原来的「聂辉微信」＋他个人替公司付的、借给公司的。期初 −237,668＝你原表聂总个人明细里借给公司的合计（到 2024-10），2025 年的请核对'),
        ('王伟个人户', '个人户', '王伟', None, -190000,
         '原来的「王伟微信」。期初 −190,000＝你原表王总个人往来里借给公司的合计，请核对'),
        ('张滨个人户', '个人户', '张滨', None, 0, '原来的「张滨微信」。2025-07-01 领的备用金 60,000 还剩多少请核对后填'),
        ('承兑汇票', '票据', None, None, 0, '收到还没贴现、没背书的承兑汇票'),
    ]
    # ── 流水 ──
    J = []
    pre = {}                                   # 建账日以前的流水：并进账户期初（工商银行 1 月的贷款、付筑强货款）
    for (dd, acc, memo, inc, exp) in d.journal():
        if dd is not None and dd < OPEN_DATE:
            a = ACC_MAP.get(acc, acc)
            pre[a] = round(pre.get(a, 0) + (inc or 0) - (exp or 0), 2)
            continue
        if (inc in (None, 0)) and exp is not None and exp < 0:
            inc, exp = -exp, None             # 原表把收款记成「支出」负数的，改到收入栏
        J.append((dd, ACC_MAP.get(acc, acc), memo, inc, exp))
    # 演示：几笔原表里看不出对象的，替你在「改…」列选好（其余的保持自动认，✗ 的留给你看怎么改）
    fixes = {'汇款退回': {'N': '杭州行消物资贸易'},
             # 机械费表里 2026 年这 3 笔登记在【应付登记】，付款那天选上单位＝付应付款（冲应付），不然会算两遍
             '邹桥派出所设备运输费': {'N': '章德吉货运'}, '邹桥派出所挖机费': {'N': '航帆增辉'}, '德安农商行挖机费': {'N': '德安县恒建工程部'}}
    J = [row + ((fixes[row[2]],) if row[2] in fixes else ()) for row in J]
    ctx['journal'] = J
    ctx['accounts'] = [(n, t, o, no, round((op or 0) + pre.get(n, 0), 2),
                        note + (f'（含 1～2 月流水 {pre[n]:,.2f}：1/12 对公贷款 281,148、1/14 付筑强 281,000 等）' if pre.get(n) else ''))
                       for (n, t, o, no, op, note) in ctx['accounts']]
    # ── 期初：项目（累计确认＝应收金额，累计收款＝已收 − 2026 流水里这个项目的工程款），应付（单位×项目） ──
    rec26, paid26 = _journal_2026(J, projects, units)
    op = []
    share_base = {'九安架空层': 1500000, '德安养老中心': 4362576.49, '科创中心': 1641809, '数字产业园': 325612.18,
                  '邹桥派出所': 91006.56, '抚州1号楼': 695372.72, '抚州2号楼': 64000, '抚州8号楼': 960000, '德安农商行总部': 8992}
    pool25 = 63079.24 + 59850 + 7366.32 + 144850
    tot = sum(share_base.values())
    lab0 = {'德安养老中心': 1073656.4, '科创中心': 402775, '九安架空层': 96985}
    inv0 = {p[0]: p[5] for p in d.PROJECTS if p[5]}
    NEW26 = {'德安农商行总部'}                 # 你发来的项目账样例：2026 年的活，产值在【收入确认】里按 2026 年确认
    for p in d.PROJECTS:
        n = p[0]
        if (n not in d.RECEIVED and n not in lab0) or n in NEW26:
            continue
        rec = round(d.RECEIVED.get(n, 0) - rec26.get(n, 0), 2)
        op.append(dict(pj=n, rev=p[3], rec=rec, inv=inv0.get(n), lab=lab0.get(n),
                       alloc=round(pool25 * share_base[n] / tot, 2) if n in share_base else None,
                       note='按应收账款总表；累计收到＝已收 − 2026 年流水里收的' + (
                           f'（{rec26[n]:,.2f}）' if rec26.get(n) else '') + ('；人工＝用工成本表 25年~26年2月合计' if n in lab0 else '')
                       + ('；已摊＝你原来待摊费用表 2025 年那部分按工程款比例' if n in share_base else '')))
    ctx['open_proj'] = op
    oap = []
    for name, full, total, per, paid in d.MAT_SUP:
        items = d.split_rest(total, per)
        p0 = round(paid - paid26.get(name, 0), 2)
        for k, (pj, amt) in enumerate(items):
            oap.append(dict(unit=name, pj=pj or None, amt=amt, paid=p0 if k == 0 else None,
                            note=('累计已付 − 2026 年流水付的' + (f'（{paid26[name]:,.2f}）' if paid26.get(name) else '')) if k == 0 else
                            ('应付总表里各项目栏合计比应付总额少的部分，项目请补' if not pj else None)))
    for name, total, per, paid in d.SUB:
        uname = '张小朋（分包）' if name == '张小朋' else name
        items = d.split_rest(total, per)
        p0 = round(paid - paid26.get(uname, 0), 2)
        for k, (pj, amt) in enumerate(items):
            oap.append(dict(unit=uname, pj=pj or None, amt=amt, paid=p0 if k == 0 else None,
                            note=('累计已付 − 2026 年流水付的' + (f'（{paid26[uname]:,.2f}）' if paid26.get(uname) else '')) if k == 0 else
                            ('分包表里各项目栏合计比应付总额少的部分，项目请补' if not pj else None)))
    for (dd, team, pj, amt, paid0, payee, inv, rate, note) in d.MACH:
        if dd < OPEN_DATE:
            oap.append(dict(unit=team, pj=pj, amt=amt, paid=paid0, inv=inv or None, note='机械费表 ' + note))
    ctx['open_ap'] = oap
    ctx['open_other'] = [281148, None, None, None, None, None, None, None]     # 短期借款：工商银行 1/12 对公贷款
    # ── 应付登记（2026 年发生的机械运输） ──
    ctx['ap_rows'] = [(dd, '机械运输', team, pj, '机械费表：' + note, amt, rate, None)
                      for (dd, team, pj, amt, paid0, payee, inv, rate, note) in d.MACH if dd >= OPEN_DATE]
    ctx['rev_rows'] = [(D(2026, 7, 31), '德安农商行总部', '进度确认', 8992,
                        '演示：你原表2「项目汇总情况」合同 8,992、已开票 8,992（确认日期原表没有，按完工 7 月底估的）')]
    ctx['off_rows'] = []
    # ── 发票登记（演示：原表 2 项目账里农商行的两张进项票） ──
    co_full = ctx['company'][1]
    ctx['inv_rows'] = [
        dict(seller='竣辉', buyer='农村商业银行股份有限公司德安县总部', time=D(2026, 7, 31), net=8249.54, tax=742.46, total=8992,
             kind='数电票（增值税专用发票）', stat='正常', pj='德安农商行总部', note='演示：原表 已开票 8,992（开票日期按估）'),
        dict(seller='德安县恒建工程部', buyer='竣辉', time=D(2026, 7, 2), net=792.08, tax=7.92, total=800, kind='数电票（普通发票）',
             stat='正常', pj='德安农商行总部', note='机械费表：已开票 800（日期按付款日）'),
        dict(seller='江西筑强建材有限公司', buyer='竣辉', time=D(2026, 6, 30), net=484.02, tax=62.92, total=546.94,
             kind='数电票（增值税专用发票）', stat='正常', pj='德安农商行总部', note='原表2 应付汇总：消防器材（日期没给，演示按 6/30）'),
    ]
    # ── 考勤：4 月（原表）＋ 2~3 月（截图里看得到的 12 人） ──
    att = []
    for a in d.april_attendance():
        pairs = list(a['days'].items())
        att.append(dict(mon=D(2026, 4, 1), name=a['name'], pairs=pairs,
                        note=None if pairs else ('月薪，没填天数＝全部算公司管理费' if a['name'] in managers else None)))
    for n, days in d.FEB_MAR:
        att.append(dict(mon=D(2026, 3, 1), name=n, pairs=list(days.items()), note='2、3 月合在一起的（截图）'))
    ctx['att_rows'] = att
    ctx['fixed_assets'] = []
    ctx['pl_default'] = '德安农商行总部'          # 项目账默认打开你发来的样例项目
    ctx['demo_note'] = ('⚠ 现在表里是你原表的数（演示用）：原表只有「累计」数，建账日定在 2026-03-01，以前的都当期初；'
                        '2026 年只有德安农商行总部录了产值确认，其他项目 3 月以后的产值确认、5～8 月的考勤都还没录，'
                        '所以本年利润表只有成本、收入很少，工人显示「多发了」。补录【收入确认】【考勤工资】后就准了。正式用之前，看说明里「开始用自己的账」清空这几张表。')
    return ctx


def _journal_2026(J, projects, units):
    """用跟表格一样的规则（最长关键词）粗算：2026 年流水里每个项目收的工程款、付给每家材料商/分包/机械的钱"""
    pkw = []
    for p in projects:
        for k in [p['name']] + p['kw']:
            pkw.append((k, p['name']))
    ukw = []
    for u in units:
        for k in (u.get('name'), u.get('al1'), u.get('al2')):
            if k:
                ukw.append((k, u['name']))
    utype = {u['name']: u['type'] for u in units}
    rec, paid = {}, {}
    for row in J:
        dd, acc, memo, inc, exp = row[:5]
        fix = row[5] if len(row) > 5 else {}
        m = memo or ''
        net = (inc or 0) - (exp or 0)
        pj = max((k for k, _ in pkw if k in m), key=len, default='')
        pj = dict(pkw).get(pj, '') if pj else ''
        dai = ''
        if '代' in m and '付' in m[m.index('代'):]:
            dai = m[m.index('代') + 1:m.index('付', m.index('代'))]
        un = max((k for k, _ in ukw if k in dai), key=len, default='') if dai else ''
        if not un:
            un = max((k for k, _ in ukw if k in m), key=len, default='')
        un = dict(ukw).get(un, '') if un else ''
        un = fix.get('N', un)
        if net > 0 and '工程款' in m and pj:
            rec[pj] = rec.get(pj, 0) + net
        if un and utype.get(un) in ('材料供应商', '分包', '机械运输') and '工程款' not in m:
            paid[un] = paid.get(un, 0) - net
    return {k: round(v, 2) for k, v in rec.items()}, {k: round(v, 2) for k, v in paid.items()}


if __name__ == '__main__':
    c = build_ctx()
    r, p = _journal_2026(c['journal'], c['projects'], c['units'])
    print('2026 工程款', r)
    print('2026 付款', p)
    print(len(c['units']), len(c['journal']), len(c['att_rows']), len(c['open_ap']))
