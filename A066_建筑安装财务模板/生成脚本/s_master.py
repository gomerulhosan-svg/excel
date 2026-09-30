# -*- coding: utf-8 -*-
"""【项目档案】【往来单位】【工人档案（含工资标准·涨薪表）】"""
import datetime as dt
from common import *
from layout import *


def build_proj(wb, ctx):
    ws = wb.create_sheet(SH_PROJ)
    widths(ws, {'A': 5, 'B': 14, 'C': 30, 'D': 26, 'E': 13, 'F': 11, 'G': 13, 'H': 8, 'I': 11, 'J': 7, 'K': 11, 'L': 11,
                'M': 10, 'N': 9, 'O': 12, 'P': 12, 'Q': 12, 'R': 13, 'S': 34, 'T': 13})
    title(ws, '项 目 档 案（一个工地一行 · 项目简称全表通用）', 'T', C_BASE,
          '💡 每接一个工地就在下面加一行。「项目简称」是全表通用的名字（流水、应付、工资、报表都用它），起短一点、别重名。'
          '「摘要关键词」：流水摘要里出现这几个词就自动认成这个项目（比如「德养」「养老中心」都是德安养老中心）；项目简称本身不用再填。'
          '合同额、变更签证、结算额都填含税价；结算了填结算额，没结算空着（按合同＋变更算）。项目不要删整行，完结了把状态改成「已完结」。'
          '建账前的累计产值、收款在【期初余额】① 填；「已收工程款」自动算到截止日（期初＋流水＋总包代发抵账）。')
    header(ws, PJ_HDR, [(PJ_SEQ, '序号'), (PJ_NAME, '项目简称'), (PJ_FULL, '项目全称（合同上的）'), (PJ_CUS, '甲方 / 总包'),
                        (PJ_AMT, '合同额\n（含税）'), (PJ_CHG, '变更签证'), (PJ_SET, '结算额\n（结算了才填）'),
                        (PJ_RET, '质保金\n比例'), (PJ_RETD, '质保\n到期日'), (PJ_TAX, '税率'), (PJ_START, '开工日期'),
                        (PJ_END, '完工日期'), (PJ_STAT, '状态'), (PJ_MGR, '负责人'), (PJ_KW1, '摘要关键词1'),
                        (PJ_KW2, '摘要关键词2'), (PJ_KW3, '摘要关键词3'), (PJ_REC0, '已收工程款\n（自动）'), (PJ_NOTE, '备注'),
                        (PJ_TOTAL, '总价\n（自动）')], C_BASE)
    for i, r in enumerate(range(PJ_R0, PJ_R1 + 1)):
        put(ws, f'{PJ_SEQ}{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in (PJ_NAME, PJ_FULL, PJ_CUS, PJ_AMT, PJ_CHG, PJ_SET, PJ_RET, PJ_RETD, PJ_TAX, PJ_START, PJ_END, PJ_STAT, PJ_MGR,
                  PJ_KW1, PJ_KW2, PJ_KW3, PJ_NOTE):
            fmt = {PJ_AMT: MONEY, PJ_CHG: MONEY, PJ_SET: MONEY, PJ_REC0: MONEY, PJ_RET: '0%', PJ_TAX: '0%',
                   PJ_RETD: DATE, PJ_START: DATE, PJ_END: DATE}.get(c)
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, fmt, AL if c in (PJ_FULL, PJ_CUS, PJ_NOTE) else AC)
        put(ws, f'{PJ_TOTAL}{r}', f'=IF({PJ_NAME}{r}="","",IF(N({PJ_SET}{r})<>0,{PJ_SET}{r},N({PJ_AMT}{r})+N({PJ_CHG}{r})))',
            F_AUTOB, FILL_AUTO, MONEY, AR)
        k = PS_R0 + i
        rec = '+'.join(f"{SH_PS}!${PS_COL[(m, p)]}${k}" for m in ('现金回款', '抵账') for p in ('前', '本'))
        put(ws, f'{PJ_REC0}{r}', f'=IF({PJ_NAME}{r}="","",SUMIFS({opr(OPJ_REC)},{opr(OPJ_PJ)},{PJ_NAME}{r})+{rec})',
            F_AUTO, FILL_AUTO, MONEY, AR)
    for i, p in enumerate(ctx['projects']):
        r = PJ_R0 + i
        vals = {PJ_NAME: p['name'], PJ_FULL: p['full'], PJ_CUS: p['cus'] or None, PJ_AMT: p['amt'], PJ_TAX: p['tax'],
                PJ_STAT: p.get('stat', '在建'), PJ_NOTE: p['note'] or None,
                PJ_START: p.get('start'), PJ_END: p.get('end')}
        kws = [k for k in p['kw'] if k != p['name']][:3]
        for c, k in zip((PJ_KW1, PJ_KW2, PJ_KW3), kws):
            vals[c] = k
        for c, v in vals.items():
            if v is not None:
                ws[f'{c}{r}'] = v
    dv_list(ws, f'{PJ_STAT}{PJ_R0}:{PJ_STAT}{PJ_R1}', '"' + ','.join(PJ_STATS) + '"')
    dv_list(ws, f'{PJ_CUS}{PJ_R0}:{PJ_CUS}{PJ_R1}', f'={SH_AUX}!$L$1:$L$100', '从【往来单位】类型是「甲方/总包」的里选，也可以直接打', stop=False)
    dv_date(ws, f'{PJ_START}{PJ_R0}:{PJ_START}{PJ_R1}')
    dv_date(ws, f'{PJ_END}{PJ_R0}:{PJ_END}{PJ_R1}')
    ws.freeze_panes = f'C{PJ_R0}'
    ws.auto_filter.ref = f'A{PJ_HDR}:{PJ_TOTAL}{PJ_R1}'
    return ws


def build_unit(wb, ctx):
    ws = wb.create_sheet(SH_UNIT)
    widths(ws, {'A': 5, 'B': 16, 'C': 10, 'D': 12, 'E': 12, 'F': 28, 'G': 20, 'H': 13, 'I': 21, 'J': 18, 'K': 11, 'L': 11,
                'M': 12, 'N': 34})
    title(ws, '往 来 单 位 及 人 员（甲方 · 材料商 · 分包 · 机械运输 · 股东 · 管理人员 · 工人，一个名字一行）', 'N', C_BASE,
          '💡 所有跟公司有钱来往的单位和人都在这一张：名称用你平时叫的（筑强、恩旗老叶、邓泽贵），流水摘要里出现名称、别名1、别名2 任何一个就自动认出来。'
          '「类型」很要紧：材料供应商、分包、机械运输（欠他们的钱在【应付登记】记）；管理人员、工人（工资在【考勤工资】算）；'
          '临时工（现结工资，不走考勤）；股东（借钱给公司的，每人在【基础资料】② 开一个个人户）。'
          '「期初欠薪」＝建账那天公司还欠他的工资（多发了、预支了填负数）。同名的人名字后面加个括号区分，比如 张伟(电工)。'
          '不用的单位不要删整行（别的表按行取数），清空内容或者在备注写「不用了」。')
    header(ws, UN_HDR, [(UN_SEQ, '序号'), (UN_NAME, '名称 / 姓名'), (UN_TYPE, '类型'), (UN_AL1, '别名1'), (UN_AL2, '别名2'),
                        (UN_FULL, '全称'), (UN_ID, '税号 / 身份证号'), (UN_TEL, '手机'), (UN_BANKNO, '银行账号'), (UN_BANK, '开户行'),
                        (UN_IN, '入职'), (UN_OUT, '离职'), (UN_OWE0, '期初欠薪\n（建账日）'), (UN_NOTE, '备注')], C_BASE)
    for i, r in enumerate(range(UN_R0, UN_R1 + 1)):
        put(ws, f'{UN_SEQ}{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in (UN_NAME, UN_TYPE, UN_AL1, UN_AL2, UN_FULL, UN_ID, UN_TEL, UN_BANKNO, UN_BANK, UN_IN, UN_OUT, UN_OWE0, UN_NOTE):
            fmt = {UN_OWE0: MONEY, UN_IN: DATE, UN_OUT: DATE, UN_ID: '@', UN_TEL: '@', UN_BANKNO: '@'}.get(c)
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, fmt, AL if c in (UN_NAME, UN_FULL, UN_BANK, UN_NOTE) else AC)
    for i, u in enumerate(ctx['units']):
        r = UN_R0 + i
        for c, k in ((UN_NAME, 'name'), (UN_TYPE, 'type'), (UN_AL1, 'al1'), (UN_AL2, 'al2'), (UN_FULL, 'full'), (UN_ID, 'id'),
                     (UN_BANKNO, 'bankno'), (UN_BANK, 'bank'), (UN_OWE0, 'owe0'), (UN_NOTE, 'note')):
            v = u.get(k)
            if v not in (None, ''):
                ws[f'{c}{r}'] = v
    dv_list(ws, f'{UN_TYPE}{UN_R0}:{UN_TYPE}{UN_R1}', '"' + ','.join(UN_TYPES) + '"',
            '材料供应商/分包/机械运输：欠钱记应付登记；管理人员/工人：工资走考勤；临时工：现结；股东：开个人户')
    dv_date(ws, f'{UN_IN}{UN_R0}:{UN_IN}{UN_R1}')
    dv_date(ws, f'{UN_OUT}{UN_R0}:{UN_OUT}{UN_R1}')
    ws.freeze_panes = f'C{UN_R0}'
    ws.auto_filter.ref = f'A{UN_HDR}:{UN_NOTE}{UN_R1}'
    return ws


def build_rate(wb, ctx):
    ws = wb.create_sheet(SH_RATE)
    widths(ws, {'A': 5, 'B': 12, 'C': 12, 'D': 10, 'E': 11, 'F': 30, 'G': 30})
    title(ws, '工 资 标 准（日薪 / 月薪 · 涨工资加一行，旧的不改）', 'G', C_BASE,
          '💡 每人一行：从哪天起、日薪或月薪（只填一个）。涨工资不要改原来那行，在下面加一行新的「从哪天起」和新单价：'
          '比如 胡家新 2026-01-01 起日薪 270，2026-07-01 起日薪 300 —— 考勤工资里 6 月及以前按 270 算，7 月起按 300 算，以前的工资不会变。'
          '按月查：用这个月最后一天有效的单价；月中涨薪的，考勤工资里这个人这个月拆两行，涨薪前那行在「单价手填」填旧单价。'
          '老板、负责人、办公室按月薪填。')
    header(ws, RT_HDR, [(RT_SEQ, '序号'), (RT_NAME, '姓名'), (RT_DATE, '从哪天起'), (RT_DAY, '日薪'), (RT_MON, '月薪'),
                        (RT_NOTE, '说明'), (RT_CHK, '校验')], C_BASE)
    for i, r in enumerate(range(RT_R0, RT_R1 + 1)):
        put(ws, f'{RT_SEQ}{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in (RT_NAME, RT_DATE, RT_DAY, RT_MON, RT_NOTE):
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, DATE if c == RT_DATE else (MONEY if c in (RT_DAY, RT_MON) else None),
                AL if c == RT_NOTE else AC)
        g = lambda c: f'{c}{r}'
        put(ws, g(RT_CHK), f'=IF(AND({g(RT_NAME)}="",{g(RT_DATE)}=""),"",IF(COUNTIF({UN_NAMES},{g(RT_NAME)})=0,"✗ 没在【往来单位】登记",'
                           f'IF(NOT(ISNUMBER({g(RT_DATE)})),"✗ 从哪天起要填日期",'
                           f'IF((N({g(RT_DAY)})<>0)=(N({g(RT_MON)})<>0),"✗ 日薪、月薪只填一个",'
                           f'IF(COUNTIFS({RT_NAMES},{g(RT_NAME)},{RT_DNS},{g(RT_DN)})>1,"✗ 同一天有两条，单价会算重","√")))))',
            F_AUTO, FILL_AUTO, align=AL)
        ws[g(RT_DN)] = f'=IF(ISNUMBER({g(RT_DATE)}),INT({g(RT_DATE)}),0)'
        ws[g(RT_DN)].font = F_HELP
    hide(ws, RT_DN)
    for i, (n, d, day, mon, note) in enumerate(ctx['rates']):
        r = RT_R0 + i
        ws[f'{RT_NAME}{r}'], ws[f'{RT_DATE}{r}'] = n, d
        if day:
            ws[f'{RT_DAY}{r}'] = day
        if mon:
            ws[f'{RT_MON}{r}'] = mon
        if note:
            ws[f'{RT_NOTE}{r}'] = note
    dv_list(ws, f'{RT_NAME}{RT_R0}:{RT_NAME}{RT_R1}', f'={UN_NAMES}', '【往来单位】登记过的人', stop=False)
    dv_date(ws, f'{RT_DATE}{RT_R0}:{RT_DATE}{RT_R1}')
    ws.freeze_panes = f'C{RT_R0}'
    ws.auto_filter.ref = f'A{RT_HDR}:G{RT_R1}'
    return ws
