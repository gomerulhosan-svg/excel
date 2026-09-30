# -*- coding: utf-8 -*-
"""【基础资料】① 公司 ② 资金账户 ③ 收支类别 ④ 摘要关键词 ⑤ 参数 ⑥ 固定资产；【_辅助】下拉清单和关键词表。"""
import datetime as dt
from common import *
from layout import *
import cats


def build_base(wb, ctx):
    ws = wb.create_sheet(SH_BASE)
    widths(ws, {'A': 2, 'B': 12, 'C': 22, 'D': 2, 'E': 2,
                'F': 5, 'G': 13, 'H': 9, 'I': 8, 'J': 20, 'K': 13, 'L': 13, 'M': 13, 'N': 11, 'O': 12, 'P': 24, 'Q': 2,
                'R': 5, 'S': 13, 'T': 6, 'U': 12, 'V': 12, 'W': 8, 'X': 40, 'Y': 2,
                'Z': 5, 'AA': 12, 'AB': 6, 'AC': 13, 'AD': 2,
                'AE': 22, 'AF': 11, 'AG': 44, 'AH': 2,
                'AI': 5, 'AJ': 14, 'AK': 11, 'AL': 12, 'AM': 7, 'AN': 7, 'AO': 12, 'AP': 10, 'AQ': 18})
    title(ws, '基 础 资 料（公司 · 资金账户 · 收支类别 · 摘要关键词 · 参数 · 固定资产）', 'AQ', C_BASE,
          '💡 淡黄格子手填，灰格子自动。① 公司和建账日期；② 资金账户：银行、现金，还有老板、负责人每人一个「个人户」（他替公司花的钱记在他的个人户，'
          '借给公司、公司还他、报销、备用金都是银行和他个人户之间的互转），承兑汇票在手的记「票据」户；'
          '③ 收支类别：每一类决定钱记到报表哪一行、要不要摊到项目（一般不用改，要加类别往下加）；'
          '④ 摘要关键词：流水摘要里出现这个词就自动认成这一类（按从上到下的顺序，具体的放上面）；⑤ 税率、分摊方法；⑥ 车、设备等固定资产（自动提折旧）。')
    # ① 公司
    section(ws, 3, 'B', 'C', '① 公司', C_HOME)
    for (lab, cell, val, fmt) in (('公司简称', CO_NAME, ctx['company'][0], None), ('公司全称', CO_FULL, ctx['company'][1], None),
                                  ('税号', CO_TAX, ctx['company'][2], '@'), ('纳税人', CO_TYPE, '一般纳税人', None),
                                  ('建账日期', CO_OPEN, ctx['open_date'], DATE)):
        r = int(cell[1:])
        put(ws, f'B{r}', lab, F_KPI_L, fill('FFD9E1F2'), align=AC)
        put(ws, cell, val, F_IN, FILL_IN, fmt, AC)
    put(ws, 'B10', '建账日期填某个月的 1 号：这天以前的账在【期初余额】一次性录，这天起每一笔按实际日期录。'
                   '小规模纳税人：⑤ 销项税率改成 3%（或 1%），项目里填的 9% 自动不算，进项票不抵扣（自动）。', F_NOTE, align=ALW, border=False)
    ws.merge_cells('B10:C14')
    dv_list(ws, CO_TYPE, '"一般纳税人,小规模纳税人"')
    # ② 资金账户
    section(ws, 3, AC_SEQ, AC_NOTE, '② 资金账户（银行 · 现金 · 个人户 · 票据）', C_CASH)
    header(ws, B_HDR, [(AC_SEQ, '序号'), (AC_NAME, '账户名称'), (AC_TYPE, '类型'), (AC_OWNER, '个人户\n是谁的'), (AC_NO, '账号'),
                       (AC_OPEN, '期初余额\n（建账日）'), (AC_NOW, '表里余额\n（自动）'), (AC_REAL, '实际余额\n（银行App）'),
                       (AC_RDATE, '核对日期'), (AC_DIFF, '差额\n（自动）'), (AC_NOTE, '备注')], C_CASH)
    for i, r in enumerate(range(AC_R0, AC_R1 + 1)):
        put(ws, f'{AC_SEQ}{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in (AC_NAME, AC_TYPE, AC_OWNER, AC_NO, AC_OPEN, AC_REAL, AC_RDATE, AC_NOTE):
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, MONEY if c in (AC_OPEN, AC_REAL) else ('@' if c == AC_NO else (DATE if c == AC_RDATE else None)),
                AL if c in (AC_NOTE, AC_NO) else AC)
        bal = lambda upto: (f'N({AC_OPEN}{r})+SUMIFS({jr(J_NET)},{jr(J_ACC)},{AC_NAME}{r}{upto})'
                            f'-SUMIFS({jr(J_NET)},{jr(J_TO)},{AC_NAME}{r},{jr(J_LINE)},"账户互转"{upto})')
        put(ws, f'{AC_NOW}{r}', f'=IF({AC_NAME}{r}="","",ROUND({bal("")},2))', F_AUTOB, FILL_AUTO, MONEY, AR)
        upto = f',{jr(J_DATE)},"<="&{AC_RDATE}{r}'
        put(ws, f'{AC_DIFF}{r}', f'=IF(OR({AC_NAME}{r}="",{AC_REAL}{r}="",NOT(ISNUMBER({AC_RDATE}{r}))),"",ROUND({bal(upto)}-{AC_REAL}{r},2))',
            F_AUTOB, FILL_AUTO, MONEY, AR)
    for i, (name, typ, owner, no, op, note) in enumerate(ctx['accounts']):
        r = AC_R0 + i
        for c, v in ((AC_NAME, name), (AC_TYPE, typ), (AC_OWNER, owner), (AC_NO, no), (AC_OPEN, op), (AC_NOTE, note)):
            if v not in (None, ''):
                ws[f'{c}{r}'] = v
    dv_date(ws, f'{AC_RDATE}{AC_R0}:{AC_RDATE}{AC_R1}')
    dv_list(ws, f'{AC_TYPE}{AC_R0}:{AC_TYPE}{AC_R1}', '"' + ','.join(AC_TYPES) + '"',
            '个人户：老板/负责人替公司收付钱用的（余额为负＝公司欠他）；票据：在手的承兑汇票')
    dv_list(ws, f'{AC_OWNER}{AC_R0}:{AC_OWNER}{AC_R1}', f'={UN_NAMES}', '个人户填是谁的（【往来单位】里的人）', stop=False)
    # ③ 收支类别
    section(ws, 3, CT_SEQ, CT_NOTE, '③ 收支类别（每一类记到报表哪一行、要不要摊到项目）', C_RPT)
    header(ws, B_HDR, [(CT_SEQ, '序号'), (CT_NAME, '收支类别'), (CT_DIR, '方向'), (CT_LINE, '报表项目'), (CT_ALLOC, '分摊归类\n（自动）'),
                       (CT_PROJ, '要选'), (CT_NOTE, '什么钱记这一类')], C_RPT)
    allcats = cats.CATS + cats.ACCRUAL_CATS
    direct = [n for n, g in cats.IS_LINES if g in ('直接成本', '项目财务费')]
    pool = [n for n, g in cats.IS_LINES if g == '间接费用']
    other = [n for n, g in cats.IS_LINES if g not in ('直接成本', '项目财务费', '间接费用', '收入', '税金')]
    inl = lambda x, names: 'OR(' + ','.join(f'{x}="{n}"' for n in names) + ')'
    for i, r in enumerate(range(CT_R0, CT_R1 + 1)):
        put(ws, f'{CT_SEQ}{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in (CT_NAME, CT_DIR, CT_LINE, CT_PROJ, CT_NOTE):
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, align=AL if c == CT_NOTE else AC)
        ln = f'{CT_LINE}{r}'
        put(ws, f'{CT_ALLOC}{r}', f'=IF({CT_NAME}{r}="","",IF({inl(ln, direct)},"{cats.AL_DIRECT}",IF({inl(ln, pool)},"{cats.AL_POOL}",'
                                  f'IF({inl(ln, other)},"{cats.AL_NOALLOC}","{cats.AL_BS}"))))', F_AUTO, FILL_AUTO, align=AC)
        if i < len(allcats):
            n, dr, line, al, need, note = allcats[i]
            ws[f'{CT_NAME}{r}'], ws[f'{CT_DIR}{r}'], ws[f'{CT_LINE}{r}'] = n, dr, line
            if need:
                ws[f'{CT_PROJ}{r}'] = need
            ws[f'{CT_NOTE}{r}'] = note
            if allcats[i] in cats.ACCRUAL_CATS:
                for c in (CT_NAME, CT_DIR, CT_LINE, CT_PROJ):
                    ws[f'{c}{r}'].fill = FILL_AUTO
    dv_list(ws, f'{CT_DIR}{CT_R0}:{CT_DIR}{CT_R1}', '"收,支,双向"', '支：支出栏填负数＝冲回；双向：借进借出、互转、保证金')
    dv_list(ws, f'{CT_LINE}{CT_R0}:{CT_LINE}{CT_R1}', f'={SH_AUX}!$H$1:$H$60',
            '这一类记到报表哪一行。「分摊归类」跟着自动变：材料人工分包机械其他直接费＝项目直接成本（要选项目）；办公招待车辆等＝公司间接费（年底按费率摊到项目）；'
            '利息手续费其他收支＝公司不分摊；借还款、互转、交税、付应付款＝不进利润')
    dv_list(ws, f'{CT_PROJ}{CT_R0}:{CT_PROJ}{CT_R1}', '"项目,人,单位,对方账户"')
    # ④ 摘要关键词
    section(ws, 3, KW_SEQ, KW_CAT, '④ 摘要关键词 → 收支类别', C_INV)
    header(ws, B_HDR, [(KW_SEQ, '序号'), (KW_WORD, '关键词'), (KW_DIR, '收/支'), (KW_CAT, '收支类别')], C_INV)
    for i, r in enumerate(range(KW_R0, KW_R1 + 1)):
        put(ws, f'{KW_SEQ}{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in (KW_WORD, KW_DIR, KW_CAT):
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, align=AC)
        if i < len(cats.KEYWORDS):
            ws[f'{KW_WORD}{r}'], ws[f'{KW_DIR}{r}'], ws[f'{KW_CAT}{r}'] = cats.KEYWORDS[i]
    dv_list(ws, f'{KW_DIR}{KW_R0}:{KW_DIR}{KW_R1}', '"收,支,全"')
    dv_list(ws, f'{KW_CAT}{KW_R0}:{KW_CAT}{KW_R1}', f'={CT_NAMES}')
    # ⑤ 参数
    section(ws, 3, PA_LBL, PA_NOTE, '⑤ 参数', C_CHK)
    header(ws, B_HDR, [(PA_LBL, '参数'), (PA_VAL, '值'), (PA_NOTE, '说明')], C_CHK)
    for i, (lab, v, fmt, note) in enumerate(PARAMS):
        r = PA_R0 + i
        put(ws, f'{PA_LBL}{r}', lab, F_TXT, align=AL)
        put(ws, f'{PA_VAL}{r}', v, F_IN, FILL_IN, fmt, AC)
        put(ws, f'{PA_NOTE}{r}', note, F_NOTE, align=ALW)
        ws.row_dimensions[r].height = 28
    dv_list(ws, PA_DRV.split('!')[1].replace('$', ''), '"施工费,直接成本"')
    # ⑥ 固定资产
    section(ws, 3, FA_SEQ, FA_NOTE, '⑥ 固定资产（车、设备，5000 元以下直接算费用；购入次月起按月提折旧）', C_AR)
    hide(ws, FA_S, FA_E)
    header(ws, B_HDR, [(FA_SEQ, '序号'), (FA_NAME, '名称'), (FA_DATE, '购入日期'), (FA_COST, '原值'), (FA_MON, '用几个月'),
                       (FA_RES, '残值率'), (FA_PJ, '使用项目\n（空＝公司）'), (FA_DEP_M, '月折旧'), (FA_NOTE, '备注')], C_AR)
    for i, r in enumerate(range(FA_R0, FA_R1 + 1)):
        put(ws, f'{FA_SEQ}{r}', i + 1, F_AUTO, FILL_AUTO, align=AC)
        for c in (FA_NAME, FA_DATE, FA_COST, FA_MON, FA_RES, FA_PJ, FA_NOTE):
            put(ws, f'{c}{r}', None, F_IN, FILL_IN, {FA_DATE: DATE, FA_COST: MONEY, FA_RES: '0%'}.get(c),
                AL if c in (FA_NAME, FA_NOTE) else AC)
        put(ws, f'{FA_DEP_M}{r}', f'=IF(OR(N({FA_COST}{r})=0,N({FA_MON}{r})=0,NOT(ISNUMBER({FA_DATE}{r}))),0,'
                                  f'ROUND({FA_COST}{r}*(1-N({FA_RES}{r}))/{FA_MON}{r},2))',
            F_AUTO, FILL_AUTO, MONEY, AR)
        # 隐藏：开始提折旧的月序号（购入次月）、提完的月序号
        ws[f'{FA_S}{r}'] = f'=IF(ISNUMBER({FA_DATE}{r}),YEAR({FA_DATE}{r})*12+MONTH({FA_DATE}{r})+1,0)'
        ws[f'{FA_E}{r}'] = f'=IF(ISNUMBER({FA_DATE}{r}),{FA_S}{r}+N({FA_MON}{r})-1,-1)'
        ws[f'{FA_S}{r}'].font = ws[f'{FA_E}{r}'].font = F_HELP
    for i, fa in enumerate(ctx.get('fixed_assets', [])):
        r = FA_R0 + i
        for c, v in zip((FA_NAME, FA_DATE, FA_COST, FA_MON, FA_RES, FA_PJ, FA_NOTE), fa):
            if v not in (None, ''):
                ws[f'{c}{r}'] = v
    dv_date(ws, f'{FA_DATE}{FA_R0}:{FA_DATE}{FA_R1}')
    dv_list(ws, f'{FA_PJ}{FA_R0}:{FA_PJ}{FA_R1}', f'={PJ_NAMES}')
    ws.freeze_panes = 'A5'
    return ws


def build_aux(wb, ctx):
    """_辅助：A/B 项目关键词 → 项目；C/D 单位人员名字和别名 → 标准名；E 公司管理＋项目（考勤下拉）；F/G 账户关键词 → 账户；H 报表项目清单；
       P～AA 上面三张关键词表按长短排好"""
    ws = wb.create_sheet(SH_AUX)
    n = PJ_R1 - PJ_R0 + 1
    for i in range(n):
        p = PJ_R0 + i
        for k, col in enumerate((PJ_NAME, PJ_KW1, PJ_KW2, PJ_KW3)):
            r = i * 4 + k + 1
            ws[f'{AX_PKW}{r}'] = f'=IF({SH_PROJ}!${PJ_NAME}${p}="","",TRIM({SH_PROJ}!${col}${p}&""))'
            ws[f'{AX_PKP}{r}'] = f'={SH_PROJ}!${PJ_NAME}${p}&""'
    nu = UN_R1 - UN_R0 + 1
    for i in range(nu):
        u = UN_R0 + i
        for k, col in enumerate((UN_NAME, UN_AL1, UN_AL2)):
            r = i * 3 + k + 1
            ws[f'{AX_UKW}{r}'] = f'=IF({SH_UNIT}!${UN_NAME}${u}="","",TRIM({SH_UNIT}!${col}${u}&""))'
            ws[f'{AX_UKN}{r}'] = f'={SH_UNIT}!${UN_NAME}${u}&""'
    np_ = PJ_R1 - PJ_R0 + 1
    ws[f'{AX_ALL}1'] = CO_PSEUDO
    for i in range(np_):
        ws[f'{AX_ALL}{i + 2}'] = f'=IF({SH_PROJ}!${PJ_NAME}${PJ_R0 + i}="","",{SH_PROJ}!${PJ_NAME}${PJ_R0 + i})'
    na = AC_R1 - AC_R0 + 1
    for i in range(na):
        a = AC_R0 + i
        ws[f'{AX_AKW}{i + 1}'] = f'=TRIM({SH_BASE}!${AC_NAME}${a}&"")'
        ws[f'{AX_AKA}{i + 1}'] = f'={SH_BASE}!${AC_NAME}${a}&""'
        ws[f'{AX_AKW}{na + i + 1}'] = f'=IF({SH_BASE}!${AC_TYPE}${a}="个人户",TRIM({SH_BASE}!${AC_OWNER}${a}&""),"")'
        ws[f'{AX_AKA}{na + i + 1}'] = f'={SH_BASE}!${AC_NAME}${a}&""'
    # P～AA：三张关键词表按长短排好（流水认项目、单位、对方账户用：排在前面的先对上＝最长的）
    for W, I, S, N, kw, val, n_ in ((AX_PW, AX_PI, AX_PS, AX_PN, AX_PKW, AX_PKP, AX_PKW_N),
                                    (AX_UW, AX_UI, AX_US, AX_UN, AX_UKW, AX_UKN, AX_UKW_N),
                                    (AX_AW, AX_AI, AX_AS, AX_AN, AX_AKW, AX_AKA, AX_AKW_N)):
        wr = f'${W}$1:${W}${n_}'
        for r in range(1, n_ + 1):
            ws[f'{W}{r}'] = f'=IF({kw}{r}="",0,LEN({kw}{r})*10000+{10000 - r})'
            ws[f'{I}{r}'] = f'=IF(LARGE({wr},{r})=0,0,MATCH(LARGE({wr},{r}),{wr},0))'
            ws[f'{S}{r}'] = f'=IF({I}{r}=0,"",INDEX(${kw}$1:${kw}${n_},{I}{r})&"")'
            ws[f'{N}{r}'] = f'=IF({I}{r}=0,"",INDEX(${val}$1:${val}${n_},{I}{r}))'
    for i in range(KW_R1 - KW_R0 + 1):
        k = KW_R0 + i
        w, d = f'{SH_BASE}!${KW_WORD}${k}', f'{SH_BASE}!${KW_DIR}${k}'
        for col, io in ((AX_KWI, '收'), (AX_KWO, '支')):
            ws[f'{col}{i + 1}'] = f'=IF(OR({d}="全",{d}="{io}"),{w}&"","")'
    lines = [n for n in cats.IS_NAMES if n not in ('营业收入', '税金估算')] + ['应收账款', '应付款', '应付职工薪酬', '短期借款', '其他应收款', '应交税费', '应付设备款', '实收资本',
                             '账户互转']
    for i, l in enumerate(lines):
        ws[f'H{i + 1}'] = l
    # J 列：报表年度、截止日推出来的几个数（各报表都用）
    last = f'MAX({jr(J_DATE)})'
    j = {
        AX_Y: f'=IF(ISNUMBER({SEL_Y}),{SEL_Y},YEAR({AX_E.replace("$J$9", "$J$10")}))',
        AX_LAST: f'=IF({last}=0,{OPEN_DATE},{last})',
        AX_E: f'=EOMONTH(IF(ISNUMBER({SEL_E}),{SEL_E},IF(YEAR({AX_LAST})={AX_Y},{AX_LAST},DATE({AX_Y},12,31))),0)',
        AX_YM0: f'={AX_Y}*100+1',
        AX_YM1: f'=IF(YEAR({AX_E})={AX_Y},YEAR({AX_E})*100+MONTH({AX_E}),IF(YEAR({AX_E})>{AX_Y},{AX_Y}*100+12,{AX_Y}*100+1))',
        AX_END: f'=EOMONTH(DATE(INT({AX_YM1}/100),MOD({AX_YM1},100),1),0)',
        AX_OYM: f'=YEAR({OPEN_DATE})*100+MONTH({OPEN_DATE})',
        AX_PYM: f'=({AX_Y}-1)*100+12',
        AX_Y0: f'=YEAR({OPEN_DATE})',
        AX_MON: f'=MOD({AX_YM1},100)',
    }
    labels = {AX_YM0: '本年1月', AX_YM1: '截止月', AX_END: '截止月末', AX_OYM: '建账年月', AX_PYM: '上年12月', AX_Y0: '建账年',
              AX_Y: '报表年度', AX_MON: '截止几月', AX_E: '截止日（月底）', AX_LAST: '最后一笔'}
    for cell, f in j.items():
        a = cell.split('!')[1].replace('$', '')
        ws[a] = f
        ws['I' + a[1:]] = labels[cell]
    nu = UN_R1 - UN_R0 + 1
    counter(ws, 'N', 1, nu, lambda i: f'{SH_UNIT}!${UN_TYPE}${UN_R0 + i}="甲方/总包"')
    for i in range(nu):
        ws[f'L{i + 1}'] = f'=IFERROR(INDEX({UN_NAMES},MATCH({i + 1},$N$1:$N${nu},0))&"","")' if i < 100 else None
    ws.sheet_state = 'hidden'
    return ws
