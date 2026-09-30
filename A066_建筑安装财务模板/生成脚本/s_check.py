# -*- coding: utf-8 -*-
"""【数据校验】哪里录错了、漏了、快满了；【首页】选年度/截止日、关键数、项目盈亏前后 5 名、提醒、各表入口。"""
import datetime as dt
from openpyxl.formatting.rule import FormulaRule
from common import *
from layout import *
import s_agg
from s_agg import ms
from s_views import _lbl, _val, GREY
import s_proj as sp
import s_fin

CHK_ROWS = {}        # 检查项 → 行号（首页提醒用）


def _cnt_mark(col_rng, mark):
    return f'COUNTIF({col_rng},"{mark}*")'


def build_chk(wb, ctx):
    ws = wb.create_sheet(SH_CHK)
    widths(ws, {'A': 5, 'B': 40, 'C': 14, 'D': 8, 'E': 60, 'F': 14})
    title(ws, '数 据 校 验（录错的、漏录的、快满了的，都在这里）', 'F', C_CHK,
          '💡 全自动。✗＝一定要改（不改报表就不准），⚠＝看一下是不是真的有问题，√＝没问题。点「去看看」跳到那张表，按校验列筛选 ✗ 就能找到是哪几行。')
    home_link(ws, 'G1')
    header(ws, 4, [('A', '序号'), ('B', '检查什么'), ('C', '结果'), ('D', '状态'), ('E', '怎么处理'), ('F', '去看看')], C_CHK)
    PE = s_fin._bs_parts('E')
    RT_CHKS = rng(SH_RATE, RT_CHK, RT_R0, RT_R1)
    dup = lambda r: f'SUMPRODUCT(({r}<>"")*(COUNTIF({r},{r})>1))'
    used = lambda r: f'SUMPRODUCT(--({r}<>""))'
    items = [
        # (检查项, 结果公式, 状态公式（用 C 格）, 处理, 表)
        ('资金流水：要改的（✗）', _cnt_mark(jr(J_CHK), '✗'), 'x', '按「校验」列筛选 ✗，照提示改（多数是在「改类别/改项目/改单位」选一下）', SH_CASH),
        ('资金流水：要看一下的（⚠）', _cnt_mark(jr(J_CHK), '⚠'), 'w', '摘要里有两个项目、单位没登记等', SH_CASH),
        ('应付登记：要改的（✗）', _cnt_mark(apr(AP_CHK), '✗'), 'x', '单位、项目要先在档案里登记', SH_AP),
        ('收入确认：要改的（✗）', _cnt_mark(rvr(RV_CHK), '✗'), 'x', '', SH_REV),
        ('代发抵账：要改的（✗）', _cnt_mark(ofr(OF_CHK), '✗'), 'x', '代发工资的对象要是工人，代付款的对象要是材料商/分包', SH_OFF),
        ('发票登记：要改的（✗）', _cnt_mark(ivr(IV_CHK), '✗'), 'x', '多数是【基础资料】① 公司全称没填对，认不出销项/进项', SH_INV),
        ('发票登记：要看一下的（⚠）', _cnt_mark(ivr(IV_CHK), '⚠'), 'w', '没选项目、号码重复、开票单位没登记', SH_INV),
        ('考勤工资：要改的（✗）', _cnt_mark(atr(AT_CHK), '✗'), 'x', '人没登记、工资标准查不到单价、项目没选', SH_ATT),
        ('考勤工资：要看一下的（⚠）', _cnt_mark(atr(AT_CHK), '⚠'), 'w', '月中调价要拆两行、天数超过当月天数', SH_ATT),
        ('工资标准：要改的（✗）', _cnt_mark(RT_CHKS, '✗'), 'x', '', SH_RATE),
        ('月薪的人有月份没录考勤', f'N(COUNTIF({SH_PAYS}!$AD$6:$AD$205,"⚠*"))', 'w', '老板、负责人每个月也要录一行（不填天数＝全算公司管理费），不然管理费少了', SH_PAYS),
        ('工资汇总对不上的人', f'COUNTIF({SH_PAYS}!$P$6:$P$205,"✗*")', 'x', '', SH_PAYS),
        ('资金账户：表里余额跟银行 App 差多少', f'SUMIF({rng(SH_BASE, AC_DIFF, AC_R0, AC_R1)},">0")-SUMIF({rng(SH_BASE, AC_DIFF, AC_R0, AC_R1)},"<0")', 'w',
         '在【基础资料】② 填银行 App 上的余额和日期；有差额就是漏记、记错金额或日期', SH_BASE),
        ('资产负债表平不平（截止月末差额）', f'{s_fin.BS_DIFF}', 'x0', '不是 0 说明公式被改坏了，请联系做表的人', SH_BS),
        ('资产负债表平不平（年初差额）', f'{s_fin.BS_DIFF_B}', 'x0', '', SH_BS),
        ('待查：没分类、没选账户的流水净额', f'{s_fin.BS_PEND}', 'w0', '去【资金流水】改 ✗ 的行', SH_CASH),
        ('收付款没分到项目（应收账款）', f'ROUND({PE["应收残"]},2)', 'w0', '工程款没选项目、收入确认/代发抵账的项目写错', SH_AR),
        ('付款没分到单位（应付账款）', f'ROUND({PE["应付残"]},2)', 'w0', '付应付款没选单位、单位类型不是材料商/分包/机械', SH_APS),
        ('工资没分到人（应付职工薪酬）', f'ROUND({PE["工资残"]},2)', 'w0', '发工资没选人、人不是工人/管理人员', SH_PAYS),
        ('项目利润 ↔ 利润表 对账差额', f'{sp.ppl_bridge_diff()}', 'x0', '不是 0 说明公式被改坏了', SH_PPL),
        ('期初未分配利润：倒算的跟老板/会计认的差', f'IF(ISNUMBER({s_fin.BS_PLUG}),{s_fin.BS_PLUG},0)', 'w0',
         '差得多说明建账日的数没填全（银行期初、老板借款、欠材料商的、甲方欠的）', SH_OPEN),
        ('项目要注意的（成本多产值少、完工没结算、亏损）', f'SUMPRODUCT(--({SH_PPL}!${sp.C_TIP}${sp.PPL_R0}:${sp.C_TIP}${sp.PPL_R0 + sp.NPJ - 1}<>""))', 'w',
         '看【项目利润表】最右边「提示」', SH_PPL),
        ('以前年度还没定稿管理费率', f'SUMPRODUCT(({SH_ALLOC}!$A${sp.AL_Y0}:$A${sp.AL_Y0 + NYEARS - 1}<{AX_Y})*({SH_ALLOC}!$E${sp.AL_Y0}:$E${sp.AL_Y0 + NYEARS - 1}=""))',
         'w', '上一年的账定了，就把费率填到【费用分摊】E 列，以前年度的项目利润就不会再变', SH_ALLOC),
        ('期初项目有产值、没填税金', f'COUNTIFS({opr(OPJ_REV)},">0",{opr(OPJ_TAX)},"")', 'w', '建账前的项目税金（实交的或估算的）填在【期初余额】①，不然开工至今利润偏高', SH_OPEN),
        ('期初项目不在项目档案里', f'SUMPRODUCT(({opr(OPJ_PJ)}<>"")*(COUNTIF({PJ_NAMES},{opr(OPJ_PJ)})=0))', 'x', '项目名要跟【项目档案】的简称一样', SH_OPEN),
        ('期初应付的单位没登记', f'COUNTIF({oapr(OAP_TYPE)},"？")', 'x', '单位名要跟【往来单位】的名称一样（类型那列显示 ？ 的）', SH_OPEN),
        ('建账日期不是某月 1 号', f'IF(AND(ISNUMBER({OPEN_DATE}),DAY({OPEN_DATE})<>1),1,0)', 'x', '【基础资料】① 建账日期填某个月的 1 号（报表按月算）', SH_BASE),
        ('应付设备款是负数（付了设备款但没在⑥登记？）', f'MIN(0,{SH_BS}!${s_fin.BS_ROWS["应付设备款"][0]}${s_fin.BS_ROWS["应付设备款"][1]})',
         'w0', '买车买设备在【基础资料】⑥ 登记；建账前赊购没付完的在【期初余额】③ 填', SH_BASE),
        ('收入确认：要看一下的（⚠）', _cnt_mark(rvr(RV_CHK), '⚠'), 'w', '累计确认超过总价（结算额整笔又录了一遍？）、没选类型', SH_REV),
        ('项目简称重名', dup(PJ_NAMES), 'x', '项目简称不能重名（加个字区分）', SH_PROJ),
        ('往来单位/人重名', dup(UN_NAMES), 'x', '同名的人名字后面加括号区分，比如 张伟(电工)', SH_UNIT),
        ('资金账户重名', dup(AC_NAMES), 'x', '', SH_BASE),
        ('资金流水用了多少行（共 2000）', used(jr(J_DATE)), 'cap2000', '快满了请联系做表的人加行（或者一年一本）', SH_CASH),
        ('应付登记用了多少行（共 1000）', used(apr(AP_DATE)), 'cap1000', '', SH_AP),
        ('考勤工资用了多少行（共 1000）', used(atr(AT_MON)), 'cap1000', '', SH_ATT),
        ('发票登记用了多少行（共 1000）', used(ivr(IV_DATE)), 'cap1000', '', SH_INV),
        ('收入确认用了多少行（共 200）', used(rvr(RV_DATE)), 'cap200', '', SH_REV),
        ('往来单位及人员用了多少行（共 300）', f'COUNTIF({UN_NAMES},"?*")', 'cap300', '', SH_UNIT),
        ('项目档案用了多少行（共 60）', f'COUNTIF({PJ_NAMES},"?*")', 'cap60', '', SH_PROJ),
    ]
    for i, (lab, f, kind, how, sh) in enumerate(items):
        r = 5 + i
        CHK_ROWS[lab] = r
        put(ws, f'A{r}', i + 1, F_AUTO, align=AC)
        put(ws, f'B{r}', lab, F_TXT, align=AL)
        put(ws, f'C{r}', f'={f}', F_AUTOB, FILL_AUTO, INT if kind in ('x', 'w') or kind.startswith('cap') else MONEY, AR)
        if kind == 'x':
            st = f'IF(N(C{r})>0,"✗","√")'
        elif kind == 'w':
            st = f'IF(N(C{r})>0,"⚠","√")'
        elif kind == 'x0':
            st = f'IF(ABS(N(C{r}))>0.01,"✗","√")'
        elif kind == 'w0':
            st = f'IF(ABS(N(C{r}))>1,"⚠","√")'
        else:
            capn = int(kind[3:])
            st = f'IF(N(C{r})>{capn}*0.9,"⚠","√")'
            ws[f'C{r}'].number_format = f'0"/{capn}"'
        put(ws, f'D{r}', f'={st}', F_TXTB, align=AC)
        put(ws, f'E{r}', how, F_NOTE, align=ALW)
        c = put(ws, f'F{r}', f'=HYPERLINK("#\'{sh}\'!A1","去看看 →")', Font(name=YH, sz=10, color='FF0563C1', underline='single'), align=AC)
        ws.row_dimensions[r].height = 20
    R1 = 5 + len(items) - 1
    ws.conditional_formatting.add(f'D5:D{R1}', FormulaRule(formula=['D5="✗"'], fill=FILL_WARN, font=F_RED))
    ws.conditional_formatting.add(f'D5:D{R1}', FormulaRule(formula=['D5="⚠"'], fill=fill('FFFFEB9C')))
    ws.conditional_formatting.add(f'D5:D{R1}', FormulaRule(formula=['D5="√"'], font=Font(name=YH, sz=10, bold=True, color='FF00B050')))
    put(ws, 'B3', f'="要改 "&COUNTIF(D5:D{R1},"✗")&" 项，要看 "&COUNTIF(D5:D{R1},"⚠")&" 项"', F_RED, align=AL, border=False)
    ws.freeze_panes = 'A5'
    global CHK_X, CHK_W, CHK_RANGE
    CHK_X, CHK_W = f'COUNTIF({SH_CHK}!$D$5:$D${R1},"✗")', f'COUNTIF({SH_CHK}!$D$5:$D${R1},"⚠")'
    CHK_RANGE = (5, R1)
    return ws


CHK_X = CHK_W = CHK_RANGE = None


# ═══════════════════════════ 首页 ═══════════════════════════
NAV = [
    ('每天 / 每笔录', 'FF2F75B5', [(SH_CASH, '银行、现金、老板个人户的每一笔收支（摘要里的项目、单位、类别自动认）'),
                                  (SH_AP, '跟材料商、分包、机械对了账就登记应付（付款在流水里记）'),
                                  (SH_ATT, '每人每月一行：在哪个工地干了几天，工资自动算'),
                                  (SH_INV, '电子税务局导出的发票整块粘进来'),
                                  (SH_REV, '甲方/总包确认的产值、结算（收入按这里算）'),
                                  (SH_OFF, '总包代发工资、代付款、甲供材、甲方扣款')]),
    ('档案（偶尔改）', 'FF7F7F7F', [(SH_PROJ, '一个工地一行：合同额、甲方、开工完工、摘要关键词'),
                                  (SH_UNIT, '甲方、材料商、分包、机械、工人、老板都在这一张'),
                                  (SH_RATE, '日工资、月薪；涨工资加一行'),
                                  (SH_OPEN, '建账日以前的数，只填一次'),
                                  (SH_BASE, '公司、资金账户、收支类别、关键词、税率、固定资产')]),
    ('查看（自动）', 'FF548235', [(SH_ACC, '选一个账户看每一笔和余额'), ('资金报表', '资金周报 / 月报'),
                               (SH_AR, '每个项目甲方还欠多少、欠了多久'), (SH_APS, '欠材料商、分包、机械多少，对账单'),
                               (SH_PER, '老板、负责人垫了多少、公司欠他多少；保证金押金台账'), (SH_PAY, '选月份出工资表'),
                               (SH_PAYS, '每人本年应发、已发、欠薪'), (SH_LAB, '各项目每月人工'), (SH_INVS, '开票、收票、欠票')]),
    ('项目（自动）', 'FFC65911', [(SH_PL, '选一个项目看全部明细（不用每个项目做一张表）'), (SH_PPL, '全部项目本年和开工至今的利润'),
                               (SH_ALLOC, '老板工资、公司开支怎么摊到项目')]),
    ('老板报表（自动）', 'FF833C0C', [(SH_IS, '按月的利润表'), (SH_BS, '资产负债表'), (SH_BE, '保本产值、项目还能花多少、报价计算器'),
                                  (SH_CHK, '哪里录错了、漏了')]),
]


def build_home(wb, ctx):
    ws = wb.create_sheet(SH_HOME, 0)
    widths(ws, {'A': 2, 'B': 16, 'C': 15, 'D': 3, 'E': 16, 'F': 15, 'G': 3, 'H': 16, 'I': 15, 'J': 3, 'K': 16, 'L': 15})
    title(ws, '建 筑 安 装 财 务 模 板（消防改造 · 水电安装 · 强弱电）', 'L', C_HOME,
          '💡 平时只录蓝色那几张（资金流水、应付登记、考勤工资、发票、收入确认、代发抵账），其余全是自动的。'
          '下面黄格子选报表年度和截止日（截止日空着＝最后一笔流水那天），所有报表跟着变。')
    ws['A1'].value = f'={SH_BASE}!${CO_NAME[0]}${CO_NAME[1:]}&" · 建筑安装财务模板（消防改造 · 水电安装 · 强弱电）"'
    _lbl(ws, 'B4', '报表年度')
    put(ws, 'C4', ctx['open_date'].year, F_SEL, FILL_SEL, '0', AC)
    _lbl(ws, 'E4', '截止月份')
    put(ws, 'F4', None, F_SEL, FILL_SEL, DATE, AC)
    dv_date(ws, 'F4')
    dv = DataValidation(type='whole', operator='between', formula1='2000', formula2='2100', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add('C4')
    put(ws, 'H4', f'=IF(N(C4)<{AX_Y0},"⚠ 早于建账年（"&{AX_Y0}&"），没有数",'
                  f'"报表到："&TEXT({AX_E},"yyyy年m月d日")&"（截止日所在月的月底；最后一笔流水 "&TEXT({AX_LAST},"m月d日")&"）")',
        F_NOTE, align=AL, border=False)
    ws.merge_cells('H4:L4')
    if ctx.get('demo_note'):
        put(ws, 'B5', ctx['demo_note'], Font(name=YH, sz=10, bold=True, color='FFC00000'), fill('FFFFF2CC'), align=ALW)
        ws.merge_cells('B5:L5')
        ws.row_dimensions[5].height = 64
    # 关键数
    section(ws, 6, 'B', 'L', '关键数（到截止月底）', C_HOME)
    Y = lambda n: ms(n, MS_YTD)
    is_net = (f'{Y("营业收入")}+{Y("其他收入")}-(' + '+'.join(Y(n) for n in __import__('cats').IS_NAMES if n not in ('营业收入', '其他收入')) + ')')
    BS = lambda lab_col: lab_col
    cash_c, cash_r = s_fin.BS_ROWS['货币资金（银行＋现金）']
    kpis = [('本年确认产值', f'={Y("营业收入")}', SH_IS), ('本年净利润', f'={is_net}', SH_IS),
            ('银行＋现金余额', f'={SH_BS}!${cash_c}${cash_r}', '资金报表'), ('甲方还欠（应收净额）', f'=SUM({s_agg.ps(PS_AR_E)})', SH_AR),
            ('欠材料分包机械（应付净额）', f'=SUM({s_agg.bx(BX_AP_E)})', SH_APS), ('欠工资（净额，含管理人员）', f'=SUM({s_agg.bx(BX_WG_E)})', SH_PAYS),
            ('公司欠老板/负责人（个人户）', f'=-SUMIFS({s_agg.ba(BA_E)},{s_agg.ba(BA_TYPE)},"个人户")', SH_PER),
            ('全年保本产值', f'={SH_BE}!$C$16', SH_BE)]
    for i, (lab, f, sh) in enumerate(kpis):
        r = 7 + (i // 4) * 2
        c1 = ['B', 'E', 'H', 'K'][i % 4]
        c2 = CL(CI(c1) + 1)
        put(ws, f'{c1}{r}', f'=HYPERLINK("#\'{sh}\'!A1","{lab}")', Font(name=YH, sz=10, bold=True, color='FF1F3864', underline='single'),
            fill('FFD9E1F2'), align=AC)
        ws.merge_cells(f'{c1}{r}:{c2}{r}')
        put(ws, f'{c1}{r + 1}', f, F_KPI_V, fill('FFFFFFFF'), MONEY, AC)
        ws.merge_cells(f'{c1}{r + 1}:{c2}{r + 1}')
        ws.row_dimensions[r + 1].height = 26
    ws.conditional_formatting.add('B8:L8', FormulaRule(formula=['AND(ISNUMBER(B8),B8<0)'], font=Font(name=YH, sz=13, bold=True, color='FF00B050')))
    put(ws, 'B11', f'="安全边际："&IF(ISNUMBER({SH_BE}!$C$22),TEXT({SH_BE}!$C$22,"0%"),"-")&"　　"&{SH_BE}!$C$24', F_NOTE, align=AL, border=False)
    ws.merge_cells('B11:L11')
    # 项目盈亏前后 5 名（开工至今）
    section(ws, 13, 'B', 'F', '赚得最多的 5 个项目（开工至今）', 'FF548235')
    section(ws, 13, 'H', 'L', '亏得最多 / 赚得最少的 5 个项目', 'FFC00000')
    header(ws, 14, [('B', '项目'), ('C', '净利'), ('E', '确认产值'), ('F', '净利率')], 'FF548235')
    header(ws, 14, [('H', '项目'), ('I', '净利'), ('K', '确认产值'), ('L', '净利率')], 'FFC00000')
    ws.merge_cells('C14:D14')
    ws.merge_cells('I14:J14')
    PPL = lambda c: f'{SH_PPL}!${c}${sp.PPL_R0}:${c}${sp.PPL_R0 + sp.NPJ - 1}'
    for i in range(sp.NPJ):
        rr = 1 + i
        pr = sp.PPL_R0 + i
        g = lambda c: f'{SH_PPL}!${c}${pr}'
        act = f'OR(N({g(sp.C_REV)})<>0,N({g(sp.C_MAT)})+N({g(sp.C_LAB)})+N({g(sp.C_SUB)})+N({g(sp.C_MACH)})+N({g(sp.C_OTH)})<>0)'
        ws[f'Z{rr}'] = f'=IF(AND({g(sp.P_NAME)}<>"",{act}),N({g(sp.C_NET)})+({sp.NPJ + 1}-{rr})/10000000,"")'
        ws[f'Z{rr}'].font = F_HELP
    K = f'$Z$1:$Z${sp.NPJ}'
    for k in range(5):
        r = 15 + k
        for (cn, cv, cr, cp), fn in ((('B', 'C', 'E', 'F'), 'LARGE'), (('H', 'I', 'K', 'L'), 'SMALL')):
            ix = f'IFERROR(MATCH({fn}({K},{k + 1}),{K},0),0)'
            if fn == 'SMALL':      # 项目少于 10 个时，下面不重复列上面已经列了的
                ix = f'IF({k + 1}>COUNT({K})-5,0,{ix})'
            put(ws, f'{cn}{r}', f'=IF({ix}=0,"",INDEX({PPL(sp.P_NAME)},{ix}))', F_TXT, align=AL)
            put(ws, f'{cv}{r}', f'=IF({ix}=0,"",INDEX({PPL(sp.C_NET)},{ix}))', F_AUTOB, fmt=MONEY, align=AR)
            ws.merge_cells(f'{cv}{r}:{CL(CI(cv) + 1)}{r}')
            put(ws, f'{cr}{r}', f'=IF({ix}=0,"",INDEX({PPL(sp.C_REV)},{ix}))', F_AUTO, fmt=MONEY, align=AR)
            put(ws, f'{cp}{r}', f'=IF({ix}=0,"",INDEX({PPL(sp.C_RATE)},{ix}))', F_AUTO, fmt=PCT, align=AC)
    ws.conditional_formatting.add('C15:C19', FormulaRule(formula=['N(C15)<0'], font=F_RED))
    ws.conditional_formatting.add('I15:I19', FormulaRule(formula=['N(I15)<0'], font=F_RED))
    hide(ws, 'Z')
    # 提醒
    section(ws, 21, 'B', 'L', '提醒', C_CHK)
    r0, r1 = CHK_RANGE
    rem = [f'="数据校验：要改 "&{CHK_X}&" 项、要看 "&{CHK_W}&" 项"&IF({CHK_X}+{CHK_W}=0,"，都没问题 √","  → 点这里去看")',
           f'=IF(N({SH_CHK}!$C${CHK_ROWS["资金流水：要改的（✗）"]})>0,"资金流水有 "&{SH_CHK}!$C${CHK_ROWS["资金流水：要改的（✗）"]}&" 行要改（没认出类别、项目或单位）","资金流水都认出来了 √")',
           f'=IF(N({SH_CHK}!$C${CHK_ROWS["项目要注意的（成本多产值少、完工没结算、亏损）"]})>0,{SH_CHK}!$C${CHK_ROWS["项目要注意的（成本多产值少、完工没结算、亏损）"]}&" 个项目要注意：成本花得多产值确认得少（该找甲方报量）、完工没结算或亏损 → 看【项目利润表】提示","")',
           f'=IF(COUNTIF({SH_AR}!$Y$6:$Y${5 + sp.NPJ},"质保金*")>0,COUNTIF({SH_AR}!$Y$6:$Y${5 + sp.NPJ},"质保金*")&" 个项目质保金到期可以催了","")'
           f'&IF(N({SH_AR}!$L$4)>0,"  开了票没回款的 "&TEXT({SH_AR}!$L$4,"#,##0")&" 元","")',
           f'=IF(N({SH_CHK}!$C${CHK_ROWS["月薪的人有月份没录考勤"]})>0,{SH_CHK}!$C${CHK_ROWS["月薪的人有月份没录考勤"]}&" 个管理人员有月份没录考勤（管理费会少算）","")',
           f'=IF(ABS(N({SH_CHK}!$C${CHK_ROWS["资金账户：表里余额跟银行 App 差多少"]}))>0.01,"账户余额跟银行 App 对不上，差 "&TEXT({SH_CHK}!$C${CHK_ROWS["资金账户：表里余额跟银行 App 差多少"]},"#,##0.00"),"")']
    for i, f in enumerate(rem):
        r = 22 + i
        put(ws, f'B{r}', f, F_TXT if i else F_TXTB, align=AL, border=False)
        ws.merge_cells(f'B{r}:L{r}')
    link(ws['B22'], SH_CHK)
    ws.conditional_formatting.add('B22:B27', FormulaRule(formula=['OR(ISNUMBER(SEARCH("要改",B22)),ISNUMBER(SEARCH("要注意",B22)),ISNUMBER(SEARCH("对不上",B22)))'],
                                                         font=Font(name=YH, sz=10, bold=True, color='FFC00000')))
    # 导航
    n0 = 29
    section(ws, n0, 'B', 'L', '各张表（点名字跳过去；每张表右上角有「← 回首页」）', C_HOME)
    r = n0 + 1
    col_sets = [('B', 'C', 'F'), ('H', 'I', 'L')]
    for gi, (grp, color, sheets) in enumerate(NAV):
        put(ws, f'B{r}', grp, F_SEC, fill(color), align=AL)
        ws.merge_cells(f'B{r}:L{r}')
        r += 1
        for j, (sh, desc) in enumerate(sheets):
            c1, c2, c3 = col_sets[j % 2]
            rr = r + j // 2
            cell = put(ws, f'{c1}{rr}', sh, Font(name=YH, sz=10, bold=True, color='FF0563C1', underline='single'), align=AL)
            link(cell, sh)
            put(ws, f'{c2}{rr}', desc, F_NOTE, align=AL)
            ws.merge_cells(f'{c2}{rr}:{c3}{rr}')
        r += (len(sheets) + 1) // 2
    # 使用步骤
    s0 = r + 1
    section(ws, s0, 'B', 'L', '怎么用（第一次看这里）', C_HOME)
    steps = ['① 先填档案：【基础资料】公司全称、建账日期、资金账户（每个老板/负责人一个「个人户」）；【项目档案】每个工地一行；【往来单位】甲方、材料商、分包、机械、工人、老板；【工资标准】每人日工资/月薪。',
             '② 建账日以前的账：在【期初余额】填一次——每个老项目累计确认的产值、收款、成本；每家材料商/分包每个项目累计应付、已付；银行期初余额在【基础资料】②。',
             '③ 平时：每笔钱进出记【资金流水】（摘要写清楚项目和对象，类别自动认；认不出的在「改类别」选）；跟材料商/分包对了账记【应付登记】；每月月底录【考勤工资】。',
             '④ 甲方/总包确认了产值（报量批了、结算了）就记【收入确认】——利润按这个算，不按收钱算。总包代发工资、甲方扣款记【代发抵账】。发票每月从电子税务局导出粘到【发票登记】。',
             '⑤ 看报表：首页选年度；【项目账】选项目看明细；【项目利润表】看哪个项目赚哪个亏；【利润表】【资产负债表】【盈亏平衡表】给老板看。每年年底在【费用分摊】填定稿费率。',
             '⑥ 每月对一次账：【基础资料】② 填银行 App 上的实际余额，差额是 0 就对了；再看一眼【数据校验】有没有 ✗。',
             '⑦ 注意：记错了就清空那几格，不要删整行、插行（别的表按行取数，删行公式会出错）；灰色格子是公式，不要往里打字。']
    for i, t in enumerate(steps):
        rr = s0 + 1 + i
        put(ws, f'B{rr}', t, F_TXT, align=ALW, border=False)
        ws.merge_cells(f'B{rr}:L{rr}')
        ws.row_dimensions[rr].height = 32
    ws.sheet_view.showGridLines = False
    return ws


def build_all(wb, ctx):
    build_chk(wb, ctx)
    build_home(wb, ctx)
