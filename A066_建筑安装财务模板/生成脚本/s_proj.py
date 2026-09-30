# -*- coding: utf-8 -*-
"""项目的三张表：【费用分摊】（每年一个管理费率，按施工费摊到项目）【项目利润表】（全部项目：本年 ＋ 开工至今）【项目账】（选一个项目看全部明细）"""
from openpyxl.formatting.rule import FormulaRule
from common import *
from layout import *
import calc
import cats
import s_agg
from s_agg import ms
from s_views import _sheet, _lbl, _val, jc, GREY, AR_R0

NPJ = PJ_R1 - PJ_R0 + 1
POOL = [n for n, g in cats.IS_LINES if g == '间接费用']
DIRECT = ['材料费', '人工费', '分包费', '机械运输费', '其他直接费']

# 费用分摊的位置（项目利润表、项目账要引用）
AL_Y0 = 8                     # ① 各年费率：第 8 行起 NYEARS 行
AL_P0 = AL_Y0 + NYEARS + 5    # ② 项目 × 年：数据起始行
AL_YC = [CL(5 + i) for i in range(NYEARS)]   # E..L 各年摊到
AL_OPEN, AL_CUM, AL_DRV, AL_NOW, AL_SHARE, AL_OLD, AL_DIFF, AL_ACT = 'D', 'M', 'N', 'O', 'P', 'Q', 'R', 'S'

# 项目利润表的列
PPL_H = 5
PPL_TOT = 6
PPL_R0 = 7
(P_SEQ, P_NAME, P_STAT, P_TOTAL,
 P_REV, P_MAT, P_LAB, P_SUB, P_MACH, P_OTH, P_TAX, P_INT, P_GROSS, P_ALLOC, P_NET, P_RATE,
 C_REV, C_MAT, C_LAB, C_SUB, C_MACH, C_OTH, C_TAX, C_INT, C_ALLOC, C_NET, C_RATE, C_PROG, C_CPROG, C_REC, C_ARB, C_TIP) = \
    [CL(i) for i in range(1, 33)]


def _pool(y):
    """某一年（到截止月为止）的间接费用合计"""
    a, b = f'MAX({y}*100+1,{AX_OYM})', f'MIN({y}*100+12,{AX_YM1})'
    return '(' + '+'.join(calc.line(n, a, b) for n in POOL) + ')'


# ═══════════════════════════ 费用分摊 ═══════════════════════════
def build_alloc(wb, ctx):
    ws = _sheet(wb, SH_ALLOC, 'FFC65911', '费 用 分 摊（老板工资、办公室开支怎么摊到每个项目）',
                '💡 办法：每年算一个「管理费率」＝这一年公司的间接费用（管理人员在公司的工资、社保、办公、招待、车辆、伙食、折旧等）÷ 这一年所有项目的施工费（人工＋分包＋机械），'
                '每个项目这一年摊到的管理费＝它这一年的施工费 × 费率。年中按到目前为止的费率先摊，年底在「定稿费率」填最后的数，以前年度就不会再变。'
                '好处：谁干的活多（工人、分包、机械用得多）谁多摊，工期长的自然摊得多，工期一周的只摊一周的；工程款几年收不到也不影响；同一年的间接费用全部摊完，一分不多一分不少。',
                'S', {'A': 8, 'B': 16, 'C': 12, 'D': 13, **{c: 12 for c in AL_YC}, 'M': 14, 'N': 14, 'O': 13, 'P': 9, 'Q': 13, 'R': 12, 'S': 8})
    put(ws, 'B3', f'="分摊依据：【基础资料】⑤ 选的「"&{PA_DRV}&"」（"&IF({PA_DRV}="直接成本","材料＋人工＋分包＋机械＋其他直接费","人工＋分包＋机械")&"）；截止 "&TEXT({AX_E},"yyyy年m月d日")',
        F_KPI_L, align=AL, border=False)
    ws.merge_cells('B3:M3')
    section(ws, AL_Y0 - 2, 'A', 'I', '① 各年管理费率', 'FFC65911')
    header(ws, AL_Y0 - 1, [('A', '年份'), ('B', '间接费用合计'), ('C', '分摊依据合计'), ('D', '算出来的\n费率'), ('E', '定稿费率\n（年底填）'),
                           ('F', '采用的费率'), ('G', '摊到项目合计'), ('H', '没摊完(＋)\n多摊了(−)'), ('I', '说明')], 'FFC65911', height=40)
    ws.merge_cells(f'I{AL_Y0 - 1}:L{AL_Y0 - 1}')
    for i in range(NYEARS):
        r = AL_Y0 + i
        y = f'$A{r}'
        put(ws, f'A{r}', f'={AX_Y0}+{i}', F_TXTB, align=AC)
        on = f'{y}<={AX_Y}'
        put(ws, f'B{r}', f'=IF({on},IF({y}={AX_Y},{"+".join(ms(n, MS_YTD) for n in POOL)},{_pool(y)}),"")', F_AUTO, fmt=MONEY, align=AR)
        dcol = PS_DRV[i]
        put(ws, f'C{r}', f'=IF({on},SUM({SH_PS}!${dcol}${PS_R0}:${dcol}${PS_R0 + NPJ - 1}),"")', F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'D{r}', f'=IF(OR(NOT({on}),N(C{r})=0),"",B{r}/C{r})', F_AUTO, fmt=PCT, align=AC)
        put(ws, f'E{r}', None, F_IN, FILL_IN, PCT, AC)
        put(ws, f'F{r}', f'=IF(NOT({on}),"",IF(ISNUMBER(E{r}),E{r},IF(ISNUMBER(D{r}),D{r},0)))', F_AUTOB, fmt=PCT, align=AC)
        put(ws, f'G{r}', f'=IF(NOT({on}),"",ROUND(SUM({AL_YC[i]}${AL_P0}:{AL_YC[i]}${AL_P0 + NPJ - 1}),2))', F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'H{r}', f'=IF(NOT({on}),"",ROUND(B{r}-G{r},2))', F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'I{r}', f'=IF(NOT({on}),"",IF(AND({y}={AX_Y},NOT(ISNUMBER(E{r}))),"今年还没完，按到"&{AX_MON}&"月的费率先摊；年底填定稿费率",'
                         f'IF(NOT(ISNUMBER(E{r})),"去年的账定了就把费率填到 E 列",IF(ABS(N(H{r}))>1,"定稿费率跟算出来的不一样，差额留在公司",""))))',
            F_NOTE, align=AL)
        ws.merge_cells(f'I{r}:L{r}')
    ws.conditional_formatting.add(f'A{AL_Y0}:L{AL_Y0 + NYEARS - 1}', FormulaRule(formula=[f'$A{AL_Y0}>{AX_Y}'], font=GREY))
    dv = DataValidation(type='decimal', operator='between', formula1='0', formula2='1', allow_blank=True, showErrorMessage=True,
                        errorTitle='费率', error='填百分比，比如 12%')
    ws.add_data_validation(dv)
    dv.add(f'E{AL_Y0}:E{AL_Y0 + NYEARS - 1}')
    # ② 项目 × 年
    h = AL_P0 - 1
    section(ws, h - 1, 'A', 'S', '② 每个项目摊到的管理费', 'FFC65911')
    header(ws, h, [('A', '序号'), ('B', '项目'), ('C', '状态'), (AL_OPEN, '建账前\n已摊'), (AL_CUM, '开工至今\n合计'), (AL_DRV, '本年\n分摊依据'),
                   (AL_NOW, '本年摊到'), (AL_SHARE, '本年\n占比'), (AL_OLD, '老办法\n(按合同额)'), (AL_DIFF, '差多少\n(新−老)'), (AL_ACT, '本年\n干过活')],
           'FFC65911', height=40)
    for i, c in enumerate(AL_YC):
        put(ws, f'{c}{h}', f'={AX_Y0}+{i}&"年"', F_HDR, fill('FFC65911'), align=AC)
    tot_pool = f'SUMIFS($B${AL_Y0}:$B${AL_Y0 + NYEARS - 1},$A${AL_Y0}:$A${AL_Y0 + NYEARS - 1},{AX_Y})'
    for i in range(NPJ):
        r, p, k = AL_P0 + i, PJ_R0 + i, PS_R0 + i
        nm = f'$B{r}'
        put(ws, f'A{r}', f'=IF({nm}="","",{i + 1})', F_AUTO, align=AC)
        put(ws, f'B{r}', f'={SH_PROJ}!${PJ_NAME}${p}&""', F_TXT, align=AL)
        put(ws, f'C{r}', f'={SH_PROJ}!${PJ_STAT}${p}&""', F_AUTO, align=AC)
        put(ws, f'{AL_OPEN}{r}', f'=IF({nm}="","",SUMIFS({opr(OPJ_ALLOC)},{opr(OPJ_PJ)},{nm}))', F_AUTO, fmt=MONEY, align=AR)
        for j, c in enumerate(AL_YC):
            put(ws, f'{c}{r}', f'=IF({nm}="","",ROUND({SH_PS}!${PS_DRV[j]}${k}*N($F${AL_Y0 + j}),2))', F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'{AL_CUM}{r}', f'=IF({nm}="","",N({AL_OPEN}{r})+SUM({AL_YC[0]}{r}:{AL_YC[-1]}{r}))', F_AUTOB, fmt=MONEY, align=AR)
        yi = f'({AX_Y}-{AX_Y0}+1)'
        put(ws, f'{AL_DRV}{r}', f'=IF({nm}="","",INDEX({SH_PS}!${PS_DRV[0]}${k}:${PS_DRV[-1]}${k},{yi}))', F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'{AL_NOW}{r}', f'=IF({nm}="","",INDEX({AL_YC[0]}{r}:{AL_YC[-1]}{r},{yi}))', F_AUTOB, fmt=MONEY, align=AR)
        put(ws, f'{AL_SHARE}{r}', f'=IF(OR({nm}="",N(${AL_NOW}${AL_P0 + NPJ})=0),"",{AL_NOW}{r}/${AL_NOW}${AL_P0 + NPJ})', F_AUTO, fmt=PCT, align=AC)
        act = (f'IF({nm}="",0,IF(N({AL_DRV}{r})+{SH_PS}!${PS_COL[("确认收入", "本")]}${k}+{SH_PS}!${PS_COL[("材料", "本")]}${k}'
               f'+{SH_PS}!${PS_COL[("其他直接", "本")]}${k}<>0,1,0))')
        put(ws, f'{AL_ACT}{r}', f'={act}', F_AUTO, fmt=INT, align=AC)
        tot = f'N({SH_PROJ}!${PJ_TOTAL}${p})'
        ws[f'T{r}'] = f'={tot}*{AL_ACT}{r}'
        ws[f'T{r}'].font = F_HELP
        den = f'SUM($T${AL_P0}:$T${AL_P0 + NPJ - 1})'
        put(ws, f'{AL_OLD}{r}', f'=IF(OR({nm}="",{AL_ACT}{r}=0,{den}=0),0,ROUND({tot_pool}*{tot}/{den},2))', F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'{AL_DIFF}{r}', f'=IF({nm}="","",{AL_NOW}{r}-{AL_OLD}{r})', F_AUTO, fmt=MONEY, align=AR)
    R1 = AL_P0 + NPJ - 1
    r = R1 + 1
    put(ws, f'B{r}', '合计', F_TXTB, FILL_TOT, align=AC)
    for c in [AL_OPEN] + AL_YC + [AL_CUM, AL_DRV, AL_NOW, AL_OLD, AL_DIFF]:
        put(ws, f'{c}{r}', f'=SUM({c}{AL_P0}:{c}{R1})', F_AUTOB, FILL_TOT, MONEY, AR)
    for c in ('A', 'C', AL_SHARE, AL_ACT):
        put(ws, f'{c}{r}', None, fill_=FILL_TOT)
    ws.conditional_formatting.add(f'{AL_YC[0]}{h}:{AL_YC[-1]}{R1}', FormulaRule(formula=[f'{AL_YC[0]}${h}>{AX_Y}&"年"'], font=GREY))
    t0 = r + 3
    section(ws, t0, 'A', 'S', '③ 为什么不按工程款、不按收入当月摊', 'FFC65911')
    notes = ['· 按收入当月摊：工程款几年收不到的项目一直不摊，收到钱那个月突然摊一大笔，项目利润忽高忽低，看不出真实盈亏。',
             '· 按合同额摊（Q 列老办法）：大合同但只做了一点点的项目摊多了，小合同但人工多、工期长的摊少了；工期一周的项目跟做一年的摊一样的比例也不合理。',
             '· 按施工费摊：管理人员的精力、公司的车和办公室，主要花在带工人、盯分包、调机械上，施工费越多说明越费管理，最贴近实际；每月都有数，不用等收款。',
             '· 月薪的管理人员在工地待了几天，【考勤工资】里填那个项目几天，那部分工资直接算项目人工（不进管理费）；在公司的天数才进管理费再摊。',
             '· 年底把 E 列「定稿费率」填上（一般就填算出来的数），这一年就定了；以后改了以前的流水，只有差额显示在「没摊完/多摊了」，项目利润不会跟着乱变。',
             '· 想改成按「直接成本」（含材料）摊：到【基础资料】⑤ 把「管理费分摊依据」改成「直接成本」。材料多的项目会多摊，一般建筑安装按施工费更公平。']
    for i, t in enumerate(notes):
        rr = t0 + 1 + i
        put(ws, f'A{rr}', t, F_TXT, align=ALW, border=False)
        ws.merge_cells(f'A{rr}:S{rr}')
        ws.row_dimensions[rr].height = 20
    hide(ws, 'T')
    ws.freeze_panes = 'C8'
    return ws


# ═══════════════════════════ 项目利润表 ═══════════════════════════
def build_ppl(wb, ctx):
    wmap = {P_SEQ: 5, P_NAME: 14, P_STAT: 9, P_TOTAL: 14, C_TIP: 40}
    for c in [P_REV, P_MAT, P_LAB, P_SUB, P_MACH, P_OTH, P_TAX, P_INT, P_GROSS, P_ALLOC, P_NET,
              C_REV, C_MAT, C_LAB, C_SUB, C_MACH, C_OTH, C_TAX, C_INT, C_ALLOC, C_NET, C_REC, C_ARB]:
        wmap[c] = 14
    for c in (P_RATE, C_RATE, C_PROG, C_CPROG):
        wmap[c] = 8
    ws = _sheet(wb, SH_PPL, 'FFC65911', '项 目 利 润 表（每个项目：本年 ＋ 开工至今，赚了还是亏了）',
                '💡 全自动。收入＝甲方确认的产值（不是收到的钱）；成本＝材料、人工、分包、机械、其他直接费（按发生，挂账没付的也算）；税金按确认产值估算；'
                '管理费按【费用分摊】的费率摊。左边「本年」＝首页选的年度到截止日，右边「开工至今」＝含建账前（【期初余额】）。'
                '最右「提示」：成本花得多、产值确认得少（该找甲方报量了）、完工没结算、亏损的，都会提醒。最下面是跟【利润表】对账。',
                C_TIP, wmap)
    put(ws, 'B3', f'="本年＝"&{AX_Y}&"年1月～"&TEXT({AX_E},"m月d日")', F_KPI_L, align=AL, border=False)
    ws.merge_cells('B3:F3')
    ws.merge_cells(f'{P_REV}4:{P_RATE}4')
    put(ws, f'{P_REV}4', '本  年', F_SEC, fill('FF2F75B5'), align=AC)
    ws.merge_cells(f'{C_REV}4:{C_TIP}4')
    put(ws, f'{C_REV}4', '开 工 至 今（含建账前）', F_SEC, fill('FFC65911'), align=AC)
    heads = [(P_SEQ, '序号'), (P_NAME, '项目'), (P_STAT, '状态'), (P_TOTAL, '总价'),
             (P_REV, '确认收入'), (P_MAT, '材料'), (P_LAB, '人工'), (P_SUB, '分包'), (P_MACH, '机械运输'), (P_OTH, '其他直接'),
             (P_TAX, '税金\n(估算)'), (P_INT, '票据贴息'), (P_GROSS, '项目毛利'), (P_ALLOC, '摊管理费'),
             (P_NET, '项目净利'), (P_RATE, '净利率'),
             (C_REV, '确认收入'), (C_MAT, '材料'), (C_LAB, '人工'), (C_SUB, '分包'), (C_MACH, '机械运输'), (C_OTH, '其他直接'),
             (C_TAX, '税金\n(估算)'), (C_INT, '票据贴息'), (C_ALLOC, '摊管理费'), (C_NET, '项目净利'), (C_RATE, '净利率'),
             (C_PROG, '产值\n确认进度'), (C_CPROG, '成本\n进度'), (C_REC, '已收款\n(含抵账)'), (C_ARB, '应收余额'), (C_TIP, '提示')]
    header(ws, PPL_H, heads, 'FFC65911', height=40)
    for c, _ in heads[4:17]:
        ws[f'{c}{PPL_H}'].fill = fill('FF2F75B5')
    for i in range(NPJ):
        r, p, k = PPL_R0 + i, PJ_R0 + i, PS_R0 + i
        nm = f'$B{r}'
        pc = lambda m, per: f'{SH_PS}!${PS_COL[(m, per)]}${k}'
        o = lambda col: f'SUMIFS({opr(col)},{opr(OPJ_PJ)},{nm})'
        g = lambda f: f'=IF({nm}="","",{f})'
        ws[f'{P_SEQ}{r}'] = f'=IF({nm}="","",{i + 1})'
        ws[f'{P_NAME}{r}'] = f'={SH_PROJ}!${PJ_NAME}${p}&""'
        ws[f'{P_STAT}{r}'] = f'={SH_PROJ}!${PJ_STAT}${p}&""'
        ws[f'{P_TOTAL}{r}'] = g(f'N({SH_PROJ}!${PJ_TOTAL}${p})')
        ws[f'{P_REV}{r}'] = g(pc('确认收入', '本'))
        ws[f'{P_MAT}{r}'] = g(pc('材料', '本'))
        ws[f'{P_LAB}{r}'] = g(pc('人工', '本'))
        ws[f'{P_SUB}{r}'] = g(pc('分包', '本'))
        ws[f'{P_MACH}{r}'] = g(pc('机械', '本'))
        ws[f'{P_OTH}{r}'] = g(f'{pc("其他直接", "本")}+{pc("甲方扣款", "本")}')
        tax = lambda per, rev: f'ROUND(({pc("销项税", per)}-{pc("进项税", per)})*(1+{PA_SURR})+{rev}*{PA_STAMP},2)'
        ws[f'{P_TAX}{r}'] = g(tax('本', f'{P_REV}{r}'))
        ws[f'{P_INT}{r}'] = g(pc('贴息', '本'))
        ws[f'{P_GROSS}{r}'] = g(f'{P_REV}{r}-SUM({P_MAT}{r}:{P_INT}{r})')
        ws[f'{P_ALLOC}{r}'] = g(f'{SH_ALLOC}!${AL_NOW}${AL_P0 + i}')
        ws[f'{P_NET}{r}'] = g(f'{P_GROSS}{r}-{P_ALLOC}{r}')
        ws[f'{P_RATE}{r}'] = f'=IF(OR({nm}="",N({P_REV}{r})=0),"",{P_NET}{r}/{P_REV}{r})'
        both = lambda m: f'{pc(m, "前")}+{pc(m, "本")}'
        ws[f'{C_REV}{r}'] = g(f'{o(OPJ_REV)}+{both("确认收入")}')
        ws[f'{C_MAT}{r}'] = g(f'{o(OPJ_MAT)}+{both("材料")}')
        ws[f'{C_LAB}{r}'] = g(f'{o(OPJ_LAB)}+{both("人工")}')
        ws[f'{C_SUB}{r}'] = g(f'{o(OPJ_SUB)}+{both("分包")}')
        ws[f'{C_MACH}{r}'] = g(f'{o(OPJ_MACH)}+{both("机械")}')
        ws[f'{C_OTH}{r}'] = g(f'{o(OPJ_OTH)}+{both("其他直接")}+{both("甲方扣款")}')
        ws[f'{C_TAX}{r}'] = g(f'{o(OPJ_TAX)}+{tax("前", pc("确认收入", "前"))}+{P_TAX}{r}')
        ws[f'{C_INT}{r}'] = g(both('贴息'))
        ws[f'{C_ALLOC}{r}'] = g(f'{SH_ALLOC}!${AL_CUM}${AL_P0 + i}')
        ws[f'{C_NET}{r}'] = g(f'{C_REV}{r}-SUM({C_MAT}{r}:{C_ALLOC}{r})')
        ws[f'{C_RATE}{r}'] = f'=IF(OR({nm}="",N({C_REV}{r})=0),"",{C_NET}{r}/{C_REV}{r})'
        ws[f'{C_PROG}{r}'] = f'=IF(OR({nm}="",N({P_TOTAL}{r})=0),"",{C_REV}{r}/{P_TOTAL}{r})'
        ws[f'{C_CPROG}{r}'] = f'=IF(OR({nm}="",N({P_TOTAL}{r})=0),"",SUM({C_MAT}{r}:{C_OTH}{r})/{P_TOTAL}{r})'
        ws[f'{C_REC}{r}'] = g(f'{SH_AR}!$G${AR_R0 + i}')
        ws[f'{C_ARB}{r}'] = g(f'{SH_AR}!$I${AR_R0 + i}')
        st = f'{P_STAT}{r}'
        ws[f'{C_TIP}{r}'] = (f'=IF({nm}="","",IF(AND(ISNUMBER({C_CPROG}{r}),N({C_CPROG}{r})-N({C_PROG}{r})>{PA_LAG}),'
                             f'"成本花了"&TEXT({C_CPROG}{r},"0%")&"，产值才确认"&TEXT(N({C_PROG}{r}),"0%")&"：找甲方报量、补【收入确认】",'
                             f'IF(AND(N({C_REV}{r})=0,SUM({C_MAT}{r}:{C_OTH}{r})>0),"有成本没收入：补【收入确认】（产值）",'
                             f'IF(AND(N({C_REV}{r})=0,N({C_REC}{r})>0),"收了钱还没确认产值：补【收入确认】",'
                             f'IF(AND(N({C_REV}{r})>0,SUM({C_MAT}{r}:{C_OTH}{r})=0),"有产值没成本：成本是不是没录？（建账前的在【期初余额】①）",'
                             f'IF({st}="完工未结算","完工了还没结算：抓紧结算",'
                             f'IF(AND(OR({st}="已结算",{st}="质保期",{st}="已完结"),N({C_REV}{r})>0,SUM({C_MAT}{r}:{C_OTH}{r})<N({C_REV}{r})*0.3),'
                             f'"已结算但成本不到产值三成：成本是不是没录全？",'
                             f'IF(N({C_NET}{r})<0,"开工至今亏损",""))))))))')
    R1 = PPL_R0 + NPJ - 1
    cols = [CL(i) for i in range(1, 33)]
    money = [c for c in cols if c not in (P_SEQ, P_NAME, P_STAT, P_RATE, C_RATE, C_PROG, C_CPROG, C_TIP)]
    style_rows(ws, PPL_R0, R1, cols, auto=cols, fmts={**{c: MONEY for c in money}, **{c: PCT for c in (P_RATE, C_RATE, C_PROG, C_CPROG)}},
               aligns={P_NAME: AL, C_TIP: AL, **{c: AR for c in money}}, bold=[P_NET, C_NET])
    for r in range(PPL_R0, R1 + 1):
        for c in cols:
            ws[f'{c}{r}'].fill = FILL_NONE
    for c in (P_NET, C_NET):
        ws.conditional_formatting.add(f'{c}{PPL_R0}:{c}{R1}', FormulaRule(formula=[f'N(${c}{PPL_R0})<0'], fill=FILL_WARN))
    ws.conditional_formatting.add(f'{C_TIP}{PPL_R0}:{C_TIP}{R1}', FormulaRule(formula=[f'${C_TIP}{PPL_R0}<>""'], font=F_RED))
    r = PPL_TOT
    put(ws, f'{P_NAME}{r}', '合计', F_TXTB, FILL_TOT, align=AC)
    for c in money:
        put(ws, f'{c}{r}', f'=SUM({c}{PPL_R0}:{c}{R1})', F_AUTOB, FILL_TOT, MONEY, AR)
    put(ws, f'{P_RATE}{r}', f'=IF(N({P_REV}{r})=0,"",{P_NET}{r}/{P_REV}{r})', F_AUTOB, FILL_TOT, PCT, AC)
    put(ws, f'{C_RATE}{r}', f'=IF(N({C_REV}{r})=0,"",{C_NET}{r}/{C_REV}{r})', F_AUTOB, FILL_TOT, PCT, AC)
    for c in (P_SEQ, P_STAT, C_PROG, C_CPROG, C_TIP):
        put(ws, f'{c}{r}', None, fill_=FILL_TOT)
    # ── 跟利润表对账（本年） ──
    b0 = R1 + 3
    section(ws, b0, 'A', 'M', '本年：各项目净利 → 公司利润表净利润（对得上才说明没漏没重）', C_RPT)
    Y = lambda n: ms(n, MS_YTD)
    dir_is = '+'.join(Y(n) for n in DIRECT + ['税金估算', '票据贴息'])
    dir_pj = '+'.join(f'{c}{PPL_TOT}' for c in (P_MAT, P_LAB, P_SUB, P_MACH, P_OTH, P_TAX, P_INT))
    fin = f'{Y("利息支出")}+{Y("手续费")}+{Y("利息收入")}+{Y("其他支出")}+{Y("其他税费")}-{Y("其他收入")}'
    net_is = (f'{Y("营业收入")}+{Y("其他收入")}-(' + '+'.join(Y(n) for n in cats.IS_NAMES if n not in ('营业收入', '其他收入')) + ')')
    lines = [('各项目净利合计', f'={P_NET}{PPL_TOT}', ''),
             ('加：没归到项目的收入', f'={Y("营业收入")}-{P_REV}{PPL_TOT}', '收入确认里项目填错、没填的'),
             ('减：没归到项目的直接成本、税金、贴息', f'=-(({dir_is})-({dir_pj}))', '流水、应付登记里项目没选的（去看 ✗）'),
             ('减：没摊完的管理费（多摊了是加）', f'=-({"+".join(Y(n) for n in POOL)}-{P_ALLOC}{PPL_TOT})', '定稿费率跟算出来的不一样时的差额'),
             ('减：公司不分摊的（利息、手续费、其他收支）', f'=-({fin})', '贷款利息、手续费、其他收支留在公司，不摊到项目'),
             ('减：所得税', f'=-{Y("所得税")}', ''),
             ('＝ 算出来的本年净利润', f'=SUM(C{b0 + 1}:C{b0 + 6})', ''),
             ('【利润表】本年净利润', f'={net_is}', ''),
             ('差额（应该是 0）', f'=ROUND(C{b0 + 7}-C{b0 + 8},2)', '')]
    for i, (lab, f, note) in enumerate(lines):
        rr = b0 + 1 + i
        put(ws, f'B{rr}', lab, F_TXTB if lab.startswith(('＝', '【', '差额')) else F_TXT, align=AL)
        put(ws, f'C{rr}', f, F_AUTOB, FILL_TOT if lab.startswith(('＝', '差额')) else FILL_AUTO, MONEY, AR)
        put(ws, f'D{rr}', note, F_NOTE, align=AL, border=False)
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    ws.conditional_formatting.add(f'C{b0 + 9}', FormulaRule(formula=[f'ABS(N(C{b0 + 9}))>0.01'], fill=FILL_WARN))
    ws.freeze_panes = f'C{PPL_R0}'
    ws.auto_filter.ref = f'A{PPL_H}:{C_TIP}{R1}'
    return ws


PPL_BRIDGE_DIFF = None   # 数据校验用：对账差额所在格（build_ppl 之后算出来）


def ppl_bridge_diff():
    return f"{SH_PPL}!$C${PPL_R0 + NPJ - 1 + 3 + 9}"


# ═══════════════════════════ 项目账（选一个项目） ═══════════════════════════
PL_UN, PL_LAB, PL_IV, PL_J = 60, 80, 80, 300


def build_pl(wb, ctx):
    ws = _sheet(wb, SH_PL, 'FFC65911', '项 目 账（选一个项目：收入、成本、利润、应付、人工、发票、流水全在这一张）',
                '💡 黄格子选项目，全部自动，不用每个项目单独做一张表、也不会漏改：流水、应付、考勤、发票一录，这里就跟着变。'
                '「开工至今」含建账前（【期初余额】）；管理费按【费用分摊】的费率摊；税金按确认产值估算。下面几块清单：这个项目的材料商/分包/机械、干过活的工人、发票、每一笔收支。',
                'AD', {'A': 5, 'B': 16, 'C': 14, 'D': 12, 'E': 14, 'F': 12, 'G': 24, 'H': 2, 'I': 30, 'J': 14, 'K': 2, 'L': 10, 'M': 13, 'N': 11, 'O': 12,
                       'P': 2, 'Q': 11, 'R': 6, 'S': 20, 'T': 13, 'U': 11, 'V': 2, 'W': 11, 'X': 11, 'Y': 30, 'Z': 12, 'AA': 12, 'AB': 11, 'AC': 14, 'AD': 12})
    first = ctx.get('pl_default') or next((p['name'] for p in ctx['projects']), '')
    _lbl(ws, 'B3', '选项目')
    put(ws, 'C3', first, F_SEL, FILL_SEL, align=AC)
    ws.merge_cells('C3:D3')
    dv_list(ws, 'C3', f'={PJ_NAMES}', '选项目')
    PJ = '$C$3'
    put(ws, 'E3', f'="截止 "&TEXT({AX_E},"yyyy年m月d日")&"　本年＝"&{AX_Y}&"年"', F_KPI_L, align=AL, border=False)
    ws.merge_cells('E3:G3')
    IX = '$AF$3'
    ws['AF3'] = f'=IFERROR(MATCH({PJ},{PJ_NAMES},0),0)'
    ws['AF3'].font = F_HELP
    put(ws, 'I3', f'=IF({IX}=0,"⚠ 项目不在【项目档案】里","")', F_RED, border=False)
    P = lambda col: f'IF({IX}=0,"",INDEX({rng(SH_PROJ, col, PJ_R0, PJ_R1)},{IX}))'
    L = lambda col: f'IF({IX}=0,"",INDEX({SH_PPL}!${col}${PPL_R0}:${col}${PPL_R0 + NPJ - 1},{IX}))'
    # ① 基本情况
    section(ws, 5, 'A', 'G', '① 基本情况', C_BASE)
    put(ws, 'B6', '项目全称', F_KPI_L, fill('FFD9E1F2'), align=AC)
    put(ws, 'C6', f'={P(PJ_FULL)}&""', F_AUTOB, FILL_AUTO, align=AL)
    ws.merge_cells('C6:G6')
    info = [[('甲方/总包', PJ_CUS, None), ('状态', PJ_STAT, None), ('负责人', PJ_MGR, None)],
            [('合同额', PJ_AMT, MONEY), ('变更签证', PJ_CHG, MONEY), ('结算额', PJ_SET, MONEY)],
            [('总价', PJ_TOTAL, MONEY), ('税率', PJ_TAX, '0%'), ('质保金比例', PJ_RET, '0%')],
            [('开工日期', PJ_START, DATE), ('完工日期', PJ_END, DATE), ('质保到期', PJ_RETD, DATE)]]
    for i, row in enumerate(info):
        r = 7 + i
        for j, (lab, col, fmt) in enumerate(row):
            lc, vc = ['B', 'D', 'F'][j], ['C', 'E', 'G'][j]
            _lbl(ws, f'{lc}{r}', lab)
            v = P(col)
            put(ws, f'{vc}{r}', f'=IF({IX}=0,"",IF({v}="","",{v}))' if fmt else f'={v}&""', F_AUTOB, FILL_AUTO, fmt, AR if fmt == MONEY else AC)
    # ② 项目利润
    section(ws, 12, 'A', 'G', '② 项目利润', 'FFC65911')
    header(ws, 13, [('B', '项目'), ('C', '开工至今'), ('D', '占收入'), ('E', '本年'), ('F', '占收入'), ('G', '说明')], 'FFC65911')
    rows = [('确认产值（含税）', C_REV, P_REV, '甲方确认的产值（【收入确认】＋期初）'),
            ('材料', C_MAT, P_MAT, '应付登记＋现买现付＋甲供材'), ('人工', C_LAB, P_LAB, '考勤拆到本项目的工资＋临时工'),
            ('分包', C_SUB, P_SUB, ''), ('机械运输', C_MACH, P_MACH, '含本项目用的设备折旧'), ('其他直接费', C_OTH, P_OTH, '检测、期间费、甲方扣款等'),
            ('直接成本小计', None, None, ''), ('税金（估算）', C_TAX, P_TAX, '增值税（销项−进项专票）＋附加＋印花'),
            ('票据贴息', C_INT, P_INT, ''), ('项目毛利', None, None, '产值 − 直接成本 − 税金 − 贴息'),
            ('摊管理费', C_ALLOC, P_ALLOC, '【费用分摊】按施工费摊的'),
            ('项目净利', C_NET, P_NET, ''), ('净利率', None, None, '')]
    r0 = 14
    for i, (lab, cc, pc_, note) in enumerate(rows):
        r = r0 + i
        bold = lab in ('直接成本小计', '项目毛利', '项目净利')
        put(ws, f'B{r}', lab, F_TXTB if bold else F_TXT, FILL_SUB if bold else None, align=AL)
        if lab == '直接成本小计':
            fc, fp = f'=SUM(C{r0 + 1}:C{r - 1})', f'=SUM(E{r0 + 1}:E{r - 1})'
        elif lab == '项目毛利':
            fc, fp = f'=C{r0}-C{r0 + 6}-C{r0 + 7}-C{r0 + 8}', f'=E{r0}-E{r0 + 6}-E{r0 + 7}-E{r0 + 8}'
        elif lab == '净利率':
            fc, fp = f'=IF(N(C{r0})=0,"",C{r - 1}/C{r0})', f'=IF(N(E{r0})=0,"",E{r - 1}/E{r0})'
        else:
            fc = f'=N({L(cc)})' if cc else '=0'
            fp = f'=N({L(pc_)})'
        fmt = PCT if lab == '净利率' else MONEY
        put(ws, f'C{r}', fc, F_AUTOB if bold else F_AUTO, FILL_SUB if bold else None, fmt, AR)
        put(ws, f'E{r}', fp, F_AUTOB if bold else F_AUTO, FILL_SUB if bold else None, fmt, AR)
        if lab != '净利率':
            put(ws, f'D{r}', f'=IF(N(C${r0})=0,"",C{r}/C${r0})', F_NOTE, FILL_SUB if bold else None, PCT, AC)
            put(ws, f'F{r}', f'=IF(N(E${r0})=0,"",E{r}/E${r0})', F_NOTE, FILL_SUB if bold else None, PCT, AC)
        else:
            put(ws, f'D{r}', None)
            put(ws, f'F{r}', None)
        put(ws, f'G{r}', note, F_NOTE, align=AL)
    rn = r0 + len(rows) - 2
    ws.conditional_formatting.add(f'C{rn}:E{rn}', FormulaRule(formula=[f'N(C{rn})<0'], fill=FILL_WARN))
    put(ws, f'B{r0 + len(rows)}', f'={L(C_TIP)}', F_RED, align=AL, border=False)
    ws.merge_cells(f'B{r0 + len(rows)}:G{r0 + len(rows)}')
    # ③ 收款开票（右上）
    section(ws, 5, 'I', 'J', '③ 收款和开票（到截止日）', C_AR)
    AR_ = lambda col: f'IF({IX}=0,0,INDEX({SH_AR}!${col}${AR_R0}:${col}${AR_R0 + NPJ - 1},{IX}))'
    k = f'{IX}'
    items = [('累计确认产值', AR_('F'), MONEY), ('累计收款（含抵账）', AR_('G'), MONEY),
             ('其中总包代发、抵账', f'IF({IX}=0,0,SUMIFS({ofr(OF_AMT)},{ofr(OF_PJ)},{PJ},{ofr(OF_YM)},">="&{AX_OYM},{ofr(OF_YM)},"<="&{AX_YM1}))', MONEY),
             ('应收余额', AR_('I'), MONEY), ('回款比例', AR_('H'), PCT), ('已开票', AR_('J'), MONEY), ('未开票', AR_('K'), MONEY),
             ('开票未回款', AR_('L'), MONEY), ('质保金', AR_('M'), MONEY), ('最后回款日', AR_('P'), DATE)]
    for i, (lab, f, fmt) in enumerate(items):
        r = 6 + i
        _lbl(ws, f'I{r}', lab)
        put(ws, f'J{r}', f'=IF({IX}=0,"",IF({f}="","",{f}))', F_AUTOB, FILL_AUTO, fmt, AR if fmt == MONEY else AC)
    # ④ 税负测算（右中）：建账前的税金取【期初余额】，建账后的按确认产值和收到的专票估
    section(ws, 17, 'I', 'J', '④ 税负测算（开工至今，估算）', C_INV)
    PSc = lambda m: f'IF({IX}=0,0,INDEX({s_agg.ps(m, "前")},{IX})+INDEX({s_agg.ps(m, "本")},{IX}))'
    rate = f'IF(ISNUMBER({P(PJ_TAX)}),{P(PJ_TAX)},{PA_VATR})'
    rec = AR_('G')
    rev_after = PSc('确认收入')
    tx = [('建账前的税金（期初填的）', f'SUMIFS({opr(OPJ_TAX)},{opr(OPJ_PJ)},{PJ})', MONEY),
          ('建账后确认产值（不含税）', f'{rev_after}-({PSc("销项税")})', MONEY),
          ('销项税', PSc('销项税'), MONEY), ('可抵扣进项税（专票）', PSc('进项税'), MONEY),
          ('增值税（销项−进项）', 'J20-J21', MONEY), ('附加税', f'ROUND(J22*{PA_SURR},2)', MONEY),
          ('印花税', f'ROUND({rev_after}*{PA_STAMP},2)', MONEY),
          ('税金合计（＝上面②的税金）', 'J18+J22+J23+J24', MONEY),
          ('税负率（税金÷产值）', f'IF(N(C{r0})=0,"",J25/C{r0})', PCT),
          ('已交增值税及附加（流水里选了这个项目的）',
           f'-SUMIFS({jr(J_NET)},{jr(J_PJ)},{PJ},{jr(J_LINE)},"应交税费",{jr(J_YM)},">="&{AX_OYM},{jr(J_YM)},"<="&{AX_YM1})', MONEY),
          ('异地预缴增值税（按累计收款估）', f'ROUND({rec}/(1+{rate})*{PA_PRE},2)', MONEY)]
    for i, (lab, f, fmt) in enumerate(tx):
        r = 18 + i
        _lbl(ws, f'I{r}', lab)
        put(ws, f'J{r}', f'=IF({IX}=0,"",{f})', F_AUTOB, FILL_AUTO, fmt, AR if fmt == MONEY else AC)
    # ⑤ 管理费分摊（按年）、⑥ 本年各月人工
    section(ws, 5, 'L', 'O', '⑤ 摊到的管理费（按年）', 'FFC65911')
    header(ws, 6, [('L', '年份'), ('M', '分摊依据'), ('N', '费率'), ('O', '摊到')], 'FFC65911')
    put(ws, 'L7', '建账前', F_TXT, align=AC)
    put(ws, 'O7', f'=IF({IX}=0,"",INDEX({SH_ALLOC}!${AL_OPEN}${AL_P0}:${AL_OPEN}${AL_P0 + NPJ - 1},{IX}))', F_AUTO, fmt=MONEY, align=AR)
    put(ws, 'M7', None)
    put(ws, 'N7', None)
    for i in range(NYEARS):
        r = 8 + i
        put(ws, f'L{r}', f'={AX_Y0}+{i}', F_TXT, align=AC)
        put(ws, f'M{r}', f'=IF({IX}=0,"",INDEX({SH_PS}!${PS_DRV[i]}${PS_R0}:${PS_DRV[i]}${PS_R0 + NPJ - 1},{IX}))', F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'N{r}', f'={SH_ALLOC}!$F${AL_Y0 + i}', F_AUTO, fmt=PCT, align=AC)
        put(ws, f'O{r}', f'=IF({IX}=0,"",INDEX({SH_ALLOC}!${AL_YC[i]}${AL_P0}:${AL_YC[i]}${AL_P0 + NPJ - 1},{IX}))', F_AUTO, fmt=MONEY, align=AR)
    ws.conditional_formatting.add(f'L8:O{7 + NYEARS}', FormulaRule(formula=[f'$L8>{AX_Y}'], font=GREY))
    r = 8 + NYEARS
    put(ws, f'L{r}', '合计', F_TXTB, FILL_TOT, align=AC)
    put(ws, f'O{r}', f'=SUM(O7:O{r - 1})', F_AUTOB, FILL_TOT, MONEY, AR)
    put(ws, f'M{r}', f'=SUM(M8:M{r - 1})', F_AUTOB, FILL_TOT, MONEY, AR)
    put(ws, f'N{r}', None, fill_=FILL_TOT)
    section(ws, 5, 'Q', 'T', '⑥ 本年各月人工', C_HOME)
    header(ws, 6, [('Q', '月份'), ('R', ''), ('S', '人工'), ('T', '累计')], C_HOME)
    ws.merge_cells('Q6:R6')
    for m in range(12):
        r = 7 + m
        put(ws, f'Q{r}', f'{m + 1}月', F_TXT, align=AC)
        ws.merge_cells(f'Q{r}:R{r}')
        put(ws, f'S{r}', f'=IF({IX}=0,"",INDEX({SH_PS}!${PS_LABM[m]}${PS_R0}:${PS_LABM[m]}${PS_R0 + NPJ - 1},{IX}))', F_AUTO, fmt=MONEY, align=AR)
        put(ws, f'T{r}', f'=SUM(S$7:S{r})', F_AUTO, fmt=MONEY, align=AR)
    ws.conditional_formatting.add('Q7:T18', FormulaRule(formula=[f'{AX_Y}*100+ROW()-6>{AX_YM1}'], font=GREY))
    # ── 下面四块清单 ──
    top = 34
    nu = UN_R1 - UN_R0 + 1
    HX = ['BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG']        # 隐藏：单位应付、计数；人的天数、金额、计数；发票计数；流水计数
    for i in range(nu):
        rr = 1 + i
        un = f'{SH_UNIT}!${UN_NAME}${UN_R0 + i}'
        ut = f'{SH_UNIT}!${UN_TYPE}${UN_R0 + i}'
        ws[f'BA{rr}'] = (f'=IF(OR({un}="",{IX}=0),0,SUMIFS({oapr(OAP_AMT)},{oapr(OAP_UNIT)},{un},{oapr(OAP_PJ)},{PJ})'
                         f'+SUMIFS({apr(AP_AMT)},{apr(AP_UNIT)},{un},{apr(AP_PJ)},{PJ},{apr(AP_YM)},">="&{AX_OYM},{apr(AP_YM)},"<="&{AX_YM1}))')
        isp = f'OR({ut}="管理人员",{ut}="工人")'
        ra = f'{atr(AT_YM)},">="&{AX_OYM},{atr(AT_YM)},"<="&{AX_YM1}'
        ws[f'BC{rr}'] = (f'=IF(OR({un}="",{IX}=0,NOT({isp})),0,' + '+'.join(
            f'SUMIFS({atr(d)},{atr(AT_NAME)},{un},{atr(p)},{PJ},{ra})' for p, d in zip(AT_PJS, AT_DDS)) + ')')
        ws[f'BD{rr}'] = (f'=IF(BC{rr}=0,0,' + '+'.join(
            f'SUMIFS({atr(a)},{atr(AT_NAME)},{un},{atr(p)},{PJ},{ra})' for p, a in zip(AT_PJS, AT_AMTS)) + ')')
        for c in ('BA', 'BC', 'BD'):
            ws[f'{c}{rr}'].font = F_HELP
    counter(ws, 'BB', 1, nu, lambda i: f'BA{1 + i}<>0')
    counter(ws, 'BE', 1, nu, lambda i: f'BC{1 + i}<>0')
    niv = IV_R1 - IV_R0 + 1
    counter(ws, 'BF', 1, niv, lambda i: (f'({SH_INV}!${IV_PJ}${IV_R0 + i}={PJ})*ISNUMBER({SH_INV}!${IV_YM}${IV_R0 + i})'
                                          f'*({SH_INV}!${IV_YM}${IV_R0 + i}>={AX_OYM})*({SH_INV}!${IV_YM}${IV_R0 + i}<={AX_YM1})'))
    nj = J_R1 - J_R0 + 1
    NJ = skey(ws, 'BG', 2, nj, lambda i: (f'AND({jc(J_PJ, J_R0 + i)}={PJ},ISNUMBER({jc(J_DATE, J_R0 + i)}),ISNUMBER({jc(J_YM, J_R0 + i)}),'
                                           f'{jc(J_YM, J_R0 + i)}>={AX_OYM},{jc(J_YM, J_R0 + i)}<={AX_YM1})'),
              lambda i: f'INT({jc(J_DATE, J_R0 + i)})')
    # ⑦ 应付汇总
    section(ws, top, 'A', 'G', '⑦ 这个项目的材料商、分包、机械（应付汇总）', C_AR)
    header(ws, top + 1, [('A', '序号'), ('B', '单位'), ('C', '本项目应付'), ('D', '本项目已付'), ('E', '本项目未付'), ('F', '本项目收票'),
                         ('G', '欠票\n（应付−收票）')], C_AR, height=40)
    # ⑧ 人工明细
    section(ws, top, 'I', 'O', '⑧ 在这个项目干过活的人（开工至今）', C_HOME)
    header(ws, top + 1, [('I', '姓名'), ('J', '工资'), ('L', '天数'), ('M', '平均日工资'), ('N', '类型'), ('O', '现在欠他')], C_HOME, height=40)
    # ⑨ 发票
    section(ws, top, 'Q', 'U', '⑨ 这个项目的发票', C_INV)
    header(ws, top + 1, [('Q', '日期'), ('R', '方向'), ('S', '对方单位'), ('T', '价税合计'), ('U', '税额')], C_INV, height=40)
    # ⑩ 流水
    section(ws, top, 'W', 'AD', '⑩ 这个项目的每一笔收支（资金流水里认成这个项目的）', C_CASH)
    header(ws, top + 1, [('W', '日期'), ('X', '账户'), ('Y', '摘要'), ('Z', '收入'), ('AA', '支出'), ('AB', '收支类别'), ('AC', '单位/人'), ('AD', '流水序号')],
           C_CASH, height=40)
    d0 = top + 2
    for kk in range(PL_J):
        r = d0 + kk
        # 流水
        ix = f'$BJ{r}'
        ws[ix] = f'={ksorted(kk + 1, "BG", 2, nj, NJ)}'
        g = lambda col: f'INDEX({jr(col)},{ix})'
        ws[f'W{r}'] = f'=IF({ix}=0,"",{g(J_DATE)})'
        ws[f'X{r}'] = f'=IF({ix}=0,"",{g(J_ACC)}&"")'
        ws[f'Y{r}'] = f'=IF({ix}=0,"",{g(J_MEMO)}&"")'
        ws[f'Z{r}'] = f'=IF({ix}=0,"",MAX({g(J_NET)},0))'
        ws[f'AA{r}'] = f'=IF({ix}=0,"",MAX(-{g(J_NET)},0))'
        ws[f'AB{r}'] = f'=IF({ix}=0,"",{g(J_CT)}&"")'
        ws[f'AC{r}'] = f'=IF({ix}=0,"",{g(J_UN)}&"")'
        ws[f'AD{r}'] = f'=IF({ix}=0,"",{ix})'
        ws[ix].font = F_HELP
        if kk < PL_IV:
            ix = f'$BI{r}'
            ws[ix] = f'={kth(kk + 1, "BF", 1, niv)}'
            ws[ix].font = F_HELP
            v = lambda col: f'INDEX({ivr(col)},{ix})'
            ws[f'Q{r}'] = f'=IF({ix}=0,"",{v(IV_DATE)})'
            ws[f'R{r}'] = f'=IF({ix}=0,"",{v(IV_DIR)}&"")'
            ws[f'S{r}'] = f'=IF({ix}=0,"",{v(IV_UNIT)}&"")'
            ws[f'T{r}'] = f'=IF({ix}=0,"",{v(IV_TOTU)})'
            ws[f'U{r}'] = f'=IF({ix}=0,"",{v(IV_TAXU)})'
        if kk < PL_LAB:
            ix = f'$BH{r}'
            ws[ix] = f'={kth(kk + 1, "BE", 1, nu)}'
            ws[ix].font = F_HELP
            ws[f'I{r}'] = f'=IF({ix}=0,"",INDEX({UN_NAMES},{ix}))'
            ws[f'J{r}'] = f'=IF({ix}=0,"",INDEX($BD$1:$BD${nu},{ix}))'
            ws[f'L{r}'] = f'=IF({ix}=0,"",INDEX($BC$1:$BC${nu},{ix}))'
            ws[f'M{r}'] = f'=IF(OR({ix}=0,N(L{r})=0),"",J{r}/L{r})'
            ws[f'N{r}'] = f'=IF({ix}=0,"",INDEX({UN_TYPES_R},{ix})&"")'
            ws[f'O{r}'] = f'=IF({ix}=0,"",INDEX({s_agg.bx(BX_WG_E)},{ix}))'
        if kk < PL_UN:
            ix = f'$AZ{r}'
            ws[ix] = f'={kth(kk + 1, "BB", 1, nu)}'
            ws[ix].font = F_HELP
            nm = f'$B{r}'
            ws[f'A{r}'] = f'=IF({ix}=0,"",{kk + 1})'
            ws[f'B{r}'] = f'=IF({ix}=0,"",INDEX({UN_NAMES},{ix}))'
            ws[f'C{r}'] = f'=IF({ix}=0,"",INDEX($BA$1:$BA${nu},{ix}))'
            paid = (f'SUMIFS({oapr(OAP_PAID)},{oapr(OAP_UNIT)},{nm},{oapr(OAP_PJ)},{PJ})'
                    f'-SUMIFS({jr(J_NET)},{jr(J_UN)},{nm},{jr(J_PJ)},{PJ},{jr(J_CT)},"付应付款",{jr(J_YM)},">="&{AX_OYM},{jr(J_YM)},"<="&{AX_YM1})'
                    f'+SUMIFS({ofr(OF_AMT)},{ofr(OF_WHO)},{nm},{ofr(OF_PJ)},{PJ},{ofr(OF_TYPE)},"总包代付材料分包款",{ofr(OF_YM)},"<="&{AX_YM1})')
            ws[f'D{r}'] = f'=IF({ix}=0,"",{paid})'
            ws[f'E{r}'] = f'=IF({ix}=0,"",C{r}-D{r})'
            ws[f'F{r}'] = (f'=IF({ix}=0,"",SUMIFS({oapr(OAP_INV)},{oapr(OAP_UNIT)},{nm},{oapr(OAP_PJ)},{PJ})'
                           f'+SUMIFS({ivr(IV_TOTU)},{ivr(IV_DIR)},"进项",{ivr(IV_UNIT)},{nm},{ivr(IV_PJ)},{PJ},{ivr(IV_YM)},">="&{AX_OYM},{ivr(IV_YM)},"<="&{AX_YM1}))')
            ws[f'G{r}'] = f'=IF({ix}=0,"",C{r}-F{r})'
    style_rows(ws, d0, d0 + PL_J - 1, ['W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD'], auto=['W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD'],
               fmts={'W': DATE, 'Z': MONEY, 'AA': MONEY}, aligns={'Y': AL, 'Z': AR, 'AA': AR, 'AC': AL})
    style_rows(ws, d0, d0 + PL_IV - 1, list('QRSTU'), auto=list('QRSTU'), fmts={'Q': DATE, 'T': MONEY, 'U': MONEY},
               aligns={'S': AL, 'T': AR, 'U': AR})
    style_rows(ws, d0, d0 + PL_LAB - 1, list('IJKLMNO'), auto=list('IJKLMNO'), fmts={'J': MONEY, 'L': '0.##', 'M': '#,##0', 'O': MONEY},
               aligns={'J': AR, 'O': AR})
    style_rows(ws, d0, d0 + PL_UN - 1, list('ABCDEFG'), auto=list('ABCDEFG'), fmts={c: MONEY for c in 'CDEFG'},
               aligns={'B': AL, **{c: AR for c in 'CDEFG'}})
    for r in range(d0, d0 + PL_J):
        for c in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'Q', 'R', 'S', 'T', 'U', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD']:
            ws[f'{c}{r}'].fill = FILL_NONE
    # 各清单合计（放在标题行右边）
    put(ws, f'H{top}', None, border=False)
    for c, f in (('C', f'=SUM(C{d0}:C{d0 + PL_UN - 1})'), ('D', f'=SUM(D{d0}:D{d0 + PL_UN - 1})'), ('E', f'=SUM(E{d0}:E{d0 + PL_UN - 1})'),
                 ('F', f'=SUM(F{d0}:F{d0 + PL_UN - 1})'), ('G', f'=SUM(G{d0}:G{d0 + PL_UN - 1})'), ('J', f'=SUM(J{d0}:J{d0 + PL_LAB - 1})'), ('Z', f'=SUM(Z{d0}:Z{d0 + PL_J - 1})'),
                 ('AA', f'=SUM(AA{d0}:AA{d0 + PL_J - 1})'), ('T', f'=SUM(T{d0}:T{d0 + PL_IV - 1})')):
        put(ws, f'{c}{top - 1}', f, F_AUTOB, FILL_TOT, MONEY, AR)
    put(ws, f'B{top - 1}', '下面清单合计 →', F_NOTE, align=AR, border=False)
    msgs = (f'IF({cnt("BB", 1, nu)}>{PL_UN},"⚠ 单位超过{PL_UN}家 ","")&IF({cnt("BE", 1, nu)}>{PL_LAB},"⚠ 人超过{PL_LAB}个 ","")'
            f'&IF({cnt("BF", 1, niv)}>{PL_IV},"⚠ 发票超过{PL_IV}张 ","")&IF({NJ}>{PL_J},"⚠ 流水超过{PL_J}笔（只列前面的）","")')
    put(ws, f'W{top - 1}', f'={msgs}', F_RED, border=False)
    hide(ws, 'AF', 'AZ', 'BH', 'BI', 'BJ', *HX)
    ws.freeze_panes = 'A5'
    return ws


def build_all(wb, ctx):
    build_alloc(wb, ctx)
    build_ppl(wb, ctx)
    build_pl(wb, ctx)
