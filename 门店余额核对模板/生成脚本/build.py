# -*- coding: utf-8 -*-
"""门店余额核对 · 生成脚本

在你的《多帐户收支登记表》里加两张表，原来的 27 张表一个字都不动：
  【门店余额核对】     门店余额 ＋ 收款 － 付款 ＋ 公司年初结余 ＝ 应存；应存跟银行余额（实存）比，差额正常是 0
  【门店余额核对辅助】（隐藏）【数据录入】每一行算三个标记：计入日期、是不是 14 个银行现金账户、是不是门店

做法：openpyxl 只用来画新表（存成一个临时工作簿），再把新表的 XML 和它用到的样式直接塞进原文件 ——
原文件不经 openpyxl 存盘，【操作流程】里的 15 张截图、5 处批注、WPS 的动态数组都原样保留。
新表的数先在副本上用 LibreOffice 整本算一遍，再把算出来的数写进新表当缓存（打开就有数，改了数据录入照常自动重算）。

跑法：python3 build.py [原表] [输出]
"""
import copy
import glob
import os
import re
import shutil
import subprocess
import sys
import zipfile

from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, '参考', '原表_数据录入_2026.9.27.xlsx')
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(ROOT, '多帐户收支登记表_加门店余额核对.xlsx')
TMP = os.path.join(HERE, '_tmp')

SH, AUX = '门店余额核对', '门店余额核对辅助'
D0, D1 = 19, 30191                  # 数据录入：第 5～18 行是 14 个账户的期初，第 19 行起是流水（跟各账户子表一样取到 30191 行）
M0, M1 = 5, 1004                    # 门店月度收支统计：第 5 行起每个门店 12 行（留到 1004 行）
NS_DEFAULT = [('公司', '公司费用、利息、税费等'), ('配送', '进货付款、收调拨等（配送中心）'), ('预收', ''),
              ('收款', ''), ('其他收款', ''), ('股东借款', '股东借进借出'),
              ('内部转款', '自己账户之间转来转去、支取现金，收付应该相抵'), ('房租及其他', ''),
              ('12月门店', '去年 12 月门店的款，没分到具体门店')]

NSX = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS = '{%s}' % NSX
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

DR = lambda c: f'数据录入!${c}${D0}:${c}${D1}'
XR = lambda c: f'{AUX}!${c}${D0}:${c}${D1}'
MR = lambda c: f'门店月度收支统计!${c}${M0}:${c}${M1}'
CUT, LAST, BANK0, BOOKY, AUTO0 = (f'{AUX}!$J${i}' for i in range(1, 6))
MY = '门店月度收支统计!$B$2'          # 那张表统计的年份（首月手加数按它算）

# ---- 版面 ----
R_CUT = 4
A1_ROWS = dict(store=9, recv=10, pay=11, co=12, should=13, bank=14, diff=15, stmt=16)
M_H, M_0 = 18, 20                   # ② 按月
X_H, X_0 = 33, 35                   # ③ 差额从哪来
N_H, N_0, N_1 = 46, 48, 62          # ④ 收款 / 付款 明细（不算门店的名字，黄格可改）
N_BL, N_TOT = 63, 64
S_H, S_0, S_1 = 66, 68, 167         # ⑤ 门店余额明细（100 行）
S_RES, S_TOT = 168, 169
A_H, A_0, A_1 = 171, 173, 192       # ⑥ 账户余额（20 行）
A_TOT = 193
NSR = f'${"B"}${N_0}:${"B"}${N_1}'

YH = '微软雅黑'
F_TITLE = Font(name=YH, size=15, bold=True, color='FF1F3864')
F_SEC = Font(name=YH, size=11, bold=True, color='FFFFFFFF')
F_HDR = Font(name=YH, size=10, bold=True, color='FF1F3864')
F_TXT = Font(name=YH, size=10)
F_TXTB = Font(name=YH, size=10, bold=True)
F_NOTE = Font(name=YH, size=9, color='FF595959')
F_BIG = Font(name=YH, size=12, bold=True, color='FF1F3864')
F_IN = Font(name=YH, size=10, bold=True, color='FF0000C0')
FILL_SEC = PatternFill('solid', fgColor='FF2F5597')
FILL_HDR = PatternFill('solid', fgColor='FFD9E1F2')
FILL_IN = PatternFill('solid', fgColor='FFFFF2CC')
FILL_AUTO = PatternFill('solid', fgColor='FFF2F2F2')
FILL_KEY = PatternFill('solid', fgColor='FFE2EFDA')
FILL_TOT = PatternFill('solid', fgColor='FFDDEBF7')
FILL_RED = PatternFill('solid', fgColor='FFFFC7CE')
FILL_ORG = PatternFill('solid', fgColor='FFFCE4D6')
THIN = Side(style='thin', color='FFBFBFBF')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
AL = Alignment(horizontal='left', vertical='center', wrap_text=True)
AC = Alignment(horizontal='center', vertical='center', wrap_text=True)
AR = Alignment(horizontal='right', vertical='center')
MONEY = '#,##0.00;[Red]-#,##0.00'
DATE = 'yyyy-mm-dd'


def put(ws, ref, v=None, font=F_TXT, fill=None, fmt=None, align=AL, border=BOX):
    c = ws[ref]
    c.value = v
    c.font = font
    if fill is not None:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    c.alignment = align
    if border is not None:
        c.border = border
    return c


def merge(ws, rng, v=None, **kw):
    first = rng.split(':')[0]
    put(ws, first, v, **kw)
    a, b = rng.split(':')
    col0, col1 = re.match(r'[A-Z]+', a).group(), re.match(r'[A-Z]+', b).group()
    row = int(re.search(r'\d+', a).group())
    from openpyxl.utils import column_index_from_string as ci, get_column_letter as cl
    for k in range(ci(col0) + 1, ci(col1) + 1):
        put(ws, f'{cl(k)}{row}', None, font=kw.get('font', F_TXT), fill=kw.get('fill'), border=kw.get('border', BOX))
    ws.merge_cells(rng)


def section(ws, row, text, last='J'):
    merge(ws, f'A{row}:{last}{row}', text, font=F_SEC, fill=FILL_SEC, border=None)
    ws.row_dimensions[row].height = 22


def heads(ws, row, items, height=30):
    for col, t in items:
        put(ws, f'{col}{row}', t, font=F_HDR, fill=FILL_HDR, align=AC)
    ws.row_dimensions[row].height = height


def le(r):
    return f'"<="&{r}'


# ============================================================ 画【门店余额核对】
def build_main():
    wb = Workbook()
    ws = wb.active
    ws.title = SH
    ws.sheet_properties.tabColor = 'FF00B050'
    ws.sheet_view.showGridLines = False
    for col, w in dict(A=6, B=40, C=17, D=17, E=17, F=17, G=17, H=19, I=17, J=12, K=4, L=4).items():
        ws.column_dimensions[col].width = w
    merge(ws, 'A1:J1', '门店余额核对　　门店余额 ＋ 收款 － 付款 ＝ 银行余额', font=F_TITLE, border=None)
    ws.row_dimensions[1].height = 30
    merge(ws, 'A2:J2',
          '全部从【数据录入】自动取数，这张表不用手填（黄格子除外）。门店余额＝【门店月度收支统计】里每个门店的年初数（连同那张表首月手加的数）'
          '＋ 这个门店在【数据录入】的收入 － 支出；收款、付款＝不算门店的收支（公司、配送、内部转款、股东借款、没填门店的……见 ④）；'
          '实存＝14 个账户（数据录入第 5～18 行）的余额。调拨、货款直抵费用、平台款分摊这些没走银行账户的记录，只是在门店和公司之间挪，'
          '收付应该正好相抵 —— 不相抵的部分就是差额，③ 里分类列出来。原来的表一格都没动。',
          font=F_NOTE, border=None)
    ws.row_dimensions[2].height = 64
    put(ws, f'A{R_CUT}', '截止日期', font=F_TXTB, fill=FILL_HDR, align=AC)
    put(ws, f'B{R_CUT}', None, font=F_IN, fill=FILL_IN, fmt=DATE, align=AC)
    merge(ws, f'C{R_CUT}:F{R_CUT}', '← 空着＝全部（到最后一笔）；填某一天＝算到那天为止（没写日期的记录按「收支所属月份」的月底算）',
          font=F_NOTE, border=None)
    put(ws, f'G{R_CUT}', '现在算的是', font=F_TXTB, fill=FILL_HDR, align=AC)
    merge(ws, f'H{R_CUT}:J{R_CUT}', f'=IF(ISNUMBER($B${R_CUT}),"截至 "&TEXT($B${R_CUT},"yyyy-mm-dd"),'
                                    f'"全部（最后一笔 "&TEXT({LAST},"yyyy-mm-dd")&"）")', font=F_TXTB, fill=FILL_KEY)
    ws.row_dimensions[R_CUT].height = 24
    put(ws, 'A5', '账本年初', font=F_TXTB, fill=FILL_HDR, align=AC)
    put(ws, 'B5', '=数据录入!$B$5', font=F_TXT, fill=FILL_AUTO, fmt=DATE, align=AC)
    merge(ws, 'C5:J5', f'="这一天 14 个账户的余额合计 "&TEXT({BANK0},"#,##0.00")&" 元（数据录入第 5～18 行的期初）"',
          font=F_NOTE, border=None)
    dv = DataValidation(type='date', operator='greaterThan', formula1='36526', allow_blank=True,
                        showErrorMessage=True, errorTitle='截止日期', error='填日期，比如 2026/9/30；空着＝全部')
    ws.add_data_validation(dv)
    dv.add(f'B{R_CUT}')

    # ---------------- ① 核对 ----------------
    r = A1_ROWS
    section(ws, 7, '① 核对：门店余额 ＋ 收款 － 付款 ＋ 公司年初结余 ＝ 应存，应存 ＝ 银行余额（实存）')
    heads(ws, 8, [('A', '序'), ('B', '项目'), ('C', '金额'), ('D', '说明')], 22)
    merge(ws, 'D8:J8', '说明', font=F_HDR, fill=FILL_HDR, align=AC)
    xc1 = f'SUMIFS({DR("F")},{XR("C")},1,{XR("A")},{le(CUT)})-SUMIFS({DR("G")},{XR("C")},1,{XR("A")},{le(CUT)})'
    rows1 = [
        ('1', '门店余额（各门店：年初 ＋ 收入 － 支出）', f'=$C${S_TOT}+$D${S_TOT}+{xc1}',
         '每个门店多少看 ⑤。门店＝【门店月度收支统计】里的门店（含自助药机、疼痛馆、关了的店），不含 ④ 里那些名字'),
        ('2', '加：收款（不算门店的收入）', f'=SUMIFS({DR("F")},{XR("C")},0,{XR("A")},{le(CUT)})',
         '公司、配送、内部转款、股东借款、没填门店的……分别多少看 ④'),
        ('3', '减：付款（不算门店的支出）', f'=SUMIFS({DR("G")},{XR("C")},0,{XR("A")},{le(CUT)})', '同上'),
        ('4', '加：公司年初结余', f'=IF(ISNUMBER($H${r["co"]}),$H${r["co"]},{AUTO0})',
         '＝年初银行余额 － 年初门店余额合计：年初那天银行里的钱扣掉属于门店的，剩下的算公司的（负数＝公司用门店的钱进了货）。'
         '想用别的数，在右边 H 格手填'),
        ('＝', '应存（1 ＋ 2 － 3 ＋ 4）', f'=ROUND($C${r["store"]}+$C${r["recv"]}-$C${r["pay"]}+$C${r["co"]},2)', ''),
        ('5', '实存（14 个账户的账上余额合计）',
         f'=ROUND({BANK0}+SUMIFS({DR("F")},{XR("B")},1,{XR("A")},{le(CUT)})-SUMIFS({DR("G")},{XR("B")},1,{XR("A")},{le(CUT)}),2)',
         '每个账户多少看 ⑥'),
        ('差', '差额（实存 － 应存，正常是 0）', f'=ROUND($C${r["bank"]}-$C${r["should"]},2)', None),
        ('6', '银行对账单余额合计（⑥ 里填了对账单的账户按对账单，没填的按账上）',
         f'=IF(COUNT($G${A_0}:$G${A_1})=0,"",ROUND(SUM($I${A_0}:$I${A_1}),2))', None),
    ]
    for (no, lab, f, note), rr in zip(rows1, range(9, 17)):
        key = rr in (r['should'], r['bank'], r['diff'])
        put(ws, f'A{rr}', no, font=F_TXTB, fill=FILL_HDR, align=AC)
        put(ws, f'B{rr}', lab, font=F_TXTB if key else F_TXT, fill=FILL_KEY if key else None)
        put(ws, f'C{rr}', f, font=F_BIG if key else F_TXTB, fill=FILL_KEY if key else FILL_AUTO, fmt=MONEY, align=AR)
        if note is not None:
            merge(ws, f'D{rr}:J{rr}', note, font=F_NOTE)
        ws.row_dimensions[rr].height = 30 if rr == r['co'] else 22
    # 公司年初结余：右边可以手填
    ws.unmerge_cells(f'D{r["co"]}:J{r["co"]}')
    merge(ws, f'D{r["co"]}:F{r["co"]}', rows1[3][3], font=F_NOTE)
    put(ws, f'G{r["co"]}', '手填（可空）→', font=F_NOTE, align=AR)
    put(ws, f'H{r["co"]}', None, font=F_IN, fill=FILL_IN, fmt=MONEY, align=AR)
    merge(ws, f'I{r["co"]}:J{r["co"]}', f'="自动算的："&TEXT({AUTO0},"#,##0.00")', font=F_NOTE)
    dvn = DataValidation(type='decimal', operator='between', formula1='-9999999999', formula2='9999999999',
                         allow_blank=True, showErrorMessage=True, errorTitle='只能填数字', error='填金额；空着＝自动算')
    ws.add_data_validation(dvn)
    dvn.add(f'H{r["co"]}')
    merge(ws, f'D{r["diff"]}:J{r["diff"]}',
          f'=IF(ABS($C${r["diff"]})<0.005,"✓ 对上了：门店余额＋收款－付款＋公司年初结余＝银行余额",'
          f'"✗ 差 "&TEXT($C${r["diff"]},"#,##0.00")&" 元：看下面 ③ 是哪些记录造成的")', font=F_TXTB)
    merge(ws, f'D{r["stmt"]}:J{r["stmt"]}',
          f'=IF($C${r["stmt"]}="","⑥ 里还没填银行对账单余额（填了就能跟银行的实际余额对）",'
          f'IF(ABS($C${r["stmt"]}-$C${r["bank"]})<0.005,"✓ 跟账上余额一致",'
          f'"✗ 跟账上余额差 "&TEXT($C${r["stmt"]}-$C${r["bank"]},"#,##0.00")&" 元：看 ⑥ 是哪个账户"))', font=F_TXTB)
    ws.conditional_formatting.add(f'C{r["diff"]}:J{r["diff"]}', FormulaRule(
        formula=[f'ABS($C${r["diff"]})>=0.005'], fill=FILL_RED, font=Font(color='FF9C0006', bold=True)))
    ws.conditional_formatting.add(f'D{r["diff"]}:J{r["diff"]}', FormulaRule(
        formula=[f'ABS($C${r["diff"]})<0.005'], font=Font(color='FF00803C', bold=True)))

    # ---------------- ② 按月 ----------------
    section(ws, M_H, '② 按月看（每个月月底：门店余额 ＋ 收款 － 付款 ＋ 公司年初结余 ＝ 银行余额）')
    heads(ws, M_H + 1, [('A', '月份'), ('B', '月底'), ('C', '门店余额'), ('D', '收款（年初起累计）'),
                        ('E', '付款（年初起累计）'), ('F', '公司年初结余'), ('G', '应存'), ('H', '实存（账上）'),
                        ('I', '差额'), ('J', '')])
    for k in range(1, 13):
        rr = M_0 + k - 1
        show = f'DATE({BOOKY},{k},1)<={LAST}'
        b = f'$B{rr}'
        put(ws, f'A{rr}', f'{k}月', font=F_TXTB, fill=FILL_HDR, align=AC)
        put(ws, f'B{rr}', f'=EOMONTH(DATE({BOOKY},{k},1),0)', fill=FILL_AUTO, fmt=DATE, align=AC)
        fs = {
            'C': f'$C${S_TOT}+$D${S_TOT}+SUMIFS({DR("F")},{XR("C")},1,{XR("A")},{le(b)})'
                 f'-SUMIFS({DR("G")},{XR("C")},1,{XR("A")},{le(b)})',
            'D': f'SUMIFS({DR("F")},{XR("C")},0,{XR("A")},{le(b)})',
            'E': f'SUMIFS({DR("G")},{XR("C")},0,{XR("A")},{le(b)})',
            'F': f'$C${r["co"]}',
            'G': f'C{rr}+D{rr}-E{rr}+F{rr}',
            'H': f'{BANK0}+SUMIFS({DR("F")},{XR("B")},1,{XR("A")},{le(b)})-SUMIFS({DR("G")},{XR("B")},1,{XR("A")},{le(b)})',
            'I': f'H{rr}-G{rr}',
        }
        for col, f in fs.items():
            put(ws, f'{col}{rr}', f'=IF(NOT({show}),"",ROUND({f},2))', fill=FILL_AUTO, fmt=MONEY, align=AR,
                font=F_TXTB if col == 'I' else F_TXT)
        put(ws, f'J{rr}', f'=IF(I{rr}="","",IF(ABS(I{rr})<0.005,"✓","✗"))', font=F_TXTB, align=AC)
        ws.row_dimensions[rr].height = 18
    ws.conditional_formatting.add(f'I{M_0}:J{M_0 + 11}', FormulaRule(
        formula=[f'AND($I{M_0}<>"",ABS(N($I{M_0}))>=0.005)'], fill=FILL_RED, font=Font(color='FF9C0006', bold=True)))
    ws.conditional_formatting.add(f'J{M_0}:J{M_0 + 11}', FormulaRule(
        formula=[f'$J{M_0}="✓"'], font=Font(color='FF00803C', bold=True)))

    # ---------------- ③ 差额从哪来 ----------------
    section(ws, X_H, '③ 差额从哪来：没走银行账户的记录，收入和支出应该正好相抵')
    heads(ws, X_H + 1, [('A', '序'), ('B', '项目'), ('C', '笔数'), ('D', '收入'), ('E', '支出'), ('F', '净额（应为 0）')])
    merge(ws, f'G{X_H + 1}:J{X_H + 1}', '说明', font=F_HDR, fill=FILL_HDR, align=AC)
    c_ = lambda *crit: (f'COUNTIFS({",".join(crit)},{DR("F")},"<>",{XR("A")},{le(CUT)})'
                        f'+COUNTIFS({",".join(crit)},{DR("G")},"<>",{XR("A")},{le(CUT)})')
    s_ = lambda col, *crit: f'SUMIFS({DR(col)},{",".join(crit)},{XR("A")},{le(CUT)})'
    blank, tiao = f'{DR("C")},""', f'{DR("E")},"*调拨*"'
    zd = f'{DR("C")},"货款直抵费用"'
    nb_named = f'{XR("B")},0,{DR("C")},"<>"'
    x = X_0
    xrows = [
        ('a', '没记账户的调拨（账户空着、摘要里有「调拨」）', f'={c_(blank, tiao)}', f'={s_("F", blank, tiao)}',
         f'={s_("G", blank, tiao)}', '配送收调拨、门店付调拨，一收一付应该相等'),
        ('b', '没记账户的其他记录（平台款分摊、报销拆分等）', f'={c_(blank)}-C{x}', f'={s_("F", blank)}-D{x}',
         f'={s_("G", blank)}-E{x}', '比如「上海汇付」聚合款先记一笔总数、再按门店拆：拆出来的合计要等于总数'),
        ('c', '货款直抵费用（门店用营业款直接付的费用）', f'={c_(zd)}', f'={s_("F", zd)}', f'={s_("G", zd)}',
         '门店拿营业款付了费用：收入记营业款、支出记费用，两边应该相等'),
        ('d', '账户名不在 14 个账户里的（写错账户名？）', f'={c_(nb_named)}-C{x + 2}', f'={s_("F", nb_named)}-D{x + 2}',
         f'={s_("G", nb_named)}-E{x + 2}', '账户名要跟【基础资料】A 列、数据录入第 5～18 行的一模一样'),
    ]
    for i, (no, lab, fc, fi, fe, note) in enumerate(xrows):
        rr = x + i
        put(ws, f'A{rr}', no, font=F_TXTB, fill=FILL_HDR, align=AC)
        put(ws, f'B{rr}', lab)
        put(ws, f'C{rr}', fc, fill=FILL_AUTO, fmt='0', align=AR)
        put(ws, f'D{rr}', fi, fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'E{rr}', fe, fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'F{rr}', f'=ROUND(D{rr}-E{rr},2)', font=F_TXTB, fill=FILL_AUTO, fmt=MONEY, align=AR)
        merge(ws, f'G{rr}:J{rr}', note, font=F_NOTE)
        ws.row_dimensions[rr].height = 30
    rs = x + 4
    put(ws, f'A{rs}', '', fill=FILL_TOT)
    put(ws, f'B{rs}', '小计：没走银行账户的记录（净额应为 0）', font=F_TXTB, fill=FILL_TOT)
    for col in 'CDEF':
        put(ws, f'{col}{rs}', f'=SUM({col}{x}:{col}{x + 3})', font=F_TXTB, fill=FILL_TOT,
            fmt='0' if col == 'C' else MONEY, align=AR)
    merge(ws, f'G{rs}:J{rs}', '这一行不是 0，差额就不是 0：去【数据录入】按账户筛选「空白」或「货款直抵费用」找', font=F_NOTE, fill=FILL_TOT)
    rp = x + 5
    put(ws, f'A{rp}', 'e', font=F_TXTB, fill=FILL_HDR, align=AC)
    put(ws, f'B{rp}', '公司年初结余手填了，跟自动算的差')
    for col in 'CDE':
        put(ws, f'{col}{rp}', None, fill=FILL_AUTO)
    put(ws, f'F{rp}', f'=IF(ISNUMBER($H${r["co"]}),ROUND($H${r["co"]}-{AUTO0},2),0)', font=F_TXTB, fill=FILL_AUTO,
        fmt=MONEY, align=AR)
    merge(ws, f'G{rp}:J{rp}', '① 的 H 格没手填就是 0', font=F_NOTE)
    rt = x + 6
    put(ws, f'A{rt}', '', fill=FILL_TOT)
    put(ws, f'B{rt}', '合计（＝ ① 的差额）', font=F_TXTB, fill=FILL_TOT)
    for col in 'CDE':
        put(ws, f'{col}{rt}', None, fill=FILL_TOT)
    put(ws, f'F{rt}', f'=ROUND(-F{rs}-F{rp},2)', font=F_TXTB, fill=FILL_TOT, fmt=MONEY, align=AR)
    merge(ws, f'G{rt}:J{rt}', f'=IF(ABS(F{rt}-$C${r["diff"]})<0.005,"✓ 跟 ① 的差额一致","✗ 跟 ① 的差额不一致（公式被改动了？）")',
          font=F_TXTB, fill=FILL_TOT)
    ws.conditional_formatting.add(f'F{x}:F{rs}', FormulaRule(formula=[f'ABS(N(F{x}))>=0.005'], fill=FILL_RED,
                                                             font=Font(color='FF9C0006', bold=True)))
    # 参考信息
    ri = x + 7
    put(ws, f'A{ri}', '参考', font=F_TXTB, fill=FILL_HDR, align=AC)
    put(ws, f'B{ri}', '年初就对不上的数（从旧表带过来的，不影响上面的差额）')
    put(ws, f'F{ri}', f'=ROUND(SUMIFS({MR("G")},{MR("B")},"配送")+SUMIFS({MR("G")},{MR("B")},"公司")-{AUTO0},2)',
        fill=FILL_AUTO, fmt=MONEY, align=AR)
    for col in 'CDE':
        put(ws, f'{col}{ri}', None, fill=FILL_AUTO)
    merge(ws, f'G{ri}:J{ri}',
          f'="【门店月度收支统计】里配送＋公司的年初是 "&TEXT(SUMIFS({MR("G")},{MR("B")},"配送")+SUMIFS({MR("G")},{MR("B")},"公司"),"#,##0.00")'
          f'&"，按年初银行余额倒推的公司年初结余是 "&TEXT({AUTO0},"#,##0.00")&"。① 用的是倒推的数；想用旧表的，在 ① 的 H 格手填"',
          font=F_NOTE)
    ws.row_dimensions[ri].height = 44
    rn = x + 8
    put(ws, f'A{rn}', '参考', font=F_TXTB, fill=FILL_HDR, align=AC)
    put(ws, f'B{rn}', '没写日期、也没写所属月份的记录')
    put(ws, f'C{rn}', f'=COUNTIFS({XR("A")},0,{DR("F")},"<>")+COUNTIFS({XR("A")},0,{DR("G")},"<>")',
        fill=FILL_AUTO, fmt='0', align=AR)
    for col in 'DEF':
        put(ws, f'{col}{rn}', None, fill=FILL_AUTO)
    merge(ws, f'G{rn}:J{rn}', '这些按年初算进去了（截止日期填哪天都算上），最好去【数据录入】补上日期', font=F_NOTE)
    rk = x + 9
    put(ws, f'A{rk}', '自查', font=F_TXTB, fill=FILL_HDR, align=AC)
    put(ws, f'B{rk}', '④⑤⑥ 明细的合计跟 ① 一致吗')
    for col in 'CDEF':
        put(ws, f'{col}{rk}', None, fill=FILL_AUTO)
    merge(ws, f'G{rk}:J{rk}',
          f'=IF(AND(ABS($G${S_TOT}-$C${r["store"]})<0.005,ABS($C${N_TOT}-$C${r["recv"]})<0.005,'
          f'ABS($D${N_TOT}-$C${r["pay"]})<0.005,ABS($F${A_TOT}-$C${r["bank"]})<0.005),"✓ 一致",'
          f'"✗ 不一致（公式被改动了？）")', font=F_TXTB)

    # ---------------- ④ 收款 / 付款 明细 ----------------
    section(ws, N_H, '④ 收款 / 付款 明细（这些名字不算门店；B 列黄格可以改、可以往下加）')
    heads(ws, N_H + 1, [('A', '序'), ('B', '名字（数据录入 I 列「公司」）'), ('C', '收款'), ('D', '付款'), ('E', '净额'),
                        ('F', '其中走银行的·收'), ('G', '其中走银行的·付'), ('H', '说明')])
    merge(ws, f'H{N_H + 1}:J{N_H + 1}', '说明', font=F_HDR, fill=FILL_HDR, align=AC)
    for i in range(N_1 - N_0 + 1):
        rr = N_0 + i
        nm, note = NS_DEFAULT[i] if i < len(NS_DEFAULT) else (None, '')
        b = f'$B{rr}'
        put(ws, f'A{rr}', i + 1, fill=FILL_HDR, align=AC)
        put(ws, f'B{rr}', nm, font=F_IN, fill=FILL_IN)
        put(ws, f'C{rr}', f'=IF({b}="","",SUMIFS({DR("F")},{DR("I")},{b},{XR("A")},{le(CUT)}))', fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'D{rr}', f'=IF({b}="","",SUMIFS({DR("G")},{DR("I")},{b},{XR("A")},{le(CUT)}))', fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'E{rr}', f'=IF({b}="","",C{rr}-D{rr})', fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'F{rr}', f'=IF({b}="","",SUMIFS({DR("F")},{DR("I")},{b},{XR("B")},1,{XR("A")},{le(CUT)}))',
            fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'G{rr}', f'=IF({b}="","",SUMIFS({DR("G")},{DR("I")},{b},{XR("B")},1,{XR("A")},{le(CUT)}))',
            fill=FILL_AUTO, fmt=MONEY, align=AR)
        merge(ws, f'H{rr}:J{rr}', note, font=F_NOTE)
    rr = N_BL
    put(ws, f'A{rr}', '', fill=FILL_HDR)
    put(ws, f'B{rr}', '没填门店的（I 列空着）')
    for col, fld, extra in (('C', 'F', ''), ('D', 'G', ''), ('F', 'F', f',{XR("B")},1'), ('G', 'G', f',{XR("B")},1')):
        put(ws, f'{col}{rr}', f'=SUMIFS({DR(fld)},{XR("C")},0{extra},{XR("A")},{le(CUT)})-SUM({col}{N_0}:{col}{N_1})',
            fill=FILL_AUTO, fmt=MONEY, align=AR)
    put(ws, f'E{rr}', f'=C{rr}-D{rr}', fill=FILL_AUTO, fmt=MONEY, align=AR)
    merge(ws, f'H{rr}:J{rr}', '比如「24小时微信回款」没分门店；分到门店的话，填上 I 列就会算到门店', font=F_NOTE)
    rr = N_TOT
    put(ws, f'A{rr}', '', fill=FILL_TOT)
    put(ws, f'B{rr}', '合计（＝ ① 的收款 / 付款）', font=F_TXTB, fill=FILL_TOT)
    for col in 'CDEFG':
        put(ws, f'{col}{rr}', f'=SUM({col}{N_0}:{col}{N_BL})', font=F_TXTB, fill=FILL_TOT, fmt=MONEY, align=AR)
    merge(ws, f'H{rr}:J{rr}', '收款付款里含不走银行的（如配送收调拨），所以比银行流水的数大', font=F_NOTE, fill=FILL_TOT)

    # ---------------- ⑤ 门店余额明细 ----------------
    section(ws, S_H, '⑤ 门店余额明细（门店名单取自【门店月度收支统计】，年初数也从那里来）')
    heads(ws, S_H + 1, [('A', '序'), ('B', '门店'), ('C', '年初余额'), ('D', '首月手加的数'), ('E', '收入'),
                        ('F', '支出'), ('G', '期末余额'), ('H', '【门店月度收支统计】\n少算的（全部）'), ('I', '提示')], 40)
    merge(ws, f'I{S_H + 1}:J{S_H + 1}', '提示', font=F_HDR, fill=FILL_HDR, align=AC)
    yfrom, yto = f'{MY}*100+1', f'{MY}*100+12'
    for i in range(S_1 - S_0 + 1):
        rr = S_0 + i
        b = f'$B{rr}'
        put(ws, f'A{rr}', f'=IF({b}="","",{i + 1})', fill=FILL_HDR, align=AC)
        put(ws, f'B{rr}', f'=IFERROR(INDEX({AUX}!$E${M0}:$E${M1},MATCH({i + 1},{AUX}!$G${M0}:$G${M1},0)),"")')
        put(ws, f'C{rr}', f'=IF({b}="","",SUMIFS({MR("G")},{MR("B")},{b}))', fill=FILL_AUTO, fmt=MONEY, align=AR)
        # K：这个门店在数据录入里、所属月份在那张表年份内的净额（＝那张表按公式取到的）；L：这个门店全部记录的净额
        put(ws, f'K{rr}', f'=IF({b}="","",SUMIFS({DR("F")},{DR("I")},{b},{DR("L")},">="&{yfrom},{DR("L")},"<="&{yto})'
                          f'-SUMIFS({DR("G")},{DR("I")},{b},{DR("L")},">="&{yfrom},{DR("L")},"<="&{yto}))',
            font=F_NOTE, border=None)
        put(ws, f'L{rr}', f'=IF({b}="","",SUMIFS({DR("F")},{DR("I")},{b})-SUMIFS({DR("G")},{DR("I")},{b}))',
            font=F_NOTE, border=None)
        put(ws, f'D{rr}', f'=IF({b}="","",ROUND(SUMIFS({MR("D")},{MR("B")},{b})-SUMIFS({MR("E")},{MR("B")},{b})-K{rr},2))',
            fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'E{rr}', f'=IF({b}="","",SUMIFS({DR("F")},{DR("I")},{b},{XR("A")},{le(CUT)}))', fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'F{rr}', f'=IF({b}="","",SUMIFS({DR("G")},{DR("I")},{b},{XR("A")},{le(CUT)}))', fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'G{rr}', f'=IF({b}="","",C{rr}+D{rr}+E{rr}-F{rr})', font=F_TXTB, fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'H{rr}', f'=IF({b}="","",ROUND(L{rr}-K{rr},2))', fill=FILL_AUTO, fmt=MONEY, align=AR)
        merge(ws, f'I{rr}:J{rr}', f'=IF({b}="","",IF(ABS(N(H{rr}))>=0.005,"所属月份没填/不是本年的，那张表没算进去",""))',
              font=F_NOTE)
    ws.conditional_formatting.add(f'H{S_0}:H{S_1}', FormulaRule(formula=[f'ABS(N(H{S_0}))>=0.005'], fill=FILL_ORG))
    ws.conditional_formatting.add(f'D{S_0}:D{S_1}', FormulaRule(formula=[f'ABS(N(D{S_0}))>=0.005'], fill=FILL_IN))
    rr = S_RES
    put(ws, f'A{rr}', '', fill=FILL_HDR)
    put(ws, f'B{rr}', '名字不在【门店月度收支统计】里的（也按门店算）')
    put(ws, f'C{rr}', None, fill=FILL_AUTO)
    put(ws, f'D{rr}', None, fill=FILL_AUTO)
    put(ws, f'E{rr}', f'=SUMIFS({DR("F")},{XR("C")},1,{XR("A")},{le(CUT)})-SUM(E{S_0}:E{S_1})', fill=FILL_AUTO, fmt=MONEY, align=AR)
    put(ws, f'F{rr}', f'=SUMIFS({DR("G")},{XR("C")},1,{XR("A")},{le(CUT)})-SUM(F{S_0}:F{S_1})', fill=FILL_AUTO, fmt=MONEY, align=AR)
    put(ws, f'G{rr}', f'=E{rr}-F{rr}', font=F_TXTB, fill=FILL_AUTO, fmt=MONEY, align=AR)
    put(ws, f'H{rr}', None, fill=FILL_AUTO)
    merge(ws, f'I{rr}:J{rr}', f'=IF(AND(ABS(E{rr})<0.005,ABS(F{rr})<0.005),"",'
                              f'"数据录入 I 列有门店名字写得跟门店月度收支统计不一样，去改一致")', font=F_NOTE)
    ws.conditional_formatting.add(f'E{rr}:G{rr}', FormulaRule(formula=[f'OR(ABS($E${rr})>=0.005,ABS($F${rr})>=0.005)'],
                                                              fill=FILL_RED))
    rr = S_TOT
    put(ws, f'A{rr}', '', fill=FILL_TOT)
    put(ws, f'B{rr}', '合计（＝ ① 的门店余额）', font=F_TXTB, fill=FILL_TOT)
    for col in 'CDEFGH':
        last = S_1 if col in 'CDH' else S_RES
        put(ws, f'{col}{rr}', f'=SUM({col}{S_0}:{col}{last})', font=F_TXTB, fill=FILL_TOT, fmt=MONEY, align=AR)
    merge(ws, f'I{rr}:J{rr}', '', fill=FILL_TOT)

    # ---------------- ⑥ 账户余额 ----------------
    section(ws, A_H, '⑥ 账户余额（14 个账户＝数据录入第 5～18 行；G 列可以手填银行对账单上的余额）')
    heads(ws, A_H + 1, [('A', '序'), ('B', '账户'), ('C', '年初余额'), ('D', '收入'), ('E', '支出'), ('F', '账上余额'),
                        ('G', '银行对账单余额\n（手填，可空）'), ('H', '差\n（对账单－账上）'), ('I', '实存按这个算')], 40)
    acc = '数据录入!$C$5:$C$18'
    for i in range(A_1 - A_0 + 1):
        rr = A_0 + i
        b = f'$B{rr}'
        put(ws, f'A{rr}', f'=IF({b}="","",{i + 1})', fill=FILL_HDR, align=AC)
        put(ws, f'B{rr}', f'=IF({i + 1}>ROWS({acc}),"",INDEX({acc},{i + 1})&"")')
        put(ws, f'C{rr}', f'=IF({b}="","",N(INDEX(数据录入!$H$5:$H$18,{i + 1})))', fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'D{rr}', f'=IF({b}="","",SUMIFS({DR("F")},{DR("C")},{b},{XR("A")},{le(CUT)}))', fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'E{rr}', f'=IF({b}="","",SUMIFS({DR("G")},{DR("C")},{b},{XR("A")},{le(CUT)}))', fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'F{rr}', f'=IF({b}="","",ROUND(C{rr}+D{rr}-E{rr},2))', font=F_TXTB, fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'G{rr}', None, font=F_IN, fill=FILL_IN, fmt=MONEY, align=AR)
        put(ws, f'H{rr}', f'=IF(OR({b}="",NOT(ISNUMBER(G{rr}))),"",ROUND(G{rr}-F{rr},2))', fill=FILL_AUTO, fmt=MONEY, align=AR)
        put(ws, f'I{rr}', f'=IF({b}="","",IF(ISNUMBER(G{rr}),G{rr},F{rr}))', fill=FILL_AUTO, fmt=MONEY, align=AR)
    dva = DataValidation(type='decimal', operator='between', formula1='-9999999999', formula2='9999999999',
                         allow_blank=True, showErrorMessage=True, errorTitle='只能填数字', error='填对账单上的余额')
    ws.add_data_validation(dva)
    dva.add(f'G{A_0}:G{A_1}')
    ws.conditional_formatting.add(f'H{A_0}:H{A_1}', FormulaRule(formula=[f'AND(H{A_0}<>"",ABS(N(H{A_0}))>=0.005)'],
                                                                fill=FILL_RED))
    rr = A_TOT
    put(ws, f'A{rr}', '', fill=FILL_TOT)
    put(ws, f'B{rr}', '合计（账上余额＝ ① 的实存）', font=F_TXTB, fill=FILL_TOT)
    for col in 'CDEFGHI':
        put(ws, f'{col}{rr}', f'=ROUND(SUM({col}{A_0}:{col}{A_1}),2)', font=F_TXTB, fill=FILL_TOT, fmt=MONEY, align=AR)
    merge(ws, f'A{A_TOT + 1}:J{A_TOT + 1}', 'G 列：截止日期那天银行 App / 对账单上的余额（现金填盘点数）。不填就按账上余额算；'
                                          '截止日期变了，这一列要跟着改。', font=F_NOTE, border=None)

    ws.column_dimensions['K'].hidden = True
    ws.column_dimensions['L'].hidden = True
    ws.freeze_panes = 'A6'
    ws.print_options.horizontalCentered = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = '1:5'
    os.makedirs(TMP, exist_ok=True)
    p = os.path.join(TMP, 'main.xlsx')
    wb.save(p)
    return p


# ============================================================ 【门店余额核对辅助】（直接写 XML，用共享公式）
def aux_cells():
    """返回 {行: [(列, 公式或None, 文字或None, 共享信息)]}"""
    rows = {}

    def add(r, col, f=None, text=None, shared=None):
        rows.setdefault(r, []).append((col, f, text, shared))

    add(1, 'A', text='【门店余额核对】用的辅助表：不要改、不要删。A～C 列跟【数据录入】一行对一行（第 19～30191 行）')
    add(3, 'A', text='计入日期')
    add(3, 'B', text='是银行账户')
    add(3, 'C', text='是门店')
    add(4, 'E', text='门店月度收支统计 B 列')
    add(4, 'F', text='是门店且第一次出现')
    add(4, 'G', text='累计第几个门店')
    labels = ['截止日期（数字）', '最后一笔的计入日期', '年初银行余额（14 个账户）', '账本年份', '公司年初结余（自动算）']
    vals = [f'IF(ISNUMBER({SH}!$B${R_CUT}),{SH}!$B${R_CUT},2958465)', f'MAX($A${D0}:$A${D1})',
            'SUM(数据录入!$H$5:$H$18)', 'YEAR(数据录入!$B$5)', f'$J$3-{SH}!$C${S_TOT}-{SH}!$D${S_TOT}']
    for i, (lab, f) in enumerate(zip(labels, vals), 1):
        add(i, 'I', text=lab)
        add(i, 'J', f=f)
    # 数据录入逐行
    d = lambda c: f'数据录入!{c}{D0}'
    fa = (f'IF(AND(ISNUMBER({d("B")}),{d("B")}>=36526,{d("B")}<=2958465),INT({d("B")}),'
          f'IF(AND(ISNUMBER({d("L")}),{d("L")}>=190001,MOD({d("L")},100)>=1,MOD({d("L")},100)<=12),'
          f'DATE(INT({d("L")}/100),MOD({d("L")},100)+1,0),0))')
    fb = f'IF({d("C")}="",0,IF(COUNTIF(数据录入!$C$5:$C$18,{d("C")})>0,1,0))'
    fc = f'IF({d("I")}="",0,IF(COUNTIF({SH}!{NSR},{d("I")})>0,0,1))'
    for col, f, si in (('A', fa, 0), ('B', fb, 1), ('C', fc, 2)):
        add(D0, col, f=f, shared=(si, f'{col}{D0}:{col}{D1}'))
        for r in range(D0 + 1, D1 + 1):
            add(r, col, shared=(si, None))
    # 门店名单（从门店月度收支统计 B 列取不重复的门店名）
    fe = f'IFERROR(IF(门店月度收支统计!B{M0}="","",门店月度收支统计!B{M0}&""),"")'
    ff = f'IF(E{M0}="",0,IF(COUNTIF({SH}!{NSR},E{M0})>0,0,IF(COUNTIF($E${M0}:E{M0},E{M0})=1,1,0)))'
    add(M0, 'E', f=fe, shared=(3, f'E{M0}:E{M1}'))
    add(M0, 'F', f=ff, shared=(4, f'F{M0}:F{M1}'))
    add(M0, 'G', f=f'F{M0}')
    add(M0 + 1, 'G', f=f'G{M0}+F{M0 + 1}', shared=(5, f'G{M0 + 1}:G{M1}'))
    for r in range(M0 + 1, M1 + 1):
        add(r, 'E', shared=(3, None))
        add(r, 'F', shared=(4, None))
        if r > M0 + 1:
            add(r, 'G', shared=(5, None))
    return rows


def col_num(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;'))


def vxml(v):
    """缓存值 → (t 属性, <v> 内容)"""
    if v is None:
        return None, None
    if isinstance(v, bool):
        return 'b', '1' if v else '0'
    if isinstance(v, (int, float)):
        return None, repr(float(v)) if isinstance(v, float) else str(v)
    if hasattr(v, 'toordinal'):          # 日期 → Excel 序列号
        import datetime as dt
        base = dt.datetime(1899, 12, 30)
        if not isinstance(v, dt.datetime):
            v = dt.datetime(v.year, v.month, v.day)
        return None, repr((v - base).total_seconds() / 86400.0).rstrip('0').rstrip('.')
    s = str(v)
    if s.startswith('#'):
        return 'e', s
    return 'str', s


def aux_xml(cache, hdr_style):
    rows = aux_cells()
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           f'<worksheet xmlns="{NSX}" xmlns:r="{RNS}">',
           '<sheetPr><tabColor rgb="FF7F7F7F"/></sheetPr>',
           f'<dimension ref="A1:J{D1}"/>',
           '<sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/>'
           '</sheetView></sheetViews>',
           '<sheetFormatPr defaultRowHeight="15"/>',
           '<cols><col min="1" max="3" width="12" customWidth="1"/><col min="5" max="5" width="28" customWidth="1"/>'
           '<col min="6" max="7" width="14" customWidth="1"/><col min="9" max="9" width="26" customWidth="1"/>'
           '<col min="10" max="10" width="16" customWidth="1"/></cols>',
           '<sheetData>']
    for r in sorted(rows):
        cells = sorted(rows[r], key=lambda x: col_num(x[0]))
        parts = [f'<row r="{r}">']
        for col, f, text, shared in cells:
            ref = f'{col}{r}'
            if text is not None:
                st = f' s="{hdr_style}"' if r in (1, 3, 4) or col == 'I' else ''
                parts.append(f'<c r="{ref}"{st} t="inlineStr"><is><t>{esc(text)}</t></is></c>')
                continue
            t, v = vxml(cache.get(ref))
            ta = f' t="{t}"' if t else ''
            if shared is not None:
                si, rng = shared
                fx = (f'<f t="shared" ref="{rng}" si="{si}">{esc(f)}</f>' if rng else f'<f t="shared" si="{si}"/>')
            else:
                fx = f'<f>{esc(f)}</f>'
            vx = f'<v>{esc(v)}</v>' if v is not None else ''
            parts.append(f'<c r="{ref}"{ta}>{fx}{vx}</c>')
        parts.append('</row>')
        out.append(''.join(parts))
    out.append('</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
    return '\n'.join(out).encode('utf-8')


# ============================================================ 样式合并
def merge_styles(orig_xml, tmp_xml):
    """把临时工作簿的样式接到原文件 styles.xml 后面，返回 (新 styles.xml, xf 映射, dxf 映射, 辅助表表头样式号)"""
    o = etree.fromstring(orig_xml)
    t = etree.fromstring(tmp_xml)

    def coll(root, tag):
        return root.find(NS + tag)

    # numFmts
    onf = coll(o, 'numFmts')
    if onf is None:
        onf = etree.Element(NS + 'numFmts', count='0')
        o.insert(0, onf)
    used = {int(n.get('numFmtId')) for n in onf}
    nxt = max(used | {200}) + 1
    nf_map = {}
    tnf = coll(t, 'numFmts')
    if tnf is not None:
        for n in tnf:
            old = int(n.get('numFmtId'))
            code = n.get('formatCode')
            same = [int(x.get('numFmtId')) for x in onf if x.get('formatCode') == code]
            if same:
                nf_map[old] = same[0]
                continue
            e = etree.SubElement(onf, NS + 'numFmt', numFmtId=str(nxt), formatCode=code)
            nf_map[old] = nxt
            nxt += 1
    onf.set('count', str(len(onf)))

    offs = {}
    for tag in ('fonts', 'fills', 'borders'):
        oc, tc = coll(o, tag), coll(t, tag)
        offs[tag] = len(oc)
        for e in tc:
            oc.append(copy.deepcopy(e))
        oc.set('count', str(len(oc)))
    ox, tx = coll(o, 'cellXfs'), coll(t, 'cellXfs')
    xf0 = len(ox)
    xf_map = {}
    for i, e in enumerate(tx):
        e = copy.deepcopy(e)
        nf = int(e.get('numFmtId', '0'))
        e.set('numFmtId', str(nf_map.get(nf, nf)))
        e.set('fontId', str(int(e.get('fontId', '0')) + offs['fonts']))
        e.set('fillId', str(int(e.get('fillId', '0')) + offs['fills']))
        e.set('borderId', str(int(e.get('borderId', '0')) + offs['borders']))
        e.set('xfId', '0')
        ox.append(e)
        xf_map[i] = xf0 + i
    # 辅助表表头：微软雅黑粗体
    of = coll(o, 'fonts')
    fe = etree.SubElement(of, NS + 'font')
    etree.SubElement(fe, NS + 'b')
    etree.SubElement(fe, NS + 'sz', val='10')
    etree.SubElement(fe, NS + 'name', val='微软雅黑')
    etree.SubElement(fe, NS + 'charset', val='134')
    of.set('count', str(len(of)))
    etree.SubElement(ox, NS + 'xf', numFmtId='0', fontId=str(len(of) - 1), fillId='0', borderId='0', xfId='0',
                     applyFont='1')
    hdr_idx = len(ox) - 1
    ox.set('count', str(len(ox)))
    od, td = coll(o, 'dxfs'), coll(t, 'dxfs')
    dxf_map = {}
    if td is not None and len(td):
        if od is None:
            od = etree.Element(NS + 'dxfs', count='0')
            coll(o, 'cellStyles').addnext(od)
        d0 = len(od)
        for i, e in enumerate(td):
            od.append(copy.deepcopy(e))
            dxf_map[i] = d0 + i
        od.set('count', str(len(od)))
    xml = etree.tostring(o, xml_declaration=True, encoding='UTF-8', standalone=True)
    return xml, xf_map, dxf_map, hdr_idx


def main_xml(tmp_path, xf_map, dxf_map, cache):
    z = zipfile.ZipFile(tmp_path)
    sst = []
    if 'xl/sharedStrings.xml' in z.namelist():
        s = etree.fromstring(z.read('xl/sharedStrings.xml'))
        for si in s.findall(NS + 'si'):
            sst.append(''.join(x.text or '' for x in si.iter(NS + 't')))
    root = etree.fromstring(z.read('xl/worksheets/sheet1.xml'))
    for sv in root.iter(NS + 'sheetView'):
        if 'tabSelected' in sv.attrib:
            del sv.attrib['tabSelected']
    for c in root.iter(NS + 'c'):
        s = c.get('s')
        c.set('s', str(xf_map.get(int(s) if s else 0, xf_map[0])))
        if c.get('t') == 's':
            v = c.find(NS + 'v')
            text = sst[int(v.text)]
            c.remove(v)
            c.set('t', 'inlineStr')
            is_ = etree.SubElement(c, NS + 'is')
            tt = etree.SubElement(is_, NS + 't')
            tt.text = text
            if text != text.strip():
                tt.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        f = c.find(NS + 'f')
        if f is not None:
            v = c.find(NS + 'v')
            if v is not None:
                c.remove(v)
            if 't' in c.attrib:
                del c.attrib['t']
            t, val = vxml(cache.get(c.get('r')))
            if val is not None:
                if t:
                    c.set('t', t)
                ve = etree.SubElement(c, NS + 'v')
                ve.text = val
    for rule in root.iter(NS + 'cfRule'):
        if rule.get('dxfId') is not None:
            rule.set('dxfId', str(dxf_map[int(rule.get('dxfId'))]))
    # 没写 s 的格子（openpyxl 默认样式）也要指到新样式 0
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True), \
        z.read('xl/styles.xml')


# ============================================================ 拼包
def assemble(src, dst, cache_main, cache_aux):
    tmp = build_main()
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    styles_new, xf_map, dxf_map, hdr_idx = merge_styles(zin.read('xl/styles.xml'), zipfile.ZipFile(tmp).read('xl/styles.xml'))
    sheet_main, _ = main_xml(tmp, xf_map, dxf_map, cache_main)
    sheet_aux = aux_xml(cache_aux, hdr_idx)

    wbx = zin.read('xl/workbook.xml').decode('utf-8')
    rels = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    ct = zin.read('[Content_Types].xml').decode('utf-8')
    app = zin.read('docProps/app.xml').decode('utf-8')
    if f'name="{SH}"' in wbx:
        raise SystemExit('原表里已经有【门店余额核对】了：请用没加过的原表')
    sheet_ids = [int(x) for x in re.findall(r'<sheet [^>]*sheetId="(\d+)"', wbx)]
    rids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    parts = [int(x) for x in re.findall(r'worksheets/sheet(\d+)\.xml', rels)]
    sid1, sid2 = max(sheet_ids) + 1, max(sheet_ids) + 2
    rid1, rid2 = f'rId{max(rids) + 1}', f'rId{max(rids) + 2}'
    p1, p2 = f'sheet{max(parts) + 1}.xml', f'sheet{max(parts) + 2}.xml'
    wbx = wbx.replace('</sheets>', f'<sheet name="{SH}" sheetId="{sid1}" r:id="{rid1}"/>'
                                   f'<sheet name="{AUX}" sheetId="{sid2}" state="hidden" r:id="{rid2}"/></sheets>', 1)
    ws_t = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
    rels = rels.replace('</Relationships>', f'<Relationship Id="{rid1}" Type="{ws_t}" Target="worksheets/{p1}"/>'
                                            f'<Relationship Id="{rid2}" Type="{ws_t}" Target="worksheets/{p2}"/></Relationships>', 1)
    ws_ct = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
    ct = ct.replace('</Types>', f'<Override PartName="/xl/worksheets/{p1}" ContentType="{ws_ct}"/>'
                                f'<Override PartName="/xl/worksheets/{p2}" ContentType="{ws_ct}"/></Types>', 1)
    m = re.search(r'<vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>(\d+)</vt:i4>', app)
    if m:
        n = int(m.group(1))
        app = app.replace(m.group(0), m.group(0).replace(f'<vt:i4>{n}</vt:i4>', f'<vt:i4>{n + 2}</vt:i4>'), 1)
        m2 = re.search(r'<TitlesOfParts><vt:vector size="(\d+)" baseType="lpstr">', app)
        if m2:
            k = int(m2.group(1))
            app = app.replace(m2.group(0), m2.group(0).replace(f'size="{k}"', f'size="{k + 2}"'), 1)
            app = app.replace('</vt:vector></TitlesOfParts>',
                              f'<vt:lpstr>{SH}</vt:lpstr><vt:lpstr>{AUX}</vt:lpstr></vt:vector></TitlesOfParts>', 1)
    repl = {'xl/workbook.xml': wbx.encode('utf-8'), 'xl/_rels/workbook.xml.rels': rels.encode('utf-8'),
            '[Content_Types].xml': ct.encode('utf-8'), 'docProps/app.xml': app.encode('utf-8'),
            'xl/styles.xml': styles_new}
    zout = zipfile.ZipFile(dst + '.part', 'w', zipfile.ZIP_DEFLATED)
    for info in zin.infolist():
        data = repl.get(info.filename, zin.read(info.filename))
        zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
        zi.compress_type = zipfile.ZIP_DEFLATED
        zi.external_attr = info.external_attr
        zout.writestr(zi, data)
    zout.writestr(f'xl/worksheets/{p1}', sheet_main, zipfile.ZIP_DEFLATED)
    zout.writestr(f'xl/worksheets/{p2}', sheet_aux, zipfile.ZIP_DEFLATED)
    zout.close()
    os.replace(dst + '.part', dst)
    return p1, p2


def recalc_values(path):
    """副本上用 LibreOffice 整本重算，读回新加两张表的数"""
    rc = (glob.glob('/root/.claude/skills/synced/*/xlsx/scripts/recalc.py') +
          glob.glob('/mnt/skills/public/xlsx/scripts/recalc.py'))[0]
    cp = os.path.join(TMP, 'calc.xlsx')
    shutil.copy(path, cp)
    r = subprocess.run([sys.executable, rc, cp, '1800', '--force'], capture_output=True, text=True)
    print('   LibreOffice：', r.stdout.strip()[:400])
    wb = load_workbook(cp, data_only=True, read_only=True)
    out = []
    for nm in (SH, AUX):
        d = {}
        for row in wb[nm].iter_rows():
            for c in row:
                if c.value is not None:
                    d[c.coordinate] = c.value
        out.append(d)
    return out


if __name__ == '__main__':
    os.makedirs(TMP, exist_ok=True)
    print('① 拼新表（还没缓存值）')
    assemble(SRC, OUT, {}, {})
    print('② LibreOffice 副本上算数')
    cm, ca = recalc_values(OUT)
    print(f'   新表 {len(cm)} 格有数，辅助表 {len(ca)} 格')
    print('③ 把算好的数写进成品当缓存')
    assemble(SRC, OUT, cm, ca)
    print('已保存：', OUT, f'{os.path.getsize(OUT) / 1e6:.1f} MB')
