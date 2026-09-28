# -*- coding: utf-8 -*-
"""《03》往来这一块：
   ① 往来业务明细：对账补充行扩到 300 行；辅助列改成「按日期排序键」；加「对账金额 / 对账已付款」两列
      （收、付、定金、退款、借款全放一起算一个净额）；
   ② 应收账款对账单 / 应付账款对账单 / 往来对账单：统一成原来「新疆果然鲜仓储有限公司 对账单」版式，
      逐笔按日期排，每行带期末余额；往来对账单不分应收应付，只要跟这个人有关的全显示；
   ③ 往来对账单明细（新）、应付账款明细、应收应付汇总表：加时段选择和合计；
   ④ 库存价值与欠款比对：应收款余额改取往来对账单明细的期末余额。"""
import copy
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from common0928 import *

OLD_END = 12104

# ─────────────────────────────── ① 往来业务明细 ───────────────────────────────
def fix_detail(wb):
    ws = wb['往来业务明细']
    # 对账补充行 104～303 行接到 12105～12304
    for i, r in enumerate(range(OLD_END + 1, WR1 + 1)):
        k = 104 + i
        row = {
            'A': '03补充', 'B': k,
            'C': f'=IFERROR(1*对账补充行!B{k},0)',
            'D': f'=IF(OR(K{r}<>0,L{r}<>0),IF(对账补充行!C{k}="","",对账补充行!C{k}),"")',
            'E': f'=IF(对账补充行!L{k}="应付","应付","应收")',
            'F': f'=IF(对账补充行!D{k}="","",对账补充行!D{k})',
            'H': f'=IF(对账补充行!E{k}="","",对账补充行!E{k})',
            'I': f'=IFERROR(1*对账补充行!F{k},0)',
            'J': f'=IFERROR(1*对账补充行!G{k},0)',
            'K': f'=IF(对账补充行!K{k}<>"",IFERROR(1*对账补充行!K{k},0),ROUND(I{r}*J{r},2))',
            'L': f'=-IFERROR(1*对账补充行!I{k},0)',
            'M': f'=K{r}+L{r}',
            'N': f'=IF(对账补充行!J{k}="","",对账补充行!J{k})',
        }
        for c, v in row.items():
            ws[f'{c}{r}'] = v
    # 辅助列：Q 单位去重编号；R 往来对账单分组键；S 往来对账单显示键；T/U 应收/应付对账单排序键；
    #          V 对账金额；W 对账已付款；X 所属显示键
    SQ = "往来对账单"; SR = "应收账款对账单"; SP = "应付账款对账单"
    rng = lambda c: f'${c}${WR0}:${c}${WR1}'
    for r in range(WR0, WR1 + 1):
        per = lambda s: f'OR(C{r}=0,AND(C{r}>={s}!$M$2,C{r}<={s}!$M$3))'
        key = f'IF(C{r}>0,C{r},{INF})*100000+ROW()'
        ws[f'Q{r}'] = f'=N(Q{r-1})+IF(D{r}="",0,IF(MATCH(D{r},{rng("D")},0)=ROW()-3,1,0))'
        ws[f'R{r}'] = (f'=IF(AND(D{r}<>"",D{r}={SQ}!$B$2,OR(K{r}<>0,L{r}<>0),{per(SQ)}),'
                       f'IF(A{r}="03资金","资金|"&B{r},E{r}&"|"&F{r}&"|"&G{r}&"|"&H{r}&"|"&TEXT(J{r},"0.########")),"")')
        mx = f'_xlfn.MAXIFS({rng("C")},{rng("R")},R{r})'
        dd = f'IF({SQ}!$M$5=1,{mx},C{r})'
        ws[f'S{r}'] = (f'=IF(R{r}="","",IF(AND({SQ}!$M$5=1,MATCH(R{r},{rng("R")},0)<>ROW()-3),"",'
                       f'IF({dd}>0,{dd},{INF})*100000+ROW()))')
        ws[f'T{r}'] = (f'=IF(AND(D{r}<>"",D{r}={SR}!$B$2,E{r}="应收",OR(K{r}<>0,L{r}<>0),{per(SR)}),{key},"")')
        ws[f'U{r}'] = (f'=IF(AND(D{r}<>"",D{r}={SP}!$B$2,E{r}="应付",OR(K{r}<>0,L{r}<>0),{per(SP)}),{key},"")')
        ws[f'V{r}'] = f'=IF(D{r}="",0,IF(A{r}="03资金",0,IF(E{r}="应付",-K{r},K{r})))'
        ws[f'W{r}'] = f'=IF(D{r}="",0,IF(A{r}="03资金",IF(E{r}="应收",-(K{r}+L{r}),K{r}+L{r}),IF(E{r}="应付",L{r},-L{r})))'
        ws[f'X{r}'] = f'=IF(R{r}="","",IF({SQ}!$M$5=1,INDEX({rng("S")},MATCH(R{r},{rng("R")},0)),S{r}))'
    # ── 样式：跟前面的表统一 ──
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    ws['A1'] = '往来业务明细 · 每笔往来业务可追溯（全部自动，勿手工改）'
    title(ws, ws['A1'].value, 'A', 'P')
    ws['A2'] = ('销售、有价领用、仓储/装卸/周转费计应收；采购、费用应计计应付；收付款只取【资金日记账】里往来分类为应收/应付的行'
                '（定金、押金、借款、货款）。K 发生额＋L 收付冲减＝M 本笔增减。V/W 两列是【往来对账单】用的收付全口径：'
                'V＝我方应收为正、应付为负；W＝对方付给我方为正、我方付给对方为负。')
    note(ws, ws['A2'].value, 'A', 'X', 2, height=36)
    heads = {'A': '来源', 'B': '源行', 'C': '日期', 'D': '往来单位', 'E': '方向', 'F': '摘要', 'G': '品名', 'H': '单位',
             'I': '数量', 'J': '单价', 'K': '发生额', 'L': '收付冲减', 'M': '本笔增减', 'N': '备注', 'O': '单据号',
             'P': '资金账户', 'Q': '单位编号', 'R': '分组键', 'S': '显示键', 'T': '应收键', 'U': '应付键',
             'V': '对账·金额', 'W': '对账·已付款', 'X': '所属键'}
    for c, t in heads.items():
        cell(ws, f'{c}3', t, F_HDR, FL_HDR, AC)
    ws.row_dimensions[3].height = 30
    fmt = {'C': 'yyyy/m/d;;;', 'I': QTY, 'J': MONEY, 'K': MONEY, 'L': MONEY, 'M': MONEY, 'V': MONEY, 'W': MONEY}
    st_txt = dict(font=F_AUTO, fl=FL_AUTO, al=ACN)
    for r in range(WR0, WR1 + 1):
        for c in 'ABCDEFGHIJKLMNOPVW':
            x = ws[f'{c}{r}']
            x.font = F_AUTO; x.fill = FL_AUTO; x.border = BOX
            x.alignment = AR if c in fmt and c != 'C' else ACN
            if c in fmt: x.number_format = fmt[c]
        for c in 'QRSTUX':
            ws[f'{c}{r}'].font = F_HELP
    widths(ws, {'A': 11, 'B': 6, 'C': 11, 'D': 14, 'E': 6, 'F': 16, 'G': 14, 'H': 6, 'I': 10, 'J': 10, 'K': 13,
                'L': 13, 'M': 13, 'N': 22, 'O': 13, 'P': 13, 'V': 13, 'W': 13})
    for c in 'QRSTUX':
        ws.column_dimensions[c].hidden = True
    ws.freeze_panes = 'E4'
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f'A3:P{WR1}'


# ─────────────────────── ② 三张对账单（统一成原来的对账单版式） ───────────────────────
S_TEAL = '275C70'
PAGE = 50
R0, R1 = 7, 7 + PAGE - 1            # 明细 7～56
FT = R1 + 1                         # 页脚第一行 57

def _f(name, size, bold=False, color='1F3864'):
    return Font(name=name, size=size, bold=bold, color=color)

_sb = Side(style='thin', color='8EA9C1')
SBOX = Border(left=_sb, right=_sb, top=_sb, bottom=_sb)


def statement(wb, name, kind):
    """kind: '全部'（往来对账单）/ '应收' / '应付'"""
    ws = wb[name]
    wipe(ws)
    ws.sheet_properties.tabColor = S_TEAL
    widths(ws, {'A': 12, 'B': 30, 'C': 7, 'D': 11, 'E': 10, 'F': 15, 'G': 14, 'H': 15, 'I': 24,
                'K': 16, 'L': 8, 'M': 14})
    hide(ws, 'J', 'K', 'L', 'M')
    ws.sheet_view.showGridLines = False
    # 第 1 行抬头
    ws.merge_cells('A1:G1')
    cell(ws, 'A1', '新疆果然鲜仓储有限公司', _f('黑体', 24, True, 'FFFFFF'), fill(S_TEAL), AC, border=NOB)
    ws.merge_cells('H1:I1')
    t = {'全部': '往来对账单', '应收': '应收对账单', '应付': '应付对账单'}[kind]
    cell(ws, 'H1', t, _f('黑体', 16, True, S_TEAL), fill('FFFFFF'), AC,
         border=Border(*(Side(style='medium', color=S_TEAL),) * 4))
    ws['I1'].border = Border(*(Side(style='medium', color=S_TEAL),) * 4)
    ws.row_dimensions[1].height = 54
    lab = lambda ref, txt: cell(ws, ref, txt, _f('宋体', 11, True), FL_NONE, AR, border=NOB)
    sm = lambda ref, txt: cell(ws, ref, txt, _f('宋体', 9, False, '595959'), fill('F2F2F2'), AC, border=SBOX)
    inp = lambda ref, v=None, fmt=None: cell(ws, ref, v, _f('宋体', 11, True, '1F4E79'), fill('DDEBF7'), AC, fmt, SBOX)
    # 第 2～4 行
    lab('A2', '客　户:'); ws.merge_cells('B2:C2'); inp('B2', {'全部': '吐尔逊', '应收': '麦麦提', '应付': '全修塑业'}[kind])
    ws['B2'].font = _f('宋体', 12, True, '1F4E79'); ws['C2'].border = SBOX
    sm('D2', '年　份:'); inp('E2', 2026, '0')
    sm('F2', '开始日期:'); inp('G2', None, 'yyyy-m-d')
    sm('H2', '结束日期:'); inp('I2', None, 'yyyy-m-d')
    lab('A3', '品　类:'); ws.merge_cells('B3:C3'); inp('B3'); ws['C3'].border = SBOX
    sm('D3', '明细方式:')
    if kind == '全部':
        inp('E3', '逐笔')
        dv_list(ws, 'E3', '"逐笔,合并同价"', err=True)
    else:
        cell(ws, 'E3', '逐笔', _f('宋体', 10, False, '595959'), fill('F2F2F2'), AC, border=SBOX)
    sm('F3', '对账笔数:')
    cell(ws, 'G3', '=$M$6', _f('宋体', 11, True, '006100'), FL_AUTO, AC, INT, SBOX)
    sm('H3', '是否含税:'); inp('I3', '不含税')
    dv_list(ws, 'I3', '"含税,不含税"', err=True)
    lab('A4', '电　话:'); ws.merge_cells('B4:C4'); inp('B4'); ws['C4'].border = SBOX
    sm('D4', '页　码:'); inp('E4', 1, '0')
    sm('F4', '筛选说明')
    ws.merge_cells('G4:I4')
    how = {'全部': '收款、付款、定金、押金、退款、借款、物料、筐子、仓储/装卸/周转费、公司购买果品全在一张单上',
           '应收': '只列应收方向（销售、仓储/装卸/周转费、物料、借出）；定金押金属应付，看【往来对账单】',
           '应付': '只列应付方向（采购、公司购买果品、定金押金、借入）；全口径看【往来对账单】'}[kind]
    cell(ws, 'G4', f'="客户必填；年份/日期留空＝不限。共 "&$M$7&" 页，每页 {PAGE} 笔，合计含全部页。{how}"',
         _f('宋体', 9, False, '7F6000'), FL_NOTE, AL, border=SBOX)
    ws['H4'].border = SBOX; ws['I4'].border = SBOX
    for r in (2, 3, 4):
        ws.row_dimensions[r].height = 24 if r < 4 else 36
    dvy = DataValidation(type='list', formula1=YEARS, allow_blank=True, showErrorMessage=False)
    dvy.add('E2'); ws.add_data_validation(dvy)
    dvd = DataValidation(type='date', operator='between', formula1='36526', formula2='73050', allow_blank=True,
                         showErrorMessage=True, errorTitle='日期', error='请填日期，如 2026/9/1')
    dvd.add('G2'); dvd.add('I2'); ws.add_data_validation(dvd)
    dvp = DataValidation(type='whole', operator='between', formula1='1', formula2='99', allow_blank=False,
                         showErrorMessage=True, errorTitle='页码', error='页码填 1、2、3……')
    dvp.add('E4'); ws.add_data_validation(dvp)
    names_dv = DataValidation(type='list', formula1='INDIRECT("往来对账单明细!$B$7:$B$206")', allow_blank=True,
                              showErrorMessage=False)
    names_dv.add('B2'); ws.add_data_validation(names_dv)
    # 第 5 行表头
    heads = ['日　期', '摘　　要', '单 位', '数量/毛重', '单　价', '金额（元）', '已付款', '期末余额', '备　　注']
    for i, h in enumerate(heads):
        cell(ws, f'{"ABCDEFGHI"[i]}5', h, _f('黑体', 11, True, 'FFFFFF'), FL_HDR, AC, border=SBOX)
    ws.row_dimensions[5].height = 26
    # 隐藏辅助格
    M = ws
    M['M2'] = '=IF(N($G$2)>0,$G$2,IF(N($E$2)>0,DATE($E$2,1,1),1))'
    M['M3'] = f'=IF(N($I$2)>0,$I$2,IF(N($E$2)>0,DATE($E$2,12,31),{INF}))'
    M['M5'] = '=IF($E$3="合并同价",1,0)' if kind == '全部' else '=0'
    keycol = {'全部': 'S', '应收': 'T', '应付': 'U'}[kind]
    member = {'全部': 'X', '应收': 'T', '应付': 'U'}[kind]
    M['M6'] = f'=COUNT({W(keycol)})'
    M['M7'] = f'=MAX(1,ROUNDUP($M$6/{PAGE},0))'
    if kind == '全部':
        M['M4'] = (f'=ROUND(SUMIFS({W("V")},{W("D")},$B$2,{W("C")},">0",{W("C")},"<"&$M$2)'
                   f'-SUMIFS({W("W")},{W("D")},$B$2,{W("C")},">0",{W("C")},"<"&$M$2),2)')
    else:
        M['M4'] = (f'=ROUND(SUMIFS({W("M")},{W("D")},$B$2,{W("E")},"{kind}",{W("C")},">0",{W("C")},"<"&$M$2),2)')
    for k, lbl in (('K1', '辅助：大写金额 / 起 / 止 / 期初 / 合并 / 笔数 / 页数'),):
        ws[k] = lbl
    for r in range(1, 8):
        ws[f'M{r}'].font = F_HELP
    # 第 6 行：上期结余
    cell(ws, 'A6', '=IF($M$2<=1,"",$M$2)', _f('宋体', 10), fill('FFF2CC'), AC, 'yyyy-m-d', SBOX)
    cell(ws, 'B6', '=IF($M$2<=1,"期初结余（建账以来）","上期结余（开始日期之前）")', _f('宋体', 10, True), fill('FFF2CC'), AL,
         border=SBOX)
    for c in 'CDEFGI':
        cell(ws, f'{c}6', None, _f('宋体', 10), fill('FFF2CC'), AC, border=SBOX)
    cell(ws, 'H6', '=$M$4', _f('宋体', 10, True), fill('FFF2CC'), AR, MONEY, SBOX)
    # 明细 7～56
    WD = lambda c: f'往来业务明细!${c}$1:${c}${WR1}'
    for r in range(R0, R1 + 1):
        band = fill('FFFFFF') if r % 2 else fill('EBF3E6')
        k = f'($E$4-1)*{PAGE}+ROW()-{R0 - 1}'
        ws[f'K{r}'] = f'=IFERROR(SMALL({W(keycol)},{k}),"")'
        ws[f'L{r}'] = f'=IF($K{r}="","",$K{r}-INT($K{r}/100000)*100000)'
        ws[f'K{r}'].font = F_HELP; ws[f'L{r}'].font = F_HELP
        g = lambda c: f'INDEX({WD(c)},$L{r})'
        f = {
            'A': f'=IF($K{r}="","",IF(INT($K{r}/100000)>={INF},"",INT($K{r}/100000)))',
            'B': f'=IF($L{r}="","",{g("F")}&IF({g("G")}&""="",""," "&{g("G")}))',
            'C': f'=IF($L{r}="","",{g("H")}&"")',
            'E': f'=IF($L{r}="","",N({g("J")}))',
            'I': f'=IF($L{r}="","",TRIM(IF({g("A")}="03资金",{g("P")}&" ","")&{g("N")}&""))',
        }
        if kind == '全部':
            f['D'] = f'=IF($L{r}="","",IF($M$5=1,SUMIFS({W("I")},{W("X")},$K{r}),N({g("I")})))'
            f['F'] = f'=IF($L{r}="","",ROUND(SUMIFS({W("V")},{W("X")},$K{r}),2))'
            f['G'] = f'=IF($L{r}="","",ROUND(SUMIFS({W("W")},{W("X")},$K{r}),2))'
            f['H'] = (f'=IF($L{r}="","",ROUND($M$4+SUMIFS({W("V")},{W("X")},"<="&$K{r})'
                      f'-SUMIFS({W("W")},{W("X")},"<="&$K{r}),2))')
        else:
            f['D'] = f'=IF($L{r}="","",N({g("I")}))'
            f['F'] = f'=IF($L{r}="","",ROUND(N({g("K")}),2))'
            f['G'] = f'=IF($L{r}="","",ROUND(-N({g("L")}),2))'
            f['H'] = f'=IF($L{r}="","",ROUND($M$4+SUMIFS({W("M")},{W(keycol)},"<="&$K{r}),2))'
        fmts = {'A': 'yyyy-m-d', 'D': '#,##0.###;[Red]\\-#,##0.###;', 'E': '#,##0.00;[Red]\\-#,##0.00;',
                'F': MONEY, 'G': MONEY, 'H': MONEY}
        for c in 'ABCDEFGHI':
            cell(ws, f'{c}{r}', f[c], _f('宋体', 10), band, AL if c in 'BI' else (AR if c in 'DEFGH' else AC),
                 fmts.get(c), SBOX)
        ws.row_dimensions[r].height = 19
    # 页脚 57～60
    a, b, c_, d = FT, FT + 1, FT + 2, FT + 3
    ff = _f('宋体', 11, True)
    for r in (a, b, c_):
        ws.merge_cells(f'A{r}:D{r}')
        ws.row_dimensions[r].height = 24
    cell(ws, f'A{a}', '=IF($I$2="","对账日期：　"&TEXT(TODAY(),"yyyy 年 m 月 d 日"),"对账日期：　"&TEXT($I$2,"yyyy 年 m 月 d 日"))',
         ff, fill('DCE6F1'), AL, border=SBOX)
    cell(ws, f'A{b}', f'="对账时段：　"&IF($M$2<=1,"不限",TEXT($M$2,"yyyy/ m/ d"))&"　--　"&IF($M$3>={INF},"不限",TEXT($M$3,"yyyy/ m/ d"))',
         ff, fill('DCE6F1'), AL, border=SBOX)
    cell(ws, f'A{c_}', '="是否含税：　"&IF($I$3="","（请在 I3 选择）",$I$3)', ff, fill('C5D9E8'), AL, border=SBOX)
    for r in (a, b, c_):
        for cc in 'BCD':
            ws[f'{cc}{r}'].border = SBOX
    tot_f = f'SUMIFS({W("V")},{W(member)},">0")' if kind == '全部' else f'SUMIFS({W("K")},{W(keycol)},">0")'
    paid_f = f'SUMIFS({W("W")},{W(member)},">0")' if kind == '全部' else f'-SUMIFS({W("L")},{W(keycol)},">0")'
    rows = [(a, '合计金额:', f'=ROUND({tot_f},2)', '上期结余:', '=$M$4'),
            (b, '已 付 款:', f'=ROUND({paid_f},2)', '本期笔数:', '=$M$6'),
            (c_, f'=IF($F${c_}<0,"应 收 款:","应 付 款:")' if kind == '应付' else f'=IF($F${c_}<0,"应 付 款:","应 收 款:")',
             f'=ROUND($M$4+F{a}-F{b},2)',
             {'全部': '', '应收': '应付方向余额:', '应付': '应收方向余额:'}[kind],
             None if kind == '全部' else
             f'=ROUND(SUMIFS({W("M")},{W("D")},$B$2,{W("E")},"{"应付" if kind == "应收" else "应收"}",{W("C")},"<="&$M$3),2)')]
    for r, l1, v1, l2, v2 in rows:
        cell(ws, f'E{r}', l1, ff, FL_NONE, AR, border=SBOX)
        cell(ws, f'F{r}', v1, ff, FL_NONE, AR, '¥#,##0.00;[Red]-¥#,##0.00;¥ -', SBOX)
        cell(ws, f'G{r}', l2, ff, FL_NONE, AR, border=SBOX)
        cell(ws, f'H{r}', v2, ff, FL_NONE, AR, INT if l2 == '本期笔数:' else '¥#,##0.00;[Red]-¥#,##0.00;¥ -', SBOX)
        cell(ws, f'I{r}', None, ff, FL_NONE, AR, border=SBOX)
    ws.merge_cells(f'A{d}:E{d}')
    until = 'IF($I$2="","今天",TEXT($I$2,"yyyy年m月d日"))'
    if kind == '应付':
        txt = f'=IF($F${c_}<0,"截止"&{until}&"，贵公司共欠我方款项：","截止"&{until}&"，我方共欠贵公司应付账款：")'
    else:
        txt = f'=IF($F${c_}<0,"截止"&{until}&"，我方共欠贵公司款项：","截止"&{until}&"，贵公司共欠我方应收账款：")'
    cell(ws, f'A{d}', txt, _f('宋体', 11, True, '000000'), fill('FFFF00'), AL, border=SBOX)
    ws.merge_cells(f'F{d}:I{d}')
    x = f'ABS(ROUND(N($F${c_}),2))'
    # 整数金额先把「零角零分」换成「整」，不然会变成「叁仟元零整」（原对账单公式就有这个毛病）
    ws['M1'] = (f'=IF({x}=0,"零元整",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TEXT(INT({x}),"[DBNum2]G/通用格式")&"元"&'
                f'TEXT(RIGHT(TEXT({x},"0.00"),2),"[DBNum2]0角0分"),"零角零分","整"),"零角",IF({x}<1,"","零")),"零分","整"))')
    ws['M1'].font = F_HELP
    cell(ws, f'F{d}', f'=IF(OR(ISNUMBER(SEARCH("通用格式",$M$1)),ISNUMBER(SEARCH("G/",$M$1)),ISERROR($M$1)),'
                      f'"人民币 "&TEXT({x},"¥#,##0.00"),$M$1)',
         _f('宋体', 12, True, '000000'), fill('F8CBCB'), AC, border=SBOX)
    for cc in 'BCDE':
        ws[f'{cc}{d}'].border = SBOX
    for cc in 'GHI':
        ws[f'{cc}{d}'].border = SBOX
    ws.row_dimensions[d].height = 26
    ws.freeze_panes = 'A6'
    ws.print_area = f'A1:I{d}'
    ws.print_title_rows = '1:5'
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.5


# ─────────────────────── ③ 往来对账单明细 / 应付账款明细 / 应收应付汇总表 ───────────────────────
def _names(r, first):
    k = f'ROW()-{first - 1}'
    return f'=IF({k}>往来业务明细!$Q${WR1},"",INDEX({W("D")},MATCH({k},{W("Q")},0)))'


def summary_detail(wb):
    ws = wb['往来对账单明细']
    wipe(ws)
    ws.sheet_properties.tabColor = S_TEAL
    title(ws, '往 来 对 账 单 明 细（每家单位收付全口径 · 与【往来对账单】同口径）', 'A', 'J')
    p = period(ws, 2, 'A', month=False, year=2026, open_ended=True)
    note(ws, '金额＝我方应收的业务（销售、仓储/装卸/周转费、物料、筐子、借出），负数＝我方应付（采购对方的果品/物料等）；'
             '已付款＝对方付给我方的钱（货款、定金、押金），负数＝我方付给对方（退定金、付货款、借给对方）。'
             '期末余额＝期初＋金额－已付款：正数＝对方欠我方，负数＝我方欠对方。年份/日期留空＝不限；期末余额算到截止日。', 'A', 'J', 3, 42)
    header(ws, 5, 'A', ['序号', '往来单位', '期初余额', '本期金额', '本期已付款', '期末余额', '其中：应收余额',
                        '其中：应付余额\n（定金/押金/货款）', '余额方向', '提示'])
    ws.row_dimensions[5].height = 36
    F0, F1 = 7, 206
    s, e = p['start'], p['end']
    for r in range(F0, F1 + 1):
        n = f'$B{r}'
        f = {
            'A': f'=IF($B{r}="","",ROW()-{F0 - 1})',
            'B': _names(r, F0),
            'C': f'=IF({n}="","",ROUND(SUMIFS({W("V")},{W("D")},{n},{W("C")},">0",{W("C")},"<"&{s})'
                 f'-SUMIFS({W("W")},{W("D")},{n},{W("C")},">0",{W("C")},"<"&{s}),2))',
            'D': f'=IF({n}="","",ROUND(SUMIFS({W("V")},{W("D")},{n},{in_period(W("C"), p)})'
                 f'+SUMIFS({W("V")},{W("D")},{n},{W("C")},0),2))',
            'E': f'=IF({n}="","",ROUND(SUMIFS({W("W")},{W("D")},{n},{in_period(W("C"), p)})'
                 f'+SUMIFS({W("W")},{W("D")},{n},{W("C")},0),2))',
            'F': f'=IF({n}="","",ROUND(C{r}+D{r}-E{r},2))',
            'G': f'=IF({n}="","",ROUND(SUMIFS({W("M")},{W("D")},{n},{W("E")},"应收",{W("C")},"<="&{e}),2))',
            'H': f'=IF({n}="","",ROUND(SUMIFS({W("M")},{W("D")},{n},{W("E")},"应付",{W("C")},"<="&{e}),2))',
            'I': f'=IF({n}="","",IF(F{r}>0,"对方欠我方",IF(F{r}<0,"我方欠对方","已结清")))',
            'J': f'=IF({n}="","",IF(ROUND(G{r}-H{r}-F{r},2)<>0,"✘ 口径不符，请核对",IF(OR(G{r}<0,H{r}<0),'
                 f'"有退无收或预收预付：可能是记账前的业务，可在【对账补充行】补期初","")))',
        }
        for c, v in f.items():
            cell(ws, f'{c}{r}', v, F_AUTOB if c == 'B' else F_AUTO, FL_AUTO,
                 AL if c in 'BJ' else (AR if c in 'CDEFGH' else AC), MONEY if c in 'CDEFGH' else None)
    for rr, lab, fl, fl_l, fl_n in ((6, '合　计', FL_TOT, F_TOT, F_TOTN), (F1 + 1, '合　计', FL_TOT, F_TOT, F_TOTN)):
        ws.merge_cells(f'A{rr}:B{rr}')
        cell(ws, f'A{rr}', lab, fl_l, fl, AC); ws[f'B{rr}'].border = BOX
        for c in 'CDEFGH':
            cell(ws, f'{c}{rr}', f'=ROUND(SUM({c}{F0}:{c}{F1}),2)', fl_n, fl, AR, MONEY)
        cell(ws, f'I{rr}', f'=COUNTIF(I{F0}:I{F1},"对方欠我方")&" 家欠我方，"&COUNTIF(I{F0}:I{F1},"我方欠对方")&" 家我方欠"',
             fl_n, fl, AC)
        cell(ws, f'J{rr}', None, fl_n, fl, AC)
    widths(ws, {'A': 6, 'B': 16, 'C': 14, 'D': 14, 'E': 14, 'F': 15, 'G': 14, 'H': 16, 'I': 22, 'J': 46})
    ws.freeze_panes = 'C7'
    return p


def payables(wb):
    ws = wb['应付账款明细']
    wipe(ws)
    title(ws, '应 付 账 款 明 细（采购 · 公司购买果品 · 借入 · 定金押金）', 'A', 'G')
    p = period(ws, 2, 'A', month=False, year=2026, open_ended=True)
    note(ws, '只列有应付业务的单位。本期发生＝采购、公司购买果品、借入、收的定金押金；本期已付＝我方付出去的（货款、还款、退定金）。'
             '应付余额算到截止日，负数＝多付/预付，或记账前的业务没补。逐笔看【应付账款对账单】，全口径看【往来对账单明细】。', 'A', 'G', 3, 36)
    header(ws, 5, 'A', ['序号', '往来单位', '期初应付', '本期发生', '本期已付', '应付余额', '说明'])
    F0, F1 = 7, 106
    # 辅助：全部往来单位 → 有应付业务的才列
    for r in range(F0, F0 + 200):
        k = r - F0 + 1
        ws[f'L{r}'] = f'=IF({k}>往来业务明细!$Q${WR1},"",INDEX({W("D")},MATCH({k},{W("Q")},0)))'
        ws[f'M{r}'] = f'=IF(L{r}="",0,IF(COUNTIFS({W("D")},L{r},{W("E")},"应付")>0,1,0))'
        ws[f'N{r}'] = f'=N(N{r - 1})+M{r}' if r > F0 else f'=M{r}'
        for c in 'LMN':
            ws[f'{c}{r}'].font = F_HELP
    last = f'$N${F0 + 199}'
    hide(ws, 'L', 'M', 'N')
    s, e = p['start'], p['end']
    for r in range(F0, F1 + 1):
        n = f'$B{r}'
        k = f'ROW()-{F0 - 1}'
        f = {
            'A': f'=IF($B{r}="","",ROW()-{F0 - 1})',
            'B': f'=IF({k}>{last},"",INDEX($L${F0}:$L${F0 + 199},MATCH({k},$N${F0}:$N${F0 + 199},0)))',
            'C': f'=IF({n}="","",ROUND(SUMIFS({W("M")},{W("D")},{n},{W("E")},"应付",{W("C")},">0",{W("C")},"<"&{s}),2))',
            'D': f'=IF({n}="","",ROUND(SUMIFS({W("K")},{W("D")},{n},{W("E")},"应付",{in_period(W("C"), p)})'
                 f'+SUMIFS({W("K")},{W("D")},{n},{W("E")},"应付",{W("C")},0),2))',
            'E': f'=IF({n}="","",-ROUND(SUMIFS({W("L")},{W("D")},{n},{W("E")},"应付",{in_period(W("C"), p)})'
                 f'+SUMIFS({W("L")},{W("D")},{n},{W("E")},"应付",{W("C")},0),2))',
            'F': f'=IF({n}="","",ROUND(C{r}+D{r}-E{r},2))',
            'G': f'=IF({n}="","",IF(F{r}>0,"我方还欠对方",IF(F{r}<0,"多付/预付，或记账前的业务未补","已结清")))',
        }
        for c, v in f.items():
            cell(ws, f'{c}{r}', v, F_AUTOB if c == 'B' else F_AUTO, FL_AUTO,
                 AL if c in 'BG' else (AR if c in 'CDEF' else AC), MONEY if c in 'CDEF' else None)
    for rr, fl, fl_l, fl_n in ((6, FL_TOT, F_TOT, F_TOTN), (F1 + 1, FL_TOT, F_TOT, F_TOTN)):
        ws.merge_cells(f'A{rr}:B{rr}')
        cell(ws, f'A{rr}', '合　计', fl_l, fl, AC); ws[f'B{rr}'].border = BOX
        for c in 'CDEF':
            cell(ws, f'{c}{rr}', f'=ROUND(SUM({c}{F0}:{c}{F1}),2)', fl_n, fl, AR, MONEY)
        cell(ws, f'G{rr}', f'=COUNTIF(B{F0}:B{F1},"?*")&" 家单位"', fl_n, fl, AC)
    widths(ws, {'A': 6, 'B': 18, 'C': 14, 'D': 14, 'E': 14, 'F': 15, 'G': 34, 'H': 12, 'I': 5, 'J': 12})
    ws.freeze_panes = 'C7'


def ar_ap_summary(wb):
    ws = wb['应收应付汇总表']
    wipe(ws)
    title(ws, '应 收 应 付 汇 总 表（每家单位分方向核对）', 'A', 'L')
    p = period(ws, 2, 'A', month=False, year=2026, open_ended=True)
    note(ws, '应收＝销售、仓储/装卸/周转费、物料、筐子、借出；应付＝采购、公司购买果品、借入、定金押金。'
             '已收/已付取【资金日记账】和【对账补充行】。余额算到截止日。净往来＝应收余额－应付余额，'
             '与【往来对账单明细】的期末余额一致。', 'A', 'L', 3, 36)
    header(ws, 5, 'A', ['序号', '往来单位', '期初应收', '本期应收发生', '本期已收', '应收余额', '期初应付', '本期应付发生',
                        '本期已付', '应付余额', '净往来\n（应收－应付）', '提示'])
    F0, F1 = 7, 206
    s = p['start']
    for r in range(F0, F1 + 1):
        n = f'$B{r}'
        def side(d, col):
            base = f'{W("D")},{n},{W("E")},"{d}"'
            return (f'SUMIFS({W(col)},{base},{in_period(W("C"), p)})+SUMIFS({W(col)},{base},{W("C")},0)')
        f = {
            'A': f'=IF($B{r}="","",ROW()-{F0 - 1})',
            'B': _names(r, F0),
            'C': f'=IF({n}="","",ROUND(SUMIFS({W("M")},{W("D")},{n},{W("E")},"应收",{W("C")},">0",{W("C")},"<"&{s}),2))',
            'D': f'=IF({n}="","",ROUND({side("应收", "K")},2))',
            'E': f'=IF({n}="","",-ROUND({side("应收", "L")},2))',
            'F': f'=IF({n}="","",ROUND(C{r}+D{r}-E{r},2))',
            'G': f'=IF({n}="","",ROUND(SUMIFS({W("M")},{W("D")},{n},{W("E")},"应付",{W("C")},">0",{W("C")},"<"&{s}),2))',
            'H': f'=IF({n}="","",ROUND({side("应付", "K")},2))',
            'I': f'=IF({n}="","",-ROUND({side("应付", "L")},2))',
            'J': f'=IF({n}="","",ROUND(G{r}+H{r}-I{r},2))',
            'K': f'=IF({n}="","",ROUND(F{r}-J{r},2))',
            'L': f'=IF({n}="","",IF(OR(F{r}<0,J{r}<0),"预收/预付或历史单据待补",""))',
        }
        for c, v in f.items():
            cell(ws, f'{c}{r}', v, F_AUTOB if c == 'B' else F_AUTO, FL_AUTO,
                 AL if c in 'BL' else (AR if c in 'CDEFGHIJK' else AC), MONEY if c in 'CDEFGHIJK' else None)
    for rr, fl, fl_l, fl_n in ((6, FL_TOT, F_TOT, F_TOTN), (F1 + 1, FL_TOT, F_TOT, F_TOTN)):
        ws.merge_cells(f'A{rr}:B{rr}')
        cell(ws, f'A{rr}', '合　计', fl_l, fl, AC); ws[f'B{rr}'].border = BOX
        for c in 'CDEFGHIJK':
            cell(ws, f'{c}{rr}', f'=ROUND(SUM({c}{F0}:{c}{F1}),2)', fl_n, fl, AR, MONEY)
        cell(ws, f'L{rr}', None, fl_n, fl, AC)
    widths(ws, {'A': 6, 'B': 16, 'C': 13, 'D': 14, 'E': 13, 'F': 14, 'G': 13, 'H': 14, 'I': 13, 'J': 14, 'K': 15,
                'L': 26})
    ws.freeze_panes = 'C7'


def stock_vs_debt(wb):
    ws = wb['库存价值与欠款比对']
    for r in range(6, 66):
        ws[f'I{r}'] = (f'=IF($B{r}="","",ROUND(IFERROR(INDEX(往来对账单明细!$F$7:$F$206,'
                       f'MATCH($B{r},往来对账单明细!$B$7:$B$206,0)),0),2))')
        ws[f'K{r}'] = (f'=IF($B{r}="","",IF(N($I{r})<0,"我方欠对方 · 不用催收",IF(N($I{r})=0,"没有欠款",'
                       f'IF(N($J{r})>=0,"货够抵账 · 可以先不收","货不够抵 · 该收钱了"))))')
        ws[f'L{r}'] = (f'=IF($B{r}="","",$B{r}&" 库里还压着 "&TEXT(N($H{r}),"#,##0.00")&" 元的货，"&IF(N($I{r})<0,'
                       f'"我方还欠对方 "&TEXT(-N($I{r}),"#,##0.00")&" 元","欠我们 "&TEXT(N($I{r}),"#,##0.00")&" 元，差额 "'
                       f'&TEXT(N($J{r}),"#,##0.00")&" 元"))')
    # 正文统一成前面老表的样子：公式格浅绿、手工格浅蓝、细灰格线、18 行高
    for row in ws.iter_rows(min_row=6, max_row=ws.max_row, max_col=12):
        for c in row:
            if c.row in (66, 67, 68) or c.row > 469:
                continue
            fg = c.fill.fgColor.rgb if c.fill.fill_type == 'solid' else None
            if fg in ('FFFFF2CC', 'FFDDEBF7'):
                c.fill = FL_IN; c.font = F_IN
            elif fg in ('FF2F5597', 'FF2E75B6', 'FFD6E4F0', 'FF9DC3E6'):
                continue
            else:
                c.fill = FL_AUTO; c.font = F_AUTOB if c.column_letter == 'B' else F_AUTO
            c.border = BOX
        if row[0].row not in (66, 67, 68):
            ws.row_dimensions[row[0].row].height = 18
    for r in (66, 469):
        for c in range(1, 13):
            x = ws.cell(r, c)
            if x.fill.fill_type == 'solid':
                x.fill = FL_TOT; x.border = BOX
    for r in range(6, 66):
        ws[f'L{r}'].alignment = AL.__class__(horizontal='left', vertical='center', wrap_text=False)
    ws.column_dimensions['L'].width = 58
    ws.row_dimensions[2].height = 30
    ws['I5'] = ('="应收款余额"&CHAR(10)&"（往来对账单明细·截至"&IF(往来对账单明细!$J$2>=2958465,"全部",'
                'TEXT(往来对账单明细!$J$2,"yyyy/m/d"))&"）"')
    ws['I5'].alignment = AC
    ws.row_dimensions[5].height = 36
    ws['A2'] = ('★ 下半部分「单价」那一列是淡黄色的手工格，你按等级填个心里价，库存价值自动算。上半部分按客户把库存价值和应收款摆在一起：'
                '差额＝库存价值−应收款。应收款余额取【往来对账单明细】的期末余额（收付全口径，截止日跟那张表走）。')
    # 标题、表头统一成前面的样式
    for m in list(ws.merged_cells.ranges):
        if str(m) == 'A1:J1':
            ws.unmerge_cells('A1:J1')
    title(ws, '库 存 价 值 与 欠 款 比 对', 'A', 'L')
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for c in row:
            if c.fill.fill_type == 'solid' and c.fill.fgColor.rgb in ('FF2F5597',):
                c.fill = FL_HDR
                c.font = F_HDR
            elif c.fill.fill_type == 'solid' and c.fill.fgColor.rgb in ('FFD6E4F0',):
                c.fill = FL_SEC
                c.font = F_SEC


def apply(wb):
    fix_detail(wb)
    # 对账补充行：方向下拉扩到 300 行
    ws = wb['对账补充行']
    for dv in ws.data_validations.dataValidation:
        if str(dv.sqref) == 'L4:L103':
            dv.sqref = type(dv.sqref)('L4:L303')
    # 对接源_02物料：起始月份 T8（2026-08-01）和标签 S8 原来藏在合并格 S6:T9 里面，openpyxl 会把它们丢掉
    # （B 列 24 个月的月份全靠它），这里拆开合并格，把起始月份放成看得见的格子
    ws = wb['对接源_02物料']
    import datetime as _dt
    if 'S6:T9' in [str(m) for m in ws.merged_cells.ranges]:
        ws.unmerge_cells('S6:T9')
        ws.merge_cells('S6:T7')
    ws['S8'] = '起始月份'
    ws['S8']._style = copy.copy(ws['S5']._style)
    ws['T8'] = _dt.datetime(2026, 8, 1)
    cell(ws, 'T8', None, F_IN, FL_IN, AC, 'yyyy-mm')
    # 对接源_02物料 里对往来业务明细的整区引用跟着扩
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and str(OLD_END) in c.value:
                c.value = c.value.replace(f'${OLD_END}', f'${WR1}')
            if isinstance(c.value, str) and 'INDIRECT("\'往来业务明细\'!' in c.value:
                # INDIRECT 每次编辑都全表重算，换成普通引用
                c.value = re.sub(r'INDIRECT\("\'往来业务明细\'!(\$[A-Z]+\$\d+:\$[A-Z]+\$\d+)"\)', r'往来业务明细!\1', c.value)
    statement(wb, '往来对账单', '全部')
    statement(wb, '应收账款对账单', '应收')
    statement(wb, '应付账款对账单', '应付')
    summary_detail(wb)
    payables(wb)
    ar_ap_summary(wb)
    stock_vs_debt(wb)
