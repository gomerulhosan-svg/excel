# -*- coding: utf-8 -*-
"""【流水格式转换】微信、支付宝、其他银行的导出格式跟农行不一样：整块粘进来，选格式，
   右边就排成资金台帐 B～J 列的顺序（账户 交易时间 收入 支出 余额 对方账号 对方户名 开户行 摘要），复制→到资金台帐「粘贴为数值」。"""
from common import *

FMT_R0, FMT_R1 = 6, 10          # 格式预设 5 行
RAW_HDR, RAW_R0, RAW_R1 = 14, 15, 1014     # 粘贴区 1000 行
RAW_COLS = [CL(i) for i in range(1, 17)]   # A..P 16 列
OUT_COLS = ['S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA']
H_DATE, H_IN, H_OUT, H_OK, H_CNT, H_SKIP = 'AC', 'AD', 'AE', 'AF', 'AG', 'AH'
FIELDS = ['交易时间', '收入金额', '支出金额', '单列金额', '收支标志', '收入标志', '支出标志', '银行余额', '对方账号',
          '对方户名', '对方开户行', '摘要 1', '摘要 2', '支付方式', '交易状态']
PRESETS = [('农行网银', 1, 2, 3, None, None, None, None, 4, 5, 6, 7, 8, None, None, None),
           ('微信账单', 1, None, None, 6, 5, '收入', '支出', None, None, 3, None, 2, 4, 7, 8),
           ('支付宝账单', 1, None, None, 7, 6, '收入', '支出', None, 4, 3, None, 5, 2, 8, 9),
           ('单列金额(正收负支)', 1, None, None, 2, None, None, None, 3, 4, 5, 6, 7, None, None, None),
           ('自定义', 1, 2, 3, None, None, None, None, 4, 5, 6, 7, 8, None, None, None)]


def build_conv(wb, ctx):
    ws = wb.create_sheet(SH_CONV)
    widths(ws, {**{c: 12 for c in RAW_COLS}, 'Q': 2, 'R': 2, 'S': 11, 'T': 18, 'U': 12, 'V': 12, 'W': 12, 'X': 18,
                'Y': 24, 'Z': 16, 'AA': 22})
    title(ws, '流 水 格 式 转 换（微信 · 支付宝 · 其他银行 → 资金台帐的列顺序）', 'AA', C_CONV,
          '💡 ① 在 B3 选格式、E3 选账户；② 把导出文件里的明细行（不要表头）整块粘到下面 A15 开始的粘贴区；'
          '③ 右边 S～AA 就排好了，选中有内容的行复制，到【资金台帐】B 列最下面的空行「选择性粘贴 → 数值」。'
          '农行网银的导出不用转，直接粘到资金台帐 C 列。别的银行列顺序不一样：在「自定义」那行填上每样东西在第几列。'
          '微信、支付宝里：零钱提现记成支出、充值记成收入（银行那边对应的一笔在资金台帐 L 列选「账户互转」）；'
          '用银行卡、信用卡付的（钱不是从零钱/余额出的，银行流水里已经有）和交易关闭的自动跳过，免得记两次；转入零钱通这类也跳过。')
    selector(ws, 'A3', '格式', 'B3', '微信账单', f'=$A${FMT_R0}:$A${FMT_R1}', prompt='导出文件是哪种格式')
    selector(ws, 'D3', '账户', 'E3', None, f'={AC_NAMES}', prompt='这批流水是哪个账户的')
    put(ws, 'G3', f'="识别出 "&MAX(${H_CNT}${RAW_R0}:${H_CNT}${RAW_R1})&" 笔，右边可以复制了"'
                  f'&IF(SUM(${H_SKIP}${RAW_R0}:${H_SKIP}${RAW_R1})>0,"；跳过 "&SUM(${H_SKIP}${RAW_R0}:${H_SKIP}${RAW_R1})'
                  f'&" 笔银行卡/信用卡付的（银行流水里已经有）或交易关闭的","")', F_RED, align=AL, border=False)
    section(ws, 4, 'A', 'P', '格式设置：每样东西在导出文件的第几列（从 1 数；没有就空着）。「自定义」那行可以改', C_CONV)
    for i, t in enumerate(['格式'] + FIELDS):
        put(ws, f'{CL(i + 1)}5', t, F_HDR, fill(C_CONV), align=ACW)
    ws.row_dimensions[5].height = 30
    for k, p in enumerate(PRESETS):
        r = FMT_R0 + k
        for i, v in enumerate(p):
            put(ws, f'{CL(i + 1)}{r}', v, F_IN if p[0] == '自定义' else F_AUTO,
                FILL_IN if p[0] == '自定义' else FILL_AUTO, align=AC)
    put(ws, 'A11', '当前用', F_TXTB, FILL_SUB, align=AC)
    for i in range(len(FIELDS)):
        c = CL(i + 2)
        put(ws, f'{c}11', f'=IFERROR(INDEX({c}${FMT_R0}:{c}${FMT_R1},MATCH($B$3,$A${FMT_R0}:$A${FMT_R1},0))&"","")',
            F_AUTOB, FILL_SUB, align=AC)
    # 粘贴区
    section(ws, 13, 'A', 'P', '粘贴区：导出文件的明细行整块粘在 A15（第 1 列对 A 列）', 'FF5B9BD5')
    section(ws, 13, 'S', 'AA', '→ 转好的：复制这里，到资金台帐 B 列「粘贴为数值」', C_CASH)
    for i, c in enumerate(RAW_COLS):
        put(ws, f'{c}{RAW_HDR}', f'第{i + 1}列', F_HDR, fill('FF5B9BD5'), align=AC)
    for i, t in enumerate(['账户', '交易时间', '收入金额', '支出金额', '银行余额', '对方账号', '对方户名', '对方开户行', '摘要']):
        put(ws, f'{OUT_COLS[i]}{RAW_HDR}', t, F_HDR, fill(C_CASH), align=ACW)
    for c, t in ((H_DATE, '日期'), (H_IN, '收'), (H_OUT, '支'), (H_OK, '有效'), (H_CNT, '计数'), (H_SKIP, '跳过')):
        ws[f'{c}{RAW_HDR}'] = t
        ws[f'{c}{RAW_HDR}'].font = F_HELP
    fld = {n: f'${CL(i + 2)}$11' for i, n in enumerate(FIELDS)}

    def cellof(r, key):
        c = fld[key]
        return f'IF(N(--({c}&"0"))=0,"",INDEX($A{r}:$P{r},1,--{c}))'

    for r in range(RAW_R0, RAW_R1 + 1):
        t = cellof(r, '交易时间')
        ws[f'{H_DATE}{r}'] = f'=IFERROR({date_parse(t)},"")'
        one = f'{num(cellof(r, "单列金额"))}'
        flag = f'TRIM({cellof(r, "收支标志")}&"")'
        single = f'N(--({fld["单列金额"]}&"0"))>0'
        hasflag = f'N(--({fld["收支标志"]}&"0"))>0'
        # 收/支是「/」「不计收支」的：零钱提现、提现 → 支出（钱转去银行卡），充值 → 收入；别的（转入零钱通等）跳过
        memo = f'({cellof(r, "摘要 1")}&" "&{cellof(r, "摘要 2")})'
        wd = f'ISNUMBER(SEARCH("提现",{memo}))'
        cz = f'ISNUMBER(SEARCH("充值",{memo}))'
        ws[f'{H_IN}{r}'] = (f'=IFERROR(IF({single},IF({hasflag},IF({flag}={fld["收入标志"]},ABS({one}),'
                            f'IF(AND({flag}<>{fld["支出标志"]},{cz},NOT({wd})),ABS({one}),0)),MAX({one},0)),'
                            f'{num(cellof(r, "收入金额"))}),0)')
        ws[f'{H_OUT}{r}'] = (f'=IFERROR(IF({single},IF({hasflag},IF({flag}={fld["支出标志"]},ABS({one}),'
                             f'IF(AND({flag}<>{fld["收入标志"]},{wd}),ABS({one}),0)),MAX(-{one},0)),'
                             f'{num(cellof(r, "支出金额"))}),0)')
        # 微信/支付宝里用银行卡、信用卡、花呗付的：钱不是从零钱/余额出的，银行流水里已经有这一笔，跳过免得记两次；交易关闭、失败的也跳过
        pay = f'TRIM({cellof(r, "支付方式")}&"")'
        payok = (f'OR({pay}="",{pay}="/",ISNUMBER(SEARCH("零钱",{pay})),ISNUMBER(SEARCH("余额",{pay})),{cz})')
        st = f'({cellof(r, "交易状态")}&"")'
        stok = f'NOT(OR(ISNUMBER(SEARCH("关闭",{st})),ISNUMBER(SEARCH("失败",{st}))))'
        has = f'AND(ISNUMBER({H_DATE}{r}),{H_IN}{r}+{H_OUT}{r}<>0)'
        ws[f'{H_OK}{r}'] = f'=IFERROR(IF(AND({has},{payok},{stok}),1,0),0)'
        ws[f'{H_SKIP}{r}'] = f'=IFERROR(IF(AND({has},{H_OK}{r}=0),1,0),0)'
        ws[f'{H_CNT}{r}'] = f'=N({H_CNT}{r - 1})+{H_OK}{r}' if r > RAW_R0 else f'={H_OK}{r}'
        for c in (H_DATE, H_IN, H_OUT, H_OK, H_CNT, H_SKIP):
            ws[f'{c}{r}'].font = F_HELP
        k = f'ROW()-{RAW_HDR}'
        src = f'MATCH({k},${H_CNT}${RAW_R0}:${H_CNT}${RAW_R1},0)+{RAW_HDR}'

        def pick(key):
            c = fld[key]
            return f'IF(N(--({c}&"0"))=0,"",INDEX($A${RAW_R0}:$P${RAW_R1},{src}-{RAW_HDR},--{c}))'
        ok = f'{k}<=MAX(${H_CNT}${RAW_R0}:${H_CNT}${RAW_R1})'

        def txt(key):
            # 微信、支付宝没有的栏目导出成「/」，当空的
            return f'IF(TRIM({pick(key)}&"")="/","",TRIM({pick(key)}&""))'
        out = [f'$E$3&""', pick('交易时间'), f'INDEX(${H_IN}${RAW_R0}:${H_IN}${RAW_R1},{src}-{RAW_HDR})',
               f'INDEX(${H_OUT}${RAW_R0}:${H_OUT}${RAW_R1},{src}-{RAW_HDR})', pick('银行余额'), txt('对方账号'),
               txt('对方户名'), txt('对方开户行'), f'TRIM({txt("摘要 1")}&" "&{txt("摘要 2")})']
        for i, v in enumerate(out):
            c = OUT_COLS[i]
            if i in (2, 3):
                ws[f'{c}{r}'] = f'=IF({ok},IF({v}=0,"",{v}),"")'
            else:
                ws[f'{c}{r}'] = f'=IF({ok},{v},"")'
    style_rows(ws, RAW_R0, RAW_R1, RAW_COLS, fills={c: FILL_PASTE for c in RAW_COLS})
    style_rows(ws, RAW_R0, RAW_R1, OUT_COLS, auto=OUT_COLS, fmts={'U': MONEY, 'V': MONEY, 'W': MONEY},
               aligns={'Y': AL, 'AA': AL, 'U': AR, 'V': AR, 'W': AR})
    hide(ws, H_DATE, H_IN, H_OUT, H_OK, H_CNT, H_SKIP)
    ws.freeze_panes = f'A{RAW_R0}'
    return ws
