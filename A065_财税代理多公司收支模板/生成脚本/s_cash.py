# -*- coding: utf-8 -*-
"""【资金台帐】混合录入（流水整块粘贴 ＋ 手工记）→ 即时余额；往来单位、收支项目、所属公司自动认。"""
from openpyxl.formatting.rule import FormulaRule
from common import *

# 往来对账单的选择格（s_ar 里用同样的地址）
ST_CO, ST_PARTY, ST_S, ST_E = f"{SH_STMT}!$C$3", f"{SH_STMT}!$F$3", f"{SH_STMT}!$R$3", f"{SH_STMT}!$R$4"


def _pt(col):
    return pr(col)


def build_cash(wb, ctx):
    ws = wb.create_sheet(SH_CASH)
    widths(ws, {'A': 6, 'B': 11, 'C': 17, 'D': 12, 'E': 12, 'F': 13, 'G': 20, 'H': 28, 'I': 18, 'J': 14,
                'K': 16, 'L': 11, 'M': 14, 'N': 11, 'O': 10, 'P': 26, 'Q': 8, 'R': 16, 'S': 14, 'T': 10, 'U': 9,
                'V': 14, 'W': 9, 'X': 30})
    title(ws, '资 金 台 帐（流水粘贴 ＋ 手工记 · 即时余额 · 往来单位和收支项目自动认）', J_LAST_VIS, C_CASH,
          '💡 银行流水：从网银导出的表里选中明细行的「交易时间～摘要」8 列（农行就是 A～H 列）复制，粘到这里 C 列最下面的空行，'
          '再在 B 列把这几行的账户选上（选一格往下拖）。微信、支付宝、其他银行的格式不一样，先到【流水格式转换】转一下再粘。'
          '现金、个人代付等手工记的：B 账户、C 日期、D/E 金额、H 对方户名（或 K 直接选往来单位）、J 摘要填上就行。'
          '右边灰色全自动：往来单位按【往来单位】里的名称/别名/账号认，收支项目按「手工 → 自家公司 → 绑定 → 摘要关键词 → 客户首笔新增、以后续费」的顺序定。'
          '认不出来的在 X 列提示，到 K、L 两列手工选一下就行（社保挂靠、费用类的就在这里手工筛选）。')
    # 第 3 行：汇总条
    put(ws, 'A3', '待处理：', F_KPI_L, fill('FFD9E1F2'), align=AR)
    put(ws, 'B3', f'=COUNTIF({jr(J_CHK)},"✗*")&" 笔"', F_KPI_V, fill('FFD9E1F2'), align=AC)
    ws.merge_cells('C3:D3')
    put(ws, 'C3', '已记笔数：', F_KPI_L, fill('FFD9E1F2'), align=AR)
    put(ws, 'E3', f'=COUNT({jr(J_DATE)})', F_KPI_V, fill('FFD9E1F2'), INT, AC)
    put(ws, 'F3', '最后日期：', F_KPI_L, fill('FFD9E1F2'), align=AR)
    put(ws, 'G3', f'=IF(COUNT({jr(J_DATE)})=0,"",MAX({jr(J_DATE)}))', F_KPI_V, fill('FFD9E1F2'), DATE, AC)
    put(ws, 'H3', f'="容量 "&TEXT(COUNTA({jr(J_TIME)})/{J_R1 - J_R0 + 1},"0%")&"（快满了：在倒数几行中间右键插入行，再把上一行复制下来）"', F_NOTE,
        fill('FFD9E1F2'), align=AL)
    put(ws, 'I3', '全部账户余额：', F_KPI_L, fill('FFD9E1F2'), align=AR)
    put(ws, 'J3', f'=ROUND(SUM({AC_OPENS})+SUM({jr(J_NET)}),2)', F_KPI_V, fill('FFD9E1F2'), MONEY, AC)
    # 第 4 行：分区色带
    bands = [('A', 'B', '账户', 'FF305496'), ('C', 'J', '← 流水粘贴区（列顺序跟农行网银导出一样）', 'FF5B9BD5'),
             ('K', 'M', '手工调整（认错了/认不出才填）', 'FFBF8F00'), ('N', 'X', '自动（不用填）', 'FF7F7F7F')]
    for c1, c2, t, col in bands:
        section(ws, 4, c1, c2, t, col)
    heads = [(J_SEQ, '序号'), (J_ACC, '账户'), (J_TIME, '交易时间'), (J_IN, '收入金额'), (J_OUT, '支出金额'),
             (J_BANKBAL, '银行余额\n（流水带的）'), (J_OACCT, '对方账号'), (J_ONAME, '对方户名'), (J_OBANK, '对方开户行'),
             (J_MEMO, '摘要'), (J_MPARTY, '往来单位\n（手工改）'), (J_MITEM, '收支项目\n（手工改）'), (J_NOTE, '备注'),
             (J_DATE, '日期'), (J_CO, '所属公司'), (J_PARTY, '往来单位'), (J_PTYPE, '往来类型'), (J_CUS, '客户'),
             (J_SUP, '供应商'), (J_ITEM, '收支项目'), (J_CLS, '收支类别'), (J_BAL, '账户余额\n（即时）'),
             (J_BCHK, '银行余额\n核对'), (J_CHK, '校验')]
    for col, t in heads:
        c = {'A': 'FF305496', 'B': 'FF305496'}.get(col)
        if c is None:
            c = 'FF5B9BD5' if CI(col) <= CI('J') else ('FFBF8F00' if CI(col) <= CI('M') else 'FF7F7F7F')
        put(ws, f'{col}{J_HDR}', t, F_HDR, fill(c), align=ACW)
    ws.row_dimensions[J_HDR].height = 34
    for col, t in ((J_INV, '收入数'), (J_OUTV, '支出数'), (J_NET, '净额'), (J_AUTO, '自动匹配'), (J_DUPK, '查重键'),
                   (J_TODO, '待登记名'), (J_TODO1, '首次'), (J_TODOC, '计数'), (J_SKEY, '对账单键'), (J_KW, '关键词'),
                   (J_APF, '冲应付')):
        ws[f'{col}{J_HDR}'] = t
        ws[f'{col}{J_HDR}'].font = F_HELP

    # 常用区域
    PN, PTY, PB, PA1, PA2, PAC, POLD, PAP = (_pt(PT_NAME), _pt(PT_TYPE), _pt(PT_BIND), _pt(PT_N1), _pt(PT_N2), _pt(PT_ACCT),
                                             _pt(PT_OLD), _pt(PT_AP))
    PN0 = _pt(PT_N0)
    KW = br(KW_WORD, KW_R0, KW_R1)
    KWD = br(KW_DIR, KW_R0, KW_R1)
    KWI = br(KW_ITEM, KW_R0, KW_R1)
    R0 = J_R0
    for r in range(J_R0, J_R1 + 1):
        g = lambda c: f'{c}{r}'
        emp = f'AND({g(J_ACC)}="",{g(J_TIME)}="",{g(J_IN)}="",{g(J_OUT)}="")'
        f = {}
        f[J_SEQ] = f'=IF({emp},"",ROW()-{J_HDR})'
        f[J_INV] = f'={num(g(J_IN))}'
        f[J_OUTV] = f'={num(g(J_OUT))}'
        f[J_NET] = f'={g(J_INV)}-{g(J_OUTV)}'
        f[J_DATE] = f'={date_parse(g(J_TIME))}'
        f[J_CO] = f'=IF({g(J_ACC)}="","",IFERROR(INDEX({AC_COS},MATCH({g(J_ACC)},{AC_NAMES},0))&"",""))'
        nh = norm(g(J_ONAME))
        f[J_AUTO] = (f'=IF(AND({g(J_ONAME)}="",{g(J_OACCT)}=""),"",'
                     f'IFERROR(INDEX({PN},MATCH({nh},{PN0},0)),IFERROR(INDEX({PN},MATCH({nh},{PA1},0)),'
                     f'IFERROR(INDEX({PN},MATCH({nh},{PA2},0)),IFERROR(INDEX({CO_NAMES},MATCH({nh},{CO_NORMS},0)),'
                     f'IF({g(J_OACCT)}="","",IFERROR(INDEX({PN},MATCH({g(J_OACCT)}&"",{PAC},0)),"")))))))')
        f[J_PARTY] = f'=IF({g(J_MPARTY)}<>"",{g(J_MPARTY)},{g(J_AUTO)})'
        f[J_PTYPE] = (f'=IF({g(J_PARTY)}="","",IF(ISNUMBER(MATCH({g(J_PARTY)},{CO_NAMES},0)),"内部公司",'
                      f'IFERROR(INDEX({PTY},MATCH({g(J_PARTY)},{PN},0))&"","未登记")))')
        f[J_CUS] = f'=IF({g(J_PTYPE)}="客户",{g(J_PARTY)},"")'
        f[J_SUP] = f'=IF({g(J_PTYPE)}="供应商",{g(J_PARTY)},"")'
        text = f'{g(J_ONAME)}&" "&{g(J_MEMO)}&" "&{g(J_NOTE)}'
        f[J_KW] = (f'=IF({g(J_INV)}+{g(J_OUTV)}=0,"",IFERROR(INDEX({KWI},MATCH(1,INDEX(({KW}<>"")*ISNUMBER(SEARCH({KW},{text}))'
                   f'*(({KWD}="全部")+({KWD}=IF({g(J_INV)}>0,"收入","支出"))>0),0),0)),""))')
        bind = f'IFERROR(INDEX({PB},MATCH({g(J_PARTY)},{PN},0))&"","")'
        old = f'IFERROR(INDEX({POLD},MATCH({g(J_PARTY)},{PN},0))&"","")'
        prior = (f'COUNTIFS({jr(J_CO)},{g(J_CO)},{jr(J_PARTY)},{g(J_PARTY)},{jr(J_INV)},">0",{jr(J_DATE)},"<"&{g(J_DATE)})'
                 f'+COUNTIFS(${J_CO}${R0}:{g(J_CO)},{g(J_CO)},${J_PARTY}${R0}:{g(J_PARTY)},{g(J_PARTY)},'
                 f'${J_INV}${R0}:{g(J_INV)},">0",${J_DATE}${R0}:{g(J_DATE)},{g(J_DATE)})-1')
        f[J_ITEM] = (f'=IF({g(J_MITEM)}<>"",{g(J_MITEM)},IF(OR({g(J_ACC)}="",{g(J_INV)}+{g(J_OUTV)}=0),"",'
                     f'IF({g(J_PTYPE)}="内部公司",IF({g(J_PARTY)}={g(J_CO)},"账户互转","内部划转"),'
                     f'IF({bind}<>"",{bind},IF({g(J_KW)}<>"",{g(J_KW)},'
                     f'IF(AND({g(J_PTYPE)}="客户",{g(J_INV)}>0,{g(J_DATE)}<>""),'
                     f'IF(OR({old}="是",{prior}>0),"续费","新增"),""))))))')
        f[J_CLS] = f'=IF({g(J_ITEM)}="","",IFERROR(INDEX({IT_CLSS},MATCH({g(J_ITEM)},{IT_NAMES},0))&"","？"))'
        # 冲应付：成本费用 / 投资类的付款，且对方是「计应付」的单位（供应商、J 列填了是、或者有计应付的进项发票）
        apf = f'IFERROR(INDEX({PAP},MATCH({g(J_PARTY)},{PN},0))&"","")'
        f[J_APF] = (f'=IF(AND(OR({g(J_CLS)}="{CLS_VAR}",{g(J_CLS)}="{CLS_FIX}",{g(J_CLS)}="{CLS_OTH}",{g(J_CLS)}="{CLS_INVT}"),'
                    f'{g(J_PARTY)}<>"",{apf}<>"否"),IF(OR({g(J_PTYPE)}="供应商",{apf}="是",'
                    f'COUNTIFS({vr(V_PARTY)},{g(J_PARTY)},{vr(V_USE)},"应付")>0),1,0),0)')
        f[J_BAL] = (f'=IF({g(J_ACC)}="","",ROUND(IFERROR(INDEX({AC_OPENS},MATCH({g(J_ACC)},{AC_NAMES},0)),0)'
                    f'+SUMIFS(${J_NET}${R0}:{g(J_NET)},${J_ACC}${R0}:{g(J_ACC)},{g(J_ACC)}),2))')
        f[J_BCHK] = (f'=IF(OR({g(J_BANKBAL)}="",{g(J_ACC)}=""),"",IF(ABS({g(J_BAL)}-{num(g(J_BANKBAL))})<0.005,"✓",'
                     f'"差 "&TEXT({g(J_BAL)}-{num(g(J_BANKBAL))},"#,##0.00")))')
        f[J_DUPK] = (f'=IF({emp},"",{g(J_ACC)}&"|"&{g(J_TIME)}&"|"&{g(J_INV)}&"|"&{g(J_OUTV)}&"|"&LEFT({g(J_ONAME)},20)'
                     f'&"|"&{g(J_BANKBAL)})')
        f[J_TODO] = (f'=IF(AND({g(J_PARTY)}="",{g(J_ONAME)}<>"",{g(J_ITEM)}=""),{norm(g(J_ONAME))},'
                     f'IF({g(J_PTYPE)}="未登记",{g(J_PARTY)},""))')
        f[J_TODO1] = f'=IF({g(J_TODO)}="",0,IF(MATCH({g(J_TODO)},{jr(J_TODO)},0)={r - R0 + 1},1,0))'
        f[J_TODOC] = f'=N({J_TODOC}{r - 1})+{g(J_TODO1)}' if r > R0 else f'={g(J_TODO1)}'
        per = (f'AND({g(J_PARTY)}={ST_PARTY},{g(J_PARTY)}<>"",OR({ST_CO}="全部",{g(J_CO)}={ST_CO}),'
               f'{g(J_DATE)}>={ST_S},{g(J_DATE)}<={ST_E},{g(J_INV)}+{g(J_OUTV)}<>0,{g(J_CLS)}<>"账户互转")')
        f[J_SKEY] = f'=IF({per},{g(J_DATE)}*100000+50000+ROW(),"")'
        chk = (f'=IF({emp},"",IF({g(J_ACC)}="","✗ 没选账户",IF(COUNTIF({AC_NAMES},{g(J_ACC)})=0,"✗ 账户不在基础资料里",'
               f'IF({g(J_DATE)}="","✗ 日期看不懂（要像 2026/9/1 这样）",'
               f'IF(AND({g(J_INV)}<>0,{g(J_OUTV)}<>0),"✗ 收入、支出只能填一边",'
               f'IF({g(J_INV)}+{g(J_OUTV)}=0,"✗ 没填金额（或金额不是数字）",'
               f'IF({g(J_ITEM)}="",IF(AND({g(J_PARTY)}="",{g(J_ONAME)}<>""),"✗ 认不出：对方户名没登记，到【往来单位】加，或 K/L 列手工选",'
               f'"✗ 请在 L 列选收支项目"),'
               f'IF({g(J_CLS)}="？","✗ 收支项目不在基础资料清单里",'
               f'IF(LEFT({g(J_BCHK)},1)="差","⚠ 余额跟银行对不上（漏记或重复？）",'
               f'IF(COUNTIF(${J_DUPK}${R0}:{g(J_DUPK)},{g(J_DUPK)})>1,"⚠ 跟上面某行一模一样，是不是重复粘贴了",'
               f'IF({g(J_PTYPE)}="未登记","⚠ 往来单位没在【往来单位】登记（统计照算）",'
               f'IF({g(J_ITEM)}="新增","√ 新客户首笔","√"))))))))))))')
        f[J_CHK] = chk
        for col, v in f.items():
            ws[f'{col}{r}'] = v
    # 演示数据：你发来的农行 9 月流水，原样
    for i, row in enumerate(ctx['bank'][2]):
        r = J_R0 + i
        ws[f'{J_ACC}{r}'] = ctx['my_acc']
        for col, v in zip([J_TIME, J_IN, J_OUT, J_BANKBAL, J_OACCT, J_ONAME, J_OBANK, J_MEMO], row):
            ws[f'{col}{r}'] = v
    # 样式
    inp = [J_ACC, J_TIME, J_IN, J_OUT, J_BANKBAL, J_OACCT, J_ONAME, J_OBANK, J_MEMO]
    man = [J_MPARTY, J_MITEM, J_NOTE]
    auto = [J_SEQ, J_DATE, J_CO, J_PARTY, J_PTYPE, J_CUS, J_SUP, J_ITEM, J_CLS, J_BAL, J_BCHK, J_CHK]
    fmts = {J_IN: MONEY, J_OUT: MONEY, J_BANKBAL: MONEY, J_BAL: MONEY, J_DATE: DATE, J_OACCT: '@'}
    aligns = {J_ONAME: AL, J_OBANK: AL, J_MEMO: AL, J_MPARTY: AL, J_NOTE: AL, J_PARTY: AL, J_CUS: AL, J_SUP: AL,
              J_CHK: AL, J_IN: AR, J_OUT: AR, J_BANKBAL: AR, J_BAL: AR}
    style_rows(ws, J_R0, J_R1, inp + man + auto, auto=auto, fmts=fmts, aligns=aligns,
               fills={**{c: FILL_PASTE for c in inp}, **{c: FILL_IN for c in man}}, bold=[J_BAL])
    for col in (J_INV, J_OUTV, J_NET, J_AUTO, J_DUPK, J_TODO, J_TODO1, J_TODOC, J_SKEY, J_KW, J_APF):
        for r in range(J_R0, J_R1 + 1):
            ws[f'{col}{r}'].font = F_HELP
        ws.column_dimensions[col].hidden = True
    rng = f'{J_CHK}{J_R0}:{J_CHK}{J_R1}'
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(${J_CHK}{J_R0},1)="✗"'], fill=FILL_WARN,
                                                   font=Font(name=YH, sz=10, bold=True, color='FFC00000')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(${J_CHK}{J_R0},1)="⚠"'], fill=fill('FFFFEB9C')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(${J_CHK}{J_R0},1)="√"'],
                                                   font=Font(name=YH, sz=10, bold=True, color='FF00B050')))
    ws.conditional_formatting.add(f'{J_ITEM}{J_R0}:{J_ITEM}{J_R1}',
                                  FormulaRule(formula=[f'${J_ITEM}{J_R0}="新增"'], font=Font(name=YH, sz=10, bold=True, color='FFC00000')))
    dv_list(ws, f'{J_ACC}{J_R0}:{J_ACC}{J_R1}', f'={AC_NAMES}', '这笔走的哪个账户（在【基础资料】登记）')
    dv_list(ws, f'{J_MPARTY}{J_R0}:{J_MPARTY}{J_R1}', f'={SH_AUX}!$C$1:$C$620',
            '认错了或认不出才选；也可以直接打名字（没登记的会提示）', stop=False)
    dv_list(ws, f'{J_MITEM}{J_R0}:{J_MITEM}{J_R1}', f'={IT_NAMES}', '自动认的不对才选：比如把「续费」改成「代办」「项目」')
    ws.auto_filter.ref = f'A{J_HDR}:{J_CHK}{J_R1}'
    ws.freeze_panes = f'C{J_R0}'
    return ws
