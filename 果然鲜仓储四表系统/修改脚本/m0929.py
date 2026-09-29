# -*- coding: utf-8 -*-
"""0929 这一轮（在 0928 成品上改）：
   ① 《03》借款汇总：业务类型 / 费用项目是「家用」「国外」的行不提取（它们各有自己的汇总表）；
   ② 库存价值与欠款比对：上下两部分都改成「结余数量 ＋ 结余净重」，
      结余净重＝结余数量 × 单箱净重（单箱净重取这一货主＋品种＋等级入库的平均件重＝入库总重 ÷ 入库数量），
      库存价值＝结余净重 × 单价（元/KG）。
      原来的「结余总重KG」是入库总重 − 出库总重，原料出库没填件重时出库重量是 0，
      就会出现结余数量 0、结余总重却还有几千 KG 的情况，所以不再用它算价值。
   《01》库存等级接口加一列 I「单箱净重KG」给《03》跨文件取。"""
import re
from common0928 import *

INV_N = 400                     # 库存等级接口 / 对接源_01库存 4～403 行
LV_D0, LV_D1, LV_T1 = 6, 65, 66   # 库存价值与欠款比对：上半部分
LV_H2, LV_E0, LV_E1, LV_T2 = 68, 69, 468, 469   # 下半部分


def iface_01(wb):
    """《01》库存等级接口：I 列 单箱净重KG＝入库总重 ÷ 入库数量（库存结余 H ÷ G）"""
    ws = wb['库存等级接口']
    ws['I3'] = '单箱净重KG'
    ws['I3']._style = ws['H3']._style
    for r in range(4, 4 + INV_N):
        ws[f'I{r}'] = (f'=IF(库存结余!$B{r}="","",IF(N(库存结余!$G{r})>0,'
                       f'ROUND(N(库存结余!$H{r})/N(库存结余!$G{r}),3),""))')
        ws[f'I{r}']._style = ws[f'G{r}']._style
        ws[f'I{r}'].number_format = '0.000'
    ws.column_dimensions['I'].width = 12
    a2 = ws['A2'].value or ''
    if '单箱净重' not in a2:
        ws['A2'] = a2 + ' I 列「单箱净重KG」＝这一行入库总重 ÷ 入库数量（平均件重），《03》用它算结余净重。'


def loans_03(wb):
    """借款汇总：R 列（类别）遇到业务类型 / 费用项目是 家用、国外 的一律空着"""
    ws = wb['借款汇总']
    n = 0
    for r in range(JR0, JR1 + 1):
        v = ws[f'R{r}'].value
        if not (isinstance(v, str) and v.startswith('=IF(资金日记账!$C')):
            continue
        F, G = f'资金日记账!$F{r}', f'资金日记账!$G{r}'
        ws[f'R{r}'] = (f'=IF(OR({F}="家用",{F}="国外",{G}="家用",{G}="国外"),"",' + v[1:] + ')')
        n += 1
    assert n == JR1 - JR0 + 1, n
    a3 = ws['A3'].value or ''
    if '家用、国外' not in a3:
        ws['A3'] = a3.replace('记账以前就有的借款', '业务类型是「家用」「国外」的不算（在家用、国外费用汇总里看）。记账以前就有的借款', 1)


def stock_03(wb):
    # 对接源_01库存：K 列取《01》库存等级接口 I 列（单箱净重）
    src = wb['对接源_01库存']
    src['K3'] = '单箱净重KG'
    src['K3']._style = src['H3']._style
    m = re.match(r"=IFERROR\(IF\((\[\d+\])库存等级接口!", src['H4'].value)
    link = m.group(1)
    for i in range(INV_N):
        r = 4 + i
        ref = f'{link}库存等级接口!$I${r}'
        src[f'K{r}'] = f'=IFERROR(IF({ref}=0,"",{ref}),"")'
        src[f'K{r}']._style = src[f'H{r}']._style
        src[f'K{r}'].number_format = '0.000'
    src.column_dimensions['K'].width = 12

    ws = wb['库存价值与欠款比对']
    # 下半部分：G＝结余净重（数量×单箱净重），K＝单箱净重，H＝净重×单价
    ws[f'G{LV_H2}'] = '结余净重KG\n（数量×单箱净重）'
    ws[f'I{LV_H2}'] = '单价（手工填）\n元/KG'
    ws[f'K{LV_H2}'] = '单箱净重KG'
    ws[f'K{LV_H2}']._style = ws[f'J{LV_H2}']._style
    ws.row_dimensions[LV_H2].height = 32
    for r in range(LV_E0, LV_E1 + 1):
        sr = r - (LV_E0 - 4)
        ws[f'K{r}'] = f'=IF($B{r}="","",N(对接源_01库存!$K{sr}))'
        ws[f'K{r}']._style = ws[f'F{r}']._style
        ws[f'K{r}'].number_format = '0.000'
        ws[f'G{r}'] = f'=IF($B{r}="","",ROUND(N($F{r})*N($K{r}),2))'
        ws[f'H{r}'] = f'=IF($B{r}="","",ROUND(N($G{r})*N($I{r}),2))'
        ws[f'L{r}'] = (f'=IF($B{r}="","",IF(N($I{r})=0,"← 这一行还没填单价",'
                       f'IF(AND(N($F{r})<>0,N($K{r})=0),"← 没有入库件重，净重算不出","")))')
    ws[f'K{LV_T2}']._style = ws[f'J{LV_T2}']._style
    # 上半部分：E 改成结余净重（按客户合计下半部分的净重），跟下半部分一个口径
    ws[f'E{LV_D0 - 1}'] = '结余净重KG'
    for r in range(LV_D0, LV_D1 + 1):
        ws[f'E{r}'] = (f'=IF($B{r}="","",ROUND(SUMIF($B${LV_E0}:$B${LV_E1},$B{r},'
                       f'$G${LV_E0}:$G${LV_E1}),2))')
    ws['A2'] = ('★ 下半部分「单价」是淡黄色手工格，按等级填每公斤的心里价：库存价值＝结余净重 × 单价；'
                '结余净重＝结余数量 × 单箱净重（单箱净重取这一货主＋品种＋等级入库的平均件重）。'
                '上半部分按客户把结余数量、结余净重、库存价值和应收款摆在一起：差额＝库存价值−应收款。'
                '应收款余额取【往来对账单明细】的期末余额（收付全口径，截止日跟那张表走）。')


def apply_01(wb):
    iface_01(wb)


def apply_03(wb):
    loans_03(wb)
    stock_03(wb)
