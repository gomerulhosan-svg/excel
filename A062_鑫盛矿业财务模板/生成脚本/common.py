# -*- coding: utf-8 -*-
"""A062 鑫盛矿业财务模板 · 共用常量（表名 / 行列地址 / 容量）与样式。

这个文件就是各张表之间的「接口约定」：凡是一张表要引用另一张表的某一列、某一段行，
一律从这里取常量，不在各模块里手写地址，免得一处改了另一处没跟上。

整体思路（写给以后改这份脚本的人）：
- 每张录入表（现金流水、采购、领用、销售、工资、手工凭证）在自己右边的隐藏列里把这一行的
  「借贷腿」算好（科目编码 + 借额 + 贷额）；月末结转、折旧在【月末结转】表里算好腿。
- 【记账分录】按固定槽位把所有腿拼成一张总表（一行一条分录），凭证号、打印、科目余额表、明细账、报表都从这里取。
- 月末要用到「截至某月的余额」的计算（调汇、结转生产成本），一律直接对各录入表的腿求和，
  不对【记账分录】求和——否则会自己引用自己（循环引用）。
"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

# ─────────────────────────── 表名 ───────────────────────────
SH_HOME  = '首页'
SH_HELP  = '使用说明'
SH_BASE  = '基础资料'
SH_RULE  = '记账规则'
SH_RATE  = '月度汇率'
SH_ACC   = '科目表'
SH_CASH  = '现金流水总表'
SH_BUY   = '采购入库'
SH_ISS   = '领用出库'
SH_PROD  = '产量登记'
SH_SALE  = '销售结算'
SH_PAY   = '工资计提'
SH_FA    = '固定资产'
SH_MAN   = '手工凭证'
SH_MEND  = '月末结转'
SH_JE    = '记账分录'
SH_VLIST = '凭证汇总'
SH_VPRT  = '记账凭证'
SH_TB    = '科目余额表'
SH_GL    = '明细账'
SH_ACCT  = '资金账户余额表'
SH_ARAP  = '往来余额表'
SH_INV   = '存货收发存'
SH_ORE   = '原矿产销存'
SH_BAL   = '资产负债表'
SH_PL    = '利润表'
SH_CF    = '现金流量表'
SH_OPS   = '经营报表'
SH_CHK   = '数据校验'
SH_VIDX  = '_凭证索引'
SH_AUX   = '_辅助'

SHEET_ORDER = [SH_HOME, SH_HELP, SH_BASE, SH_RULE, SH_RATE, SH_ACC,
               SH_CASH, SH_BUY, SH_ISS, SH_PROD, SH_SALE, SH_PAY, SH_FA, SH_MAN,
               SH_MEND, SH_JE, SH_VLIST, SH_VPRT,
               SH_TB, SH_GL, SH_ACCT, SH_ARAP, SH_INV, SH_ORE,
               SH_BAL, SH_PL, SH_CF, SH_OPS, SH_CHK, SH_VIDX, SH_AUX]


def q(sh):
    """跨表引用时的表名写法：带数字开头、下划线开头或特殊字符的表名要加单引号"""
    return f"'{sh}'" if (sh[0].isdigit() or sh[0] == '_' or any(ch in sh for ch in ' -()（）')) else sh


def rng(sh, col, r0, r1):
    return f"{q(sh)}!${col}${r0}:${col}${r1}"


# ─────────────────────────── 期间 ───────────────────────────
N_MONTHS = 32                      # 2025-05 起 32 个月（到 2027-12）；月份序号 n = 1..32

# ─────────────────────────── 首页（全表共用的报表年度 / 月份） ───────────────────────────
HOME_YEAR, HOME_MONTH = f'{SH_HOME}!$C$4', f'{SH_HOME}!$F$4'

# ─────────────────────────── 基础资料 ───────────────────────────
# 参数 A:B（第 5 行起）
PARAMS = [  # (键, 名称, 定义名称)
    ('company',  '公司名称', '公司名称'),
    ('cur',      '本位币代码', '本位币'),
    ('curname',  '本位币名称', '本位币名'),
    ('start',    '建账起始月份', '起始月份'),
    ('golive',   '业务模块启用日', '启用日'),
    ('vat',      '增值税率', '增值税率'),
    ('vatin',    '销售单价是否含税', '单价含税'),
    ('vnoprint', '凭证号是否打印', '凭证号打印'),
    ('fxclose',  '外币兑换差额月末转汇兑损益', '兑换差额结转'),
    ('fxstart',  '兑换差额从哪个月开始转', '兑换结转起始月'),
    ('maker',    '制单人', '制单人'),
    ('boss',     '会计主管', '会计主管'),
    ('cashier',  '出纳', '出纳'),
    ('checker',  '复核', '复核人'),
    # 业务模块自动凭证用的科目（改这里就行，填末级科目编码）
    ('acc_vatin',  '进项税科目', '进项税科目'),
    ('acc_ap',     '应付账款科目（采购入库）', '应付科目'),
    ('acc_ar',     '应收账款科目（销售结算）', '应收科目'),
    ('acc_rev',    '主营业务收入科目', '收入科目'),
    ('acc_vatout', '销项税科目', '销项税科目'),
    ('acc_cogs',   '主营业务成本科目', '销售成本科目'),
    ('acc_ore',    '库存商品-原矿科目', '原矿科目'),
    ('acc_dep',    '累计折旧科目', '累计折旧科目'),
    ('acc_fx',     '汇兑损益科目', '汇兑损益科目'),
    ('acc_frt',    '运杂清关关税挂账科目', '运杂税费科目'),
    ('acc_pre',    '启用日以前的生产成本、折旧记到', '启用日前成本科目'),
]
ACC_PARAMS = ['acc_vatin', 'acc_ap', 'acc_ar', 'acc_rev', 'acc_vatout', 'acc_cogs', 'acc_ore', 'acc_dep', 'acc_fx', 'acc_frt', 'acc_pre']
EXCH_CODE = '122105'               # 外币兑换待对冲（【记账规则】里「购汇/结汇」用的科目，月末结转⑤按它算）
P_ROW = {k: 5 + i for i, (k, _, _) in enumerate(PARAMS)}

def P(key):
    return f"{SH_BASE}!$B${P_ROW[key]}"

# 资金账户 D:L（20 个）
ACC_R0, ACC_R1 = 5, 24
(A_SEQ, A_NAME, A_BANK, A_CUR, A_OPEN, A_OPENB, A_CODE, A_NOWO, A_NOWB, A_NOTE) = \
    'D E F G H I J K L M'.split()
# 项目/业务 O:P（40 个）
PRJ_R0, PRJ_R1 = 5, 44
PJ_SEQ, PJ_NAME = 'O', 'P'
# 往来单位 R:V（300 个）
CP_R0, CP_R1 = 5, 304
(CP_SEQ, CP_NAME, CP_TYPE, CP_CAP, CP_NOTE) = 'R S T U V'.split()
# 物料档案 X:AH（200 个）
MAT_R0, MAT_R1 = 5, 204
(M_SEQ, M_NAME, M_SPEC, M_UNIT, M_CAT, M_ACC, M_STOCK, M_Q0, M_P0, M_A0, M_NOTE) = \
    'X Y Z AA AB AC AD AE AF AG AH'.split()

# ─────────────────────────── 记账规则 ───────────────────────────
# 收支类别 A:I（60 个）
RC_R0, RC_R1 = 5, 64
(RC_SEQ, RC_NAME, RC_ATTR, RC_IN, RC_OUT, RC_HIST, RC_CFIN, RC_CFOUT, RC_NOTE) = 'A B C D E F G H I'.split()
# 物料类别 K:O（30 个）
MC_R0, MC_R1 = 5, 34
(MC_NAME, MC_ACC, MC_STOCK, MC_PROD, MC_NOTE) = 'K L M N O'.split()
# 领用用途 Q:S（20 个）
US_R0, US_R1 = 5, 24
(US_NAME, US_ACC, US_NOTE) = 'Q R S'.split()
# 部门 U:X（20 个）：部门 / 工资计入 / 折旧计入 / 说明
DP_R0, DP_R1 = 5, 24
(DP_NAME, DP_WAGE, DP_DEP, DP_NOTE) = 'U V W X'.split()
# 人员类别 Z:AA
PT_R0, PT_R1 = 5, 10
(PT_NAME, PT_ACC) = 'Z AA'.split()
# 固定资产类别 AC:AG（20 个）
FC_R0, FC_R1 = 5, 24
(FC_NAME, FC_ACC, FC_YEARS, FC_SALV, FC_NOTE) = 'AC AD AE AF AG'.split()
# 现金流量项目 AI:AK
CFI_R0, CFI_R1 = 5, 26
(CFI_NO, CFI_NAME, CFI_SIDE) = 'AI AJ AK'.split()

# ─────────────────────────── 月度汇率 ───────────────────────────
RATE_HDR = 5
RATE_R0, RATE_R1 = 6, 6 + N_MONTHS - 1
RATE_CURS = ['UZS', 'CNY', 'USD', 'EUR', 'RUB', 'KZT', 'JPY', 'GBP']    # B..I

# ─────────────────────────── 科目表 ───────────────────────────
AC_HDR = 4
AC_R0, AC_R1 = 5, 454              # 450 个科目
(AC_CODE, AC_NAME, AC_CLASS, AC_DIR, AC_AUX, AC_FX, AC_CASHF, AC_STAT, AC_PARA,
 AC_LVL, AC_FULL, AC_LEAF, AC_OPEN, AC_KEY) = 'A B C D E F G H I J K L M N'.split()

def acc(col):
    return rng(SH_ACC, col, AC_R0, AC_R1)

# ─────────────────────────── 现金流水总表（沿用原表 A:S 列序，第 5 行表头、第 6 行起数据） ───────────────────────────
CASH_HDR = 5
CASH_R0, CASH_R1 = 6, 4005         # 4000 行（原来 2402 行）
(K_DATE, K_ACC, K_BANK, K_CUR, K_CAT, K_ATTR, K_MEMO, K_IN, K_OUT, K_XR, K_RATE, K_NET, K_UZS,
 K_BAL, K_BBAL, K_PRJ, K_CP, K_NOTE, K_CHK) = 'A B C D E F G H I J K L M N O P Q R S'.split()
# 新加的（原表 S 列后面）
K_OVR, K_CTR, K_MERGE, K_ATT = 'T U V W'.split()          # 指定对方科目 / 实际对方科目（自动）/ 凭证合并号 / 附单据张数
# 隐藏辅助列
(K_ACODE, K_RAW, K_CCODE, K_DR, K_CR, K_VALID, K_CFI, K_GSEQ, K_GHEAD, K_VID, K_VMEMO) = \
    'X Y Z AA AB AC AD AE AF AG AH'.split()
# X 本账户科目 Y 规则给的对方科目（原样）Z 对方科目（最终编码）AA 本账户借额（收入折苏姆）AB 本账户贷额
# AC 有效（有金额且两边科目都有=1）AD 现金流量项目 AE 合并组内第几笔 AF 合并组首笔 AG 凭证ID AH 凭证摘要
K_LAST = K_ATT

def cash(col):
    return rng(SH_CASH, col, CASH_R0, CASH_R1)

# ─────────────────────────── 录入模块（第 3 行表头、第 4 行起数据） ───────────────────────────
MOD_HDR, MOD_R0 = 3, 4
BUY_R1 = MOD_R0 + 600 - 1          # 600 行
(B_SEQ, B_DATE, B_NO, B_SUP, B_MAT, B_SPEC, B_UNIT, B_QTY, B_CUR, B_UP, B_AMT0, B_RATE, B_AMT,
 B_FRT, B_VAT, B_COST, B_TOT, B_ACC, B_PRJ, B_NOTE, B_CHK) = \
    'A B C D E F G H I J K L M N O P Q R S T U'.split()
# 隐藏：V 有效  W 凭证分组键（单据号，没填＝本行自己）  X 本行分录条数  Y 同一凭证前面几行的分录条数  Z 凭证ID
B_VALID, B_GKEY, B_LEGS, B_GPRE, B_VID = 'V W X Y Z'.split()
ISS_R1 = MOD_R0 + 1000 - 1         # 1000 行
(I_SEQ, I_DATE, I_NO, I_MAT, I_SPEC, I_UNIT, I_QTY, I_USE, I_PRICE, I_AMT, I_DR, I_CR, I_PRJ,
 I_NOTE, I_CHK) = 'A B C D E F G H I J K L M N O'.split()
I_VALID, I_GKEY, I_LEGS, I_GPRE, I_VID = 'P Q R S T'.split()
PROD_R1 = MOD_R0 + 1000 - 1
(PD_SEQ, PD_DATE, PD_FACE, PD_TON, PD_DEV, PD_GRADE, PD_DIL, PD_NOTE, PD_CHK) = 'A B C D E F G H I'.split()
SALE_R1 = MOD_R0 + 300 - 1         # 300 行
(S_SEQ, S_DATE, S_NO, S_CUS, S_ITEM, S_TON, S_PRICE, S_AMT, S_RATE, S_NET, S_TAX, S_TOT, S_UC,
 S_COGS, S_PRJ, S_NOTE, S_CHK) = 'A B C D E F G H I J K L M N O P Q'.split()
S_VALID, S_GKEY, S_LEGS, S_GPRE, S_VID = 'R S T U V'.split()
PAY_R1 = MOD_R0 + 200 - 1          # 200 行
(W_SEQ, W_DATE, W_DEPT, W_TYPE, W_HEAD, W_AMT, W_DR, W_CR, W_NOTE, W_CHK) = 'A B C D E F G H I J'.split()
W_VALID, W_GKEY, W_LEGS, W_GPRE, W_VID = 'K L M N O'.split()      # 工资按月并成一张凭证
FA_R1 = MOD_R0 + 200 - 1           # 200 张卡片
(F_NO, F_NAME, F_CAT, F_ACC, F_COST, F_DATE, F_YEARS, F_SALV, F_DEPT, F_DACC, F_MDEP, F_START,
 F_MONTHS, F_ACCUM, F_NBV, F_REMK, F_CHK) = 'A B C D E F G H I J K L M N O P Q'.split()   # F_REMK 备注（别叫 F_NOTE，跟字体重名）
F_YEFF, F_SEFF, F_END, F_LASTAMT, F_OK = 'R S T U V'.split()   # 隐藏：实际年限 / 实际残值率 / 最后折旧月 / 最后一个月的折旧额（补尾差）/ 卡片没错（=1 才提折旧）
MAN_R1 = MOD_R0 + 600 - 1          # 600 行分录
(MN_GRP, MN_DATE, MN_MEMO, MN_CODE, MN_NAME, MN_DR, MN_CR, MN_CP, MN_PRJ, MN_ATT, MN_CHK) = \
    'A B C D E F G H I J K'.split()
# 隐藏：L 组内第几条  M 科目编码  N 有金额  O 凭证日期（组内第一行的日期）  P 这一行有错  Q 这张凭证能记账  R 凭证ID
MN_SEQ, MN_CODE2, MN_HAS, MN_VDATE, MN_BAD, MN_OK, MN_VID, MN_KEY, MN_GKEY = 'L M N O P Q R S T'.split()
# S：能记账的行的分组键（找凭证第一行用）  T：分组键＝组号|年月（组号每月可以重新从 1 编）
MAN_GROUPS = 150

def mod(sh, col):
    r1 = {SH_BUY: BUY_R1, SH_ISS: ISS_R1, SH_PROD: PROD_R1, SH_SALE: SALE_R1, SH_PAY: PAY_R1,
          SH_FA: FA_R1, SH_MAN: MAN_R1}[sh]
    return rng(sh, col, MOD_R0, r1)

# ─────────────────────────── 月末结转（全自动） ───────────────────────────
# 三块长表，按「月份序号 n」一行：
#  ① 调汇：每个资金账户 × 每个月一行（20 × 32）   ② 生产成本结转：每个成本科目 × 每个月（N_COSTLEAF × 32）
#  ③ 每月汇总行（折旧 / 调汇 / 成本 / 兑换差额 凭证金额）
ME_M0 = 6                                   # 月汇总块：第 6..37 行（其余各块的位置见 s_mend）
ME_M1 = ME_M0 + N_MONTHS - 1
N_ACC = ACC_R1 - ACC_R0 + 1                 # 20 个账户

# ─────────────────────────── 记账分录（总表） ───────────────────────────
JE_HDR, JE_R0 = 3, 4
# 看得见的：A 日期 B 凭证号（当月第几号） C 摘要 D 科目编码 E 科目全称 F 借方 G 贷方 H 往来单位 I 项目 J 来源
# 隐藏：K 源位置 L 凭证ID M 凭证内第几条 N 查找键(ID×1000+条) O 明细账排序键 P 有金额 Q 本行之前同一凭证已有几条（每个源位置第一行）
(J_DATE, J_VNO, J_MEMO, J_CODE, J_NAME, J_DR, J_CR, J_CP, J_PRJ, J_SRC,
 J_POS, J_VID, J_SEQ, J_KEY, J_GLK, J_NZ, J_START) = 'A B C D E F G H I J K L M N O P Q'.split()
# 来源代码（凭证ID = 代码 × 100000 + 源行位置；月末几种凭证的「位置」= 月份序号）
SRC = {'流水': 1, '采购': 2, '领用': 3, '销售': 4, '工资': 5, '手工': 6,
       '折旧': 7, '调汇': 8, '成本': 9, '销售成本': 10, '兑换': 11}

# ─────────────────────────── 凭证索引 / 凭证汇总 / 打印 ───────────────────────────
VX_R0 = 4
# A 凭证ID B 来源 C 有这张凭证 D 日期 E 排序键 F 分录条数 G 页数 H 月份键 I 当月凭证号 J 摘要 K 附单据 L 查找键(月份×10000+号)
(VX_ID, VX_SRC, VX_ON, VX_DATE, VX_KEY, VX_LINES, VX_PAGES, VX_MKEY, VX_NO, VX_MEMO, VX_ATT, VX_FIND) = \
    'A B C D E F G H I J K L'.split()
VL_HDR, VL_R0, VL_N = 5, 6, 450             # 每月最多 450 张凭证
VP_ROWS, VP_LINES, VP_PAGES = 14, 8, 500    # 每页 14 行、8 条分录；每月最多印 500 页

# ─────────────────────────── 科目余额表 / 明细账 ───────────────────────────
TB_HDR, TB_R0 = 6, 7
# A 编码 B 科目名称 C 方向 D 级次 E 末级 F/G 月初余额借/贷 H/I 本月发生借/贷 J/K 本年累计借/贷 L/M 期末余额借/贷
# 隐藏：N 年初净额(借正贷负) O 月初净额 P 期末净额 Q 科目表里第几行
(T_CODE, T_NAME, T_DIR, T_LVL, T_LEAF, T_OB_D, T_OB_C, T_M_D, T_M_C, T_Y_D, T_Y_C, T_CB_D, T_CB_C,
 T_YNET, T_ONET, T_CNET, T_POS) = 'A B C D E F G H I J K L M N O P Q'.split()
TB_R1 = TB_R0 + (AC_R1 - AC_R0)
GL_R0, GL_N = 8, 1500

# ─────────────────────────── 月份工具 ───────────────────────────
def mstart(n):
    """第 n 个月（1 起）的第一天：以【基础资料】起始月份为准"""
    return f'EDATE(起始月份,{n}-1)'


# 报表期间（首页选的年度、月份）
REP_M0 = 'DATE(报表年度,报表月份,1)'          # 本月第一天
REP_M1 = 'DATE(报表年度,报表月份+1,0)'        # 本月最后一天
REP_Y0 = 'DATE(报表年度,1,1)'                 # 本年第一天
REP_N = '((报表年度-YEAR(起始月份))*12+报表月份-MONTH(起始月份)+1)'   # 报表月是第几个月


# ─────────────────────────── 样式（微软雅黑，跟 A064 一个样） ───────────────────────────
YH = '微软雅黑'
F_TITLE = Font(name=YH, sz=16, bold=True, color='FFFFFFFF')
F_TIP   = Font(name=YH, sz=10, color='FFFF6600')
F_HDR   = Font(name=YH, sz=11, bold=True, color='FFFFFFFF')
F_HDR_D = Font(name=YH, sz=11, bold=True, color='FF1F3864')
F_TXT   = Font(name=YH, sz=10, color='FF000000')
F_TXTB  = Font(name=YH, sz=10, bold=True, color='FF000000')
F_IN    = Font(name=YH, sz=10, color='FF1F4E79')
F_AUTO  = Font(name=YH, sz=10, color='FF404040')
F_NOTE  = Font(name=YH, sz=9, color='FF808080')
F_RED   = Font(name=YH, sz=10, bold=True, color='FFC00000')
F_KPI_L = Font(name=YH, sz=11, bold=True, color='FF1F3864')
F_KPI_V = Font(name=YH, sz=13, bold=True, color='FFC00000')
F_SEC   = Font(name=YH, sz=12, bold=True, color='FF000000')
F_SEL   = Font(name=YH, sz=12, bold=True, color='FF1F4E79')

def fill(rgb):
    return PatternFill('solid', fgColor=rgb)

# 标题色（标签颜色用同一个色）
C_HOME, C_BASE, C_CASH, C_BUY, C_ISS, C_PROD, C_SALE, C_PAY, C_FA, C_MAN = \
    'FFC00000', 'FF595959', 'FF2F75B5', 'FF7030A0', 'FF8064A2', 'FF806000', 'FFC65911', 'FF2E75B6', 'FF7F6000', 'FF404040'
C_JE, C_VCH, C_LED, C_RPT, C_CHK, C_OPS = 'FF548235', 'FF375623', 'FF1F4E79', 'FF375623', 'FF833C0C', 'FFBF8F00'
C_INCOME, C_EXPENSE = 'FF00B050', 'FFFFC000'

FILL_TIP  = fill('FFFFF2CC')
FILL_IN   = fill('FFFFF7E0')   # 手工录入：淡黄
FILL_AUTO = fill('FFF2F2F2')   # 自动计算：淡灰
FILL_KPI  = fill('FFD9E1F2')
FILL_BAND = fill('FF8EA9DB')
FILL_SUBH = fill('FFDDEBF7')
FILL_TOT  = fill('FFFCE4D6')
FILL_OK   = fill('FFE2EFDA')
FILL_WARN = fill('FFFFC7CE')
FILL_SEL  = fill('FFFFFF00')
NOFILL    = PatternFill(fill_type=None)

thin = Side(style='thin', color='FFBFBFBF')
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
NOBD = Border()

AC  = Alignment(horizontal='center', vertical='center')
ACW = Alignment(horizontal='center', vertical='center', wrap_text=True)
AL  = Alignment(horizontal='left', vertical='center')
ALW = Alignment(horizontal='left', vertical='center', wrap_text=True)
AR_ = Alignment(horizontal='right', vertical='center')

MONEY = '#,##0.00;[Red]-#,##0.00;"-"'      # 苏姆金额：负数红、零显示 -
MONEY0 = '#,##0.00'
QTY   = '[=0]"-";General'                  # 数量：整数不带小数点，零显示 -
QTY2  = '#,##0.00;[Red]-#,##0.00;"-"'
PRICE = '#,##0.00##'
RATE  = '#,##0.00##'
DATE  = 'yyyy/mm/dd'
MONTH = 'yyyy"年"m"月"'
PCT   = '0.0%;[Red]-0.0%;"-"'
PCT2  = '0.00%;[Red]-0.00%;"-"'
CNT   = 'General;[Red]-General;0'


def at(rng_, k):
    """按位置取第 k 格：INDEX(区域,k)。用户在明细表里插行/删行也不会串位或变 #REF!"""
    return f'INDEX({rng_},{k})'


def upto(rng_, k):
    """区域开头到第 k 格：$C$4:INDEX($C$4:$C$1003,k)（累计用）"""
    first = rng_.split(':')[0]
    return f'{first}:INDEX({rng_},{k})'


def put(ws, coord, value=None, font=None, fill_=None, fmt=None, align=None, border=True):
    c = ws[coord]
    if value is not None:
        c.value = value
    if font is not None:
        c.font = font
    if fill_ is not None:
        c.fill = fill_
    if fmt is not None:
        c.number_format = fmt
    if align is not None:
        c.alignment = align
    if border:
        c.border = BD
    return c


def title(ws, text, last_col, color, tip=None, h1=33, h2=None):
    """第 1 行标题带 + 第 2 行 💡 提示；提示长就把第 2 行加高"""
    if h2 is None:
        width = sum((ws.column_dimensions[CL(i)].width or 9) for i in range(1, CI(last_col) + 1))
        per_line = max(20, int(width / 1.9))
        lines = -(-len(tip or '') // per_line)
        h2 = max(24, 16 * lines + 8)
    ws.merge_cells(f'A1:{last_col}1')
    put(ws, 'A1', text, F_TITLE, fill(color), align=AC, border=False)
    ws.row_dimensions[1].height = h1
    if tip:
        ws.merge_cells(f'A2:{last_col}2')
        put(ws, 'A2', tip, F_TIP, FILL_TIP, align=ALW, border=False)
        ws.row_dimensions[2].height = h2
    ws.sheet_properties.tabColor = color[2:]


def header(ws, row, cols_texts, color, font=F_HDR, height=34):
    for col, t in cols_texts:
        put(ws, f'{col}{row}', t, font, fill(color), align=ACW)
    ws.row_dimensions[row].height = height


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def hide(ws, cols):
    for c in cols:
        ws.column_dimensions[c].hidden = True


def link(cell, sheet, ref='A1'):
    """表内跳转链接：写成 location（Excel 自己存内部链接的方式），WPS / Excel 都能点过去"""
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!{ref}",
                               display=None if cell.value is None else str(cell.value))
    return cell


def style_rows(ws, r0, r1, cols, auto=(), fmts=None, aligns=None, bold=(), height=None, font_note=()):
    """数据区统一样式：白底细灰框；手填列黑字，公式列灰底灰字"""
    fmts, aligns = fmts or {}, aligns or {}
    for r in range(r0, r1 + 1):
        for col in cols:
            c = ws[f'{col}{r}']
            is_auto = col in auto
            c.font = F_NOTE if col in font_note else (F_TXTB if col in bold else (F_AUTO if is_auto else F_TXT))
            c.fill = FILL_AUTO if is_auto else NOFILL
            c.border = BD
            c.alignment = aligns.get(col, AC)
            if col in fmts:
                c.number_format = fmts[col]
        if height:
            ws.row_dimensions[r].height = height


def lock_formula_sheet(ws):
    """全是公式的表：锁上（不设密码），只能筛选、不能排序——排序会把公式打乱。要解锁：审阅→撤销工作表保护"""
    p = ws.protection
    p.sheet = True
    p.autoFilter = False          # False＝允许筛选
    p.sort = True                 # True＝不许排序
    p.formatColumns = False
    p.formatRows = False
    p.selectLockedCells = False
    p.selectUnlockedCells = False


def add_date_dv(ws, sqref):
    dv = DataValidation(type='date', operator='between', formula1='起始月份', formula2='DATE(2099,12,31)',
                        allow_blank=True, showErrorMessage=True, errorStyle='warning')
    dv.errorTitle, dv.error = '日期不对', '日期要像 2026/9/25 这样录，而且不能早于建账起始月份'
    ws.add_data_validation(dv)
    dv.add(sqref)
    return dv


def add_list_dv(ws, sqref, formula, prompt=None, allow_blank=True, stop=True):
    dv = DataValidation(type='list', formula1=formula, allow_blank=allow_blank,
                        showErrorMessage=True, errorStyle='stop' if stop else 'warning')
    if prompt:
        dv.promptTitle, dv.prompt = '提示', prompt
        dv.showInputMessage = True
    dv.error = '请从下拉里选；清单里没有的，先到【基础资料】/【记账规则】/【科目表】里登记'
    dv.errorTitle = '不在清单里'
    ws.add_data_validation(dv)
    dv.add(sqref)
    return dv
