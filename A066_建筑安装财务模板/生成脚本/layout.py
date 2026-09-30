# -*- coding: utf-8 -*-
"""A066 各张表的名字、行列地址、容量。表与表之间互相引用的地址都从这里取，不在各模块里手写。"""
from common import CL, CI

# ─────────────────────────── 表名 ───────────────────────────
SH_HOME = '首页'
SH_BASE = '基础资料'
SH_PROJ = '项目档案'
SH_UNIT = '往来单位'
SH_CASH = '资金流水'
SH_AP = '应付登记'
SH_REV = '收入确认'
SH_OFF = '代发抵账'
SH_ATT = '考勤工资'
SH_INV = '发票登记'
SH_ACC = '账户明细'
SH_WEEK = '资金周报'
SH_MON = '资金月报'
SH_AR = '应收账款总表'
SH_APS = '应付账款总表'
SH_PER = '个人往来'
SH_PAY = '工资表'
SH_PAYS = '工资汇总'
SH_LAB = '用工成本表'
SH_PL = '项目账'
SH_PPL = '项目利润表'
SH_ALLOC = '费用分摊'
SH_INVS = '发票统计'
SH_IS = '利润表'
SH_BS = '资产负债表'
SH_BE = '盈亏平衡表'
SH_CHK = '数据校验'
SH_AUX = '_辅助'


def rng(sheet, col, r0, r1):
    return f"{sheet}!${col}${r0}:${col}${r1}"


# ─────────────────────────── 基础资料 ───────────────────────────
B_HDR = 4
# ① 公司（竖着几行）
CO_NAME, CO_FULL, CO_TAX, CO_TYPE, CO_OPEN = 'C5', 'C6', 'C7', 'C8', 'C9'      # 简称、全称、税号、纳税人、建账日期
# ② 资金账户
AC_R0, AC_R1 = 5, 24
AC_SEQ, AC_NAME, AC_TYPE, AC_OWNER, AC_NO, AC_OPEN, AC_NOW, AC_REAL, AC_RDATE, AC_DIFF, AC_NOTE = \
    'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'     # M/N 手填：银行 App 上的实际余额和日期；O 自动差额
AC_TYPES = ['银行', '现金', '个人户', '票据']
# ③ 收支类别：类别 方向 报表项目 分摊归类 要选什么 说明
CT_R0, CT_R1 = 5, 64
CT_SEQ, CT_NAME, CT_DIR, CT_LINE, CT_ALLOC, CT_PROJ, CT_NOTE = 'R', 'S', 'T', 'U', 'V', 'W', 'X'
# ④ 摘要关键词
KW_R0, KW_R1 = 5, 164
KW_SEQ, KW_WORD, KW_DIR, KW_CAT = 'Z', 'AA', 'AB', 'AC'
# ⑤ 参数（税率、分摊方法）
PA_R0 = 5
PA_LBL, PA_VAL, PA_NOTE = 'AE', 'AF', 'AG'
PARAMS = [('销项税率（工程款）', 0.09, '0%', '一般计税 9%；简易计税的项目在【项目档案】税率填 3%'),
          ('附加税费率（按增值税）', 0.12, '0%', '城建 7%＋教育 3%＋地方教育 2%；小微企业减半的按实际填（比如 6%）'),
          ('印花税率（按确认产值）', 0.0003, '0.00%', '建筑安装合同万分之三'),
          ('异地预缴增值税率', 0.02, '0%', '一般计税 2%、简易计税 3%（只在【项目账】税负测算里显示，预缴的钱冲应交税费）'),
          ('管理费分摊依据', '施工费', '@', '施工费＝人工＋分包＋机械（默认）；直接成本＝再加材料和其他直接费'),
          ('收入确认落后提醒', 0.3, '0%', '直接成本占总价的比例比确认产值的比例高出这么多，就提醒补录产值确认'),
          ('税金占产值（估，没数据时用）', 0.03, '0.0%', '【盈亏平衡表】【报价计算器】估税金用；有了项目数据就自动按实际比例算')]
PA_VATR, PA_SURR, PA_STAMP, PA_PRE, PA_DRV, PA_LAG, PA_TAXB = [f"{SH_BASE}!${PA_VAL}${PA_R0 + i}" for i in range(7)]
# ⑥ 固定资产：名称 购入日期 原值 月数 残值率 使用项目（空＝公司） 建账前已提折旧 月折旧(自动) 备注
FA_R0, FA_R1 = 5, 34
FA_SEQ, FA_NAME, FA_DATE, FA_COST, FA_MON, FA_RES, FA_PJ, FA_DEP_M, FA_NOTE, FA_S, FA_E = \
    'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS'     # AR/AS 隐藏：开始提折旧、提完的月序号

AC_NAMES = rng(SH_BASE, AC_NAME, AC_R0, AC_R1)
AC_TYPES_R = rng(SH_BASE, AC_TYPE, AC_R0, AC_R1)
AC_OWNERS = rng(SH_BASE, AC_OWNER, AC_R0, AC_R1)
AC_OPENS = rng(SH_BASE, AC_OPEN, AC_R0, AC_R1)
CT_NAMES = rng(SH_BASE, CT_NAME, CT_R0, CT_R1)
CT_NEEDS = rng(SH_BASE, CT_PROJ, CT_R0, CT_R1)
CT_DIRS = rng(SH_BASE, CT_DIR, CT_R0, CT_R1)
CT_LINES = rng(SH_BASE, CT_LINE, CT_R0, CT_R1)
CT_ALLOCS = rng(SH_BASE, CT_ALLOC, CT_R0, CT_R1)
KW_WORDS = rng(SH_BASE, KW_WORD, KW_R0, KW_R1)
KW_DIRS = rng(SH_BASE, KW_DIR, KW_R0, KW_R1)
KW_CATS = rng(SH_BASE, KW_CAT, KW_R0, KW_R1)
OPEN_DATE = f"{SH_BASE}!${CO_OPEN[0]}${CO_OPEN[1:]}"

# ─────────────────────────── 项目档案 ───────────────────────────
PJ_HDR = 3
PJ_R0, PJ_R1 = 4, 63            # 60 个项目
# A 序号 B 简称 C 全称 D 甲方/总包 E 合同额 F 变更签证 G 结算额 H 质保金比例 I 质保到期 J 税率 K 开工 L 完工 M 状态 N 负责人
# O/P/Q 摘要关键词 R 建账前已收 S 备注 ｜ 自动 T 总价（结算额，没结算＝合同＋变更）
(PJ_SEQ, PJ_NAME, PJ_FULL, PJ_CUS, PJ_AMT, PJ_CHG, PJ_SET, PJ_RET, PJ_RETD, PJ_TAX, PJ_START, PJ_END, PJ_STAT, PJ_MGR,
 PJ_KW1, PJ_KW2, PJ_KW3, PJ_REC0, PJ_NOTE, PJ_TOTAL) = [CL(i) for i in range(1, 21)]
PJ_STATS = ['在建', '完工未结算', '已结算', '质保期', '已完结', '投标中']
PJ_NAMES = rng(SH_PROJ, PJ_NAME, PJ_R0, PJ_R1)

# ─────────────────────────── 往来单位及人员（甲方、材料商、分包、机械、股东、管理人员、工人都在这一张） ───────────────────────────
UN_HDR = 3
UN_R0, UN_R1 = 4, 303           # 300 个
# A 序号 B 名称 C 类型 D 别名1 E 别名2 F 全称 G 税号/身份证 H 手机 I 银行账号 J 开户行 K 入职 L 离职 M 期初欠薪 N 备注
(UN_SEQ, UN_NAME, UN_TYPE, UN_AL1, UN_AL2, UN_FULL, UN_ID, UN_TEL, UN_BANKNO, UN_BANK, UN_IN, UN_OUT, UN_OWE0, UN_NOTE) = \
    [CL(i) for i in range(1, 15)]
UN_TYPES = ['甲方/总包', '材料供应商', '分包', '机械运输', '管理人员', '工人', '临时工', '股东', '税务/银行', '其他']
PERSON_TYPES = ['管理人员', '工人']                     # 走考勤工资的人
AP_TYPES = ['材料', '分包', '机械运输', '其他直接费']      # 应付登记的类型
UN_NAMES = rng(SH_UNIT, UN_NAME, UN_R0, UN_R1)
UN_TYPES_R = rng(SH_UNIT, UN_TYPE, UN_R0, UN_R1)

# ─────────────────────────── 工资标准（涨薪表）：一条一行，涨薪加一行 ───────────────────────────
SH_RATE = '工资标准'
RT_HDR = 3
RT_R0, RT_R1 = 4, 303           # 300 条
RT_SEQ, RT_NAME, RT_DATE, RT_DAY, RT_MON, RT_NOTE, RT_CHK = 'A', 'B', 'C', 'D', 'E', 'F', 'G'
RT_DN = 'H'                     # 隐藏：从哪天起（数字；不是日期＝0）
RT_KEY = 'I'                    # 隐藏：姓名|这个人的第几条（按日期排，1 起）——查「某天有效的那条」用 COUNTIFS＋MATCH，WPS 也能算
RT_NAMES = rng(SH_RATE, RT_NAME, RT_R0, RT_R1)
RT_DATES = rng(SH_RATE, RT_DATE, RT_R0, RT_R1)
RT_DAYS = rng(SH_RATE, RT_DAY, RT_R0, RT_R1)
RT_MONS = rng(SH_RATE, RT_MON, RT_R0, RT_R1)
RT_DNS = rng(SH_RATE, RT_DN, RT_R0, RT_R1)
RT_KEYS = rng(SH_RATE, RT_KEY, RT_R0, RT_R1)

# ─────────────────────────── 资金流水 ───────────────────────────
J_HDR = 5
J_R0, J_R1 = 6, 2005            # 2000 笔
# 录入：B 日期 C 账户 D 摘要 E 收入 F 支出 ｜ 自动：G 结存 H 项目 I 收支类别 J 单位/人 K 校验 ｜ 手工改：L 改项目 M 改类别 N 改单位 O 对方账户 P 备注
(J_SEQ, J_DATE, J_ACC, J_MEMO, J_IN, J_OUT, J_BAL, J_PJ, J_CT, J_UN, J_CHK, J_MPJ, J_MCT, J_MUN, J_TOACC, J_NOTE) = \
    [CL(i) for i in range(1, 17)]
# 隐藏：R 净额 S 年月 T 认项目 U 认类别 V 认单位 W 分摊归类 X 报表项目 Y 单位有应付 Z 单位类型 AA 走考勤的人 AB 代X付
#       AC 认对方账户 AD 对方账户（用的） AE 账户类型 AF 对方账户类型 AG 关键词类别 AH 收支（报表用：收/支） AI 项目命中词 AJ 别的项目命中数
(J_NET, J_YM, J_APJ, J_ACT, J_AUN, J_ALLOC, J_LINE, J_HANG, J_UTYPE, J_WKIND, J_DAI, J_ATO, J_TO, J_ATYPE, J_TOTYPE, J_KWC,
 J_IO, J_PKW, J_PKN, J_PKH, J_UMH, J_UDH, J_DN, J_RK1, J_RK2, J_RK3) = \
    'R S T U V W X Y Z AA AB AC AD AE AF AG AH AI AJ AK AL AM AN AO AP AQ'.split()
# AN 隐藏：日期（数字；不是日期＝0，给「最后一次」这类公式用）
# AK～AM 隐藏：项目关键词、单位（摘要里）、单位（代X付里）在按长短排好的关键词表（_辅助）里第几个对上（只算一遍）
# AO～AQ 隐藏：「最后一次」用的排号：项目|第几次收工程款、单位|第几次付款、人|第几笔保证金押金（按日期，COUNTIFS 算，WPS 也能算）
J_CAP = J_R1 - J_R0 + 1


def jr(col):
    return rng(SH_CASH, col, J_R0, J_R1)

# ─────────────────────────── _辅助（下拉清单、关键词表） ───────────────────────────
# A/B 项目关键词 → 项目（每个项目 4 行：简称 ＋ 3 个关键词）
AX_PKW, AX_PKP = 'A', 'B'
AX_PKW_N = (PJ_R1 - PJ_R0 + 1) * 4
# C/D 单位、人员的名字和别名 → 标准名（每个 3 行：名称 别名1 别名2）
AX_UKW, AX_UKN = 'C', 'D'
AX_UKW_N = (UN_R1 - UN_R0 + 1) * 3
# E 项目＋「公司管理」（考勤的项目下拉）
AX_ALL = 'E'
AX_ALL_N = (PJ_R1 - PJ_R0 + 1) + 1          # 「公司管理」＋项目
# F/G 账户关键词（账户名、个人户的主人名）→ 账户（账户互转自动认对方账户）；H 报表项目清单
AX_AKW, AX_AKA = 'F', 'G'
AX_AKW_N = (AC_R1 - AC_R0 + 1) * 2
AX_PKW_R = rng(SH_AUX, AX_PKW, 1, AX_PKW_N)
AX_PKP_R = rng(SH_AUX, AX_PKP, 1, AX_PKW_N)
AX_UKW_R = rng(SH_AUX, AX_UKW, 1, AX_UKW_N)
AX_UKN_R = rng(SH_AUX, AX_UKN, 1, AX_UKW_N)
AX_ALL_R = f"{SH_AUX}!${AX_ALL}$1:${AX_ALL}${AX_ALL_N}"
# P～AA：上面三张关键词表按「长的在前（一样长表里靠前的在前）」排好，空的排在最后。
#   流水认项目/单位/对方账户＝MATCH(1, INDEX(ISNUMBER(SEARCH(排好的词, 摘要))*(排好的词<>""), 0), 0)：第一个对上的就是最长的（WPS 也能算）
AX_PW, AX_PI, AX_PS, AX_PN = 'P', 'Q', 'R', 'S'      # 项目：权重、排第几的是哪行、词、项目
AX_UW, AX_UI, AX_US, AX_UN = 'T', 'U', 'V', 'W'      # 单位
AX_AW, AX_AI, AX_AS, AX_AN = 'X', 'Y', 'Z', 'AA'     # 账户
AX_KWI, AX_KWO = 'AB', 'AC'                          # 收支类别关键词：收款能用的、付款能用的（不能用的放空），跟【基础资料】③ 一行对一行
AX_PS_R = rng(SH_AUX, AX_PS, 1, AX_PKW_N)
AX_PN_R = rng(SH_AUX, AX_PN, 1, AX_PKW_N)
AX_US_R = rng(SH_AUX, AX_US, 1, AX_UKW_N)
AX_UN_R = rng(SH_AUX, AX_UN, 1, AX_UKW_N)

# ─────────────────────────── 应付登记（材料 · 分包 · 机械运输 · 其他直接费） ───────────────────────────
AP_HDR = 4
AP_R0, AP_R1 = 5, 1004          # 1000 行
(AP_SEQ, AP_DATE, AP_TYPE, AP_UNIT, AP_PJ, AP_MEMO, AP_AMT, AP_RATE, AP_NOTE, AP_YM, AP_CHK, AP_UTYPE) = \
    [CL(i) for i in range(1, 13)]     # L 隐藏：单位类型


def apr(col):
    return rng(SH_AP, col, AP_R0, AP_R1)


# ─────────────────────────── 收入确认（产值 / 结算确认） ───────────────────────────
RV_HDR = 4
RV_R0, RV_R1 = 5, 204
(RV_SEQ, RV_DATE, RV_PJ, RV_TYPE, RV_AMT, RV_NOTE, RV_CUS, RV_YM, RV_CHK, RV_VAT) = [CL(i) for i in range(1, 11)]   # J 隐藏：销项税额
RV_TYPES = ['进度确认', '签证变更', '结算调整']


def rvr(col):
    return rng(SH_REV, col, RV_R0, RV_R1)


# ─────────────────────────── 代发抵账（没经过我们账户的收付） ───────────────────────────
OF_HDR = 4
OF_R0, OF_R1 = 5, 304
(OF_SEQ, OF_DATE, OF_PJ, OF_TYPE, OF_WHO, OF_AMT, OF_NOTE, OF_YM, OF_WTYPE, OF_CHK, OF_DN, OF_RK) = [CL(i) for i in range(1, 13)]
# K 隐藏：日期（数字）；L 隐藏：项目|第几次代发代付（按日期，应收账龄「最后回款」用）
OF_TYPES = ['总包代发工资', '总包代付材料分包款', '甲供材扣款', '甲方扣款']


def ofr(col):
    return rng(SH_OFF, col, OF_R0, OF_R1)


# ─────────────────────────── 发票登记（A～S 跟电子税务局「发票查询 → 导出 → 发票基础信息」一模一样，整块粘） ───────────────────────────
IV_HDR = 4
IV_R0, IV_R1 = 5, 1004          # 1000 张
IV_RAW = ['序号', '发票代码', '发票号码', '数电发票号码', '销方识别号', '销方名称', '购方识别号', '购买方名称', '开票日期',
          '金额', '税额', '价税合计', '发票来源', '发票票种', '发票状态', '是否正数发票', '发票风险等级', '开票人', '备注']
(IV_SEQ, IV_CODE, IV_NO, IV_ENO, IV_STAX, IV_SELLER, IV_BTAX, IV_BUYER, IV_TIME, IV_NET, IV_TAX, IV_TOTAL,
 IV_SRC, IV_KIND, IV_STAT, IV_POS, IV_RISK, IV_ISSUER, IV_NOTE) = [CL(i) for i in range(1, 20)]
IV_PJ, IV_DIR, IV_UNIT, IV_CHK = 'T', 'U', 'V', 'W'          # T 手选项目；U～W 自动：方向、对方单位、校验
IV_DATE, IV_YM, IV_SPEC, IV_TAXU, IV_TOTU, IV_RATE = 'X', 'Y', 'Z', 'AA', 'AB', 'AC'
# X～AC 隐藏：开票日期（认出来的）、年月、可抵扣（专票·一般纳税人·一般计税项目＝1）、税额（用的）、价税合计（用的，作废＝0）、税率


def ivr(col):
    return rng(SH_INV, col, IV_R0, IV_R1)

# ─────────────────────────── 期初余额（建账日那天的数，只填一次） ───────────────────────────
SH_OPEN = '期初余额'
OP_HDR = 4
# 左：项目期初（一个项目一行，项目从下拉选）
OPJ_R0, OPJ_R1 = 5, 64
(OPJ_SEQ, OPJ_PJ, OPJ_REV, OPJ_REC, OPJ_INV, OPJ_LAB, OPJ_OTH, OPJ_TAX, OPJ_ALLOC, OPJ_MAT, OPJ_SUB, OPJ_MACH, OPJ_NOTE) = \
    [CL(i) for i in range(1, 14)]        # J～L 自动：材料、分包、机械（从右边应付期初按项目汇总）
# 右：应付期初（单位 × 项目）
OAP_R0, OAP_R1 = 5, 204
(OAP_SEQ, OAP_UNIT, OAP_PJ, OAP_AMT, OAP_PAID, OAP_INV, OAP_NOTE, OAP_TYPE) = [CL(i) for i in range(15, 23)]   # O..V
# 下：其他期初（竖着一项一行）
OO_R0 = OPJ_R1 + 4
OO_LBL, OO_VAL, OO_NOTE = 'B', 'C', 'D'
OO_ITEMS = ['短期借款', '保证金押金（付出去还没退的）', '其他应收款（别的）', '应交税费（欠税为正、多交为负）', '其他应付款（别的）',
            '实收资本（实缴）', '应付设备款（建账前赊购、还没付完的）', '老板确认的期初未分配利润（可不填）']
OO_BOSS = 7                      # 老板确认值是第几项（0 起）


def opr(col):
    return rng(SH_OPEN, col, OPJ_R0, OPJ_R1)


def oapr(col):
    return rng(SH_OPEN, col, OAP_R0, OAP_R1)


def oo(i):
    return f"{SH_OPEN}!${OO_VAL}${OO_R0 + i}"
AX_AKW_R = rng(SH_AUX, AX_AKW, 1, AX_AKW_N)
AX_AKA_R = rng(SH_AUX, AX_AKA, 1, AX_AKW_N)
AX_AS_R = rng(SH_AUX, AX_AS, 1, AX_AKW_N)
AX_AN_R = rng(SH_AUX, AX_AN, 1, AX_AKW_N)

# ─────────────────────────── 考勤工资（每人每月一行：项目＋天数成对填，一行 4 对，多了再加一行） ───────────────────────────
AT_HDR = 4
AT_R0, AT_R1 = 5, 1004          # 1000 行
AT_SEQ, AT_MON, AT_NAME = 'A', 'B', 'C'
AT_PJS = ['D', 'F', 'H', 'J']    # 项目1..4
AT_DDS = ['E', 'G', 'I', 'K']    # 天数1..4
(AT_URATE, AT_ALLOW, AT_ADJ, AT_TAX, AT_SOC, AT_NOTE) = 'L M N O P Q'.split()
(AT_DAYS, AT_RATE, AT_WAY, AT_BASE, AT_PAY, AT_NETPAY, AT_CHK) = 'R S T U V W X'.split()
# 隐藏：Y 年月 Z 生效日 AA 标准日薪 AB 标准月薪 AC..AF 项目1..4 金额 AG 间接部分（公司管理） AH 月中调价日 AI 人员类型
(AT_YM, AT_EFF, AT_SDAY, AT_SMON) = 'Y Z AA AB'.split()
AT_AMTS = ['AC', 'AD', 'AE', 'AF']
AT_COAMT, AT_MIDRAISE, AT_KIND = 'AG', 'AH', 'AI'
CO_PSEUDO = '公司管理'            # 考勤里「项目」可以填这个：算公司管理费


def atr(col):
    return rng(SH_ATT, col, AT_R0, AT_R1)


# ─────────────────────────── 首页的报表年度、截止日（所有报表跟着变） ───────────────────────────
SEL_Y = f"{SH_HOME}!$C$4"        # 报表年度
SEL_E = f"{SH_HOME}!$F$4"        # 截止日（空＝最后一笔流水那天）
# _辅助 J 列里算好的几个数
AX_YM0 = f"{SH_AUX}!$J$1"        # 本年 1 月（年月数 YYYYMM）
AX_YM1 = f"{SH_AUX}!$J$2"        # 截止月（年月数）
AX_END = f"{SH_AUX}!$J$3"        # 截止月月末（资产负债表日）
AX_OYM = f"{SH_AUX}!$J$4"        # 建账年月
AX_PYM = f"{SH_AUX}!$J$5"        # 上年 12 月（年月数）
AX_Y0 = f"{SH_AUX}!$J$6"         # 建账年份
AX_Y = f"{SH_AUX}!$J$7"          # 报表年度
AX_MON = f"{SH_AUX}!$J$8"        # 截止月是几月（本年已过月数）
AX_E = f"{SH_AUX}!$J$9"          # 截止日（用的）
AX_LAST = f"{SH_AUX}!$J$10"      # 最后一笔流水日期

# ─────────────────────────── 隐藏的汇总层（各报表都从这里取数，只算一遍） ───────────────────────────
SH_MS = '_月汇总'                # 公司利润表项目 × 月
MS_YMROW = 2
MS_R0 = 3                        # 从第 3 行起：一个利润表项目一行（顺序同 cats.IS_NAMES）
MS_NAME, MS_PREV = 'A', 'B'      # B：建账起到上年底
MS_MCOLS = [CL(3 + i) for i in range(12)]      # C..N：本年 1～12 月
MS_YTD = 'O'                     # 本年累计（到截止月）
SH_PS = '_项目汇总'              # 每个项目 × 各种数
PS_R0 = 3                        # 第 3～62 行对应【项目档案】第 4～63 行
PS_NAME = 'A'
PS_MEAS = ['确认收入', '现金回款', '抵账', '开票', '销项税', '进项税', '材料', '人工', '分包', '机械', '其他直接', '甲方扣款', '贴息']
PS_COL = {}                      # (量, '前'/'本') → 列；前＝建账起到上年底，本＝本年到截止月
_c = 2
for _m in PS_MEAS:
    PS_COL[(_m, '前')] = CL(_c)
    PS_COL[(_m, '本')] = CL(_c + 1)
    _c += 2
PS_LABM = [CL(_c + i) for i in range(12)]      # 本年各月人工（用工成本表）
_c += 12
NYEARS = 8
PS_DRV = [CL(_c + i) for i in range(NYEARS)]    # 建账年起每年的分摊依据（到截止月为止）
_c += NYEARS
PS_AR_E, PS_AR_B, PS_OPEN = CL(_c), CL(_c + 1), CL(_c + 2)   # 应收余额（截止月末、年初）、期初应收（建账前）
_c += 3
PS_LAST = CL(_c - 1)
SH_BALX = '_余额'               # 每个单位/人、每个账户在截止月末和年初的余额
BX_R0 = 3                        # 第 3～302 行对应【往来单位】
(BX_NAME, BX_TYPE, BX_AP_E, BX_AP_B, BX_WG_E, BX_WG_B, BX_APAMT, BX_PAID, BX_INVIN, BX_NETPAY, BX_WPAID, BX_OFFW, BX_REFUND) = \
    [CL(i) for i in range(1, 14)]
BA_R0 = 3                        # 右边：第 3～22 行对应【基础资料】② 资金账户
(BA_NAME, BA_TYPE, BA_OWNER, BA_E, BA_B) = [CL(i) for i in range(16, 21)]   # P..T
