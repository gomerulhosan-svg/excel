# -*- coding: utf-8 -*-
"""A066 第二版 · 全书地址契约。所有模块只从这里拿表名、行列、容量、定义名称。
   规则见 设计.md §0：录入表只放录入列；逐行计算在隐藏表（按第 n 条取数）；报表只用定义名称做 SUMIFS/COUNTIFS。"""
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI

# ───────────── 表名 ─────────────
SH_HOME = '首页'
SH_CASH, SH_AR_IN, SH_AP_IN, SH_ATT, SH_INV = '收支登记', '应收登记', '应付登记', '考勤工资', '发票登记'
SH_PJ, SH_CUS, SH_SUP, SH_PER, SH_RATE, SH_BASE, SH_OPEN = '项目档案', '甲方信息', '供应商信息', '人员信息', '工资标准', '基本信息', '期初余额'
SH_FUND, SH_ACCT = '资金报表', '账户明细'
SH_AR, SH_ARD, SH_AP, SH_APD = '应收账款', '应收对账', '应付账款', '应付对账'
SH_PERSON, SH_PAY, SH_PAYSUM, SH_LABOR, SH_ARREAR, SH_INVSUM = '个人往来', '工资表', '工资汇总', '用工成本表', '欠薪补发', '发票统计'
SH_PJBOOK, SH_PJPL, SH_ALLOC = '项目账', '项目利润表', '费用分摊'
SH_PL, SH_BS, SH_BE, SH_CHK = '利润表', '资产负债表', '盈亏平衡表', '数据校验'
H_PAR, H_SHOU, H_YS, H_YF, H_KQ, H_PIAO, H_LIST, H_SUM = '_参数', '_收', '_应', '_付', '_考', '_票', '_表', '_汇'

C_HOME, C_IN, C_MASTER, C_VIEW, C_PROJ, C_RPT = 'FF1F3864', 'FF2F75B5', 'FF7F7F7F', 'FF548235', 'FFC65911', 'FF833C0C'
GROUPS = [
    ('首页', C_HOME, [SH_HOME]),
    ('录入', C_IN, [SH_CASH, SH_AR_IN, SH_AP_IN, SH_ATT, SH_INV]),
    ('基本信息', C_MASTER, [SH_PJ, SH_CUS, SH_SUP, SH_PER, SH_RATE, SH_BASE, SH_OPEN]),
    ('查看', C_VIEW, [SH_FUND, SH_ACCT, SH_AR, SH_ARD, SH_AP, SH_APD, SH_PERSON, SH_PAY, SH_PAYSUM, SH_LABOR, SH_ARREAR, SH_INVSUM]),
    ('项目', C_PROJ, [SH_PJBOOK, SH_PJPL, SH_ALLOC]),
    ('报表', C_RPT, [SH_PL, SH_BS, SH_BE, SH_CHK]),
]
HIDDEN = [H_PAR, H_SHOU, H_YS, H_YF, H_KQ, H_PIAO, H_LIST, H_SUM]

HOME_YEAR, HOME_END = 'C4', 'F4'           # 首页：报表年度、截止日期（黄格；空＝最后一笔流水那年/那天）

# ───────────── 容量 ─────────────
N_CASH, N_AR, N_AP, N_ATT, N_INV = 5000, 1000, 4000, 2000, 2000
N_PJ, N_CUS, N_SUP, N_PER, N_RATE, N_OPEN = 200, 200, 500, 500, 1000, 300
N_ACC, N_INC, N_EXP = 40, 40, 100
NPAIR = 8                                   # 考勤每行「项目＋天数」对数
YEAR0, NYEARS = 2023, 8                     # 年度 2023～2030
REF_END = 20000                             # 录入表按第 n 条取数时，锚定区域的末行（远大于容量，插行也够用）

# ───────────── 录入表：表头行、首条数据行、列 ─────────────
# 收支登记
CASH_HDR, CASH_R0 = 7, 8
CASH_COLS = dict(日期='A', 类别='B', 账户='C', 收支项目='D', 摘要='E', 收入='F', 支出='G', 结余='H', 项目='I', 负责人='J',
                 客户='K', 供应商='L', 人员='M', 对方账户='N', 费用归属='O', 已开票='P', 开票日期='Q', 备注='R')
CASH_LAST = 'R'
# 应收登记
AR_HDR, AR_R0 = 4, 5
AR_COLS = dict(日期='A', 项目='B', 客户='C', 类型='D', 金额='E', 税率='F', 发票号='G', 备注='H')
AR_LAST = 'H'
AR_TYPES = ['补充协议', '签证', '扣款', '结算调整', '确认产值', '开票']
# 应付登记
AP_HDR, AP_R0 = 4, 5
AP_COLS = dict(日期='A', 项目='B', 供应商='C', 类型='D', 摘要='E', 数量='F', 单价='G', 应付='H', 已开票='I', 开票日期='J',
               税率='K', 发票类型='L', 备注='M')
AP_LAST = 'M'
AP_TYPES = ['材料', '分包', '机械', '其他']
# 考勤工资
ATT_HDR, ATT_R0 = 4, 5
ATT_MON, ATT_NAME, ATT_TEAM = 'A', 'B', 'C'
ATT_PJS = [CL(4 + 2 * i) for i in range(NPAIR)]          # 项目1～8：D F H J L N P R
ATT_DAYS = [CL(5 + 2 * i) for i in range(NPAIR)]         # 天数1～8：E G I K M O Q S
_a = 4 + 2 * NPAIR                                        # 20 → T
ATT_RATE, ATT_EXTRA, ATT_DED, ATT_NOTE = CL(_a), CL(_a + 1), CL(_a + 2), CL(_a + 3)          # T U V W
ATT_SHOW_DAYS, ATT_SHOW_RATE, ATT_SHOW_AMT = CL(_a + 4), CL(_a + 5), CL(_a + 6)             # X Y Z（只显示）
ATT_LAST = ATT_SHOW_AMT
# 发票登记（同第一版：电子税务局「发票基础信息」A～S）
INV_HDR, INV_R0 = 4, 5
INV_HEAD = ['序号', '发票代码', '发票号码', '数电发票号码', '销方识别号', '销方名称', '购方识别号', '购买方名称', '开票日期',
            '金额', '税额', '价税合计', '发票来源', '发票票种', '发票状态', '是否正数发票', '发票风险等级', '开票人', '备注']
INV_LAST = CL(len(INV_HEAD))                              # S
INV_PJ = CL(len(INV_HEAD) + 1)                            # T 选项目（唯一要手选的列）
INV_COLS = dict(序号='A', 销方税号='E', 销方='F', 购方税号='G', 购方='H', 日期='I', 金额='J', 税额='K', 价税合计='L', 状态='O', 项目=INV_PJ)

# ───────────── 基本信息类表（表头第 4 行，数据第 5 行起） ─────────────
M_HDR, M_R0 = 4, 5
PJ_COLS = dict(名称='A', 全称='B', 甲方='C', 合同='D', 保证金='E', 质保比例='F', 质保金额='G', 质保到期='H', 开工='I', 完工='J',
               状态='K', 类型='L', 负责人='M', 税率='N', 备注='O')
PJ_LAST = 'O'
PJ_TYPES = ['工程', '公司']
PJ_STATES = ['在建', '完工', '已结算', '停工']
CUS_COLS = dict(名称='A', 简称='B', 税号='C', 联系人='D', 电话='E', 开户='F', 备注='G')
CUS_LAST = 'G'
SUP_COLS = dict(名称='A', 类型='B', 开票单位='C', 税号='D', 税率='E', 发票类型='F', 联系人='G', 电话='H', 收款账户='I', 备注='J')
SUP_LAST = 'J'
PER_COLS = dict(姓名='A', 类别='B', 工种='C', 计薪='D', 过账='E', 期初欠薪='F', 身份证='G', 电话='H', 银行卡='I', 备注='J')
PER_LAST = 'J'
PER_KINDS = ['工人', '班组长', '管理人员', '老板', '临时工']
PAY_MODES = ['日薪', '月薪']
RATE_COLS = dict(姓名='A', 生效='B', 日单价='C', 月薪='D', 备注='E')
RATE_LAST = 'E'
OPEN_COLS = dict(类型='A', 对象='B', 项目='C', 金额='D', 说明='E')
OPEN_LAST = 'E'
OPEN_TYPES = ['应收账款', '应付账款', '应付工资', '个人借款', '个人垫付', '备用金', '保证金押金', '代收代付', '其他应收',
              '其他应付', '短期借款', '固定资产', '应交税费', '实收资本', '未分配利润']

# 基本信息（参数、账户、收支项目）
BASE_PAR_R0 = 5                                            # ① 参数：B 列名称、C 列值
PARAMS = [  # (键, 显示名, 默认值, 格式, 下拉)
    ('公司', '公司名称', '九江竣辉建设工程有限公司（请改成全称）', None, None),
    ('税号', '公司税号（发票登记分销项/进项用）', '', None, None),
    ('建账日', '建账日期', '2023-05-01', 'date', None),
    ('纳税人', '纳税人类型', '一般纳税人', None, '一般纳税人,小规模纳税人'),
    ('销项税率', '销项税率（默认）', 0.09, 'pct', None),
    ('征收率', '小规模征收率', 0.03, 'pct', None),
    ('附加税率', '附加税费率（城建＋教育）', 0.12, 'pct', None),
    ('收入口径', '收入确认口径', '开票与收款取大', None, '开票与收款取大,开票,收款,确认产值'),
    ('人工口径', '人工成本口径', '考勤应发', None, '考勤应发,实际发放'),
    ('分摊依据', '管理费分摊依据', '施工费', None, '施工费,直接成本'),
    ('考勤起算', '考勤起算月份', '2025-03-01', 'date', None),
]
PAR_ROW = {k: BASE_PAR_R0 + i for i, (k, *_r) in enumerate(PARAMS)}   # 参数在 基本信息!C{row}
BASE_ACC_HDR = 18                                          # ② 账户：A 名称 B 类型 C 所属人 D 期初余额 E 备注
BASE_ACC_R0 = BASE_ACC_HDR + 1
ACC_COLS = dict(名称='A', 类型='B', 所属人='C', 期初='D', 备注='E')
ACC_TYPES = ['银行', '现金', '专户', '个人户']
BASE_INC_HDR = BASE_ACC_R0 + N_ACC + 2                     # ③ 收入类收支项目：G 名称 H 归类
BASE_EXP_HDR = BASE_INC_HDR                                # ④ 支出类收支项目：J 名称 K 归类 L 默认费用归属
INC_COLS = dict(名称='G', 归类='H')
EXP_COLS = dict(名称='J', 归类='K', 归属='L')
# 为了好看，③④放在账户表右边同一高度
BASE_INC_HDR = BASE_EXP_HDR = BASE_ACC_HDR
BASE_INC_R0 = BASE_EXP_R0 = BASE_ACC_HDR + 1
BASE_CAT_HDR = BASE_EXP_R0 + N_EXP + 2                     # ⑤ 归类说明（只读，下拉来源）：A 归类 B 方向 C 去向 D 默认归属
BASE_CAT_R0 = BASE_CAT_HDR + 1
CATS = [  # (归类, 方向, 去向说明, 默认费用归属, 是否成本类)
    ('工程款收款', '收', '冲应收（按项目）', '', 0),
    ('其他收入', '收', '利润表·营业外收入（利息收入、杂项）', '', 0),
    ('材料款', '付', '供应商有应付登记→冲应付；没有→项目成本·材料', '项目成本', 1),
    ('分包款', '付', '供应商有应付登记→冲应付；没有→项目成本·分包', '项目成本', 1),
    ('机械费', '付', '供应商有应付登记→冲应付；没有→项目成本·机械', '项目成本', 1),
    ('工资发放', '付', '有考勤的人（考勤起算月以后）→冲应付工资；否则→人工成本/管理工资（工资退回记负数或记收入）', '项目成本', 1),
    ('项目其他费用', '付', '项目成本·其他直接费', '项目成本', 1),
    ('管理费用', '付', '管理费用（办公、招待、车辆、伙食、公关红包…）', '待摊费用', 1),
    ('财务费用', '付', '财务费用（贷款利息、手续费、贴息）', '公司费用', 1),
    ('营业外支出', '付', '营业外支出（罚款、滞纳金）', '公司费用', 1),
    ('税费', '付', '冲应交税费（交增值税及附加）', '', 0),
    ('个人借款', '双向', '其他应付·个人借款：收＝借入、付＝归还（人员空＝收支项目名「个人借款-XX」的 XX）', '', 0),
    ('银行借款', '双向', '短期借款：收＝借入、付＝还本', '', 0),
    ('保证金押金', '双向', '其他应收·保证金：付＝交出、收＝退回', '', 0),
    ('备用金', '双向', '其他应收·备用金（人员）：付＝领、收＝退', '', 0),
    ('代收代付', '双向', '其他应收·过账（人员，如替员工交社保、员工退回）', '', 0),
    ('往来款', '双向', '其他应收/其他应付（单位或人员）', '', 0),
    ('报销还款', '付', '冲个人垫付（人员）：公司把垫付的钱还给他', '', 0),
    ('内部转账', '双向', '账户之间倒钱（填对方账户），不进报表', '', 0),
    ('股东投入', '收', '实收资本', '', 0),
    ('固定资产购置', '付', '固定资产（管理口径不提折旧）', '', 0),
]
CAT_NAMES = [c[0] for c in CATS]
COST_CATS = [c[0] for c in CATS if c[4]]
GUISHU = ['项目成本', '待摊费用', '公司费用']
BASE_LAST = 'L'

# ───────────── 隐藏表 _收（收支登记逐条；第 n 条在第 n+1 行，n=1..N_CASH） ─────────────
SHOU_COLS = [  # (列, 字段, 说明)   金额类列只在「有效」时有值，报表直接 SUMIFS 不用再判有效
    ('A', 'n', '第几条（常数）'),
    ('B', '日期', '真日期；无效＝0'),
    ('C', '年月', 'YYYYMM；无效＝0'),
    ('D', '类别', ''),
    ('E', '账户', ''),
    ('F', '收支项目', ''),
    ('G', '摘要', ''),
    ('H', '收入', '数值（有效才有）'),
    ('I', '支出', '数值（有效才有）'),
    ('J', '净额', '收入−支出（钱进为正；有效才有）'),
    ('K', '项目原', '录入的项目名称'),
    ('L', '客户', '录入的客户；空且有工程项目＝项目档案的甲方'),
    ('M', '供应商', ''),
    ('N', '人员原', '录入的人员'),
    ('O', '对方账户', ''),
    ('P', '归属原', '录入的费用归属'),
    ('Q', '已开票', '数值（有效才有）'),
    ('R', '开票日期', '无＝0'),
    ('S', '有效', '1＝日期、账户（在账户表里）都有且金额≠0'),
    ('T', '归类', '生效的归类：内部转账但对方账户不对＝""（按未分类进利润表）'),
    ('U', '账户类型', '银行/现金/专户/个人户'),
    ('V', '账户人', '个人户的所属人'),
    ('W', '项目', '项目名称是「工程」类项目才保留，否则""'),
    ('X', '人员', '人员原；空且归类＝个人借款时取收支项目名「-」后面的字'),
    ('Y', '冲应付', '1＝归类是材料款/分包款/机械费且供应商在应付登记里有行'),
    ('Z', '冲工资', '1＝归类是工资发放、人员在考勤里有行、日期≥考勤起算月、不是过账'),
    ('AA', '过账', '1＝人员是过账人员且归类是工资发放/代收代付'),
    ('AB', '成本类', '材料/分包/机械/人工/其他直接费/管理费用/财务费用/营业外支出/未分类；不进利润表＝""'),
    ('AC', '归属', '项目成本/待摊费用/公司费用（只对成本类行）'),
    ('AD', '校验', '""＝没问题；✗/⚠ 开头的文字'),
    ('AE', '录入行', '收支登记里的实际行号'),
    ('AF', '往来对象', '人员优先，其次供应商、客户'),
    ('AG', '归类原', '收支项目查到的归类（不管对方账户）'),
    ('AH', '净额原', '收入−支出（不管有效）'),
    ('AI', '开票年月', ''),
    ('AJ', '基数', '这笔算进管理费分摊依据的金额（项目成本行：施工费＝人工/分包/机械；直接成本＝全部）'),
    ('AK', '摊费', '基数×这一年的管理费率（_汇 年度表）'),
]
SHOU = {f: c for c, f, _ in SHOU_COLS}
SHOU_LAST = SHOU_COLS[-1][0]

# _应（应收登记逐条，第 n 条在 n+1 行）
YS_COLS = [('A', 'n', ''), ('B', '日期', ''), ('C', '年月', ''), ('D', '项目', '项目名称（工程类才保留）'),
           ('E', '客户', '空＝项目档案甲方'), ('F', '类型', ''), ('G', '金额', '数值'), ('H', '税率', '空＝项目税率或默认'),
           ('I', '应收额', '补充协议/签证/结算调整＝金额，扣款＝−金额，其他 0'), ('J', '开票额', '类型＝开票时＝金额'),
           ('K', '产值额', '类型＝确认产值时＝金额'), ('L', '销项税', '开票额×税率/(1+税率)'), ('M', '有效', ''),
           ('N', '校验', ''), ('O', '录入行', '')]
YS = {f: c for c, f, _ in YS_COLS}
# _付（应付登记逐条）
YF_COLS = [('A', 'n', ''), ('B', '日期', ''), ('C', '年月', ''), ('D', '项目', '工程类才保留'), ('E', '供应商', ''),
           ('F', '类型', '空＝供应商信息的类型'), ('G', '应付额', '空＝数量×单价'), ('H', '已开票', ''), ('I', '开票日期', '空＝日期'),
           ('J', '税率', '空＝供应商默认'), ('K', '发票类型', '空＝供应商默认'), ('L', '进项税', '专票：已开票×税率/(1+税率)'),
           ('M', '开票年月', ''), ('N', '有效', ''), ('O', '校验', ''), ('P', '录入行', ''), ('Q', '摘要', ''),
           ('R', '基数', '有工程项目：施工费口径＝分包/机械的应付额；直接成本口径＝全部应付额'), ('S', '摊费', '基数×这一年的费率')]
YF = {f: c for c, f, _ in YF_COLS}
# _考（考勤逐条；8 对展开）
KQ_FIXED = [('A', 'n', ''), ('B', '月', '当月 1 号；无效＝0'), ('C', '年月', ''), ('D', '姓名', ''), ('E', '过账', '1/0'),
            ('F', '计薪', '日薪/月薪'), ('G', '总天数', ''), ('H', '人月天数', '同人同月各行天数合计（月薪分摊用）'),
            ('I', '单价', '手填或工资标准（日单价或月薪）'),
            ('J', '应发', '有效且不是过账：日薪＝单价×天数＋补贴−扣款；月薪＝月薪×本行天数/人月天数＋补贴−扣款'),
            ('K', '有效', ''), ('L', '未分摊', '应发里没对上工程项目的部分（总天数 0、或项目是公司/不认识）→ 公司管理工资'),
            ('M', '校验', ''), ('N', '录入行', ''), ('O', '过账应发', '过账人员按同样算法的应发（只看，不进成本）')]
KQ_P0 = 16                                                  # 第 i 对（0 起）：项目 CL(16+3i)，天数 CL(17+3i)，金额 CL(18+3i)（金额只算工程项目）
KQ_PJ = [CL(KQ_P0 + 3 * i) for i in range(NPAIR)]
KQ_DAY = [CL(KQ_P0 + 1 + 3 * i) for i in range(NPAIR)]
KQ_AMT = [CL(KQ_P0 + 2 + 3 * i) for i in range(NPAIR)]
KQ = {f: c for c, f, _ in KQ_FIXED}
_x = KQ_P0 + 3 * NPAIR
KQ_JT, KQ_BASE = CL(_x), CL(_x + 1)                         # 计提（1＝进成本和应付工资）、基数（计提×工程项目金额合计）
KQ_TAN = [CL(_x + 2 + i) for i in range(NPAIR)]             # 摊1..摊8＝计提×金额i×这一年的费率
KQ['计提'], KQ['基数'] = KQ_JT, KQ_BASE
KQ_LAST = KQ_TAN[-1]
# _票（发票登记逐条）
PIAO_COLS = [('A', 'n', ''), ('B', '日期', ''), ('C', '年月', ''), ('D', '方向', '销项/进项（按销方税号是不是本公司）'),
             ('E', '对方', ''), ('F', '金额', '不含税'), ('G', '税额', ''), ('H', '价税合计', ''), ('I', '税率', '税额/金额'),
             ('J', '项目', ''), ('K', '有效', '作废、表头、合计行＝0'), ('L', '录入行', '')]
PIAO = {f: c for c, f, _ in PIAO_COLS}

# ───────────── _表：基本信息逐行镜像（第 i 个在第 i+1 行） ─────────────
# 项目：A 名称 B 类型 C 甲方 D 合同 E 保证金 F 质保比例 G 质保金额 H 质保到期 I 开工 J 完工 K 状态 L 税率 M 开工(有效：空＝最早一笔业务日)
# 供应商：O 名称 P 类型 Q 默认税率 R 发票类型；人员：T 姓名 U 类别 V 计薪 W 过账 X 期初欠薪；
# 账户：Z 名称 AA 类型 AB 所属人 AC 期初；甲方：AE 名称；收入项目：AG 名称 AH 归类；支出项目：AJ 名称 AK 归类 AL 默认归属
LIST_PJ = dict(名称='A', 类型='B', 甲方='C', 合同='D', 保证金='E', 质保比例='F', 质保金额='G', 质保到期='H', 开工='I', 完工='J',
               状态='K', 税率='L', 开工有效='M')
LIST_SUP = dict(名称='O', 类型='P', 税率='Q', 发票类型='R')
LIST_PER = dict(姓名='T', 类别='U', 计薪='V', 过账='W', 期初欠薪='X')
LIST_ACC = dict(名称='Z', 类型='AA', 所属人='AB', 期初='AC')
LIST_CUS = dict(名称='AE')
LIST_INC = dict(名称='AG', 归类='AH')
LIST_EXP = dict(名称='AJ', 归类='AK', 归属='AL')
LIST_RATE = dict(姓名='AN', 生效='AO', 日单价='AP', 月薪='AQ', 序='AR')   # 工资标准：序＝同一人按生效日排的第几条
LIST_N = 500                                                 # _表 有 500 行（每类按各自容量；工资标准到 N_RATE 行）

# ───────────── _参数：全书共用期间（B 列值） ─────────────
PAR = dict(年度='B2', 截止='B3', 年初='B4', 最后流水日='B5', 建账日='B6', 考勤起算='B7', 起算年月='B8', 收入口径='B9',
           人工口径='B10', 分摊依据='B11', 纳税人='B12', 销项税率='B13', 附加税率='B14', 征收率='B15', 公司='B16',
           截止年月='B17', 年初年月='B18', 税号='B19')

# ───────────── 定义名称（报表只用这些） ─────────────
NAMES = {}
for c, f, _d in SHOU_COLS:
    NAMES[f'收_{f}'] = f'{H_SHOU}!${c}$2:${c}${N_CASH + 1}'
for c, f, _d in YS_COLS:
    NAMES[f'应_{f}'] = f'{H_YS}!${c}$2:${c}${N_AR + 1}'
for c, f, _d in YF_COLS:
    NAMES[f'付_{f}'] = f'{H_YF}!${c}$2:${c}${N_AP + 1}'
for c, f, _d in KQ_FIXED:
    NAMES[f'考_{f}'] = f'{H_KQ}!${c}$2:${c}${N_ATT + 1}'
for i in range(NPAIR):
    NAMES[f'考_项目{i + 1}'] = f'{H_KQ}!${KQ_PJ[i]}$2:${KQ_PJ[i]}${N_ATT + 1}'
    NAMES[f'考_天数{i + 1}'] = f'{H_KQ}!${KQ_DAY[i]}$2:${KQ_DAY[i]}${N_ATT + 1}'
    NAMES[f'考_金额{i + 1}'] = f'{H_KQ}!${KQ_AMT[i]}$2:${KQ_AMT[i]}${N_ATT + 1}'
NAMES['考_计提'] = f'{H_KQ}!${KQ_JT}$2:${KQ_JT}${N_ATT + 1}'
NAMES['考_基数'] = f'{H_KQ}!${KQ_BASE}$2:${KQ_BASE}${N_ATT + 1}'
for i in range(NPAIR):
    NAMES[f'考_摊{i + 1}'] = f'{H_KQ}!${KQ_TAN[i]}$2:${KQ_TAN[i]}${N_ATT + 1}'
for c, f, _d in PIAO_COLS:
    NAMES[f'票_{f}'] = f'{H_PIAO}!${c}$2:${c}${N_INV + 1}'
for k, c in LIST_PJ.items():
    NAMES[f'项目_{k}'] = f'{H_LIST}!${c}$2:${c}${N_PJ + 1}'
for k, c in LIST_SUP.items():
    NAMES[f'供应商_{k}'] = f'{H_LIST}!${c}$2:${c}${N_SUP + 1}'
for k, c in LIST_PER.items():
    NAMES[f'人员_{k}'] = f'{H_LIST}!${c}$2:${c}${N_PER + 1}'
for k, c in LIST_ACC.items():
    NAMES[f'账户_{k}'] = f'{H_LIST}!${c}$2:${c}${N_ACC + 1}'
NAMES['甲方_名称'] = f'{H_LIST}!$AE$2:$AE${N_CUS + 1}'
NAMES['收入项目_名称'] = f'{H_LIST}!$AG$2:$AG${N_INC + 1}'
NAMES['收入项目_归类'] = f'{H_LIST}!$AH$2:$AH${N_INC + 1}'
NAMES['支出项目_名称'] = f'{H_LIST}!$AJ$2:$AJ${N_EXP + 1}'
NAMES['支出项目_归类'] = f'{H_LIST}!$AK$2:$AK${N_EXP + 1}'
NAMES['支出项目_归属'] = f'{H_LIST}!$AL$2:$AL${N_EXP + 1}'
NAMES['工资标准_姓名'] = f"{SH_RATE}!$A${M_R0}:$A${M_R0 + N_RATE - 1}"
NAMES['工资标准_生效'] = f"{SH_RATE}!$B${M_R0}:$B${M_R0 + N_RATE - 1}"
NAMES['工资标准_日单价'] = f"{SH_RATE}!$C${M_R0}:$C${M_R0 + N_RATE - 1}"
NAMES['工资标准_月薪'] = f"{SH_RATE}!$D${M_R0}:$D${M_R0 + N_RATE - 1}"
for k, c in LIST_RATE.items():
    NAMES[f'标准_{k}'] = f'{H_LIST}!${c}$2:${c}${N_RATE + 1}'
NAMES['期初_类型'] = f"{SH_OPEN}!$A${M_R0}:$A${M_R0 + N_OPEN - 1}"
NAMES['期初_对象'] = f"{SH_OPEN}!$B${M_R0}:$B${M_R0 + N_OPEN - 1}"
NAMES['期初_项目'] = f"{SH_OPEN}!$C${M_R0}:$C${M_R0 + N_OPEN - 1}"
NAMES['期初_金额'] = f"{SH_OPEN}!$D${M_R0}:$D${M_R0 + N_OPEN - 1}"
for k, cell in PAR.items():
    col, row = cell[0], cell[1:]
    NAMES[f'P_{k}'] = f'{H_PAR}!${col}${row}'
# ───────────── _汇：全书共用的汇总（收入确认要逐项目取大、应收应付按对象分正负，只能先逐对象算） ─────────────
# 报表自己的日期输入格（_汇 直接读这几个固定格，所以 s_fin 必须把黄格放在这里）
PL_IN_START, PL_IN_END = f"{SH_PL}!$C$3", f"{SH_PL}!$E$3"     # 利润表「本期」起、止（空＝年初、截止）
BS_IN_DATE = f"{SH_BS}!$C$3"                                   # 资产负债表日期（空＝截止）
SUM_Y0 = 3                                                     # ① 年度费率表：第 3～10 行＝YEAR0..YEAR0+7
SUM_YCOLS = dict(年='A', 待摊池收='B', 待摊池考='C', 待摊池='D', 基数='E', 费率='F', 已摊='G', 未摊='H')
SUM_D0 = 14                                                    # ② 日期表：A 名称 B 日期（第 14～29 行）
SUM_DATES = ['年初前'] + [f'{m}月' for m in range(1, 13)] + ['利润起前', '利润止', '负债日']
SUM_DROW = {k: SUM_D0 + i for i, k in enumerate(SUM_DATES)}
SUM_EFF = dict(利润起='B31', 利润止='B32', 负债日='B33')       # 有效日期（报表显示用）
SUM_PJ_HDR, SUM_PJ0 = 39, 40                                   # ③ 项目块：第 40 行起 N_PJ 行
SUM_PJ_LIST, SUM_PJ_ALL, SUM_PJ_NONE, SUM_PJ_CO = 40 + N_PJ, 41 + N_PJ, 42 + N_PJ, 43 + N_PJ
SUM_PJ_K = {k: [CL(3 + 4 * i + j) for j in range(4)] for i, k in enumerate(SUM_DATES)}   # 每个日期：开票、收款、产值、收入
SUM_PJ_AR0, SUM_PJ_AR1 = CL(3 + 4 * len(SUM_DATES)), CL(4 + 4 * len(SUM_DATES))         # 应收余额：年初前、负债日
SUM_SUP_HDR = SUM_PJ_CO + 6                                    # ④ 供应商块：A 名称 B 期初应付 C 年初前余额 D 负债日余额
SUM_SUP0 = SUM_SUP_HDR + 1
SUM_SUP_LIST, SUM_SUP_ALL, SUM_SUP_NONE = SUM_SUP0 + N_SUP, SUM_SUP0 + N_SUP + 1, SUM_SUP0 + N_SUP + 2
SUM_ACC_HDR = SUM_SUP_NONE + 6                                 # ⑤ 账户块：A 名称 B 类型 C 所属人 D 年初前余额 E 负债日余额
SUM_ACC0 = SUM_ACC_HDR + 1
SUM_PER_HDR = SUM_ACC0 + N_ACC + 5                             # ⑥ 人员块（垫付/持有）：A 姓名 B 期初垫付 C 年初前 D 负债日（正＝他手上有公司的钱）
SUM_PER0 = SUM_PER_HDR + 1
SUM_PER_LIST, SUM_PER_ALL, SUM_PER_NONE = SUM_PER0 + N_PER, SUM_PER0 + N_PER + 1, SUM_PER0 + N_PER + 2
SUM_RES = ['应收年初前', '预收年初前', '应收负债日', '预收负债日', '应付年初前', '预付年初前', '应付负债日', '预付负债日',
           '货币资金年初前', '货币资金负债日', '个人户余额年初前', '个人户余额负债日', '个人持有年初前', '个人垫付年初前',
           '个人持有负债日', '个人垫付负债日']
SUM_RES0 = 3                                                   # ⑦ 结果：K 名称 L 值，第 3 行起
SUM_RROW = {k: SUM_RES0 + i for i, k in enumerate(SUM_RES)}
for k, c in SUM_YCOLS.items():
    NAMES[f'汇_{k}'] = f'{H_SUM}!${c}${SUM_Y0}:${c}${SUM_Y0 + NYEARS - 1}'
for k, r in SUM_DROW.items():
    NAMES[f'汇_日{k}'] = f'{H_SUM}!$B${r}'
    NAMES[f'汇_收入{k}'] = f'{H_SUM}!${SUM_PJ_K[k][3]}${SUM_PJ_CO}'
for k, cell in SUM_EFF.items():
    NAMES[f'汇_{k}'] = f'{H_SUM}!${cell[0]}${cell[1:]}'
for k, r in SUM_RROW.items():
    NAMES[f'汇_{k}'] = f'{H_SUM}!$L${r}'
NAMES['汇_账户名'] = f'{H_SUM}!$A${SUM_ACC0}:$A${SUM_ACC0 + N_ACC - 1}'
NAMES['汇_账户余额年初前'] = f'{H_SUM}!$D${SUM_ACC0}:$D${SUM_ACC0 + N_ACC - 1}'
NAMES['汇_账户余额负债日'] = f'{H_SUM}!$E${SUM_ACC0}:$E${SUM_ACC0 + N_ACC - 1}'


def rev_formula(kp, sk, cz):
    """按【基本信息】收入口径算累计确认收入：kp 累计开票、sk 累计工程款收款、cz 累计确认产值（都是公式片段或单元格）"""
    return (f'IF(P_收入口径="开票",{kp},IF(P_收入口径="收款",{sk},IF(P_收入口径="确认产值",{cz},'
            f'IF({cz}<>0,{cz},MAX({kp},{sk})))))')


def cum_kp(pj, d):
    """项目 pj 到 d（含）为止的累计开票（应收登记 类型＝开票）"""
    return f'SUMIFS(应_开票额,应_项目,{pj},应_日期,"<="&{d})'


def cum_sk(pj, d):
    """项目 pj 到 d 为止的累计工程款收款（收支登记 归类＝工程款收款，含专户）"""
    return f'SUMIFS(收_净额,收_归类,"工程款收款",收_项目,{pj},收_日期,"<="&{d})'


def cum_cz(pj, d):
    return f'SUMIFS(应_产值额,应_项目,{pj},应_日期,"<="&{d})'


def rate_of(date_expr):
    """某天所在年度的管理费率"""
    return f'IFERROR(INDEX(汇_费率,YEAR({date_expr})-{YEAR0}+1),0)'


# 下拉来源（OFFSET 动态，只给数据验证用）
DV_NAMES = {
    '收入': f"OFFSET({SH_BASE}!${INC_COLS['名称']}${BASE_INC_R0},0,0,MAX(1,COUNTA({SH_BASE}!${INC_COLS['名称']}${BASE_INC_R0}:${INC_COLS['名称']}${BASE_INC_R0 + N_INC - 1})),1)",
    '支出': f"OFFSET({SH_BASE}!${EXP_COLS['名称']}${BASE_EXP_R0},0,0,MAX(1,COUNTA({SH_BASE}!${EXP_COLS['名称']}${BASE_EXP_R0}:${EXP_COLS['名称']}${BASE_EXP_R0 + N_EXP - 1})),1)",
    '账户列表': f"OFFSET({SH_BASE}!$A${BASE_ACC_R0},0,0,MAX(1,COUNTA({SH_BASE}!$A${BASE_ACC_R0}:$A${BASE_ACC_R0 + N_ACC - 1})),1)",
    '项目列表': f"OFFSET({SH_PJ}!$A${M_R0},0,0,MAX(1,COUNTA({SH_PJ}!$A${M_R0}:$A${M_R0 + N_PJ - 1})),1)",
    '甲方列表': f"OFFSET({SH_CUS}!$A${M_R0},0,0,MAX(1,COUNTA({SH_CUS}!$A${M_R0}:$A${M_R0 + N_CUS - 1})),1)",
    '供应商列表': f"OFFSET({SH_SUP}!$A${M_R0},0,0,MAX(1,COUNTA({SH_SUP}!$A${M_R0}:$A${M_R0 + N_SUP - 1})),1)",
    '人员列表': f"OFFSET({SH_PER}!$A${M_R0},0,0,MAX(1,COUNTA({SH_PER}!$A${M_R0}:$A${M_R0 + N_PER - 1})),1)",
}


def inref(sh, col, last=REF_END, hdr=None):
    """录入表某列的锚定区域（从表头行开始），配合 INDEX(区域, n+1) 用"""
    h = hdr if hdr is not None else {SH_CASH: CASH_HDR, SH_AR_IN: AR_HDR, SH_AP_IN: AP_HDR, SH_ATT: ATT_HDR, SH_INV: INV_HDR}.get(sh, M_HDR)
    return f"{sh}!${col}${h}:${col}${last}"


def std_rate(name, eom, monthly):
    """某人在某天（月末）有效的工资标准：生效日<=那天的有几条＝k，取他按生效日排第 k 条的单价/月薪（不用数组公式，WPS 也一样算）"""
    k = f'COUNTIFS(标准_姓名,{esc(name)},标准_生效,"<="&{eom})'
    return (f'IF({monthly},SUMIFS(标准_月薪,标准_姓名,{esc(name)},标准_序,{k}),'
            f'SUMIFS(标准_日单价,标准_姓名,{esc(name)},标准_序,{k}))')


def esc(x):
    """当 SUMIFS/COUNTIFS 的条件用时，把名字里的 ~ * ? 转义成普通字符"""
    return f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({x},"~","~~"),"*","~*"),"?","~?")'


def mref(sh, col, n):
    """基本信息表某列的锚定区域（从表头行开始，容量 n）"""
    return f"{sh}!${col}${M_HDR}:${col}${M_R0 + n + 200}"
