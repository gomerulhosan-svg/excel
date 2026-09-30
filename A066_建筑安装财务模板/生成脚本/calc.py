# -*- coding: utf-8 -*-
"""取数公式：利润表、项目利润、分摊、资产负债表都用这里的同一套公式，口径一致、互相对得上。
   区间都用「年月数」（YYYYMM）：ym0～ym1（含两头）。pj 给了就只算这个项目。"""
from layout import *
import cats


def _r(ym0, ym1, col_ym):
    return f'{col_ym},">="&{ym0},{col_ym},"<="&{ym1}'


def jline(line, ym0, ym1, pj=None, sign=-1):
    """流水里记到某个报表项目的钱：费用类取 −净额（付出去为正）"""
    s = f'SUMIFS({jr(J_NET)},{jr(J_LINE)},"{line}",{_r(ym0, ym1, jr(J_YM))}' + (f',{jr(J_PJ)},{pj}' if pj else '') + ')'
    return s if sign > 0 else f'(-{s})'


def aplog(typ, ym0, ym1, pj=None):
    return f'SUMIFS({apr(AP_AMT)},{apr(AP_TYPE)},"{typ}",{_r(ym0, ym1, apr(AP_YM))}' + (f',{apr(AP_PJ)},{pj}' if pj else '') + ')'


def rev(ym0, ym1, pj=None):
    return f'SUMIFS({rvr(RV_AMT)},{_r(ym0, ym1, rvr(RV_YM))}' + (f',{rvr(RV_PJ)},{pj}' if pj else '') + ')'


def deduct(ym0, ym1, pj=None, typ='甲方扣款'):
    """代发抵账里的甲方扣款（算项目其他直接费）、甲供材扣款（算项目材料费）"""
    return f'SUMIFS({ofr(OF_AMT)},{ofr(OF_TYPE)},"{typ}",{_r(ym0, ym1, ofr(OF_YM))}' + (f',{ofr(OF_PJ)},{pj}' if pj else '') + ')'


def lab_direct(ym0, ym1, pj=None):
    """考勤里算到项目的人工（应发合计按天数拆到各项目的部分）"""
    if pj is None:
        return f'(SUMIFS({atr(AT_PAY)},{_r(ym0, ym1, atr(AT_YM))})-SUMIFS({atr(AT_COAMT)},{_r(ym0, ym1, atr(AT_YM))}))'
    return '(' + '+'.join(f'SUMIFS({atr(a)},{atr(p)},{pj},{_r(ym0, ym1, atr(AT_YM))})' for a, p in zip(AT_AMTS, AT_PJS)) + ')'


def lab_company(ym0, ym1):
    return f'SUMIFS({atr(AT_COAMT)},{_r(ym0, ym1, atr(AT_YM))})'


def att_sum(col, ym0, ym1):
    return f'SUMIFS({atr(col)},{_r(ym0, ym1, atr(AT_YM))})'


def _idx(ym):
    return f'(INT({ym}/100)*12+MOD({ym},100))'


def fa_dep(ym0, ym1, pj=None, company=False):
    """固定资产折旧：从购入次月起，按月数提完为止；区间内的月数 × 月折旧"""
    S = rng(SH_BASE, FA_S, FA_R0, FA_R1)
    E = rng(SH_BASE, FA_E, FA_R0, FA_R1)
    M = rng(SH_BASE, FA_DEP_M, FA_R0, FA_R1)
    P = rng(SH_BASE, FA_PJ, FA_R0, FA_R1)
    a, b = _idx(ym0), _idx(ym1)
    hi = f'(({E}+{b})-ABS({E}-{b}))/2'          # MIN(E, b)
    lo = f'(({S}+{a})+ABS({S}-{a}))/2'          # MAX(S, a)
    n = f'({hi}-{lo}+1)'
    cond = f'*({P}={pj})' if pj else (f'*({P}="")' if company else '')
    return f'SUMPRODUCT({M}*(({n}+ABS({n}))/2){cond})'


def vat_out(ym0, ym1, pj=None):
    return f'SUMIFS({rvr(RV_VAT)},{_r(ym0, ym1, rvr(RV_YM))}' + (f',{rvr(RV_PJ)},{pj}' if pj else '') + ')'


def vat_in(ym0, ym1, pj=None):
    return (f'SUMIFS({ivr(IV_TAXU)},{ivr(IV_DIR)},"进项",{ivr(IV_SPEC)},1,{_r(ym0, ym1, ivr(IV_YM))}'
            + (f',{ivr(IV_PJ)},{pj}' if pj else '') + ')')


def tax_est(ym0, ym1, pj=None):
    return (f'(({vat_out(ym0, ym1, pj)}-{vat_in(ym0, ym1, pj)})*(1+{PA_SURR})+{rev(ym0, ym1, pj)}*{PA_STAMP})')


def line(name, ym0, ym1, pj=None):
    """一个利润表项目在区间里的金额（收入、其他收入为正，其余费用为正）"""
    if name == '营业收入':
        return rev(ym0, ym1, pj)
    if name == '材料费':
        return f'({aplog("材料", ym0, ym1, pj)}+{jline("材料费", ym0, ym1, pj)}+{deduct(ym0, ym1, pj, "甲供材扣款")})'
    if name == '人工费':
        return f'({lab_direct(ym0, ym1, pj)}+{jline("人工费", ym0, ym1, pj)})'
    if name == '分包费':
        return f'({aplog("分包", ym0, ym1, pj)}+{jline("分包费", ym0, ym1, pj)})'
    if name == '机械运输费':
        return f'({aplog("机械运输", ym0, ym1, pj)}+{jline("机械运输费", ym0, ym1, pj)}+{fa_dep(ym0, ym1, pj=pj) if pj else fa_dep(ym0, ym1) + "-" + fa_dep(ym0, ym1, company=True)})'
    if name == '其他直接费':
        return f'({aplog("其他直接费", ym0, ym1, pj)}+{jline("其他直接费", ym0, ym1, pj)}+{deduct(ym0, ym1, pj)})'
    if name == '税金估算':
        return tax_est(ym0, ym1, pj)
    if name == '票据贴息':
        return jline('票据贴息', ym0, ym1, pj)
    assert pj is None, name
    if name == '管理人员工资':
        return f'({lab_company(ym0, ym1)}+{jline(name, ym0, ym1)})'   # 流水里选不了这一类（校验会报 ✗），加上是为了万一选了也对得平
    if name == '社保费':
        return f'({jline("社保费", ym0, ym1)}-{att_sum(AT_SOC, ym0, ym1)})'
    if name == '折旧费':
        return f'({fa_dep(ym0, ym1, company=True)}+{jline(name, ym0, ym1)})'
    if name == '其他收入':
        return jline('其他收入', ym0, ym1, sign=1)
    return jline(name, ym0, ym1)


def driver(ym0, ym1, pj, kind):
    """分摊依据：施工费＝人工＋分包＋机械；直接成本＝再加材料和其他直接费（不含甲方扣款、税金、贴息）"""
    parts = [line('人工费', ym0, ym1, pj), line('分包费', ym0, ym1, pj), line('机械运输费', ym0, ym1, pj)]
    if kind == 'A':
        parts += [line('材料费', ym0, ym1, pj), f'({aplog("其他直接费", ym0, ym1, pj)}+{jline("其他直接费", ym0, ym1, pj)})']
    return '(' + '+'.join(parts) + ')'


def profit(ym0, ym1):
    """区间利润（管理口径，所得税后）"""
    inc = [line('营业收入', ym0, ym1), line('其他收入', ym0, ym1)]
    exp = [line(n, ym0, ym1) for n in cats.IS_NAMES if n not in ('营业收入', '其他收入')]
    return '(' + '+'.join(inc) + '-(' + '+'.join(exp) + '))'
