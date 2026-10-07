# -*- coding: utf-8 -*-
"""容量体检：四本里所有「引用录入表/清单表/接口表的区域」，结束行够不够到那张表的容量（一年用量）。

扫单元格公式、数据验证（下拉来源）、条件格式、定义名称，也认跨文件 [1]/[2] 引用；
本表内不带表名的引用也查；区域要「从数据区第一行（或表头行）起、伸进数据区」才算，表头 A3:Q3 这种不再误报；
「$H$4:$H4」这种累计区域（起点绝对、终点相对）、「B7:M7」这种同一行的横向区域不算。

跑法：python3 cap_audit.py [目录]      （默认成品目录）
"""
import os, re, sys, collections, zipfile
import openpyxl
from openpyxl.utils import column_index_from_string as CI

DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
BOOK = {'01': '01_水果进销存台账模板.xlsx', '02': '02_物料与周转物台账模板.xlsx',
        '03': '03_财务账套与报表模板.xlsx', '04': '04_综合查询对账单模板.xlsx'}
ALL = 'A:XFD'
# (本, 表) → [(列组, 数据起始行, 数据结束行)]
CAP = {
    ('01', '原料入库明细'): [(ALL, 4, 5003)], ('01', '原料出库明细'): [(ALL, 4, 3004)], ('01', '成品出库明细'): [(ALL, 4, 3003)],
    ('01', '其他代存明细'): [(ALL, 4, 803)],
    ('02', '周转筐出入库明细'): [(ALL, 4, 3003)], ('02', '包装物料出入库明细'): [(ALL, 4, 2003)],
    ('02', '客户自备加工筐明细'): [(ALL, 4, 1203)], ('02', '次果筐+托盘明细'): [(ALL, 4, 603)],
    ('03', '资金日记账'): [(ALL, 4, 4003)], ('03', '科目表'): [('A:U', 4, 203)], ('03', '成本费用登记表'): [(ALL, 4, 303)],
    ('03', '对账补充行'): [(ALL, 4, 303)],
    # ↓ cap03 这一轮加的
    ('03', '科目余额表'): [('A:J', 4, 153)],
    ('03', '期初余额'): [('A:F', 4, 153), ('H:I', 4, 43)],
    ('03', '收入成本映射'): [('A:G', 4, 83)],
    ('03', '往来业务明细'): [(ALL, 4, 19003)],
    ('03', '往来对账单明细'): [('A:J', 7, 306)],
    ('03', '应收应付汇总表'): [('A:L', 7, 306)],
    ('03', '资金与往来余额表'): [('A:I', 6, 25), ('A:I', 33, 332)],
    ('03', '收入月度汇总'): [('A:N', 7, 306), ('X:Z', 4, 4003)],
    ('03', '支出月度汇总'): [('A:N', 7, 306), ('X:Z', 4, 4003)],
    ('03', '费用月度支出汇总'): [('A:N', 7, 206), ('X:Z', 4, 4003)],
    ('03', '费用月度明细汇总'): [('A:C', 7, 206), ('E:L', 7, 2006), ('O:R', 4, 4003), ('S:S', 7, 2006)],
    ('03', '家用月度支出汇总'): [('A:N', 9, 58), ('A:N', 64, 263), ('P:U', 4, 4003)],
    ('03', '家用月度明细汇总'): [('A:D', 7, 56), ('A:D', 60, 259), ('F:N', 7, 2006), ('Q:W', 4, 4003), ('X:X', 7, 2006)],
    ('03', '应付账款明细'): [('A:G', 7, 206), ('L:N', 7, 306)],
    ('03', '国外费用汇总'): [('A:D', 7, 36), ('F:N', 7, 606), ('Q:T', 4, 4003), ('U:U', 7, 606)],
    ('03', '借款汇总'): [('A:G', 8, 37), ('A:G', 42, 71), ('H:O', 8, 307), ('R:AA', 4, 4023), ('AB:AB', 4, 4003),
                        ('AC:AC', 8, 307)],
    ('03', '结转损益'): [('Q:Q', 4, 203)],
    ('03', '_自动清单'): [('A:C', 4, 5325), ('E:E', 4, 253), ('AF:AH', 4, 4003), ('AJ:AJ', 4, 203)],
    # 《01》《02》接口表、列表表（cap0102；对账明细接口按一年用量 7,000 / 4,000）
    ('01', '对账明细接口'): [(ALL, 4, 7003)], ('01', '库存结余'): [(ALL, 4, 403)], ('01', '实盘库存盘点'): [(ALL, 5, 404)],
    ('01', '库存等级接口'): [(ALL, 4, 403)], ('01', '库存总结余'): [(ALL, 4, 303)], ('01', '入库吨位汇总'): [(ALL, 4, 303)],
    ('01', '成品出库吨位汇总'): [(ALL, 4, 303)], ('01', '其他代存汇总'): [(ALL, 4, 203)], ('01', '筐子进出接口'): [(ALL, 4, 303)],
    ('01', '公司购买接口'): [(ALL, 4, 403)], ('01', '果然鲜采购明细'): [(ALL, 4, 403)], ('01', '果然鲜销售明细'): [(ALL, 4, 403)],
    ('01', '基础资料'): [(ALL, 4, 203)],
    ('02', '对账明细接口'): [(ALL, 4, 4003)], ('02', '装筐费接口'): [(ALL, 4, 1503)], ('02', '基础资料'): [(ALL, 4, 103)],
    ('02', '物料库存结余'): [(ALL, 5, 54)], ('02', '筐子库存汇总'): [(ALL, 5, 64)], ('02', '筐子往来汇总'): [(ALL, 5, 254)],
    ('02', '筐子采购汇总'): [(ALL, 5, 204)], ('02', '筐子销售汇总'): [(ALL, 5, 204)], ('02', '自备筐汇总'): [(ALL, 5, 204)],
    ('02', '次果筐汇总'): [(ALL, 5, 204)], ('02', '押金汇总'): [(ALL, 5, 204)], ('02', '客户领用物料汇总'): [(ALL, 5, 204)],
    ('02', '包装物料采购汇总'): [(ALL, 5, 204)], ('02', '实盘库存盘点'): [(ALL, 5, 184)], ('02', '对接源_01水果筐数'): [(ALL, 5, 304)],
    ('04', '对接源_01水果'): [(ALL, 4, 7003)], ('04', '对接源_02物料'): [(ALL, 4, 4003)],
}
REF = re.compile(r'"(?:[^"]|"")*"|(?<![A-Za-z0-9_.$!:一-鿿\]\'])'
                 r"(?:'(?:\[(\d+)\])?([^']+)'|(?:\[(\d+)\])?([A-Za-z_一-鿿][\w一-鿿+.]*))?(!)?"
                 r"(\$?)([A-Z]{1,3})(\$?)(\d+):(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])")


def ext_map(path):
    z = zipfile.ZipFile(path)
    out = {}
    for nm in z.namelist():
        m = re.match(r'xl/externalLinks/_rels/externalLink(\d+)\.xml\.rels$', nm)
        if m:
            t = re.search(r'Target="([^"]+)"', z.read(nm).decode()).group(1)
            for k, b in BOOK.items():
                if t.replace('\\', '/').split('/')[-1] == b:
                    out[m.group(1)] = k
    return out


def cols(spec):
    a, b = spec.split(':')
    return CI(a), CI(b)


def scan(k):
    path = os.path.join(DIR, BOOK[k])
    ext = ext_map(path)
    wb = openpyxl.load_workbook(path)
    hits = collections.Counter()
    samples = {}

    def look(where, text, own_sheet):
        for m in REF.finditer(text):
            if m.group(0).startswith('"'):
                continue
            n1, s1, n2, s2, bang, a1, c1, ar1, r1, a2, c2, ar2, r2 = m.groups()
            if bang:
                idx, sh = (n1 or n2), (s1 or s2)
            else:
                if s1 or s2:
                    continue
                idx, sh = None, own_sheet
            if sh is None:
                continue
            book = ext.get(idx) if idx else k
            blocks = CAP.get((book, sh))
            if not blocks:
                continue
            r1, r2 = int(r1), int(r2)
            if ar1 and not ar2:            # $H$4:$H4 累计区域
                continue
            if r1 == r2:                   # B7:M7 这种同一行的横向区域
                continue
            lo, hi = sorted((CI(c1), CI(c2)))
            for spec, s0, s1_ in blocks:
                clo, chi = cols(spec)
                if not (clo <= lo and hi <= chi):
                    continue
                if s0 - 1 <= r1 <= s0 + 1 and s0 <= r2 < s1_:
                    key = (where, f'{book}:{sh}[{spec} {s0}-{s1_}]', r2)
                    hits[key] += 1
                    samples.setdefault(key, text[:160])
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith('='):
                    look(f'{ws.title}', v, ws.title)
        for dv in ws.data_validations.dataValidation:
            if dv.formula1:
                look(f'{ws.title}[下拉 {dv.sqref}]', str(dv.formula1), ws.title)
                for s in re.findall(r'INDIRECT\("([^"]+)"\)', str(dv.formula1)):
                    look(f'{ws.title}[下拉 INDIRECT {dv.sqref}]', s, ws.title)
        for cf in ws.conditional_formatting:
            for r in cf.rules:
                for f in (r.formula or []):
                    look(f'{ws.title}[条件格式]', str(f), ws.title)
    for nm, dn in wb.defined_names.items():
        if dn.attr_text:
            look(f'[定义名称 {nm}]', dn.attr_text, None)
    wb.close()
    return hits, samples


if __name__ == '__main__':
    total = 0
    for k in BOOK:
        hits, samples = scan(k)
        print(f'== 《{k}》 {BOOK[k]}：{sum(hits.values())} 处不够')
        for (where, tgt, r2), n in sorted(hits.items()):
            print(f'   {where:<34} → {tgt} 只到第 {r2} 行（{n} 处）  例：{samples[(where, tgt, r2)]}')
        total += sum(hits.values())
    print('合计不够：', total)
