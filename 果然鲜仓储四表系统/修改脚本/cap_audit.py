# -*- coding: utf-8 -*-
"""容量体检：四本里所有「引用录入表/清单表的区域」，结束行够不够到那张表的容量。

扫：单元格公式、数据验证（下拉来源）、条件格式、定义名称，也认跨文件 [1]/[2] 引用。
只看区域引用（A4:A603 这种）；单格逐行引用不算。区域起点在数据区之内、终点却比容量小的，就是「到某一行就不算了」。

跑法：python3 cap_audit.py [目录]        （默认成品目录；输出按 引用方表 × 被引用表 × 结束行 汇总）
"""
import os, re, sys, collections, zipfile
import openpyxl
from openpyxl.utils import column_index_from_string as CI

HERE = os.path.dirname(os.path.abspath(__file__))
DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..')
BOOK = {'01': '01_水果进销存台账模板.xlsx', '02': '02_物料与周转物台账模板.xlsx',
        '03': '03_财务账套与报表模板.xlsx', '04': '04_综合查询对账单模板.xlsx'}

# 录入表/清单表：(本, 表) → (数据起始行, 数据结束行＝这一轮约定的容量)
CAP = {
    ('01', '原料入库明细'): (4, 5003), ('01', '原料出库明细'): (4, 3004), ('01', '成品出库明细'): (4, 3003),
    ('01', '其他代存明细'): (4, 803),
    ('02', '周转筐出入库明细'): (4, 3003), ('02', '包装物料出入库明细'): (4, 2003),
    ('02', '客户自备加工筐明细'): (4, 1203), ('02', '次果筐+托盘明细'): (4, 603),
    ('03', '资金日记账'): (4, 4003), ('03', '科目表'): (4, 203), ('03', '成本费用登记表'): (4, 303),
    ('03', '对账补充行'): (4, 303),
}
REF = re.compile(r"(?:'(?:\[(\d+)\])?([^']+)'|(?:\[(\d+)\])?([A-Za-z_一-鿿][\w一-鿿+.]*))!"
                 r"\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")


def ext_map(path):
    """[n] → 被链接的本（01/02）"""
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


def scan(k):
    path = os.path.join(DIR, BOOK[k])
    ext = ext_map(path)
    wb = openpyxl.load_workbook(path)
    hits = collections.Counter()
    samples = {}

    def look(where, text, own_sheet):
        for m in REF.finditer(text):
            n1, s1, n2, s2, c1, r1, c2, r2 = m.groups()
            idx, sh = (n1 or n2), (s1 or s2)
            if r2 is None:
                continue
            book = ext.get(idx) if idx else k
            cap = CAP.get((book, sh))
            if not cap:
                continue
            r1, r2 = int(r1), int(r2)
            if r1 <= cap[0] + 1 and r2 < cap[1]:
                key = (where, f'{book}:{sh}', r2)
                hits[key] += 1
                samples.setdefault(key, text[:160])

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith('=') and '!' in v:
                    look(f'{ws.title}', v, ws.title)
        for dv in ws.data_validations.dataValidation:
            if dv.formula1:
                look(f'{ws.title}[下拉 {dv.sqref}]', str(dv.formula1), ws.title)
        for cf in ws.conditional_formatting:
            for r in cf.rules:
                for f in (r.formula or []):
                    look(f'{ws.title}[条件格式]', str(f), ws.title)
    for dn in wb.defined_names.values() if hasattr(wb.defined_names, 'values') else wb.defined_names.definedName:
        if dn.attr_text:
            look(f'[定义名称 {dn.name}]', dn.attr_text, None)
    for ws in wb.worksheets:
        for nm, dn in getattr(ws, 'defined_names', {}).items():
            if dn.attr_text and not nm.startswith('_xlnm'):
                look(f'{ws.title}[表内名称 {nm}]', dn.attr_text, ws.title)
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
