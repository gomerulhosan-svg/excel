# -*- coding: utf-8 -*-
"""在 LibreOffice 里量「改资金台帐一格要等几秒」（build.py 调；也可以单独跑：python3 -I lo_time.py 算好的副本.xlsx）。
做法：打开 → 关自动计算 → 整本算一遍 → 改台帐中间一行的收入、摘要、日期各一次，每次只算脏格（calculate），记秒数。
LibreOffice 比 Excel 慢几倍（A072 同口径：LO 3 秒 ≈ Excel 不到 1 秒）；WPS 没法在这里量，交付前要在 WPS 真机再量一次。
输出最后一行是 JSON：{"load":秒,"full":秒,"edits":{"收入":秒,…},"max":秒}"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main(path):
    import uno
    from com.sun.star.beans import PropertyValue
    from layout import SH_LEDGER, LEDGER_COLS, R0
    path = os.path.abspath(path)
    prof = tempfile.mkdtemp(prefix='a074_lo_')
    pipe = 'a074time%d' % os.getpid()
    p = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', '--nodefault',
                          f'-env:UserInstallation=file://{prof}', f'--accept=pipe,name={pipe};urp;'],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    res = {}
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = None
        for _ in range(240):
            try:
                ctx = resolver.resolve(f'uno:pipe,name={pipe};urp;StarOffice.ComponentContext')
                break
            except Exception:
                time.sleep(0.5)
        desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
        pv = PropertyValue()
        pv.Name, pv.Value = 'Hidden', True
        t = time.time()
        doc = desk.loadComponentFromURL('file://' + path, '_blank', 0, (pv,))
        res['load'] = round(time.time() - t, 1)
        doc.enableAutomaticCalculation(False)
        t = time.time()
        doc.calculateAll()
        res['full'] = round(time.time() - t, 1)
        sh = doc.Sheets.getByName(SH_LEDGER)
        used = 0
        col_b = LEDGER_COLS['账户']
        for r in range(R0, R0 + 20000, 500):                 # 粗找已用到第几行
            if sh.getCellRangeByName(f'{col_b}{r}').getString():
                used = r
        row = max(R0, (used + R0) // 2)
        edits = {}
        for lab, col, val in (('收入', LEDGER_COLS['收入'], 1234.56), ('摘要', LEDGER_COLS['摘要'], '测试摘要租金'),
                              ('日期', LEDGER_COLS['日期'], 46200)):
            c = sh.getCellRangeByName(f'{col}{row}')
            if isinstance(val, str):
                c.setString(val)
            else:
                c.setValue(val)
            t = time.time()
            doc.calculate()
            edits[lab] = round(time.time() - t, 2)
        res['edits'] = edits
        res['row'] = row
        res['max'] = max(edits.values())
        doc.close(True)
    finally:
        p.terminate()
        shutil.rmtree(prof, ignore_errors=True)
    return res


if __name__ == '__main__':
    print(json.dumps(main(sys.argv[1]), ensure_ascii=False))
