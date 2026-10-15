# -*- coding: utf-8 -*-
"""用 LibreOffice 在副本上整本重算（关掉多线程分组计算），数报错、数公式；可选计时。

为什么要关多线程：这台机器上 LibreOffice 24.2 开着「公式组多线程计算」算这本册子会冒出假的 Err:522（循环引用），
关掉就没有。设置项：/org.openoffice.Office.Calc/Formula/Calculation/UseThreadedCalculationForFormulaGroups=false。

    python3 -I lo_recalc.py <副本.xlsx> [超时秒数] [--time]
      --time：另外记一次 calculateAll 本身用了几秒（不含打开、存盘），结果在返回 JSON 的 calc_seconds。

返回（打印 JSON）：{"formulas": 公式格数, "errors": 报错格数, "error_cells": {错误: [前 30 个位置]}, "wall_seconds": …, "calc_seconds": …}
每次用一个新的临时配置目录，不碰别的 LibreOffice 进程（这台机器上别的任务也在用）。
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from lxml import etree

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
Sub RecalcAndSave()
  Dim t0 As Double, t1 As Double, f As Integer, p As String
  t0 = Timer
  ThisComponent.calculateAll()
  t1 = Timer
  p = Environ("A076_CALC_TIME_OUT")
  If p &lt;&gt; "" Then
    f = FreeFile
    Open p For Output As #f
    Print #f, Format(t1 - t0, "0.000")
    Close #f
  End If
  ThisComponent.store()
  ThisComponent.close(True)
End Sub
</script:module>"""

XCU_ITEM = ('<item oor:path="/org.openoffice.Office.Calc/Formula/Calculation">'
            '<prop oor:name="UseThreadedCalculationForFormulaGroups" oor:op="fuse"><value>false</value></prop></item>')


def _env():
    env = os.environ.copy()
    env['SAL_USE_VCLPLUGIN'] = 'svp'
    try:                                    # 沙箱里 AF_UNIX 被禁时要用 skill 带的垫片
        sys.path.insert(0, '/mnt/skills/public/xlsx/scripts')
        from office.soffice import get_soffice_env
        env = get_soffice_env()
    except Exception:
        pass
    return env


def _profile(threaded=False):
    d = Path(tempfile.mkdtemp(prefix='lo_a076_'))
    url = d.as_uri()
    subprocess.run(['soffice', '--headless', '--terminate_after_init', f'-env:UserInstallation={url}'],
                   capture_output=True, timeout=180, env=_env())
    xcu = d / 'user' / 'registrymodifications.xcu'
    if not threaded:
        t = xcu.read_text(encoding='utf-8')
        assert '</oor:items>' in t, 'LibreOffice 配置文件不对'
        xcu.write_text(t.replace('</oor:items>', XCU_ITEM + '</oor:items>'), encoding='utf-8')
    mdir = d / 'user' / 'basic' / 'Standard'
    assert mdir.exists(), 'LibreOffice 没建好配置目录'
    (mdir / 'Module1.xba').write_text(MACRO, encoding='utf-8')
    return d, url


def count(path, maxloc=30):
    """数公式格（含共享公式子格）和报错格（t="e"）"""
    z = zipfile.ZipFile(path)
    wbx = etree.fromstring(z.read('xl/workbook.xml'))
    rels = etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid = {r.get('Id'): r.get('Target') for r in rels}
    nf = ne = 0
    errs = {}
    for s in wbx.iter(NS + 'sheet'):
        t = rid[s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')].lstrip('/')
        t = t if t.startswith('xl/') else 'xl/' + t
        for _, el in etree.iterparse(z.open(t), tag=NS + 'c', huge_tree=True):
            if el.find(NS + 'f') is not None:
                nf += 1
            if el.get('t') == 'e':
                ne += 1
                v = el.findtext(NS + 'v')
                lst = errs.setdefault(v, [])
                if len(lst) < maxloc:
                    lst.append(f"{s.get('name')}!{el.get('r')}")
            el.clear()
    z.close()
    return nf, ne, errs


def unhide_cols(path):
    """把隐藏列都显示出来（LibreOffice 才不会漏算隐藏列尾巴上的公式；照 A062 build_1008 的做法）"""
    tmp = path + '.tmp'
    zi = zipfile.ZipFile(path)
    zo = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for it in zi.infolist():
        d = zi.read(it.filename)
        if it.filename.startswith('xl/worksheets/sheet') and it.filename.endswith('.xml'):
            d = re.sub(rb'(<col\b[^>]*?)\s+hidden="(?:1|true)"', rb'\1', d)
        zo.writestr(it, d)
    zo.close()
    zi.close()
    shutil.move(tmp, path)


def recalc(path, timeout=3600, timing=False, threaded=False):
    path = os.path.abspath(path)
    prof, url = _profile(threaded)
    try:
        before = os.stat(path).st_mtime_ns
        env = _env()
        tfile = None
        if timing:
            tfile = str(prof / 'calc_time.txt')
            env['A076_CALC_TIME_OUT'] = tfile
        t0 = time.time()
        r = subprocess.run(['timeout', str(timeout), 'soffice', '--headless', '--norestore', f'-env:UserInstallation={url}',
                            'vnd.sun.star.script:Standard.Module1.RecalcAndSave?language=Basic&location=application', path],
                           capture_output=True, text=True, env=env, timeout=timeout + 60)
        wall = time.time() - t0
        if r.returncode != 0 or os.stat(path).st_mtime_ns == before:
            return {'error': f'LibreOffice 没算成（返回码 {r.returncode}）：{(r.stderr or r.stdout)[:500]}'}
        nf, ne, errs = count(path)
        out = {'formulas': nf, 'errors': ne, 'error_cells': errs, 'wall_seconds': round(wall, 1)}
        if tfile and os.path.exists(tfile):
            out['calc_seconds'] = float(open(tfile).read().strip().replace(',', '.'))
        return out
    finally:
        shutil.rmtree(prof, ignore_errors=True)


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    res = recalc(args[0], int(args[1]) if len(args) > 1 else 3600, timing='--time' in sys.argv)
    print(json.dumps(res, ensure_ascii=False, indent=1))
    sys.exit(0 if res.get('errors') == 0 else 1)
