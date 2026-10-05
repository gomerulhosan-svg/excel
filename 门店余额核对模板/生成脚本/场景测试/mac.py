# 用法：python3 mac.py <场景名> <文件>  在 LibreOffice 里打开文件，按场景改几格，整本重算后存盘（给 verify.py --values 核对）
import glob, sys, tempfile, subprocess
from pathlib import Path
sys.path.insert(0, glob.glob('/root/.claude/skills/synced/*/xlsx/scripts')[0])
from office.soffice import run_soffice, get_soffice_env
END = """  ThisComponent.calculateAll()
  ThisComponent.store()
  ThisComponent.close(True)
End Sub
"""
MAC = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
Sub InsRow
  Dim sh As Object
  sh = ThisComponent.Sheets.getByName("数据录入")
  sh.Rows.insertByIndex(10999, 1)
  sh.getCellByPosition(1, 10999).setValue(46290)
  sh.getCellByPosition(2, 10999).setString("邮政个人")
  sh.getCellByPosition(5, 10999).setValue(500)
  sh.getCellByPosition(8, 10999).setString("博景分店（60分店）")
  sh.getCellByPosition(11, 10999).setValue(202609)
""" + END + """
Sub InsTop
  Dim sh As Object
  sh = ThisComponent.Sheets.getByName("数据录入")
  sh.Rows.insertByIndex(18, 1)
  sh.getCellByPosition(1, 18).setValue(46200)
  sh.getCellByPosition(4, 18).setString("测试调拨")
  sh.getCellByPosition(5, 18).setValue(1000)
  sh.getCellByPosition(8, 18).setString("配送")
  sh.getCellByPosition(11, 18).setValue(202606)
""" + END + """
Sub DelRow
  Dim sh As Object
  sh = ThisComponent.Sheets.getByName("数据录入")
  sh.Rows.removeByIndex(10999, 1)
""" + END + """
Sub CutEtc
  Dim s As Object
  s = ThisComponent.Sheets.getByName("门店余额核对")
  s.getCellRangeByName("B4").setValue(46203)
  s.getCellRangeByName("G157").setValue(400000)
  s.getCellRangeByName("G158").setValue(1137972.14)
  s.getCellRangeByName("B60").setString("疼痛馆")
""" + END + """
Sub TextL
  Dim d As Object
  d = ThisComponent.Sheets.getByName("数据录入")
  d.getCellByPosition(2, 11088).setString("邮政个人")
  d.getCellByPosition(4, 11088).setString("测试")
  d.getCellByPosition(5, 11088).setValue(100)
  d.getCellByPosition(8, 11088).setString("博景分店（60分店）")
  d.getCellByPosition(11, 11088).setString("9月")
""" + END + """
Sub BigL
  Dim d As Object
  d = ThisComponent.Sheets.getByName("数据录入")
  d.getCellByPosition(2, 11088).setString("邮政个人")
  d.getCellByPosition(4, 11088).setString("测试")
  d.getCellByPosition(5, 11088).setValue(100)
  d.getCellByPosition(8, 11088).setString("博景分店（60分店）")
  d.getCellByPosition(11, 11088).setValue(2026009)
""" + END + """
Sub Closed
  Dim d As Object
  d = ThisComponent.Sheets.getByName("数据录入")
  d.getCellByPosition(1, 11088).setValue(46290)
  d.getCellByPosition(2, 11088).setString("邮政公户")
  d.getCellByPosition(4, 11088).setString("测试")
  d.getCellByPosition(6, 11088).setValue(1000)
  d.getCellByPosition(8, 11088).setString("金店（52店）")
  d.getCellByPosition(11, 11088).setValue(202609)
""" + END + """
Sub NameSpace1
  Dim d As Object
  d = ThisComponent.Sheets.getByName("数据录入")
  d.getCellByPosition(1, 11088).setValue(46290)
  d.getCellByPosition(2, 11088).setString("邮政个人")
  d.getCellByPosition(4, 11088).setString("测试")
  d.getCellByPosition(6, 11088).setValue(300)
  d.getCellByPosition(8, 11088).setString("配送 ")
  d.getCellByPosition(11, 11088).setValue(202609)
""" + END + """
Sub TextDate
  Dim d As Object
  d = ThisComponent.Sheets.getByName("数据录入")
  d.getCellByPosition(1, 11088).setString("2026.9.15")
  d.getCellByPosition(2, 11088).setString("现金")
  d.getCellByPosition(4, 11088).setString("测试")
  d.getCellByPosition(6, 11088).setValue(100)
  d.getCellByPosition(8, 11088).setString("公司")
""" + END + """
Sub YearView
  ThisComponent.Sheets.getByName("门店月度收支统计").getCellByPosition(1, 1).setValue(2025)
""" + END + """
Sub AcctDel
  Dim d As Object
  d = ThisComponent.Sheets.getByName("数据录入")
  d.Rows.removeByIndex(17, 1)
""" + END + """
Sub CopySh
  ThisComponent.Sheets.copyByName("门店余额核对", "核对副本", 30)
  ThisComponent.Sheets.getByName("核对副本").getCellRangeByName("B4").setValue(46203)
""" + END + """
</script:module>"""
name, f = sys.argv[1], str(Path(sys.argv[2]).absolute())
with tempfile.TemporaryDirectory(prefix='lo-mac-') as pd:
    pd = Path(pd)
    run_soffice(['--headless', '--terminate_after_init', f'-env:UserInstallation={pd.as_uri()}'], capture_output=True, timeout=120)
    (pd / 'user/basic/Standard/Module1.xba').write_text(MAC, encoding='utf-8')
    r = subprocess.run(['timeout', '1200', 'soffice', '--headless', '--norestore', f'-env:UserInstallation={pd.as_uri()}',
                        f'vnd.sun.star.script:Standard.Module1.{name}?language=Basic&location=application', f],
                       capture_output=True, text=True, env=get_soffice_env())
    print(name, r.returncode, r.stdout[-300:], r.stderr[-300:])
