# -*- coding: utf-8 -*-
"""读「导入」目录里用户的数据（清洗好的 json，只有用户自己的数据），换成生成脚本 ctx 的格式。

  u4_资金台帐.json  用户 10/15 上传的资金台帐：三个资金账户的期初 ＋ 2026-01-23～06-30 的 999 笔流水
  u2_货柜.json      2026 出入库台账里的「货柜单号台账」9 行

怎么洗出来的、洗了什么写在 json 的「清洗」里。原上传文件不进仓库（第三方模板的说明、署名一个字都不能带）。
类别按方案甲沿用她自己的叫法（原材料、管理费用、生活支出…），只有「其他货币资金」（她账户之间换汇）改用模板的「购汇/结汇」；
她「备注」列写的二级（公关费、PAYE、宽体车…）放进流水 R 列「二级/备注」，【记账规则】的「二级明细对照」按它记到明细科目。
"""
import datetime as dt
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
CAT_MAP = {'其他货币资金': '购汇/结汇'}


def _load(name):
    with open(os.path.join(HERE, name), encoding='utf-8') as f:
        return json.load(f)


def _date(s):
    return dt.datetime.strptime(s, '%Y-%m-%d')


def accounts():
    """[{name, bank, cur, open, openb}]，顺序＝【基础资料】资金账户第 5、6、7 行（人民币、美元、FBC）"""
    return _load('u4_资金台帐.json')['accounts']


def ledger():
    """ctx['ledger'] 的格式：每笔一个 dict（row＝现金流水总表行号，按日期排好）"""
    out = []
    for f in _load('u4_资金台帐.json')['flows']:
        out.append(dict(row=f['row'], date=_date(f['date']), account=f['account'], category=CAT_MAP.get(f['cat'], f['cat']),
                        memo=f['memo'], inc=f['inc'], exp=f['exp'], xrate=f['xrate'], project=None, cp=None,
                        note=f['sub'], ovr=None, merge=None, att=None))
    return out


def containers():
    """货柜登记：[{date, box, bl, seal, goods, note}]"""
    return [dict(date=_date(x['date']), box=x['box'], bl=x['bl'], seal=x['seal'], goods=None, note=None)
            for x in _load('u2_货柜.json')['rows']]
