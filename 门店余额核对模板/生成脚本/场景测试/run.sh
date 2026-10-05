#!/bin/bash
# 场景测试：复制成品 12 份，每份在 LibreOffice 里做一种操作、重算存盘，再逐格核对。跑法：bash 场景测试/run.sh（约 10 分钟）
set -o pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
OUT="$HERE/../../多帐户收支登记表_加门店余额核对.xlsx"
W="$HERE/../_tmp/场景"; mkdir -p "$W"; cd "$W"
find . -maxdepth 1 -name ".~lock.*" -delete          # 上次没跑完留下的锁文件会让 LibreOffice 卡住
run() { for s in "$@"; do cp "$OUT" $s.xlsx; (python3 "$HERE/mac.py" $s $s.xlsx > $s.log 2>&1) & done; wait; }
run InsRow InsTop DelRow CutEtc           # 中间插行、第一笔前插行、删行、截止日期＋对账单＋④加名字
run TextL BigL Closed NameSpace1          # 所属月份写成文字 / 7 位数、已关的店来账、名字带空格
run TextDate YearView AcctDel CopySh      # 日期写成文字、门店月度看别的年份、删账户期初行、复制核对表
bad=0
for s in InsRow InsTop DelRow CutEtc TextL BigL Closed NameSpace1 TextDate YearView CopySh; do
  echo "== $s"; python3 "$HERE/../verify.py" $s.xlsx --values | tail -3 || bad=1
done
python3 "$HERE/extra.py" || bad=1
exit $bad
