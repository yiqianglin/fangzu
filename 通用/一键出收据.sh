#!/bin/bash
# 一键三连：生成收据 xlsx → 渲染 jpg → 拼 A4 打印图
# 用法：bash 每月租金生成/通用/一键出收据.sh 2026 6 [官田村]
set -e

YEAR="${1:?用法: bash 每月租金生成/通用/一键出收据.sh <年> <月> [小区，默认官田村]}"
MONTH="${2:?用法: bash 每月租金生成/通用/一键出收据.sh <年> <月> [小区，默认官田村]}"
VILLAGE="${3:-官田村}"

DIR="$(cd "$(dirname "$0")" && pwd)"

echo "═══ Step 1/3：生成收据 xlsx ═══"
python3 "$DIR/生成收据.py" --year "$YEAR" --month "$MONTH" --village "$VILLAGE"
echo
echo "═══ Step 2/3：渲染收据 jpg ═══"
python3 "$DIR/导出收据图片.py" --year "$YEAR" --month "$MONTH" --village "$VILLAGE"
echo
echo "═══ Step 3/3：拼 A4 打印图 ═══"
python3 "$DIR/生成打印图.py" --year "$YEAR" --month "$MONTH" --village "$VILLAGE"
echo
echo "✅ 全部完成"
