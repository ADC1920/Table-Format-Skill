# -*- coding: utf-8 -*-
"""check_delivery.py — 表格交付前最小门禁（下层专项检查器）。

层级关系（防同名误判为漂移）：本文件只查 xlsx 交付形态，是 table-format 技能自带的下层检查器；
本机另有全格式统一入口 `<工作区>/scripts/check_delivery.py`（按扩展名分派 docx/xlsx/pptx/pdf），
它在 xlsx 分支调用本文件。两者同名不同物、分工并存——本文件不依赖 officecli 与其它技能，
不可被全格式版替换（会破坏技能独立性）；全格式版也不可下沉到技能（会使技能依赖外部工具）。

交付形态硬伤检查，与 verify_table.py（样式规格门禁）分工：
    verify_table.py  查「排版规格」——表头样式 / 冻结 / 底纹 / 边框等
    check_delivery.py 查「能不能交出去」——默认表名残留 / 空表 / 公式错误值

检查项：
    1. 文件可打开、工作表存在数据（数据行 > 0）
    2. 工作表名非模板默认残留（Sheet / Sheet1 / 工作表1 等）
    3. 无公式错误值（#REF! / #VALUE! / #DIV/0! / #N/A / #NAME? / #NULL! / #NUM!）

用法：
    python check_delivery.py 成品.xlsx [--sheet 表名]
    退出码 0 = 可交付；1 = 存在硬伤。
"""

import argparse
import os
import re
import sys

from openpyxl import load_workbook

DEFAULT_SHEET_NAMES = {"sheet", "sheet1", "sheet2", "sheet3", "工作表1", "工作表2", "工作表3"}
ERROR_VALUES = ("#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#GETTING_DATA")


def check(path, sheet=None):
    fails = []
    warns = []

    if not os.path.isfile(path):
        print("ERROR: 文件不存在 %s" % path)
        return 1
    if os.path.getsize(path) == 0:
        print("ERROR: 文件大小为 0")
        return 1

    try:
        wb = load_workbook(path)
    except Exception as e:
        print("ERROR: 无法打开（%s）" % e)
        return 1

    ws = wb[sheet] if sheet else wb.active

    # 1. 数据行 > 0（粗略口径：max_row >= 2 视为有数据；仅表头的表另由 verify 判）
    if ws.max_row < 2:
        warns.append("工作表仅 %d 行，疑似空表" % ws.max_row)

    # 2. 默认表名残留
    if ws.title.strip().lower() in DEFAULT_SHEET_NAMES:
        fails.append("工作表名仍是模板默认值「%s」，交付前请改有意义的名字" % ws.title)

    # 3. 公式错误值扫描
    err_cells = []
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str):
                s = v.strip().upper()
                if s in ERROR_VALUES:
                    err_cells.append(c.coordinate)
                    if len(err_cells) >= 10:
                        break
        if len(err_cells) >= 10:
            break
    if err_cells:
        fails.append("存在公式错误值：%s" % "、".join(err_cells))

    print("文件: %s | sheet=%s | max_row=%d max_col=%d" % (path, ws.title, ws.max_row, ws.max_column))
    for w in warns:
        print("  [WARN] %s" % w)
    for f in fails:
        print("  [FAIL] %s" % f)
    if not fails:
        print("  [OK]   交付形态检查通过")
    print("DELIVERY: %s" % ("PASS" if not fails else "FAIL"))
    return 0 if not fails else 1


def main():
    ap = argparse.ArgumentParser(description="表格交付前最小门禁")
    ap.add_argument("xlsx", help="待检查的 xlsx")
    ap.add_argument("--sheet", help="工作表名（默认活动表）")
    a = ap.parse_args()
    sys.exit(check(a.xlsx, a.sheet))


if __name__ == "__main__":
    main()
