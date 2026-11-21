#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第二轮：编码与源文件形态边界。

上轮未覆盖：BOM / GBK / CRLF / 控制字符 / md 分隔行畸形 / 列数不齐 /
CSV 引号嵌套 / xlsx 合并单元格与公式单元格。
这些是真实用户文件的常见形态，不是刻意构造的墙角。
"""
import os
import subprocess
import sys

from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _env import DATA, MT, OUT, SKILL, VT  # noqa: E402

results = []


def check(name, cond, detail=""):
    results.append((cond, name, detail))


def run(args, **kw):
    r = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=SKILL)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def gen(src, tag, profile="business-table", extra=None):
    out = os.path.join(OUT, "enc-" + tag + ".xlsx")
    rc, txt = run([MT, "--profile", profile, "--src", src, "--out", out] + (extra or []))
    return rc, txt, out


def verify(out, profile="business-table"):
    rc, txt = run([VT, "--profile", profile, out])
    return rc, txt


# ============ 1. 编码边界 ============
# UTF-8 BOM 的 CSV
p = os.path.join(DATA, "bom.csv")
with open(p, "w", encoding="utf-8-sig", newline="") as f:
    f.write("型号,数量\nTEL-001,10\n")
rc, txt, out = gen(p, "bom-csv")
if rc == 0:
    ws = load_workbook(out).active
    check("UTF-8 BOM CSV：表头不被 BOM 污染", ws.cell(1, 1).value == "型号",
          "A1=%r" % (ws.cell(1, 1).value,))
else:
    check("UTF-8 BOM CSV", False, txt.strip()[:120])

# GBK 编码 CSV（中文 Windows 用户用 Excel 另存的常见形态）
p = os.path.join(DATA, "gbk.csv")
with open(p, "w", encoding="gbk", newline="") as f:
    f.write("型号,数量\n甲型,10\n")
rc, txt, out = gen(p, "gbk-csv")
check("GBK 编码 CSV 的处理（当前行为）", rc in (0, 1),
      "rc=%d %s" % (rc, txt.strip()[:110]))

# CRLF 换行
p = os.path.join(DATA, "crlf.csv")
with open(p, "w", encoding="utf-8", newline="") as f:
    f.write("型号,数量\r\nTEL-001,10\r\nTEL-002,20\r\n")
rc, txt, out = gen(p, "crlf-csv")
if rc == 0:
    ws = load_workbook(out).active
    vals = [ws.cell(r, 1).value for r in (2, 3)]
    check("CRLF CSV：值不带 \\r 残留", vals == ["TEL-001", "TEL-002"], "实得 %r" % (vals,))
else:
    check("CRLF CSV", False, txt.strip()[:120])

# 控制字符（制表符在单元格内、零宽字符）
p = os.path.join(DATA, "ctrl.md")
with open(p, "w", encoding="utf-8") as f:
    f.write("| 名称 | 备注 |\n| --- | --- |\n| A\tB | 含制表符 |\n| C\u200bD | 含零宽空格 |\n")
rc, txt, out = gen(p, "ctrl-md")
check("含控制字符的 md 不崩", rc in (0, 1), "rc=%d %s" % (rc, txt.strip()[:110]))

# ============ 2. md 源形态边界 ============
# 无分隔行（用户手写的伪表格）
p = os.path.join(DATA, "no-sep.md")
with open(p, "w", encoding="utf-8") as f:
    f.write("| 名称 | 数量 |\n| 甲 | 1 |\n| 乙 | 2 |\n")
rc, txt, out = gen(p, "no-sep")
check("md 无分隔行：能识别为首行表头", rc == 0,
      "rc=%d %s" % (rc, txt.strip()[:110]))

# 列数不齐
p = os.path.join(DATA, "ragged.md")
with open(p, "w", encoding="utf-8") as f:
    f.write("| A | B | C |\n| --- | --- | --- |\n| 1 | 2 |\n| 4 | 5 | 6 | 7 |\n")
rc, txt, out = gen(p, "ragged")
if rc == 0:
    ws = load_workbook(out).active
    check("md 列数不齐：按表头列数对齐不崩", ws.max_column == 3,
          "max_column=%d" % ws.max_column)
else:
    check("md 列数不齐", False, txt.strip()[:120])

# 多表格块（应取第一个）
p = os.path.join(DATA, "multi-block.md")
with open(p, "w", encoding="utf-8") as f:
    f.write("| 第一 | 表 |\n| --- | --- |\n| a | 1 |\n\n说明文字\n\n"
            "| 第二 | 表 |\n| --- | --- |\n| b | 2 |\n")
rc, txt, out = gen(p, "multi-block")
if rc == 0:
    ws = load_workbook(out).active
    check("md 多表格块：取第一个", ws.cell(1, 1).value == "第一",
          "A1=%r" % (ws.cell(1, 1).value,))
else:
    check("md 多表格块", False, txt.strip()[:120])

# 空表头单元格
p = os.path.join(DATA, "empty-head.md")
with open(p, "w", encoding="utf-8") as f:
    f.write("| 名称 |  | 数量 |\n| --- | --- | --- |\n| 甲 | x | 1 |\n")
rc, txt, out = gen(p, "empty-head")
check("md 空表头单元格不崩", rc == 0, "rc=%d %s" % (rc, txt.strip()[:110]))

# 表格前后有大量文本
p = os.path.join(DATA, "surround.md")
with open(p, "w", encoding="utf-8") as f:
    f.write("# 标题\n\n" + "正文段落。\n" * 20 + "\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n"
            + "结尾文字。\n" * 10)
rc, txt, out = gen(p, "surround")
check("md 表格被长文本包围仍能提取", rc == 0, "rc=%d %s" % (rc, txt.strip()[:110]))

# ============ 3. CSV 引号与特殊字段 ============
p = os.path.join(DATA, "quoted.csv")
with open(p, "w", encoding="utf-8", newline="") as f:
    f.write('名称,说明\n"含,逗号",1\n"含""引号",2\n"含\n换行",3\n')
rc, txt, out = gen(p, "quoted")
if rc == 0:
    ws = load_workbook(out).active
    v1 = ws.cell(2, 1).value
    v2 = ws.cell(3, 1).value
    v3 = ws.cell(4, 1).value
    check("CSV 引号字段：含逗号不切列", v1 == "含,逗号", "实得 %r" % (v1,))
    check("CSV 引号字段：转义引号还原", v2 == '含"引号', "实得 %r" % (v2,))
    check("CSV 引号字段：字段内换行保留", str(v3).replace("\r", "") == "含\n换行",
          "实得 %r" % (v3,))
else:
    check("CSV 引号字段", False, txt.strip()[:120])

# 单列 CSV
p = os.path.join(DATA, "onecol.csv")
with open(p, "w", encoding="utf-8", newline="") as f:
    f.write("只有一列\n甲\n乙\n")
rc, txt, out = gen(p, "onecol")
if rc == 0:
    ws = load_workbook(out).active
    check("单列 CSV 正常生成", ws.max_column == 1,
          "max_column=%d" % ws.max_column)
else:
    check("单列 CSV", False, txt.strip()[:120])

# ============ 4. xlsx 源表形态 ============
# 合并单元格
p = os.path.join(DATA, "merged.xlsx")
wb = Workbook()
ws = wb.active
ws["A1"] = "合并标题"
ws.merge_cells("A1:C1")
ws.append(["型号", "数量", "备注"])
ws.append(["A1", 1, "x"])
ws.append(["A2", 2, "y"])
wb.save(p)
rc, txt, out = gen(p, "merged", extra=["--header-row", "2"])
check("源 xlsx 含合并单元格：不崩", rc == 0, "rc=%d %s" % (rc, txt.strip()[:110]))

# 公式单元格（data_only 读取缓存值）
p = os.path.join(DATA, "formula.xlsx")
wb = Workbook()
ws = wb.active
ws.append(["项目", "金额"])
ws.append(["A", 100])
ws.append(["B", "=B2*2"])
wb.save(p)
rc, txt, out = gen(p, "formula")
check("源 xlsx 含公式单元格：不崩", rc == 0, "rc=%d %s" % (rc, txt.strip()[:110]))
if rc == 0:
    ws2 = load_workbook(out).active
    check("公式单元格无缓存值时留空而非报错文本",
          ws2.cell(3, 2).value in (None, ""),
          "实得 %r" % (ws2.cell(3, 2).value,))

# 日期与布尔
p = os.path.join(DATA, "datebool.xlsx")
from datetime import date
wb = Workbook()
ws = wb.active
ws.append(["日期", "启用", "值"])
ws.append([date(2026, 11, 21), True, 1])
ws.append([date(2026, 1, 1), False, 0])
wb.save(p)
rc, txt, out = gen(p, "datebool")
if rc == 0:
    ws2 = load_workbook(out).active
    check("日期单元格正常写入", ws2.cell(2, 1).value is not None,
          "实得 %r" % (ws2.cell(2, 1).value,))
    check("布尔单元格正常写入", ws2.cell(2, 2).value in (True, False),
          "实得 %r" % (ws2.cell(2, 2).value,))
else:
    check("日期布尔", False, txt.strip()[:120])

# 空白行交错（表头在中间）
p = os.path.join(DATA, "sparse.xlsx")
wb = Workbook()
ws = wb.active
ws.append(["说明A"])
ws.append([])
ws.append([])
ws.append(["型号", "数量"])
ws.append(["A1", 1])
ws.append([])
ws.append(["A2", 2])
wb.save(p)
rc, txt, out = gen(p, "sparse", extra=["--header-row", "4"])
if rc == 0:
    ws2 = load_workbook(out).active
    check("空白行交错：表头在第 4 行取对", ws2.cell(1, 1).value == "型号",
          "A1=%r" % (ws2.cell(1, 1).value,))
    check("空白行交错：中间空行被跳过，两行数据都在",
          ws2.max_row == 3, "max_row=%d" % ws2.max_row)
else:
    check("空白行交错", False, txt.strip()[:120])

# ============ 5. 数据值边界 ============
p = os.path.join(DATA, "extreme.csv")
with open(p, "w", encoding="utf-8", newline="") as f:
    f.write("类型,值\n")
    f.write("极大整数,12345678901234567890\n")
    f.write("极大浮点,1.7976931348623157e308\n")
    f.write("极小浮点,5e-324\n")
    f.write("负极大,-12345678901234567890\n")
    f.write("纯空格,   \n")
    f.write("零,0\n")
rc, txt, out = gen(p, "extreme")
if rc == 0:
    ws = load_workbook(out).active
    check("极大整数不丢精度或按文本保留",
          str(ws.cell(2, 2).value).startswith("12345678901"),
          "实得 %r" % (ws.cell(2, 2).value,))
    check("极大浮点写入成功", ws.cell(3, 2).value is not None,
          "实得 %r" % (ws.cell(3, 2).value,))
    check("极小浮点写入成功", ws.cell(4, 2).value is not None,
          "实得 %r" % (ws.cell(4, 2).value,))
else:
    check("数据值边界", False, txt.strip()[:120])

print("=" * 84)
print("第二轮：编码与源文件形态边界")
print("=" * 84)
npass = 0
for ok, name, detail in results:
    if ok:
        npass += 1
    print("[%s] %-44s %s" % ("OK  " if ok else "FAIL", name, detail.replace("\n", " ")[:95]))
print("-" * 84)
print("ENCODING/SOURCE: %d/%d PASS" % (npass, len(results)))
sys.exit(0 if npass == len(results) else 1)
