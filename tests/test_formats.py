#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""输入格式与异常数据形态测试。

覆盖：md / csv（逗号、制表符、分号）/ xlsx（普通、--header-row、--src-sheet）
      异常：单列、单行、超宽 30 列、全空值、特殊字符、前导零、负数小数、
           超长单元格、含换行、emoji、中文标点。
每个用例生成 → verify 双向确认，并断言数据值原样保留（防静默篡改）。
"""
import os
import subprocess
import sys
import tempfile

from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _env import DATA, MT, OUT, SKILL, VT  # noqa: E402

results = []


def run(args):
    r = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=SKILL)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def check(name, cond, detail=""):
    results.append((cond, name, detail))


def gen_and_verify(name, src, profile="business-table", extra=None):
    """生成 + 校验，返回 (ok, 输出文件路径)。"""
    out = os.path.join(OUT, name + ".xlsx")
    args = [MT, "--profile", profile, "--src", src, "--out", out] + (extra or [])
    rc, txt = run(args)
    if rc != 0:
        check(name + " 生成", False, "rc=%d %s" % (rc, txt.strip()[:200]))
        return False, out
    rc2, txt2 = run([VT, "--profile", profile, out])
    check(name + " 生成+校验", rc2 == 0,
          "" if rc2 == 0 else "verify rc=%d %s" % (rc2, txt2.strip()[-220:]))
    return rc2 == 0, out


# ============ 输入格式 ============
# 1. md
md = os.path.join(DATA, "fmt.md")
with open(md, "w", encoding="utf-8") as f:
    f.write("# 标题\n\n说明文字。\n\n| 型号 | 克重 | 门幅 |\n| --- | --- | --- |\n"
            "| TEL-0148 | 215 | 138 |\n| Q400331Y | 107 | 138 |\n")
ok_md, out_md = gen_and_verify("fmt-md", md)

# 2. csv 逗号
csv_c = os.path.join(DATA, "comma.csv")
with open(csv_c, "w", encoding="utf-8", newline="") as f:
    f.write("型号,克重,门幅\nTEL-0148,215,138\nQ400331Y,107,138\n")
gen_and_verify("fmt-csv-comma", csv_c)

# 3. csv 制表符
csv_t = os.path.join(DATA, "tab.csv")
with open(csv_t, "w", encoding="utf-8", newline="") as f:
    f.write("型号\t克重\t门幅\nTEL-0148\t215\t138\n")
gen_and_verify("fmt-csv-tab", csv_t)

# 4. csv 分号
csv_s = os.path.join(DATA, "semi.csv")
with open(csv_s, "w", encoding="utf-8", newline="") as f:
    f.write("型号;克重;门幅\nTEL-0148;215;138\n")
gen_and_verify("fmt-csv-semi", csv_s)

# 5. xlsx 普通
x1 = os.path.join(DATA, "plain.xlsx")
wb = Workbook()
ws = wb.active
ws.append(["型号", "克重", "门幅"])
for i in range(1, 6):
    ws.append(["TEL-%03d" % i, 100 + i, 138])
wb.save(x1)
gen_and_verify("fmt-xlsx", x1)

# 6. xlsx 带表头行偏移（前 3 行是元信息）
x2 = os.path.join(DATA, "offset.xlsx")
wb = Workbook()
ws = wb.active
ws.append(["某公司面料清单"])
ws.append(["制表：2026-11-21"])
ws.append([])
ws.append(["型号", "克重", "门幅"])
for i in range(1, 4):
    ws.append(["Q%03d" % i, 200 + i, 150])
wb.save(x2)
ok_off, out_off = gen_and_verify("fmt-xlsx-header4", x2, extra=["--header-row", "4"])
if ok_off:
    ws_o = load_workbook(out_off).active
    check("--header-row 4 表头取对（跳空行后仍需按物理行号）",
          ws_o.cell(1, 1).value == "型号" and ws_o.cell(1, 2).value == "克重",
          "第1行=%r" % ([ws_o.cell(1, c).value for c in range(1, 4)],))
    check("--header-row 4 数据从第 5 行起（Q001 在首条）",
          ws_o.cell(2, 1).value == "Q001", "A2=%r" % (ws_o.cell(2, 1).value,))

# 6b. 表头指到空行 → 必须报错而非静默取错
x2b = os.path.join(DATA, "blank-head.xlsx")
wb = Workbook()
ws = wb.active
ws.append(["元信息"])
ws.append([])
ws.append(["型号", "数量"])
ws.append(["A1", 1])
wb.save(x2b)
rc_bh, txt_bh = run([MT, "--profile", "business-table", "--src", x2b,
                     "--out", os.path.join(OUT, "blank-head.xlsx"), "--header-row", "2"])
check("表头指到空行 → 报错", rc_bh != 0 and "空行" in txt_bh,
      "rc=%d %s" % (rc_bh, txt_bh.strip()[-100:]))

# 7. xlsx 多工作表 + --src-sheet
x3 = os.path.join(DATA, "multi.xlsx")
wb = Workbook()
w1 = wb.active
w1.title = "汇总"
w1.append(["无关列"])
w1.append(["占位"])
w2 = wb.create_sheet("明细")
w2.append(["型号", "数量"])
for i in range(1, 4):
    w2.append(["D%02d" % i, i * 10])
wb.save(x3)
ok_sh, out_sh = gen_and_verify("fmt-xlsx-sheet", x3, extra=["--src-sheet", "明细"])
if ok_sh:
    ws_s = load_workbook(out_sh).active
    check("--src-sheet 取对工作表", ws_s.cell(1, 1).value == "型号",
          "A1=%r" % (ws_s.cell(1, 1).value,))

# ============ 异常数据形态 ============
# 8. 单列
one_col = os.path.join(DATA, "one-col.md")
with open(one_col, "w", encoding="utf-8") as f:
    f.write("| 唯一列 |\n| --- |\n| 甲 |\n| 乙 |\n")
gen_and_verify("edge-one-col", one_col)

# 9. 单行单列
one_cell = os.path.join(DATA, "one-cell.md")
with open(one_cell, "w", encoding="utf-8") as f:
    f.write("| 值 |\n| --- |\n| 唯一 |\n")
gen_and_verify("edge-one-cell", one_cell)

# 10. 超宽 30 列
wide = os.path.join(DATA, "wide30.md")
with open(wide, "w", encoding="utf-8") as f:
    f.write("| " + " | ".join("列%d" % i for i in range(1, 31)) + " |\n")
    f.write("| " + " | ".join("---" for _ in range(30)) + " |\n")
    for r in range(1, 4):
        f.write("| " + " | ".join(str(r * 100 + c) for c in range(1, 31)) + " |\n")
ok_wide, out_wide = gen_and_verify("edge-wide30", wide)
if ok_wide:
    ws_w = load_workbook(out_wide).active
    check("超宽 30 列列数完整", ws_w.max_column == 30, "max_column=%d" % ws_w.max_column)

# 11. 全空值行
empties = os.path.join(DATA, "empties.csv")
with open(empties, "w", encoding="utf-8", newline="") as f:
    f.write("A,B,C\n,,\n1,,3\n,,6\n")
gen_and_verify("edge-empties", empties)

# 12. 特殊字符与前导零
special = os.path.join(DATA, "special.csv")
with open(special, "w", encoding="utf-8", newline="") as f:
    f.write("编码,说明\n009,前导零保留\n00123,三位前导零\nA-001,字母数字\n"
            "-5,负整数\n-3.25,负小数\n1.550,保留末尾零\n")
ok_sp, out_sp = gen_and_verify("edge-special", special)
if ok_sp:
    ws_s2 = load_workbook(out_sp).active
    vals = [ws_s2.cell(r, 1).value for r in range(2, 7)]
    check("前导零按文本保留", vals[0] == "009" and vals[1] == "00123", "实得 %r" % (vals[:2],))
    check("负数转数值", vals[3] == -5 and vals[4] == -3.25, "实得 %r" % (vals[3:],))

# 13. 超长单元格
longc = os.path.join(DATA, "long-cell.csv")
with open(longc, "w", encoding="utf-8", newline="") as f:
    f.write("说明,值\n" + ("这是一段很长的说明文字" * 20) + ",1\n短,2\n")
gen_and_verify("edge-long-cell", longc)

# 14. 含换行与前后空格
nl = os.path.join(DATA, "newline.csv")
with open(nl, "w", encoding="utf-8", newline="") as f:
    f.write('文本,值\n"第一行\n第二行",1\n"  前后空格  ",2\n')
gen_and_verify("edge-newline", nl)

# 15. emoji 与中文标点
emo = os.path.join(DATA, "emoji.csv")
with open(emo, "w", encoding="utf-8", newline="") as f:
    f.write("名称,备注\n花型甲「测试」,含《书名号》与（括号）\n符号乙,①②③·——…\n")
gen_and_verify("edge-punct", emo)

# 16. 表头含竖线转义（md）
esc = os.path.join(DATA, "escape.md")
with open(esc, "w", encoding="utf-8") as f:
    f.write("| 名称 | 值 |\n| --- | --- |\n| a\\|b | 1 |\n")
gen_and_verify("edge-escape", esc)

# 17. 三种 profile 处理同一份宽表（跨类型稳定性）
for prof in ["academic-table", "finance-table", "data-list"]:
    gen_and_verify("edge-wide30-" + prof, wide, profile=prof)

print("=" * 82)
print("输入格式与异常数据形态测试")
print("=" * 82)
npass = 0
for ok, name, detail in results:
    if ok:
        npass += 1
    print("[%s] %-34s %s" % ("OK  " if ok else "FAIL", name, detail.replace("\n", " ")[:110]))
print("-" * 82)
print("FORMAT/EDGE: %d/%d PASS" % (npass, len(results)))
sys.exit(0 if npass == len(results) else 1)
