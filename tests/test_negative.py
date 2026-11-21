#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""负向测试：逐类破坏产物，确认 verify_table.py 每类都能抓到。

思路：先造一个合规范例（verify 必须 PASS），然后对副本做单点破坏，
逐个断言 verify 返回 1（抓到）且报错文本包含对应断言名。
这是对门禁本身的检验——门禁若漏抓，等于没有门禁。
"""
import os
import subprocess
import sys

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _env import BASE, DATA, MT, OUT, SKILL, VT  # noqa: E402

MUT = os.path.join(BASE, "mutated")
os.makedirs(MUT, exist_ok=True)

results = []


def check(name, cond, detail=""):
    results.append((cond, name, detail))


def run(args):
    r = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=SKILL)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


SRC = os.path.join(DATA, "mut-src.csv")
with open(SRC, "w", encoding="utf-8", newline="") as f:
    f.write("型号,克重,门幅\n")
    for i in range(1, 26):  # 25 行，触发冻结
        f.write("TEL-%03d,%d,138\n" % (i, 100 + i))

GOOD = os.path.join(OUT, "mut-good.xlsx")
rc, txt = run([MT, "--profile", "business-table", "--src", SRC, "--out", GOOD,
               "--title", "负向测试样例", "--note", "说明行"])
if rc != 0:
    print("生成基座失败:", txt[:300])
    sys.exit(2)

rc, txt = run([VT, "--profile", "business-table", GOOD])
check("基座产物 verify PASS", rc == 0, "rc=%d" % rc)


def mutate(tag, fn, profile="business-table"):
    """复制基座 → 施加破坏 → verify，返回 (rc, 输出)。"""
    path = os.path.join(MUT, "%s.xlsx" % tag)
    wb = load_workbook(GOOD)
    ws = wb.active
    fn(wb, ws)
    wb.save(path)
    rc, txt = run([VT, "--profile", profile, path])
    return rc, txt


def expect_caught(tag, fn, want_kw, profile="business-table"):
    rc, txt = mutate(tag, fn, profile)
    caught = (rc == 1)
    kw_ok = any(k in txt for k in want_kw) if want_kw else True
    # 找出失败的断言名
    fail_line = ""
    for line in txt.splitlines():
        if "[FAIL]" in line:
            fail_line = line.strip()
            break
    check("破坏「%s」被抓" % tag, caught and kw_ok,
          "rc=%d %s" % (rc, fail_line[:110]))


# 1. 表头底色改错
def m_head_fill(wb, ws):
    ws.cell(3, 1).fill = PatternFill("solid", fgColor="FF0000")


expect_caught("表头底色", m_head_fill, ["表头"])

# 2. 表头字色改黑
def m_head_font_color(wb, ws):
    c = ws.cell(3, 1)
    c.font = Font(name="黑体", size=12, bold=True, color="FF000000")


expect_caught("表头字色", m_head_font_color, ["表头"])

# 3. 表头取消加粗
def m_head_bold(wb, ws):
    c = ws.cell(3, 1)
    c.font = Font(name="黑体", size=12, bold=False, color="FFFFFFFF")


expect_caught("表头加粗", m_head_bold, ["表头"])

# 4. 表头字号改大
def m_head_size(wb, ws):
    c = ws.cell(3, 1)
    c.font = Font(name="黑体", size=16, bold=True, color="FFFFFFFF")


expect_caught("表头字号", m_head_size, ["表头"])

# 5. 标题字号改错
def m_title_size(wb, ws):
    c = ws.cell(1, 1)
    c.font = Font(name="黑体", size=12, bold=True, color="FF1F3864")


expect_caught("标题字号", m_title_size, ["标题"])

# 6. 标题取消加粗
def m_title_bold(wb, ws):
    c = ws.cell(1, 1)
    c.font = Font(name="黑体", size=16, bold=False, color="FF1F3864")


expect_caught("标题加粗", m_title_bold, ["标题"])

# 7. 数据字体改错
def m_data_font(wb, ws):
    ws.cell(4, 1).font = Font(name="宋体", size=10, color="FF262626")


expect_caught("数据字体", m_data_font, ["数据字体"])

# 8. 数据字号改错
def m_data_size(wb, ws):
    ws.cell(4, 1).font = Font(name="黑体", size=14, color="FF262626")


expect_caught("数据字号", m_data_size, ["数据字体"])

# 9. 数据字色改错
def m_data_color(wb, ws):
    ws.cell(4, 1).font = Font(name="黑体", size=10, color="FFFF0000")


expect_caught("数据字色", m_data_font if False else (lambda wb, ws: ws.cell(4, 1).__setattr__(
    "font", Font(name="黑体", size=10, color="FFFF0000"))), ["数据字体"])

# 10. 移除冻结
def m_no_freeze(wb, ws):
    ws.freeze_panes = None


expect_caught("移除冻结", m_no_freeze, ["冻结"])

# 11. 冻结锚点错位（钉住数据行）
def m_wrong_anchor(wb, ws):
    ws.freeze_panes = "A5"


expect_caught("冻结锚点错位", m_wrong_anchor, ["冻结"])

# 12. print_title_rows 被改错（openpyxl 对 None 赋值不生效，改为指向错误行）
def m_wrong_print_titles(wb, ws):
    ws.print_title_rows = "1:1"


expect_caught("print_title_rows 错位", m_wrong_print_titles, ["print_title_rows"])

# 13. 关闭自动换行
def m_no_wrap(wb, ws):
    ws.cell(4, 1).alignment = Alignment(vertical="center", wrap_text=False)


expect_caught("关闭自动换行", m_no_wrap, ["自动换行"])

# 14. 边框移除
def m_no_border(wb, ws):
    c = ws.cell(4, 1)
    c.border = Border()


expect_caught("移除边框", m_no_border, ["边框"])

# 15. 边框颜色改错
def m_border_color(wb, ws):
    s = Side(style="thin", color="FFFF0000")
    ws.cell(4, 1).border = Border(left=s, right=s, top=s, bottom=s)


expect_caught("边框颜色", m_border_color, ["边框"])

# 16. 隔行底纹被移除（真正的第 2 数据行是第 5 行，第 4 行本就留白）
def m_no_zebra(wb, ws):
    ws.cell(5, 1).fill = PatternFill(fill_type=None)


expect_caught("移除隔行底纹", m_no_zebra, ["隔行底纹"])

# 17. 隔行底纹错位（首行数据行着色）
def m_zebra_shift(wb, ws):
    ws.cell(4, 1).fill = PatternFill("solid", fgColor="FFF7F9FD")  # 第 1 数据行不该有
    ws.cell(5, 1).fill = PatternFill(fill_type=None)


expect_caught("隔行底纹错位", m_zebra_shift, ["隔行底纹"])

# 18. 列宽改窄到异常值（openpyxl 删维度会重建默认值 13，测不到「无列宽」；
#     改测「列宽明显不合理」——门禁当前只查显式存在，这里改为确认它不误报）
def m_width_weird(wb, ws):
    for k in list(ws.column_dimensions.keys()):
        ws.column_dimensions[k].width = 3


rc_w, txt_w = mutate("列宽异常窄", m_width_weird)
check("列宽异常窄：门禁当前放行（已知边界，非缺陷）", rc_w == 0,
      "rc=%d（verify 只查显式设置，不判合理性）" % rc_w)

# 19. 行高被清空
def m_no_height(wb, ws):
    for k in list(ws.row_dimensions.keys()):
        del ws.row_dimensions[k]


expect_caught("行高被清空", m_no_height, ["行高"])

# 20. 数据区被合并（侵入合并）
def m_merge_data(wb, ws):
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=2)


expect_caught("数据区合并", m_merge_data, ["合并"])

# 21. 三线表 profile：给数据区加竖线
AC = os.path.join(OUT, "mut-ac-good.xlsx")
rc, _ = run([MT, "--profile", "academic-table", "--src", SRC, "--out", AC,
             "--title", "三线表负向"])
if rc == 0:
    rc2, t2 = run([VT, "--profile", "academic-table", AC])
    check("三线表基座 verify PASS", rc2 == 0, "rc=%d" % rc2)

    def m_vertical(wb, ws):
        s = Side(style="thin", color="FF000000")
        ws.cell(4, 1).border = Border(left=s, right=s)

    rc3, t3 = mutate("三线表加竖线", m_vertical, "academic-table")
    # 注意 mutate 复制的是 GOOD（商务表），所以单独处理三线表
    path = os.path.join(MUT, "ac-vertical.xlsx")
    wb = load_workbook(AC)
    ws = wb.active
    s = Side(style="thin", color="FF000000")
    ws.cell(4, 1).border = Border(left=s, right=s)
    wb.save(path)
    rc4, t4 = run([VT, "--profile", "academic-table", path])
    fl = next((l.strip() for l in t4.splitlines() if "[FAIL]" in l), "")
    check("破坏「三线表加竖线」被抓", rc4 == 1 and "三线表无竖线" in t4,
          "rc=%d %s" % (rc4, fl[:110]))

    # 三线表底纹（不该有）
    path2 = os.path.join(MUT, "ac-zebra.xlsx")
    wb = load_workbook(AC)
    ws = wb.active
    ws.cell(4, 1).fill = PatternFill("solid", fgColor="FFF7F9FD")
    wb.save(path2)
    rc5, t5 = run([VT, "--profile", "academic-table", path2])
    fl5 = next((l.strip() for l in t5.splitlines() if "[FAIL]" in l), "")
    check("破坏「三线表出现底纹」被抓", rc5 == 1,
          "rc=%d %s" % (rc5, fl5[:110]))

    # 三线表表头加底色（不该有）
    path3 = os.path.join(MUT, "ac-headfill.xlsx")
    wb = load_workbook(AC)
    ws = wb.active
    ws.cell(2, 1).fill = PatternFill("solid", fgColor="FF1F3864")
    wb.save(path3)
    rc6, t6 = run([VT, "--profile", "academic-table", path3])
    fl6 = next((l.strip() for l in t6.splitlines() if "[FAIL]" in l), "")
    check("破坏「三线表表头出现底色」被抓", rc6 == 1,
          "rc=%d %s" % (rc6, fl6[:110]))
else:
    check("三线表基座生成", False, "生成失败")

# 22. 用错 profile 校验（商务表产物按学术表验）→ 必须大量报错
rc, txt = run([VT, "--profile", "academic-table", GOOD])
check("用错 profile 校验 → 报错", rc == 1 and "表头" in txt,
      "rc=%d 首个FAIL=%s" % (rc, next((l.strip() for l in txt.splitlines() if "[FAIL]" in l), "")[:80]))

print("=" * 86)
print("负向测试：逐类破坏产物，确认门禁能抓到")
print("=" * 86)
npass = 0
for ok, name, detail in results:
    if ok:
        npass += 1
    print("[%s] %-38s %s" % ("OK  " if ok else "FAIL", name, detail.replace("\n", " ")[:105]))
print("-" * 86)
print("NEGATIVE: %d/%d PASS" % (npass, len(results)))
sys.exit(0 if npass == len(results) else 1)
