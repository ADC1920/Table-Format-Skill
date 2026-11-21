#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""八个 profile 的样式真值核验。

不依赖 verify_table.py 的判断（那是被测对象自己的断言），而是独立地：
读 profiles/<类型>.yaml 的原始取值 → 生成表格 → 直接打开 xlsx 逐单元格比对。

核验项（每 profile 全查）：
  header 底色/字色/字体/字号/加粗/对齐
  title 底色/字色/字号/加粗/合并范围
  note 字号/字色
  body 字体/字号/字色/边框或三线表线条/隔行底纹开关与颜色
  页眉打印行、纸张方向、网格线
"""
import os
import subprocess
import sys

import yaml
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _env import DATA, MT, OUT, PROFILES, SKILL  # noqa: E402
SRC = os.path.join(DATA, "style-src.csv")
with open(SRC, "w", encoding="utf-8", newline="") as f:
    f.write("型号,克重,门幅,备注\n")
    for i in range(1, 7):
        # 第 3、6 行克重为 0，用于验证 finance-table 的零值短横
        # 与 stats-table 的零值填 0（两者规范依据不同，见规范库 t2-047 / t1-027）
        kz = 0 if i in (3, 6) else 100 + i
        f.write("TEL-%03d,%d,138,第%d行备注\n" % (i, kz, i))
    # 加一行负数与上千万的金额列，验证千分位与负号
    f.write("TEL-007,-52341.5,138,负数行\n")
    f.write("TEL-008,1234567.89,138,大额行\n")

results = []


def check(prof, name, cond, detail=""):
    results.append((cond, prof, name, detail))


def hexof(c):
    try:
        v = c.rgb
        return str(v).upper() if v is not None else None
    except Exception:
        return None


def ends(c, h):
    x = hexof(c)
    return x is not None and h is not None and x.endswith(str(h).upper())


def num(v):
    return None if v is None else float(v)


def run(args):
    r = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=SKILL)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def load_raw(prof):
    """读 profile 原始 yaml，并把基线 table-format.yaml 作为缺省来源。

    profile 只写差异字段，未写的字段由 spec_loader 从基线继承——所以核验时
    缺省项必须回落到基线值，否则会把「正确的继承」误判成失败。
    """
    with open(os.path.join(SKILL, "table-format.yaml"), encoding="utf-8") as f:
        base = yaml.safe_load(f) or {}
    with open(os.path.join(PROFILES, prof + ".yaml"), encoding="utf-8") as f:
        own = yaml.safe_load(f) or {}

    def merge(b, o):
        out = dict(b)
        for k, v in o.items():
            if isinstance(v, dict) and isinstance(out.get(k), dict):
                out[k] = merge(out[k], v)
            else:
                out[k] = v
        return out

    return merge(base, own)


def g(d, *path, default=None):
    cur = d
    for p in path:
        if not isinstance(cur, dict) or p not in cur:
            return default
        cur = cur[p]
    return cur


names = sorted(os.path.splitext(f)[0] for f in os.listdir(PROFILES)
               if f.endswith((".yaml", ".yml")))

for prof in names:
    raw = load_raw(prof)
    out = os.path.join(OUT, "style-" + prof + ".xlsx")
    rc, txt = run([MT, "--profile", prof, "--src", SRC, "--out", out,
                   "--title", "样式核验样例", "--note", "说明：样式真值核验"])
    if rc != 0:
        check(prof, "生成", False, "rc=%d %s" % (rc, txt.strip()[:160]))
        continue

    ws = load_workbook(out).active

    # --- 定位表头行（标题 1 + 说明 1 → 第 3 行）---
    head_row = 3
    # 三线表无底色，用结构确认表头在第 3 行
    if g(raw, "header", "fill") is None:
        check(prof, "表头行位置", ws.cell(head_row, 1).value == "型号",
              "A3=%r" % (ws.cell(head_row, 1).value,))

    # --- 表头 ---
    hc = ws.cell(head_row, 1)
    want_fill = g(raw, "header", "fill")
    if want_fill:
        check(prof, "表头底色 %s" % want_fill, ends(hc.fill.fgColor, want_fill),
              "实得 %s" % hexof(hc.fill.fgColor))
    else:
        check(prof, "表头无底色（yaml 为 null）",
              hexof(hc.fill.fgColor) in (None, "00000000"),
              "实得 %s" % hexof(hc.fill.fgColor))
    check(prof, "表头字色 %s" % g(raw, "header", "font_color"),
          ends(hc.font.color, g(raw, "header", "font_color")),
          "实得 %s" % hexof(hc.font.color))
    check(prof, "表头字体 %s" % (g(raw, "header", "font_name") or g(raw, "font", "name")),
          hc.font.name == (g(raw, "header", "font_name") or g(raw, "font", "name")),
          "实得 %s" % hc.font.name)
    check(prof, "表头字号 %s" % g(raw, "header", "size"),
          num(hc.font.size) == num(g(raw, "header", "size")), "实得 %s" % hc.font.size)
    check(prof, "表头加粗 %s" % g(raw, "header", "bold"),
          bool(hc.font.bold) == bool(g(raw, "header", "bold")), "实得 %s" % hc.font.bold)
    check(prof, "表头水平对齐 %s" % g(raw, "header", "align"),
          (hc.alignment.horizontal or "") == g(raw, "header", "align"),
          "实得 %s" % hc.alignment.horizontal)

    # --- 标题 ---
    tc = ws.cell(1, 1)
    tf = g(raw, "title", "fill")
    if tf:
        check(prof, "标题底色 %s" % tf, ends(tc.fill.fgColor, tf),
              "实得 %s" % hexof(tc.fill.fgColor))
    check(prof, "标题字色 %s" % g(raw, "title", "font_color"),
          ends(tc.font.color, g(raw, "title", "font_color")),
          "实得 %s" % hexof(tc.font.color))
    check(prof, "标题字号 %s" % g(raw, "title", "size"),
          num(tc.font.size) == num(g(raw, "title", "size")), "实得 %s" % tc.font.size)
    check(prof, "标题加粗 %s" % g(raw, "title", "bold"),
          bool(tc.font.bold) == bool(g(raw, "title", "bold")), "实得 %s" % tc.font.bold)
    # 合并范围 = A1:末列1
    merged = [str(r) for r in ws.merged_cells.ranges if r.min_row == 1]
    want_merge = "A1:%s1" % chr(ord("A") + ws.max_column - 1)
    check(prof, "标题整行合并 %s" % want_merge, any(m == want_merge for m in merged),
          "实得 %s" % merged)

    # --- 说明行 ---
    nc = ws.cell(2, 1)
    check(prof, "说明字号 %s" % g(raw, "note", "size"),
          num(nc.font.size) == num(g(raw, "note", "size")), "实得 %s" % nc.font.size)
    check(prof, "说明字色 %s" % g(raw, "note", "color"),
          ends(nc.font.color, g(raw, "note", "color")), "实得 %s" % hexof(nc.font.color))

    # --- 数据区 ---
    d1 = ws.cell(head_row + 1, 1)
    check(prof, "数据字体 %s" % g(raw, "font", "name"),
          d1.font.name == g(raw, "font", "name"), "实得 %s" % d1.font.name)
    check(prof, "数据字号 %s" % g(raw, "font", "data_size"),
          num(d1.font.size) == num(g(raw, "font", "data_size")), "实得 %s" % d1.font.size)
    check(prof, "数据字色 %s" % g(raw, "font", "data_color"),
          ends(d1.font.color, g(raw, "font", "data_color")), "实得 %s" % hexof(d1.font.color))

    # --- 三线表 vs 普通边框 ---
    three = bool(g(raw, "rules", "three_line", default=False))
    if three:
        check(prof, "三线表：表头顶线存在",
              getattr(ws.cell(head_row, 1).border.top, "style", None) is not None,
              "实得 %s" % getattr(ws.cell(head_row, 1).border.top, "style", None))
        check(prof, "三线表：表头下横线存在",
              getattr(ws.cell(head_row, 1).border.bottom, "style", None) is not None,
              "实得 %s" % getattr(ws.cell(head_row, 1).border.bottom, "style", None))
        check(prof, "三线表：数据区无竖线",
              getattr(ws.cell(head_row + 1, 1).border.left, "style", None) is None
              and getattr(ws.cell(head_row + 1, 1).border.right, "style", None) is None,
              "left=%s right=%s" % (getattr(ws.cell(head_row + 1, 1).border.left, "style", None),
                                    getattr(ws.cell(head_row + 1, 1).border.right, "style", None)))
        check(prof, "三线表：末行有底线",
              getattr(ws.cell(ws.max_row, 1).border.bottom, "style", None) is not None,
              "实得 %s" % getattr(ws.cell(ws.max_row, 1).border.bottom, "style", None))
    else:
        bs = g(raw, "body", "border_style")
        if bs:
            check(prof, "数据区边框样式 %s" % bs,
                  getattr(d1.border.left, "style", None) == bs,
                  "实得 %s" % getattr(d1.border.left, "style", None))
            check(prof, "数据区边框色 %s" % g(raw, "body", "border_color"),
                  ends(d1.border.left.color, g(raw, "body", "border_color")),
                  "实得 %s" % hexof(d1.border.left.color))

    # --- 隔行底纹 ---
    zebra = bool(g(raw, "body", "zebra", default=False))
    zc = g(raw, "body", "zebra_color")
    row2 = ws.cell(head_row + 2, 1)
    if zebra and zc:
        start = int(g(raw, "body", "zebra_start", default=2))
        first = ws.cell(head_row + 1, 1)
        if start >= 2:
            check(prof, "首行数据行留白", not ends(first.fill.fgColor, zc),
                  "实得 %s" % hexof(first.fill.fgColor))
        check(prof, "第%d数据行有底纹 %s" % (start, zc), ends(row2.fill.fgColor, zc),
              "实得 %s" % hexof(row2.fill.fgColor))
    else:
        check(prof, "无隔行底纹（yaml 关闭）", not ends(row2.fill.fgColor, "F7F9FD"),
              "实得 %s" % hexof(row2.fill.fgColor))

    # --- 数字格式（spec.number）---
    # 这组断言是为抓「配置写了但代码不读」的假实现：2026-11-21 实测发现
    # finance/stats 宣称千分位与零值短横，而 number_format 恒为 General。
    ncfg = raw.get("number") or {} if isinstance(raw, dict) else {}
    if "number" in raw:
        ncfg = raw["number"] or {}
    fmt_on = bool(ncfg.get("format_auto"))
    zero_dash = bool(ncfg.get("zero_as_dash"))
    num_cells = []
    for rr in range(head_row + 1, ws.max_row + 1):
        for j in range(1, ws.max_column + 1):
            v = ws.cell(rr, j).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                num_cells.append((rr, j))
    if num_cells:
        if fmt_on:
            nofmt = [(rr, j) for rr, j in num_cells
                     if (ws.cell(rr, j).number_format or "General") == "General"]
            check(prof, "数字格式：format_auto 已落到单元格", not nofmt,
                  "N=%d 未设格式 %s" % (len(num_cells), nofmt[:3]))
        else:
            wrongly = [(rr, j) for rr, j in num_cells
                       if (ws.cell(rr, j).number_format or "General") != "General"]
            check(prof, "数字格式：未开 format_auto 时保持 General", not wrongly,
                  "N=%d 误设 %s" % (len(num_cells), wrongly[:3]))
    # 零值短横：源数据里第 3、6 行克重为 0，开了 zero_as_dash 的 profile
    # 应把这些 0 变成文字「-」；没开的应仍为数值 0。
    if zero_dash:
        dash_cells = [(rr, j)
                      for rr in range(head_row + 1, ws.max_row + 1)
                      for j in range(1, ws.max_column + 1)
                      if ws.cell(rr, j).value == "-"]
        check(prof, "零值短横：源有 0 时产出为「-」", len(dash_cells) > 0,
              "检出短横 %d 处" % len(dash_cells))
    else:
        dash_cells = [(rr, j)
                      for rr in range(head_row + 1, ws.max_row + 1)
                      for j in range(1, ws.max_column + 1)
                      if ws.cell(rr, j).value == "-"]
        check(prof, "零值不替短横：未开 zero_as_dash 时 0 原样保留",
              len(dash_cells) == 0, "误检出短横 %d 处" % len(dash_cells))

    # --- 打印设置 ---
    check(prof, "print_title_rows=表头行",
          (ws.print_title_rows or "").replace("$", "") == "%d:%d" % (head_row, head_row),
          "实得 %s" % ws.print_title_rows)
    want_or = g(raw, "print", "orientation")
    if want_or:
        check(prof, "纸张方向 %s" % want_or,
              (ws.page_setup.orientation or "").lower() == want_or.lower(),
              "实得 %s" % ws.page_setup.orientation)

    # --- 全表自动换行 ---
    bad = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            if cell.value is None:
                continue
            if not cell.alignment.wrap_text:
                bad.append("R%dC%d" % (r, c))
    check(prof, "全表自动换行", not bad, "未开启 %s" % bad[:5])

    # --- 数据区无合并 ---
    invade = [str(r) for r in ws.merged_cells.ranges if r.min_row > 2]
    check(prof, "数据区无合并", not invade, "侵入 %s" % invade)

print("=" * 84)
print("八个 profile 样式真值核验（独立读 yaml 对照，不依赖 verify 自述）")
print("=" * 84)
by_prof = {}
for ok, prof, name, detail in results:
    by_prof.setdefault(prof, []).append((ok, name, detail))
total_ok = 0
for prof in sorted(by_prof):
    items = by_prof[prof]
    npass = sum(1 for ok, _, _ in items if ok)
    total_ok += npass
    status = "PASS" if npass == len(items) else "FAIL"
    print("\n== %s == %d/%d %s" % (prof, npass, len(items), status))
    for ok, name, detail in items:
        if not ok:
            print("   [FAIL] %-32s %s" % (name, detail))
print("\n" + "-" * 84)
print("STYLE: %d/%d PASS" % (total_ok, len(results)))
sys.exit(0 if total_ok == len(results) else 1)
