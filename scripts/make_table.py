# -*- coding: utf-8 -*-
"""make_table.py — table-format 规格化 xlsx 的唯一构建通道。

把 md / CSV / xlsx 表格按指定表格类型 profile 重排为 xlsx 工作簿。
规格唯一权威源是 `profiles/<类型>.yaml`（差异字段）+ `table-format.yaml`
（基线与继承来源）。本文件不含任何样式常量——所有颜色、字号、边框、
行高、冻结、纸张全部从 spec 读取。

**v2.0 起无默认样式**：必须用 `--profile` 指定表格类型，不指定则 exit 2。
八个类型：business-table / academic-table / finance-table / stats-table /
gov-table / quality-table / hr-table / data-list。
列全部：`python make_table.py --list-profiles`。

规格要点（各 profile 的具体取值见其 yaml，此处只列结构）：
    - 表头行：fill / font_color / size / bold / align 全部规格化
    - 标题行（可选）：可关闭底纹与合并（学术/财务/政务表多为无底色）
    - 数据行：字体色、隔行底纹（可关）、边框（可无，三线表用 rules 画线）
    - 三线表：rules.three_line=true 时只画顶线、栏目线、底线，无竖线无底色
    - 全表自动换行：所有单元格 wrap_text=True（含标题/说明行）
    - 冻结：按 spec.freeze.enabled（auto/always/never）+ min_rows + anchor
    - print_title_rows / 纸张 / 方向 / 页边距 / 网格线按 spec.print

输入约定：
    - md：取第一个 GFM 表格块（连续的 | 行），首行为表头
    - CSV：自动嗅探分隔符，首行为表头
    - xlsx：默认第一个工作表，--header-row 指定表头行（默认 1）；合并单元格不保留

已知坑（勿回退，2026-10-15 / 10-16 两次实测踩坑）：
    openpyxl 的 read_only 模式能力受限，读源工作簿一律不要 read_only=True：
      - ReadOnlyWorksheet 无 dimensions / column_dimensions / row_dimensions 属性
        （AttributeError: 'ReadOnlyWorksheet' object has no attribute 'dimensions'）
      - read_only 迭代产生的 EmptyCell 无 .row 属性（稀疏行上直接报错）
    需要读源数据时用普通模式加载，或 iter_rows(values_only=True) 一次性取值。

用法：
    python make_table.py --profile finance-table --src 数据.csv --out 成品.xlsx
        [--sheet 表名] [--title "标题"] [--note "说明行"]
        [--src-sheet 源表名] [--header-row 1] [--no-freeze]
    python make_table.py --list-profiles   # 列出全部表格类型
    python make_table.py --selftest        # 自检：建样例→生成→回读断言→退出码 0/1
"""

import argparse
import csv
import math
import os
import re
import sys
import tempfile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from spec_loader import (ProfileNotFound, ProfileNotSpecified, list_profiles,
                         load_spec, profile_meta, resolve_profile)


def _num(v):
    """spec 里的字号可能是 int 或 float，统一成数值便于比较。"""
    return None if v is None else float(v)


def display_width(text):
    """显示宽度：中日韩字符计 2 宽，其余计 1。"""
    w = 0
    for ch in str(text):
        w += 2 if ord(ch) > 0x2E7F else 1
    return w


def maybe_num(s):
    """纯数字字符串转 int/float；保护前导零（如编码 009）与带分隔符文本。"""
    t = str(s).strip()
    if re.fullmatch(r"[+-]?\d+", t):
        if len(t.lstrip("+-")) > 1 and t.lstrip("+-")[0] == "0":
            return s  # 前导零，按文本保留
        return int(t)
    if re.fullmatch(r"[+-]?\d+\.\d+", t):
        return float(t)
    return s


# ---------- 输入解析 ----------

def load_md_table(path):
    """取 Markdown 中第一个 GFM 表格块。返回 (headers, rows)。"""
    with open(path, encoding="utf-8") as f:
        lines = f.read().splitlines()
    block = []
    for ln in lines:
        s = ln.strip()
        if s.startswith("|") and s.endswith("|") and len(s) > 2:
            block.append(s)
        elif block:
            break
    if not block:
        raise SystemExit("ERROR: %s 中未找到 Markdown 表格（| 行）" % path)

    def cells(row):
        # \| 为转义竖线字符，先占位后还原，避免被当列分隔符切错列
        return [c.strip().replace("\x00", "|")
                for c in row.strip("|").replace("\\|", "\x00").split("|")]

    table = [cells(b) for b in block]
    table = [r for r in table if not all(re.fullmatch(r":?-{2,}:?", c or "-") for c in r)]
    if not table:
        raise SystemExit("ERROR: %s 表格块为空" % path)
    headers = table[0]
    rows = [[maybe_num(c) for c in r] for r in table[1:]]
    return headers, rows


def load_csv_table(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
        except csv.Error:
            dialect = csv.excel
        reader = csv.reader(f, dialect)
        table = [list(r) for r in reader if any(str(c).strip() for c in r)]
    if not table:
        raise SystemExit("ERROR: %s 为空" % path)
    headers = [str(c).strip() for c in table[0]]
    rows = [[maybe_num(c) for c in r] for r in table[1:]]
    return headers, rows


def load_xlsx_table(path, src_sheet=None, header_row=1):
    """读源 xlsx。注意：不用 read_only=True（见文件头「已知坑」）。"""
    wb = load_workbook(path, data_only=True)
    ws = wb[src_sheet] if src_sheet else wb[wb.sheetnames[0]]
    values = []
    for r in ws.iter_rows(values_only=True):  # 一次性取值，避开 read_only/EmptyCell 类问题
        if r is None or all(v is None or str(v).strip() == "" for v in r):
            continue
        values.append(list(r))
    if len(values) < header_row:
        raise SystemExit("ERROR: 源表行数不足 --header-row=%d" % header_row)
    headers = [("" if v is None else str(v).strip()) for v in values[header_row - 1]]
    while headers and headers[-1] == "":
        headers.pop()  # 去掉尾部空列
    ncol = len(headers)
    rows = []
    for r in values[header_row:]:
        r = list(r[:ncol]) + [None] * max(0, ncol - len(r))
        rows.append(r[:ncol])
    return headers, rows


# ---------- 生成 ----------

def est_lines(value, col_width):
    if value is None:
        return 1
    w = display_width(value)
    usable = max(col_width - 2, 4)
    return max(1, math.ceil(w / usable))


def _side(hexv, style):
    """按 spec 画一条边；style 为 None 表示不画线。"""
    if not style or not hexv:
        return Side(style=None)
    return Side(style=style, color=hexv)


# A4 纸宽 21.0 cm，扣掉左右页边距后是可用打印宽度。
# Excel 列宽单位是「字符宽」；实测在该字体与缩放组合下 1 单位 ≈ 0.20 cm
# （取 0.185 会压得过紧，行高因换行暴涨、页数反增）。
_CM_PER_WIDTH_UNIT = 0.205
_A4_WIDTH_CM = 21.0
_DEFAULT_MARGIN_CM = 1.5


def _fit_widths_to_page(widths, pr, cw):
    """列宽合计超过纸张可用宽度时等比压缩，避免打印横向拆页。

    等比而不是砍最宽列的原因：砍单列会把那一列压成竖排文字，
    等比压缩只是让所有列都窄一点，配合自动换行仍可读。
    压到最小列宽仍放不下时（列数过多）无法挽救，返回时保持最小列宽，
    由调用方在打印设置里改用横向或更大纸张。
    """
    orientation = str(pr.get("orientation") or "landscape").lower()
    paper = str(pr.get("paper") or "A4").upper()
    page_w = _A4_WIDTH_CM
    if paper == "A3":
        page_w = 29.7
    # 横向时纸宽与纸高互换
    usable_w = (29.7 if page_w == 21.0 else 42.0) if orientation == "landscape" else page_w
    margin = pr.get("margin_cm")
    if margin:
        usable_w -= 2 * float(margin)
    else:
        usable_w -= 2 * _DEFAULT_MARGIN_CM
    if usable_w <= 0:
        return widths
    max_units = usable_w / _CM_PER_WIDTH_UNIT
    total = sum(widths)
    if total <= max_units:
        return widths
    scale = max_units / float(total)
    shrunk = [max(int(w * scale), cw["min"]) for w in widths]
    # 最小列宽兜底后若仍超宽，按超出比例再压一次（可能低于 min，这是必要取舍：
    # 宁可列窄也不能让表散成两页）
    total2 = sum(shrunk)
    if total2 > max_units:
        scale2 = max_units / float(total2)
        shrunk = [max(int(w * scale2), 4) for w in shrunk]
    return shrunk


def build_workbook(headers, rows, out_path, sheet_name, title=None, note=None,
                   force_no_freeze=False, spec=None):
    ncol = len(headers)
    if ncol == 0:
        raise SystemExit("ERROR: 表头为空，无法生成")
    if spec is None:
        raise SystemExit("ERROR: 未提供 spec（内部错误，build_workbook 必须带规格）")

    f = spec["font"]
    h = spec["header"]
    t = spec["title"]
    nt = spec["note"]
    b = spec["body"]
    ru = spec["rules"]
    rh = spec["row_height"]
    cw = spec["col_width"]
    fz = spec["freeze"]
    pr = spec["print"]

    font_name = f["name"]
    data_size = _num(f["data_size"])
    head_size = _num(h["size"])
    title_size = _num(t["size"])

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # ---- 填充与字体（全部来自 spec；null 表示不填充）----
    head_fill = PatternFill("solid", fgColor=h["fill"]) if h.get("fill") else None
    head_font = Font(name=h.get("font_name") or font_name, size=head_size,
                     bold=bool(h.get("bold")), color=h.get("font_color"))
    title_fill = PatternFill("solid", fgColor=t["fill"]) if t.get("fill") else None
    title_font = Font(name=font_name, size=title_size,
                      bold=bool(t.get("bold")), color=t.get("font_color"))
    note_font = Font(name=font_name, size=_num(nt["size"]), color=nt.get("color"))
    data_font = Font(name=font_name, size=data_size, color=f.get("data_color"))
    alt_fill = PatternFill("solid", fgColor=b["zebra_color"]) if b.get("zebra_color") else None

    # ---- 边框：三线表走 rules，普通表走 body/header 的 border_* ----
    three_line = bool(ru.get("three_line"))
    if three_line:
        # 三线表只有横线：顶线、栏目线、底线，无竖线
        top_side = _side(f.get("data_color"), "medium")
        head_bottom = _side(f.get("data_color"), "thin")
        bottom_side = _side(f.get("data_color"), "medium")
        head_border = Border(top=top_side, bottom=head_bottom)
        data_border = Border()   # 数据行不画线，末行另补底线
    else:
        hs = h.get("border_style")
        hc = h.get("border_color")
        bs = b.get("border_style")
        bc = b.get("border_color")
        head_side = _side(hc, hs)
        head_border = Border(left=head_side, right=head_side, top=head_side, bottom=head_side)
        body_side = _side(bc, bs)
        data_border = Border(left=body_side, right=body_side,
                             top=body_side, bottom=body_side)

    r = 1
    title_row = note_row = None
    if title:
        if t.get("merge", True):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        c = ws.cell(r, 1, str(title))
        if title_fill:
            c.fill = title_fill
        c.font = title_font
        c.alignment = Alignment(horizontal=t.get("align") or "center",
                                vertical="center", wrap_text=True)
        if t.get("border"):
            c.border = head_border
        title_row = r
        r += 1
    if note:
        if nt.get("merge", True):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        c = ws.cell(r, 1, str(note))
        c.font = note_font
        c.alignment = Alignment(horizontal=nt.get("align") or "left",
                                vertical="center", indent=1, wrap_text=True)
        note_row = r
        r += 1
    head_row = r

    for j, hd in enumerate(headers, 1):
        c = ws.cell(head_row, j, str(hd))
        if head_fill:
            c.fill = head_fill
        c.font = head_font
        c.alignment = Alignment(horizontal=h.get("align") or "center",
                                vertical=h.get("valign") or "center",
                                wrap_text=bool(h.get("wrap", True)))
        if head_border is not None:
            c.border = head_border
    r = head_row + 1

    # 列宽：按表头与数据的最大显示宽度显式设置
    widths = []
    for j in range(1, ncol + 1):
        wmax = display_width(headers[j - 1]) * cw["head_factor"]
        for row in rows:
            if j - 1 < len(row) and row[j - 1] is not None:
                wmax = max(wmax, min(display_width(row[j - 1]), cw["text_cap"]))
        widths.append(int(min(max(wmax + cw["pad"], cw["min"]), cw["max"])))

    # 页宽约束：列宽合计超过纸张可用宽度时必须等比压缩。
    # 不压的直接后果是打印时横向拆页，标题合并行跟着跨页，
    # 第二页顶部留下一条空的标题色带（2026-11-21 真实业务表视觉验收实测）。
    widths = _fit_widths_to_page(widths, pr, cw)

    body_align = (b.get("align") or "auto").lower()
    zebra_start = int(b.get("zebra_start") or 2)
    nrows = len(rows)

    for i, row in enumerate(rows):
        fill = None
        # zebra_start 为 1 基的数据行序号；i 为 0 基
        if b.get("zebra") and alt_fill and (i + 1) >= zebra_start and \
                (i + 1 - zebra_start) % 2 == 0:
            fill = alt_fill
        est = 1
        is_total = bool(b.get("total_row_bold")) and i == nrows - 1
        for j in range(1, ncol + 1):
            v = row[j - 1] if j - 1 < len(row) else None
            c = ws.cell(r, j, v)
            c.font = data_font
            if fill:
                c.fill = fill
            if three_line:
                # 三线表：仅最后一行补底线
                c.border = Border(bottom=bottom_side) if i == nrows - 1 else data_border
            else:
                c.border = data_border
            if is_total:
                c.font = Font(name=font_name, size=data_size, bold=True,
                              color=f.get("data_color"))
            al = body_align
            if al == "auto":
                al = "right" if isinstance(v, (int, float)) else "left"
            c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=True)
            est = max(est, est_lines(v, widths[j - 1]))
        ws.row_dimensions[r].height = max(rh["min_data"], est * rh["line"] + 4)
        r += 1

    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    total_w = sum(widths)
    if title_row:
        # 字宽随字号线性放大（12 号 = 1.2 倍宽），否则大字号标题折行后行高不足
        scale = 1.2 * title_size / 12.0
        est_t = est_lines(title, max(int(total_w / scale), 8))
        ws.row_dimensions[title_row].height = max(rh["title_base"],
                                                  est_t * rh["title_line"] + 4)
    if note_row:
        est_n = est_lines(note, total_w)
        ws.row_dimensions[note_row].height = max(rh["note_base"], est_n * rh["line"] + 4)
    ws.row_dimensions[head_row].height = rh["head_base"]

    # ---- 冻结：按 spec.freeze.enabled / min_rows / anchor ----
    freeze = None
    mode = str(fz.get("enabled") or "auto").lower()
    if not force_no_freeze and mode != "never":
        if mode == "always" or nrows >= int(fz.get("min_rows") or 20):
            anchor = fz.get("anchor") or "auto"
            freeze = "A%d" % (head_row + 1) if str(anchor).lower() == "auto" else str(anchor)
    if freeze:
        ws.freeze_panes = freeze

    # ---- 打印设置 ----
    if str(pr.get("title_rows") or "auto").lower() == "auto":
        ws.print_title_rows = "%d:%d" % (head_row, head_row)
    elif pr.get("title_rows"):
        ws.print_title_rows = str(pr["title_rows"])
    if pr.get("orientation"):
        ws.page_setup.orientation = pr["orientation"]
    if pr.get("paper"):
        ws.page_setup.paperSize = ws.PAPERSIZE_A4 if str(pr["paper"]).upper() == "A4" else None
    if pr.get("gridlines") is not None:
        ws.print_options.gridLines = bool(pr["gridlines"])
    m = pr.get("margin_cm")
    if m:
        ws.page_margins.left = ws.page_margins.right = m / 2.54
        ws.page_margins.top = ws.page_margins.bottom = m / 2.54

    wb.save(out_path)

    # 落盘后回读一次冻结值核对（SKILL.md 明确要求）
    chk = load_workbook(out_path)
    got = chk[chk.sheetnames[0]].freeze_panes
    ok = (got == freeze)
    return {
        "out": out_path, "sheet": sheet_name, "rows": nrows, "cols": ncol,
        "head_row": head_row, "freeze": freeze, "freeze_readback": got, "freeze_ok": ok,
    }


def run(src=None, out=None, sheet=None, title=None, note=None, src_sheet=None,
        header_row=1, no_freeze=False, profile=None):
    spec = load_spec(profile=profile)
    prof_name = resolve_profile(profile)
    ext = os.path.splitext(src)[1].lower()
    if ext in (".xlsx", ".xlsm") and \
            os.path.normcase(os.path.abspath(src)) == os.path.normcase(os.path.abspath(out)):
        print("输出与输入同路径会覆盖源数据（源数据只读纪律），拒绝执行", file=sys.stderr)
        return 2
    if ext == ".md":
        headers, rows = load_md_table(src)
    elif ext == ".csv":
        headers, rows = load_csv_table(src)
    elif ext in (".xlsx", ".xlsm"):
        headers, rows = load_xlsx_table(src, src_sheet, header_row)
    else:
        raise SystemExit("ERROR: 不支持的输入扩展名 %s（支持 .md / .csv / .xlsx）" % ext)
    if not sheet:
        stem = os.path.splitext(os.path.basename(src))[0]
        sheet = re.sub(r'[\[\]:*?/\\]', "_", stem)[:31] or "数据"
    info = build_workbook(headers, rows, out, sheet, title, note, no_freeze, spec)
    meta = profile_meta(profile)
    print("profile=%s（%s）" % (prof_name, meta.get("title", "")))
    print("输出: %(out)s" % info)
    print("sheet=%(sheet)s 行数=%(rows)d 列数=%(cols)d 表头行=%(head_row)d" % info)
    print("freeze 期望=%s 回读=%s %s" %
          (info["freeze"], info["freeze_readback"], "OK" if info["freeze_ok"] else "**不符**"))
    return 0 if info["freeze_ok"] else 1


# ---------- 自检 ----------

SELFTEST_MD = """# 样例
| 名称 | 数量 | 备注 |
| --- | --- | --- |
| 甲 | 3 | 短 |
| 乙 | 5 | 一段足够长的备注文字，用于验证换行与行高估算是否正常 |
| 丙 | 7 | |
"""


def selftest():
    """自检：默认跑 business-table（与 v1.x 样式等价），覆盖小表/大表/CSV/长标题/转义。"""
    tmp = tempfile.mkdtemp(prefix="table_format_selftest_")
    fails = []
    prof = "business-table"
    spec = load_spec(profile=prof)
    h = spec["header"]
    t = spec["title"]
    b = spec["body"]
    f = spec["font"]
    rh = spec["row_height"]

    def check(name, cond, detail=""):
        print("  [%s] %s %s" % ("OK" if cond else "FAIL", name, detail))
        if not cond:
            fails.append(name)

    # 样例 1：小表（6 数据行）→ 不冻结
    md1 = os.path.join(tmp, "small.md")
    with open(md1, "w", encoding="utf-8") as fh:
        fh.write("| 名称 | 数量 |\n| --- | --- |\n" +
                 "".join("| 项%d | %d |\n" % (i, i) for i in range(1, 7)))
    out1 = os.path.join(tmp, "small.xlsx")
    rc = run(md1, out1, sheet="小表", profile=prof)
    wb = load_workbook(out1)
    ws = wb.active
    check("小表不冻结（数据行<20）", ws.freeze_panes is None,
          "freeze=%s" % ws.freeze_panes)
    check("小表表头样式",
          ws.cell(1, 1).fill.fgColor.rgb.endswith(h["fill"])
          and ws.cell(1, 1).font.color.rgb.endswith(h["font_color"])
          and ws.cell(1, 1).font.bold)
    check("首行数据行留白", ws.cell(2, 1).fill.fgColor.rgb in ("00000000", None) or
          not str(ws.cell(2, 1).fill.fgColor.rgb).endswith(b["zebra_color"]))
    check("第二数据行有隔行底纹",
          ws.cell(3, 1).fill.fgColor.rgb.endswith(b["zebra_color"]))
    check("数据字体规格", ws.cell(2, 1).font.name == f["name"]
          and ws.cell(2, 1).font.size == _num(f["data_size"]))
    check("表头字体规格", ws.cell(1, 1).font.name == (h.get("font_name") or f["name"])
          and ws.cell(1, 1).font.size == _num(h["size"]))

    # 样例 2：大表（25 数据行，带标题）→ 冻结 A4（标题1+说明1+表头1）
    md2 = os.path.join(tmp, "big.md")
    with open(md2, "w", encoding="utf-8") as fh:
        fh.write("| 编码 | 价格 |\n| --- | --- |\n" +
                 "".join("| QT-%03d | %d |\n" % (i, i * 100) for i in range(1, 26)))
    out2 = os.path.join(tmp, "big.xlsx")
    rc2 = run(md2, out2, sheet="大表", title="大表样例", note="自检样例", profile=prof)
    wb2 = load_workbook(out2)
    ws2 = wb2.active
    check("大表冻结 A4（锚点=表头末行下一行）",
          ws2.freeze_panes == "A4", "freeze=%s" % ws2.freeze_panes)
    check("print_title_rows=表头行3", ws2.print_title_rows in ("3:3", "$3:$3"),
          "=%s" % ws2.print_title_rows)
    check("标题行样式",
          ws2.cell(1, 1).fill.fgColor.rgb.endswith(t["fill"])
          and ws2.cell(1, 1).font.color.rgb.endswith(t["font_color"])
          and ws2.cell(1, 1).font.size == _num(t["size"]))
    check("前导零编码保持文本", ws2.cell(4, 1).value == "QT-001",
          "=%r" % ws2.cell(4, 1).value)
    check("纯数字转数值", isinstance(ws2.cell(4, 2).value, int),
          "=%r" % ws2.cell(4, 2).value)
    check("freeze 回读逻辑", rc == 0 and rc2 == 0)
    check("全表自动换行（标题/说明/表头/数据）",
          all(ws2.cell(rr, 1).alignment.wrap_text is True for rr in (1, 2, 3, 4)))

    # 样例 3：CSV 输入
    csv1 = os.path.join(tmp, "t.csv")
    with open(csv1, "w", encoding="utf-8", newline="") as fh:
        fh.write("型号,数量\nA-001,2\nB-002,3\n")
    out3 = os.path.join(tmp, "t.xlsx")
    run(csv1, out3, profile=prof)
    wb3 = load_workbook(out3)
    check("CSV 解析", wb3.active.cell(2, 1).value == "A-001"
          and wb3.active.cell(2, 1).value != "A-001\r")

    # 样例 4：长标题 / 长说明 → 全表换行 + 行高自适应
    md4 = os.path.join(tmp, "longtitle.md")
    with open(md4, "w", encoding="utf-8") as fh:
        fh.write("| 名称 | 数量 |\n| --- | --- |\n| 甲 | 1 |\n| 乙 | 2 |\n")
    out4 = os.path.join(tmp, "longtitle.xlsx")
    long_title = "这是一条用于验证全表自动换行与行高自适应的超长标题" * 3
    long_note = ("这是一段很长的说明文字，用来验证说明行在超过整表总宽后自动折行"
                 "并相应增高，而不是被固定行高截断显示。") * 2
    run(md4, out4, sheet="长标题", title=long_title, note=long_note, profile=prof)
    wb4 = load_workbook(out4)
    ws4 = wb4.active
    check("长标题行高自适应（>基线）",
          (ws4.row_dimensions[1].height or 0) > rh["title_base"],
          "h=%s" % ws4.row_dimensions[1].height)
    check("长说明行高自适应（>基线）",
          (ws4.row_dimensions[2].height or 0) > rh["note_base"],
          "h=%s" % ws4.row_dimensions[2].height)

    # 样例 5：窄表 + 中等长度标题 → 行高必须按折行数给足
    md_narrow = os.path.join(tmp, "narrowtitle.md")
    with open(md_narrow, "w", encoding="utf-8") as fh:
        fh.write("| 序号 | 项目名称 | 负责人 | 状态 |\n| --- | --- | --- | --- |\n| 1 | 甲 | 张三 | 进行中 |\n")
    out5 = os.path.join(tmp, "narrowtitle.xlsx")
    run(md_narrow, out5, sheet="窄标题",
        title="规范同步验证与字号对调复核及全表样式回归检查", profile=prof)
    ws5 = load_workbook(out5).active
    t5_w = sum(ws5.column_dimensions[c].width for c in "ABCD")
    t5_title = "规范同步验证与字号对调复核及全表样式回归检查"
    scale = 1.2 * _num(t["size"]) / 12.0
    t5_lines = est_lines(t5_title, max(int(t5_w / scale), 8))
    check("标题折行行高给足（按字号折算字宽）",
          t5_lines >= 2 and (ws5.row_dimensions[1].height or 0) >= t5_lines * rh["title_line"] + 4,
          "lines=%d h=%s" % (t5_lines, ws5.row_dimensions[1].height))

    # 样例 6：md 转义竖线不切列
    md5 = os.path.join(tmp, "esc.md")
    with open(md5, "w", encoding="utf-8") as fh:
        fh.write("| 名称 | 值 |\n| --- | --- |\n| a\\|b | 1 |\n")
    h5, r5 = load_md_table(md5)
    check("md 转义竖线不切列", h5 == ["名称", "值"] and r5 == [["a|b", 1]],
          "headers=%r rows=%r" % (h5, r5))

    # 样例 7：xlsx 输出与输入同路径 → 拒绝（源数据只读纪律）
    same = os.path.join(tmp, "same.xlsx")
    rc_ok = run(md1, same, sheet="同路径", profile=prof)
    rc_block = run(same, same, profile=prof)
    check("xlsx 同路径防覆盖（exit=2）", rc_ok == 0 and rc_block == 2,
          "rc_ok=%s rc_block=%s" % (rc_ok, rc_block))

    # 样例 8：三线表 profile 不画竖线、无底色、无隔行底纹
    out6 = os.path.join(tmp, "academic.xlsx")
    spec6 = load_spec(profile="academic-table")
    h6, r6 = load_md_table(md2)
    info6 = build_workbook(h6, r6, out6, "三线表", "学术表样例", None, False, spec6)
    ws6 = load_workbook(out6).active
    hr6 = info6["head_row"]
    check("三线表表头无填充",
          ws6.cell(hr6, 1).fill.fgColor.rgb in (None, "00000000"),
          "fill=%r" % (ws6.cell(hr6, 1).fill.fgColor.rgb,))
    check("三线表数据行无竖线",
          getattr(ws6.cell(hr6 + 1, 1).border.left, "style", None) is None,
          "left=%r" % (getattr(ws6.cell(hr6 + 1, 1).border.left, "style", None),))
    check("三线表顶线存在",
          getattr(ws6.cell(hr6, 1).border.top, "style", None) is not None,
          "top=%r" % (getattr(ws6.cell(hr6, 1).border.top, "style", None),))
    check("三线表末行有底线",
          getattr(ws6.cell(ws6.max_row, 1).border.bottom, "style", None) is not None,
          "bottom=%r" % (getattr(ws6.cell(ws6.max_row, 1).border.bottom, "style", None),))
    check("三线表不冻结（enabled=never）", ws6.freeze_panes is None,
          "freeze=%s" % ws6.freeze_panes)

    # 样例 9：未指定 profile → ProfileNotSpecified（exit 2 由 main 负责）
    try:
        load_spec(profile=None)
        check("无 profile 时报错", False, "未抛异常")
    except ProfileNotSpecified:
        check("无 profile 时报错（ProfileNotSpecified）", True)
    except Exception as exc:
        check("无 profile 时报错", False, "抛了 %r" % exc)

    # 样例 10：未知 profile → ProfileNotFound
    try:
        load_spec(profile="not-exist-type")
        check("未知 profile 时报错", False, "未抛异常")
    except ProfileNotFound:
        check("未知 profile 时报错（ProfileNotFound）", True)
    except Exception as exc:
        check("未知 profile 时报错", False, "抛了 %r" % exc)

    print("SELFTEST:", "PASS" if not fails else "FAIL %s" % fails)
    print("样例目录:", tmp)
    return 0 if not fails else 1


def main():
    ap = argparse.ArgumentParser(description="table-format 规格化 xlsx 生成通道（必须指定表格类型）")
    ap.add_argument("--profile", help="表格类型 profile（如 business-table / finance-table）")
    ap.add_argument("--list-profiles", action="store_true", help="列出全部表格类型后退出")
    ap.add_argument("--src", help="输入文件（.md / .csv / .xlsx）")
    ap.add_argument("--out", help="输出 xlsx 路径")
    ap.add_argument("--sheet", help="输出工作表名（默认取输入文件名）")
    ap.add_argument("--title", help="可选标题行文字")
    ap.add_argument("--note", help="可选说明行文字")
    ap.add_argument("--src-sheet", help="xlsx 输入时指定源工作表")
    ap.add_argument("--header-row", type=int, default=1, help="xlsx 输入时表头行号（默认 1）")
    ap.add_argument("--no-freeze", action="store_true", help="强制不冻结")
    ap.add_argument("--selftest", action="store_true", help="运行自检")
    a = ap.parse_args()

    if a.list_profiles:
        for name in list_profiles():
            print("%-16s %s" % (name, profile_meta(name).get("title", "")))
        raise SystemExit(0)
    if a.selftest:
        sys.exit(selftest())
    if not a.src or not a.out:
        ap.error("--src 与 --out 必填（或使用 --selftest / --list-profiles）")

    try:
        rc = run(a.src, a.out, a.sheet, a.title, a.note, a.src_sheet,
                 a.header_row, a.no_freeze, a.profile)
    except ProfileNotSpecified as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        raise SystemExit(2)
    except ProfileNotFound as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        raise SystemExit(2)
    sys.exit(rc)


if __name__ == "__main__":
    main()
