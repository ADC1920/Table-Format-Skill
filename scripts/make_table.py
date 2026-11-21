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
import io
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
        # Excel 数值只保证 15 位有效数字，更长的整数写入后会被按 float 存
        # 而丢精度（12345678901234567890 → 1.234567890123457e+19）。
        # 这类值多为订单号/身份证/条码，按文本保留比「看起来是数字但对不上」更安全。
        digits = t.lstrip("+-")
        if len(digits) > 15:
            return s
        return int(t)
    if re.fullmatch(r"[+-]?\d+\.\d+", t):
        return float(t)
    return s


def _read_text(path):
    """读文本文件，UTF-8 优先，失败依次回退 GBK / GB18030 / latin-1。

    中文 Windows 上 Excel 另存的 CSV 常是 GBK，直接 open(encoding="utf-8")
    会抛 UnicodeDecodeError 并打出一整屏堆栈（2026-11-21 第二轮测试实测）。
    GB18030 是 GBK 的超集，放它前面会掩盖真正的 GBK 问题，故按 UTF-8 →
    GBK → GB18030 → latin-1 顺序尝试，最后一个保证不失败。
    """
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xef\xbb\xbf"):   # UTF-8 BOM
        return raw[3:].decode("utf-8", errors="replace")
    for enc in ("utf-8", "gbk", "gb18030", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


# ---------- 输入解析 ----------

def load_md_table(path):
    """取 Markdown 中第一个 GFM 表格块。返回 (headers, rows)。"""
    lines = _read_text(path).splitlines()
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
    # 用 _read_text 统一处理编码回退（UTF-8 / GBK / GB18030 / latin-1），
    # 不能直接 open(encoding="utf-8-sig")——GBK 存的中文 CSV 会抛堆栈。
    text = _read_text(path)
    if not text.strip():
        raise SystemExit("ERROR: %s 为空" % path)
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
    except csv.Error:
        dialect = csv.excel
    reader = csv.reader(io.StringIO(text, newline=""), dialect)
    table = [list(r) for r in reader if any(str(c).strip() for c in r)]
    if not table:
        raise SystemExit("ERROR: %s 为空" % path)
    headers = [str(c).strip() for c in table[0]]
    rows = [[maybe_num(c) for c in r] for r in table[1:]]
    return headers, rows


def load_xlsx_table(path, src_sheet=None, header_row=1):
    """读源 xlsx。注意：不用 read_only=True（见文件头「已知坑」）。

    `--header-row` 是**物理行号**（用户在 Excel 里看到的行号），与跳不跳空行无关。
    早先实现先丢弃空白行再按序号取第 N 条，导致表头上方有空白行时取错行
    （2026-11-21 全面测试实测：加一行空行后 --header-row 4 取到了首条数据行，
    表头整行丢失且被当成数据）。改为保留 (行号, 值) 配对后按物理行号定位。
    """
    wb = load_workbook(path, data_only=True)
    ws = wb[src_sheet] if src_sheet else wb[wb.sheetnames[0]]
    numbered = []  # [(物理行号, 值列表)]
    for idx, r in enumerate(ws.iter_rows(values_only=True), 1):
        if r is None or all(v is None or str(v).strip() == "" for v in r):
            continue
        numbered.append((idx, list(r)))
    if not numbered:
        raise SystemExit("ERROR: 源表为空（无任何非空行）：%s" % path)
    if numbered[-1][0] < header_row or header_row < 1:
        raise SystemExit("ERROR: 源表只有 %d 行，不足 --header-row=%d" % (
            numbered[-1][0], header_row))
    head = None
    for pos, (rowno, vals) in enumerate(numbered):
        if rowno == header_row:
            head = (pos, vals)
            break
    if head is None:
        raise SystemExit(
            "ERROR: --header-row=%d 指向的是一空行（该行没有内容），无法作表头" % header_row)
    pos, head_vals = head
    headers = [("" if v is None else str(v).strip()) for v in head_vals]
    while headers and headers[-1] == "":
        headers.pop()  # 去掉尾部空列
    ncol = len(headers)
    rows = []
    for rowno, r in numbered[pos + 1:]:
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


def _is_numeric(v):
    """是否数值单元格。bool 是 int 的子类，要单独排除（True/False 不该套数字格式）。"""
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _number_format(num):
    """按 spec.number 生成 Excel 数字格式串；不需要时返回 None。

    `format_auto: true` 的语义是「财务/统计场景的规范数字格式」，具体为：
      - 千分位分组
      - 固定两位小数（财务报表按分计）
      - 负数用负号：中国官方报表口径（规范库 t2-046），
        **不用英美惯例的圆括号**——这一条推翻了本技能早期 SKILL.md 里
        「负数括号」的写法，那份文档本身写错了
      - 零值显示「-」：由 zero_as_dash 单独控制（它是把 0 换成文字「-」，
        不走格式串，因为格式串的第三段只影响显示、单元格值仍是 0，
        下游做汇总时会与「无此项」混淆）

    负数段显式重复正数段的格式，是为了避免 Excel 把负数渲染成
    `-1234.5`（无千分位）。
    """
    if not num.get("format_auto"):
        return None
    # 三段式：正数;负数;零。负数段显式带千分位，否则 Excel 渲染成 -1234.5。
    # 零值段用 "-"，与 zero_as_dash 的语义一致（两者都表达「这一格是 0」）。
    return '#,##0.00;-#,##0.00;-'


def _display_width_of(v, num_fmt):
    """单元格显示宽度：数值且会套千分位格式时，按格式化后的文本估算。

    只处理本项目会生成的格式 `#,##0.00;-#,##0.00;-`：
    千分位每组加一个逗号，固定两位小数，负数多一个负号。
    其他格式或文本原样返回原始显示宽度。
    """
    if not num_fmt or not _is_numeric(v):
        return display_width(v)
    if v == 0:
        return 1                      # 零值格式段渲染成单个「-」
    s = "%.2f" % abs(float(v))        # 固定两位小数
    intpart, _, frac = s.partition(".")
    groups = (len(intpart) - 1) // 3
    width = len(intpart) + groups + 1 + len(frac)   # 整数位 + 逗号 + 小数点 + 小数位
    if float(v) < 0:
        width += 1                    # 负号
    return width


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


def _col_min_widths(headers, rows, widths, cw):
    """每列的内容下限：装得下表头、或装得下多数数据值。

    等比压缩只看总宽比例，会把「天丝锦麻」这种 4 字值所在的窄列压到 6 字符，
    迫使文字折成两行、末行只剩一个字（2026-11-21 视觉验收实测）。
    这里给每列算一个下限，压缩时优先压长列、不动短列。

    取法：表头宽（中文按 2 计）与「80 分位数据宽」取较大者，再加内边距。
    用分位数而非最大值，避免个别超长单元格把整列下限抬高。
    """
    import math as _m
    mins = []
    for j in range(len(headers)):
        head_w = display_width(headers[j]) if j < len(headers) else 0
        vals = []
        for r in rows:
            if j < len(r) and r[j] is not None:
                vals.append(display_width(r[j]))
        if vals:
            vals.sort()
            q80 = vals[min(len(vals) - 1, int(_m.ceil(len(vals) * 0.8)) - 1)]
        else:
            q80 = 0
        need = max(head_w, q80) + cw["pad"]
        mins.append(int(min(max(need, cw["min"]), widths[j] if j < len(widths) else need)))
    return mins


def _fit_widths_to_page(widths, pr, cw, headers=None, rows=None):
    """列宽合计超过纸张可用宽度时压缩，避免打印横向拆页。

    压缩策略（从「等比」改为「保住下限后按富余等比」）：
      1. 先用 _col_min_widths 算出每列的内容下限；
      2. 下限之和已超页宽 → 说明列太多，只能等比硬压（无解，接受折行）；
      3. 否则把「页宽 − 下限之和」的额度按各列富余量等比分配。

    这样短列（如「类型」装 4 个汉字）不会被长列拖累压到折行，
    而长列（如「面料特点」）本来就要折行，压它代价最小。
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

    def hard_shrink(ws_list, cap, floor_list=None):
        """兜底：等比压缩到 cap 以内。

        floor_list 给定时按「下限」等比——保留下限的相对比例，
        使短列不会被长列拖累压到只剩几个字符（实测「类型」列内容 4 个汉字，
        等比硬压后只剩 5 字符，文字折成两行末行一孤字）。
        """
        base = floor_list if floor_list else ws_list
        t = sum(base)
        if t <= 0:
            return ws_list
        s = cap / float(t)
        return [max(int(w * s), 4) for w in base]

    if headers is None or rows is None:
        return hard_shrink(widths, max_units)

    mins = _col_min_widths(headers, rows, widths, cw)
    if sum(mins) >= max_units:
        # 列太多，连内容下限都装不下——按「下限的比例」等比压，
        # 这样宽内容列缩得多、短列缩得少，视觉上更接近均衡。
        # 同时给出明确提示：这种表在一页里本来就排不开，压列宽只是止损，
        # 真正的解法是换横版或缩字号（脚本不自动降字号，会连锁改行高与页数）。
        forced = hard_shrink(widths, max_units, mins)
        head_short = []
        for j, hd in enumerate(headers):
            need = display_width(hd) + cw["pad"]
            if forced[j] < need:
                head_short.append("%s(需%d实%d)" % (hd, need, forced[j]))
        print("[提示] 内容总需求超出纸张可用宽度（下限合计 %.0f > 可用 %.0f），"
              "已按比例压缩列宽；该表建议改用横版或减小字号。"
              % (sum(mins), max_units), file=sys.stderr)
        if head_short:
            print("[提示] 下列表头在压缩后已有列装不下，必然折行：%s"
                  % "、".join(head_short), file=sys.stderr)
        return forced

    # 富余额度：先全给下限，再把剩余按各列「上限-下限」比例分配
    budget = max_units - sum(mins)
    slack = [max(w - m, 0) for w, m in zip(widths, mins)]
    total_slack = sum(slack)
    if total_slack <= 0:
        return mins
    out = []
    for m, s in zip(mins, slack):
        got = int(m + budget * (s / float(total_slack)))
        # 分配额可能因取整低于下限（budget 很小时尤其明显），这里兜住
        out.append(max(got, m))
    # 兜底：取整加总可能略超页宽，按超出量从最宽列扣（不碰下限）
    while sum(out) > max_units:
        idx = max(range(len(out)), key=lambda k: out[k] - mins[k])
        if out[idx] <= mins[idx]:
            break
        out[idx] -= 1
    return out


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
    # 数值列要按「格式化之后」的宽度算，不是原始值宽度：加上千分位与两位小数后
    # 1234567.89 会显示成 1,234,567.89（12 字符 vs 原始 10 字符），按原始宽度
    # 算列宽会让数字贴边框（2026-11-21 财务表渲染实测）。
    num_fmt_preview = _number_format(spec.get("number") or {})
    widths = []
    for j in range(1, ncol + 1):
        wmax = display_width(headers[j - 1]) * cw["head_factor"]
        for row in rows:
            if j - 1 < len(row) and row[j - 1] is not None:
                v = row[j - 1]
                disp = _display_width_of(v, num_fmt_preview)
                wmax = max(wmax, min(disp, cw["text_cap"]))
        widths.append(int(min(max(wmax + cw["pad"], cw["min"]), cw["max"])))

    # 页宽约束：列宽合计超过纸张可用宽度时必须等比压缩。
    # 不压的直接后果是打印时横向拆页，标题合并行跟着跨页，
    # 第二页顶部留下一条空的标题色带（2026-11-21 真实业务表视觉验收实测）。
    widths = _fit_widths_to_page(widths, pr, cw, headers, rows)

    body_align = (b.get("align") or "auto").lower()
    zebra_start = int(b.get("zebra_start") or 2)
    nrows = len(rows)

    # 样式对象预建并复用：openpyxl 给单元格赋 font/border/alignment 时会做
    # 样式去重比较（Serialisable.__eq__/__hash__），每格新建实例会让 10 万行
    # 的表从「线性 60 秒」退化到「215 秒」。数据只有左中右三种对齐，样式各一种，
    # 预建后整表复用同一实例即可。
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fixed_align = {"left": align_left, "right": align_right, "center": align_center}
    total_font = Font(name=font_name, size=data_size, bold=True,
                      color=f.get("data_color"))
    last_border = Border(bottom=bottom_side) if three_line else data_border

    # 数字格式（spec.number）：此前整套 number 规格只写在 yaml 里、代码从不读，
    # 导致 finance-table / stats-table 宣称的「千分位、零值短横」全是空头承诺
    # （2026-11-21 实测 number_format 恒为 General）。这里按 spec 真正落到
    # 单元格的 number_format 上。
    # 依据规范库：t2-048 千分位属实务惯例、全表统一；t2-047 零值短横；
    # t2-046 **负数用负号不用圆括号**（中国官方报表口径，英美惯例才用括号）。
    num = spec.get("number") or {}
    num_fmt = _number_format(num)
    zero_as_dash = bool(num.get("zero_as_dash"))

    for i, row in enumerate(rows):
        fill = None
        # zebra_start 为 1 基的数据行序号；i 为 0 基
        if b.get("zebra") and alt_fill and (i + 1) >= zebra_start and \
                (i + 1 - zebra_start) % 2 == 0:
            fill = alt_fill
        est = 1
        is_total = bool(b.get("total_row_bold")) and i == nrows - 1
        row_border = last_border if (three_line and i == nrows - 1) else data_border
        row_font = total_font if is_total else data_font
        for j in range(1, ncol + 1):
            v = row[j - 1] if j - 1 < len(row) else None
            was_zero = False
            # 零值短横：把数值 0 换成文字「-」。只换真正的 0，
            # 不动空单元格（空白表示「无此项」，与「实测为零」语义不同，
            # 见规范库 t1-049 的空值三态）。
            if zero_as_dash and _is_numeric(v) and v == 0:
                v = "-"
                was_zero = True
            c = ws.cell(r, j, v)
            c.font = row_font
            if fill:
                c.fill = fill
            c.border = row_border
            if body_align == "auto":
                # 换成短横的零值原是数字，仍按右对齐，否则财务表里
                # 同一列会出现「数字右对齐、零值左对齐」的错位。
                al = align_right if (_is_numeric(v) or was_zero) else align_left
            else:
                al = fixed_align.get(body_align, align_left)
            c.alignment = al
            # 数字格式只给真正的数值单元格；文本列（含前导零编码、
            # 换成短横的零值）保持 General，否则会被当成数字显示成科学计数。
            if num_fmt and _is_numeric(v):
                c.number_format = num_fmt
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

    # 落盘后回读一次冻结值核对（SKILL.md 明确要求）。
    # 直接从 xlsx 的 zip 里读第一个工作表的 <pane> 标签，而不是 load_workbook
    # 重解析整张表——普通模式回读 10 万行要多花约 30 秒，read_only 模式又
    # 没有 sheet_view（ReadOnlyWorksheet 缺该属性，属已知坑）。读 zip 内
    # 一小段 XML 是唯一既快又能拿到 topLeftCell 的方式。
    got = _read_freeze_from_zip(out_path)
    ok = (got == freeze)
    return {
        "out": out_path, "sheet": sheet_name, "rows": nrows, "cols": ncol,
        "head_row": head_row, "freeze": freeze, "freeze_readback": got, "freeze_ok": ok,
    }


def _read_freeze_from_zip(path):
    """从 xlsx 的 zip 直接读第一个工作表的冻结锚点（topLeftCell）。

    返回 "A3" 形式的字符串；无冻结返回 None。
    解析失败时退回普通 load_workbook（慢但保证正确）。
    """
    try:
        import zipfile
        with zipfile.ZipFile(path) as z:
            # 找出第一个 worksheet 部件：xl/worksheets/sheet1.xml 通常就是它
            names = [n for n in z.namelist()
                     if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
            if not names:
                return None
            # 取编号最小的那个，与 workbook 里第一个 sheet 对应
            def num(n):
                m = re.search(r"sheet(\d+)\.xml$", n)
                return int(m.group(1)) if m else 10 ** 6
            names.sort(key=num)
            xml = z.read(names[0]).decode("utf-8", errors="replace")
        pane = re.search(r"<pane[^>]*topLeftCell=\"([A-Z]+\d+)\"", xml)
        if pane:
            return pane.group(1)
        return None
    except Exception:
        chk = load_workbook(path)
        return chk[chk.sheetnames[0]].freeze_panes


def run(src=None, out=None, sheet=None, title=None, note=None, src_sheet=None,
        header_row=1, no_freeze=False, profile=None):
    spec = load_spec(profile=profile)
    prof_name = resolve_profile(profile)
    # 输入校验前置：文件不存在时 open() 会抛 FileNotFoundError 并把完整堆栈
    # 打到用户面前（2026-11-21 全面测试实测），这里先给一句人话再退出。
    if not os.path.isfile(src):
        print("ERROR: 源文件不存在：%s" % src, file=sys.stderr)
        return 1
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
        print("ERROR: 不支持的输入扩展名 %s（支持 .md / .csv / .xlsx）" % ext, file=sys.stderr)
        return 1
    if not sheet:
        stem = os.path.splitext(os.path.basename(src))[0]
        sheet = re.sub(r'[\[\]:*?/\\]', "_", stem)[:31] or "数据"
    else:
        # 显式传入的表名也要清洗：Excel 禁止 [ ] : * ? / \ 七个字符，
        # 且上限 31 字符。不清洗时 xlsxwriter/openpyxl 会抛 ValueError 或
        # 生成 Excel 打不开的文件（2026-11-21 全面测试实测抛栈）。
        cleaned = re.sub(r'[\[\]:*?/\\]', "_", str(sheet))
        if cleaned != sheet:
            print("[提示] 工作表名含 Excel 禁用字符（[ ] : * ? / \\），已替换为下划线：%s" % cleaned,
                  file=sys.stderr)
        if len(cleaned) > 31:
            print("[提示] 工作表名超 31 字符上限，已截断：%s" % cleaned[:31], file=sys.stderr)
            cleaned = cleaned[:31]
        sheet = cleaned or "数据"
    # 输出目录必须存在：openpyxl 保存不会自动建目录，不存在时抛 FileNotFoundError
    out_dir = os.path.dirname(os.path.abspath(out))
    if not os.path.isdir(out_dir):
        print("ERROR: 输出目录不存在：%s（请先创建或改用已存在的目录）" % out_dir,
              file=sys.stderr)
        return 1
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
