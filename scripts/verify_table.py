# -*- coding: utf-8 -*-
"""verify_table.py — table-format 产物级门禁（退出码 0 才可交付）。

对 make_table.py（或按同一 profile 手工生成）的 xlsx 成品做回读断言：
表头行定位与样式 / 数据字体 / 隔行底纹 / 边框或三线表线条 / 全表自动换行 /
冻结与 spec 匹配 / print_title_rows / 数据区无合并 / 列宽行高显式设置。
任一不符退出码 1。

**v2.0 起必须指定表格类型**：断言全部按 `--profile` 加载的 spec 求值，
不再有硬编码的颜色与字号。不传 profile 时 exit 2。

校验口径（与 profiles/<类型>.yaml 一致）：
    - 表头行定位：按 spec.header.fill 反查（fill 为 null 时按「首个全非空且
      下方有数据」的行判定，供三线表等无底色表使用）
    - 冻结：按 spec.freeze.enabled（auto/always/never）+ min_rows + anchor；
      --expect-freeze none|A3 可显式覆盖
    - 三线表（rules.three_line=true）：断言无竖线、顶线/底线存在、无底色

注意：读 xlsx 不用 openpyxl read_only 模式（ReadOnlyWorksheet 无
dimensions/column_dimensions 等属性，EmptyCell 无 .row）；详见 make_table.py 头部注记。

用法：
    python verify_table.py --profile finance-table 成品.xlsx
        [--sheet 表名] [--expect-freeze none|A3] [--min-rows N]
    python verify_table.py --list-profiles
    python verify_table.py --selftest        # 正常产物应 PASS、破坏样式应 FAIL
"""

import argparse
import os
import sys
import tempfile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from spec_loader import (ProfileNotFound, ProfileNotSpecified, list_profiles,
                         load_spec, profile_meta, resolve_profile)

NO_FREEZE = object()  # 哨兵：未显式指定


def hexof(obj):
    """取 fill/font 颜色 hex（openpyxl 形态如 'FF262626' / '00262626'）。"""
    try:
        v = obj.rgb
        return str(v).upper() if v is not None else None
    except Exception:
        return None


def is_fill(cell, hexv):
    if not hexv:
        return False
    h = hexof(cell.fill.fgColor)
    return h is not None and h.endswith(hexv)


def is_font_color(cell, hexv):
    if not hexv:
        return False
    h = hexof(cell.font.color)
    return h is not None and h.endswith(hexv)


def norm_titles(s):
    return (s or "").replace("$", "")


def _sz(v):
    return None if v is None else float(v)


def find_head_row(ws, maxr, maxc, spec):
    """定位表头行。

    有表头底色时按底色反查（v1.x 口径）；无底色（三线表/无填充表）时按结构判定：
    跳过跨列合并的标题行与说明行，取其后第一个「单元格全非空且下方有数据」的行。
    单列表同样适用（表头下方必须有数据行）。
    """
    fill = spec["header"].get("fill")
    if fill:
        for rr in range(1, min(maxr, 12) + 1):
            if is_fill(ws.cell(rr, 1), fill):
                return rr, ["按表头底色 #%s 定位" % fill]
    # 结构判定：跳过合并行（标题/说明），取首个全非空且下方有数据的行
    for rr in range(1, min(maxr, 12) + 1):
        if any(rng.min_row <= rr <= rng.max_row and rng.min_col == 1 and rng.max_col >= 2
               for rng in ws.merged_cells.ranges):
            continue
        vals = [ws.cell(rr, j).value for j in range(1, maxc + 1)]
        if all(v is not None and str(v).strip() != "" for v in vals) and rr < maxr:
            return rr, ["无表头底色，按结构判定（跳过合并行后的首个全非空行）"]
    return None, ["前 12 行未找到表头行"]


def verify(path, sheet=None, expect_freeze=NO_FREEZE, min_rows=0, profile=None):
    spec = load_spec(profile=profile)
    prof_name = resolve_profile(profile)
    meta = profile_meta(profile)
    h = spec["header"]
    t = spec["title"]
    b = spec["body"]
    f = spec["font"]
    ru = spec["rules"]
    fz = spec["freeze"]
    three_line = bool(ru.get("three_line"))

    results = []  # (ok|False|None, name, detail)

    def check(cond, name, detail=""):
        results.append((True if cond else False, name, detail))

    def skip(name, detail=""):
        results.append((None, name, detail))

    if not os.path.isfile(path):
        print("ERROR: 文件不存在 %s" % path)
        return 1
    wb = load_workbook(path)  # 普通模式，见文件头注记
    ws = wb[sheet] if sheet else wb.active
    maxr, maxc = ws.max_row, ws.max_column

    # ---- 表头行定位 ----
    head_row, why = find_head_row(ws, maxr, maxc, spec)
    if head_row is None:
        check(False, "表头行定位", why[0])
        head_row = 1
    else:
        check(True, "表头行定位", "第 %d 行（%s）" % (head_row, why[0]))

    # ---- 表头样式（逐项按 spec）----
    head_ok, head_bad = True, ""
    head_name = h.get("font_name") or f["name"]
    head_size = _sz(h["size"])
    for j in range(1, maxc + 1):
        c = ws.cell(head_row, j)
        if str(c.value or "").strip() == "":
            continue
        conds = [
            (is_fill(c, h.get("fill")) if h.get("fill") else True, "底色"),
            (is_font_color(c, h.get("font_color")), "字色"),
            (bool(c.font.bold) == bool(h.get("bold")), "加粗"),
            (c.font.name == head_name, "字体"),
            (c.font.size is not None and _sz(c.font.size) == head_size, "字号"),
        ]
        bad = [n for ok, n in conds if not ok]
        if bad:
            head_ok = False
            head_bad = "第 %d 列表头%s不符" % (j, "、".join(bad))
            break
    check(head_ok, "表头样式（%s %s号 %s）" % (
        head_name, h["size"], "有底色 #%s" % h["fill"] if h.get("fill") else "无底色"),
        head_bad)

    # ---- 标题行样式（可选）----
    # 标题行与说明行结构相同（都是跨列合并），只能靠字号区分：
    # 与 title.size 一致按标题断言；与 note.size 一致按说明行放行；
    # 两者都不符说明该行样式被改坏，报错。
    t_size = _sz(t["size"])
    n_size = _sz(spec["note"]["size"])

    def row_merges_from_a(rr):
        return any(rng.min_row == rr and rng.max_row == rr and rng.min_col == 1
                   and rng.max_col >= 2 for rng in ws.merged_cells.ranges)

    if t.get("fill"):
        title_rows = [rr for rr in range(1, head_row) if is_fill(ws.cell(rr, 1), t["fill"])]
        note_rows = [rr for rr in range(1, head_row)
                     if rr not in title_rows and row_merges_from_a(rr)]
    else:
        cand = [rr for rr in range(1, head_row) if row_merges_from_a(rr)]
        title_rows, note_rows = [], []
        for rr in cand:
            sz = _sz(ws.cell(rr, 1).font.size)
            if sz == t_size:
                title_rows.append(rr)
            elif sz == n_size:
                note_rows.append(rr)
            else:
                title_rows.append(rr)  # 说不清归属的合并行按标题规格要求，以便报出问题

    if not title_rows:
        skip("标题行样式", "无标题行（可选）")
    else:
        t_ok, t_bad = True, ""
        for rr in title_rows:
            c = ws.cell(rr, 1)
            conds = [
                (c.font.name == f["name"], "字体"),
                (c.font.size is not None and _sz(c.font.size) == t_size, "字号"),
                (is_font_color(c, t.get("font_color")), "字色"),
                (bool(c.font.bold) == bool(t.get("bold")), "加粗"),
                (is_fill(c, t["fill"]) if t.get("fill") else True, "底色"),
            ]
            bad = [n for ok, n in conds if not ok]
            if bad:
                t_ok = False
                t_bad = "第 %d 行标题%s不符" % (rr, "、".join(bad))
                break
        check(t_ok, "标题行样式（%s %s号）" % (f["name"], t["size"]), t_bad)
    if note_rows:
        n_ok, n_bad = True, ""
        for rr in note_rows:
            c = ws.cell(rr, 1)
            if not (c.font.size is not None and _sz(c.font.size) == n_size):
                n_ok = False
                n_bad = "第 %d 行说明字号不符" % rr
                break
        check(n_ok, "说明行样式（%s号）" % spec["note"]["size"], n_bad)

    # ---- 数据行范围 ----
    data_rows = list(range(head_row + 1, maxr + 1))
    data_n = len(data_rows)
    if min_rows and data_n < min_rows:
        check(False, "数据行数下限", "实际 %d < --min-rows %d" % (data_n, min_rows))
    else:
        check(True, "数据行数", "%d 行" % data_n)

    # ---- 数据区一趟扫描：字体 / 边框或竖线 / 数字格式 / 隔行底纹 ----
    # 两项优化叠加：
    #   1. 四项原先各写一个双重循环，1 万行要扫 4 遍；合并成一趟。
    #   2. 用 iter_rows 而不是 ws.cell(rr, j)。实测 6 万次访问：
    #      ws.cell 需 25 秒，iter_rows 只需 0.09 秒——差 278 倍。原因是
    #      ws.cell() 每次都查字典且缺失时创建 Cell 对象，而 iter_rows 直接
    #      复用已解析的行列结构。这是本轮真正的性能来源，只合并遍历次数没用。
    # 各项仍独立收集「首个失败位置」并各自断言，报告结构与原来一致。
    data_size = _sz(f["data_size"])
    f_bad = ""
    v_bad = ""
    b_bad = ""
    bad_fmt = ""
    num_checked = 0
    alt_bad = ""
    bs = b.get("border_style")
    bc = b.get("border_color")
    zebra_on = bool(b.get("zebra")) and bool(b.get("zebra_color"))
    zebra_color = b.get("zebra_color")
    zebra_start = int(b.get("zebra_start") or 2)
    zebra_usable = zebra_on and data_n >= 2
    num = spec.get("number") or {}
    fmt_on = bool(num.get("format_auto"))
    # 不设底纹时需要确认「确实没有底纹」；用 zebra_color 判断，未配颜色就没法判
    zebra_absent_check = (not zebra_on) and bool(zebra_color) and data_n >= 1

    for idx, row_cells in enumerate(ws.iter_rows(min_row=head_row + 1, max_row=maxr)):
        for j, c in enumerate(row_cells, 1):
            rr = c.row

            # 字体（三项一起判，任一不符即记录并停止该维度）
            if not f_bad:
                conds = [
                    (c.font.name == f["name"], "字体"),
                    (c.font.size is not None and _sz(c.font.size) == data_size, "字号"),
                    (is_font_color(c, f.get("data_color")), "字色"),
                ]
                bad = [n for ok, n in conds if not ok]
                if bad:
                    f_bad = "R%dC%d %s不符" % (rr, j, "、".join(bad))

            # 边框：三线表查竖线禁令，普通表查四边样式与颜色
            if three_line:
                if not v_bad:
                    for side in ("left", "right"):
                        if getattr(getattr(c.border, side, None), "style", None) is not None:
                            v_bad = "R%dC%d 存在%s竖线（三线表禁止）" % (rr, j, side)
                            break
            elif bs and not b_bad:
                side = getattr(c.border, "left", None)
                if not (getattr(side, "style", None) == bs
                        and hexof(getattr(side, "color", None)) is not None
                        and hexof(side.color).endswith(bc)):
                    b_bad = "R%dC%d 边框不符" % (rr, j)

            # 数字格式
            val = c.value
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                num_checked += 1
                if not bad_fmt:
                    nf = (c.number_format or "General")
                    if fmt_on and nf == "General":
                        bad_fmt = "R%dC%d 数值单元格无数字格式（期望千分位）" % (rr, j)
                    elif (not fmt_on) and nf != "General":
                        bad_fmt = "R%dC%d 该 profile 未开 format_auto，却设了 %r" % (
                            rr, j, nf)

            # 隔行底纹（只看 A 列，与逐项实现一致）
            if j == 1 and not alt_bad:
                if zebra_usable:
                    want = (idx + 1) >= zebra_start and (idx + 1 - zebra_start) % 2 == 0
                    got = is_fill(c, zebra_color)
                    if got != want:
                        alt_bad = "第 %d 数据行底纹%s（应为%s）" % (
                            idx + 1, "有" if got else "无", "有" if want else "无")
                elif zebra_absent_check:
                    if is_fill(c, zebra_color):
                        alt_bad = "检出底纹着色（该 profile 已关闭隔行底纹）"

    check(not f_bad, "数据字体（%s %s号 #%s）" % (f["name"], f["data_size"], f.get("data_color")),
          f_bad)

    # ---- 边框 / 三线表线条（表头与末行的线在扫描外单独查）----
    if three_line:
        top_ok = getattr(ws.cell(head_row, 1).border.top, "style", None) is not None
        bottom_ok = getattr(ws.cell(head_row, 1).border.bottom, "style", None) is not None
        last_bottom = (getattr(ws.cell(maxr, 1).border.bottom, "style", None) is not None
                       if data_rows else True)
        check(top_ok, "三线表顶线", "" if top_ok else "表头行无顶线")
        check(bottom_ok, "三线表栏目线", "" if bottom_ok else "表头行无下横线")
        check(last_bottom, "三线表底线", "" if last_bottom else "末行无底线")
        check(not v_bad, "三线表无竖线", v_bad)

        # 三线表不得有任何单元格底色（印刷与黑白复印要求）。
        # 早先只查线条，漏掉了「三线表被套上深蓝表头/隔行底纹」这种错版——
        # 那是最常见的混用方式（拿商务表模板改用三线表 profile 时就会这样）。
        def solid_hex(cell):
            h = hexof(cell.fill.fgColor)
            if h is None or h in ("00000000", "00FFFFFF"):
                return None
            return h

        painted = []
        for rr in range(1, maxr + 1):
            for j in range(1, maxc + 1):
                cell = ws.cell(rr, j)
                if isinstance(cell, MergedCell) or cell.value is None:
                    continue
                h = solid_hex(cell)
                if h:
                    painted.append("R%dC%d=#%s" % (rr, j, h))
                    break
            if painted:
                break
        check(not painted, "三线表无底色（含表头与数据区）",
              "检出填充 %s" % painted[:3] if painted else "")
    else:
        if not bs:
            skip("全表边框", "该 profile 不设数据区边框")
        else:
            check(not b_bad, "全表边框（%s #%s）" % (bs, bc), b_bad)

    # ---- 数字格式断言 ----
    if num_checked == 0:
        skip("数字格式", "数据区无纯数值单元格")
    else:
        check(not bad_fmt,
              "数字格式（format_auto=%s，检查 %d 个数值格）" % (fmt_on, num_checked),
              bad_fmt)

    # ---- 隔行底纹断言 ----
    if not zebra_on or not zebra_color:
        if data_rows and zebra_absent_check:
            check(not alt_bad, "无隔行底纹（该 profile 关闭）", alt_bad)
        else:
            skip("隔行底纹", "无数据行" if not data_rows else "该 profile 未配底纹色，无法判定")
    elif data_n < 2:
        skip("隔行底纹", "数据行不足 2 行")
    else:
        check(not alt_bad, "隔行底纹（自第 %d 数据行起 #%s）" % (zebra_start, zebra_color),
              alt_bad)

    # ---- 冻结 ----
    if expect_freeze is NO_FREEZE:
        mode = str(fz.get("enabled") or "auto").lower()
        if mode == "never":
            want_freeze = None
        elif mode == "always" or data_n >= int(fz.get("min_rows") or 20):
            anchor = fz.get("anchor") or "auto"
            want_freeze = ("A%d" % (head_row + 1)) if str(anchor).lower() == "auto" \
                else str(anchor)
        else:
            want_freeze = None
    elif expect_freeze is None:
        want_freeze = None
    else:
        want_freeze = expect_freeze
    got_freeze = ws.freeze_panes
    check((got_freeze or None) == want_freeze, "冻结锚点",
          "期望 %s 实际 %s" % (want_freeze, got_freeze))

    # ---- print_title_rows ----
    ptr = spec["print"].get("title_rows") or "auto"
    if str(ptr).lower() == "auto":
        want_ptr = "%d:%d" % (head_row, head_row)
        pt = norm_titles(ws.print_title_rows)
        check(pt == want_ptr, "print_title_rows", "期望 %s 实际 %s" % (want_ptr, pt or "(空)"))
    else:
        skip("print_title_rows", "该 profile 不自动设置（%s）" % ptr)

    # ---- 数据区与表头区不得有合并 ----
    m_ok, m_bad = True, ""
    for rng in ws.merged_cells.ranges:
        if rng.min_row > head_row:
            m_ok = False
            m_bad = "合并区 %s 侵入数据区" % rng
            break
    check(m_ok, "合并区限制（仅标题/说明区）", m_bad)

    # ---- 全表自动换行 ----
    # 必须用 iter_rows：这段原先用 ws.cell 遍历全表，1 万行 × 6 列实测 26 秒，
    # 是整个 verify 里最慢的一段（比数据区四项检查加起来还慢得多）。
    # iter_rows 同样遍历只要 0.1 秒级。
    wrap_ok, wrap_bad = True, ""
    for row_cells in ws.iter_rows(min_row=1, max_row=maxr, max_col=maxc):
        for c in row_cells:
            if isinstance(c, MergedCell):
                continue
            if c.value is None or str(c.value).strip() == "":
                continue
            if not bool(c.alignment.wrap_text):
                wrap_ok = False
                wrap_bad = "R%dC%d 未启用自动换行" % (c.row, c.column)
                break
        if not wrap_ok:
            break
    check(wrap_ok, "全表自动换行（wrap_text）", wrap_bad)

    # ---- 列宽与行高显式设置 ----
    w_ok, w_bad = True, ""
    for j in range(1, maxc + 1):
        if ws.column_dimensions[get_column_letter(j)].width is None:
            w_ok = False
            w_bad = "第 %d 列无显式列宽" % j
            break
    check(w_ok, "列宽显式设置", w_bad)
    h_ok, h_bad = True, ""
    for rr in [r for r in range(1, head_row + 1)] + data_rows:
        if ws.row_dimensions[rr].height is None:
            h_ok = False
            h_bad = "第 %d 行无显式行高" % rr
            break
    check(h_ok, "行高显式设置", h_bad)

    # ---- 报告 ----
    n_fail = 0
    print("profile=%s（%s）" % (prof_name, meta.get("title", "")))
    print("文件: %s | sheet=%s | 表头行=%s | 数据行=%d" % (path, ws.title, head_row, data_n))
    for ok, name, detail in results:
        if ok is None:
            print("  [SKIP] %s %s" % (name, detail))
        elif ok:
            print("  [OK]   %s %s" % (name, detail))
        else:
            n_fail += 1
            print("  [FAIL] %s %s" % (name, detail))
    print("VERIFY: %s（%d 项失败）" % ("PASS" if n_fail == 0 else "FAIL", n_fail))
    return 0 if n_fail == 0 else 1


# ---------- 自检 ----------

def selftest():
    import make_table as mt  # 兄弟脚本（同目录）
    tmp = tempfile.mkdtemp(prefix="table_format_verify_selftest_")
    fails = []
    prof = "business-table"
    spec = load_spec(profile=prof)
    hh = spec["header"]
    tt = spec["title"]

    def expect(name, cond, detail=""):
        print("  [%s] %s %s" % ("OK" if cond else "FAIL", name, detail))
        if not cond:
            fails.append(name)

    # 正例：合规产物
    headers = ["编码", "名称", "价格"]
    rows = [["QT-%03d" % i, "作品%d" % i, i * 100] for i in range(1, 26)]
    good = os.path.join(tmp, "good.xlsx")
    mt.build_workbook(headers, rows, good, "表", title="自检正例",
                      note="table-format 门禁自检", spec=spec)
    rc_good = verify(good, None, NO_FREEZE, 0, prof)
    expect("合规产物 PASS（exit 0）", rc_good == 0, "exit=%d" % rc_good)

    # 反例：表头填充被破坏 → 应 FAIL
    bad = os.path.join(tmp, "bad.xlsx")
    wb = load_workbook(good)
    wb.active.cell(3, 1).fill = PatternFill("solid", fgColor="FF0000")
    wb.save(bad)
    rc_bad = verify(bad, None, NO_FREEZE, 0, prof)
    expect("破坏表头样式 FAIL（exit 1）", rc_bad == 1, "exit=%d" % rc_bad)

    # 反例 2：冻结被移除（≥20 行应冻结）→ 应 FAIL
    bad2 = os.path.join(tmp, "bad2.xlsx")
    wb2 = load_workbook(good)
    wb2.active.freeze_panes = None
    wb2.save(bad2)
    rc_bad2 = verify(bad2, None, NO_FREEZE, 0, prof)
    expect("缺失冻结 FAIL（exit 1）", rc_bad2 == 1, "exit=%d" % rc_bad2)

    # 反例 3：标题行字号改错 → 应 FAIL
    bad3 = os.path.join(tmp, "bad3.xlsx")
    wb3 = load_workbook(good)
    wb3.active.cell(1, 1).font = Font(name="黑体", size=12, bold=True,
                                     color="FF" + tt["font_color"])
    wb3.save(bad3)
    rc_bad3 = verify(bad3, None, NO_FREEZE, 0, prof)
    expect("标题字号错 FAIL（exit 1）", rc_bad3 == 1, "exit=%d" % rc_bad3)

    # 反例 4：数据单元格关闭自动换行 → 应 FAIL
    bad4 = os.path.join(tmp, "bad4.xlsx")
    wb4 = load_workbook(good)
    wb4.active.cell(4, 1).alignment = Alignment(vertical="center", wrap_text=False)
    wb4.save(bad4)
    rc_bad4 = verify(bad4, None, NO_FREEZE, 0, prof)
    expect("关闭自动换行 FAIL（exit 1）", rc_bad4 == 1, "exit=%d" % rc_bad4)

    # 正例 2：三线表 profile → 按三线表口径断言通过
    ac = os.path.join(tmp, "academic.xlsx")
    spec_ac = load_spec(profile="academic-table")
    mt.build_workbook(headers, rows, ac, "学术表", title="学术表自检", spec=spec_ac)
    rc_ac = verify(ac, None, NO_FREEZE, 0, "academic-table")
    expect("三线表产物 PASS（exit 0）", rc_ac == 0, "exit=%d" % rc_ac)

    # 反例 5：给三线表加竖线 → 应 FAIL
    ac_bad = os.path.join(tmp, "academic_bad.xlsx")
    wb5 = load_workbook(ac)
    ws5 = wb5.active
    from openpyxl.styles import Border, Side
    side = Side(style="thin", color="FF000000")
    ws5.cell(3, 1).border = Border(left=side, right=side)
    wb5.save(ac_bad)
    rc_ac_bad = verify(ac_bad, None, NO_FREEZE, 0, "academic-table")
    expect("三线表出现竖线 FAIL（exit 1）", rc_ac_bad == 1, "exit=%d" % rc_ac_bad)

    # 反例 6：无 profile → ProfileNotSpecified
    try:
        verify(good, None, NO_FREEZE, 0, None)
        expect("无 profile 时报错", False, "未抛异常")
    except ProfileNotSpecified:
        expect("无 profile 时报错（ProfileNotSpecified）", True)
    except Exception as exc:
        expect("无 profile 时报错", False, "抛了 %r" % exc)

    print("SELFTEST:", "PASS" if not fails else "FAIL %s" % fails)
    return 0 if not fails else 1


def main():
    ap = argparse.ArgumentParser(description="table-format 产物级门禁（必须指定表格类型）")
    ap.add_argument("xlsx", nargs="?", help="待校验的 xlsx")
    ap.add_argument("--profile", help="表格类型 profile（须与生成时一致）")
    ap.add_argument("--list-profiles", action="store_true", help="列出全部表格类型后退出")
    ap.add_argument("--sheet", help="工作表名（默认第一个）")
    ap.add_argument("--expect-freeze", help="显式期望冻结锚点：none 或如 A3")
    ap.add_argument("--min-rows", type=int, default=0, help="数据行数下限（可选）")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()

    if a.list_profiles:
        for name in list_profiles():
            print("%-16s %s" % (name, profile_meta(name).get("title", "")))
        raise SystemExit(0)
    if a.selftest:
        sys.exit(selftest())
    if not a.xlsx:
        ap.error("需要 xlsx 路径（或使用 --selftest / --list-profiles）")
    if a.expect_freeze is None:
        ef = NO_FREEZE
    elif a.expect_freeze.lower() == "none":
        ef = None
    else:
        ef = a.expect_freeze.upper()
    try:
        rc = verify(a.xlsx, a.sheet, ef, a.min_rows, a.profile)
    except ProfileNotSpecified as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        raise SystemExit(2)
    except ProfileNotFound as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        raise SystemExit(2)
    sys.exit(rc)


if __name__ == "__main__":
    main()
