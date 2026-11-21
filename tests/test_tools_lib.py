#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""check_delivery.py 与 render_pages.py 覆盖测试 + 规范库质量检查。

check_delivery：正常产物放行、默认表名拦截、空表拦截、公式错误值拦截、文件不存在
render_pages：正常产物出图数与页数一致、--pages 选页生效
规范库：395 条字段完整性、id 唯一、domain/profile 合法、检索六种方式可用
"""
import json
import os
import subprocess
import sys

from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _env import CD, DATA, MT, OUT, PNG, QS, RP, SKILL, VT  # noqa: E402

results = []


def check(name, cond, detail=""):
    results.append((cond, name, detail))


def run(args, cwd=SKILL):
    r = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=cwd)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


SRC = os.path.join(DATA, "delivery.csv")
with open(SRC, "w", encoding="utf-8", newline="") as f:
    f.write("型号,数量\n")
    for i in range(1, 31):
        f.write("TEL-%03d,%d\n" % (i, i * 10))

# ============ check_delivery ============
good = os.path.join(OUT, "delivery-good.xlsx")
rc, txt = run([MT, "--profile", "business-table", "--src", SRC, "--out", good])
check("正例生成", rc == 0, txt.strip()[:100])

rc, txt = run([CD, good])
check("check_delivery 正例放行", rc == 0, txt.strip()[-100:])

# 默认表名残留
bad_name = os.path.join(OUT, "delivery-badsheet.xlsx")
wb = Workbook()
ws = wb.active  # 默认 Sheet
ws.append(["A", "B"])
ws.append([1, 2])
wb.save(bad_name)
rc, txt = run([CD, bad_name])
check("默认表名被拦", rc == 1 and ("表名" in txt or "Sheet" in txt),
      "rc=%d %s" % (rc, txt.strip()[-90:]))

# 空表
empty = os.path.join(OUT, "delivery-empty.xlsx")
wb = Workbook()
ws = wb.active
ws.title = "数据"
ws.append(["A", "B"])  # 只有表头，无数据行
wb.save(empty)
rc, txt = run([CD, empty])
check("空表告警（WARN 不阻断，设计如此）", "WARN" in txt and "空表" in txt,
      "rc=%d %s" % (rc, txt.strip()[-90:]))

# 公式错误值
errf = os.path.join(OUT, "delivery-formulaerr.xlsx")
wb = Workbook()
ws = wb.active
ws.title = "数据"
ws.append(["A", "B"])
ws.append([1, "#REF!"])
ws.append([2, "#DIV/0!"])
wb.save(errf)
rc, txt = run([CD, errf])
check("公式错误值被拦", rc == 1 and "公式错误值" in txt and "B2" in txt,
      "rc=%d %s" % (rc, txt.strip()[-100:]))

# 文件不存在
rc, txt = run([CD, os.path.join(OUT, "no-such.xlsx")])
check("check_delivery 文件不存在 → exit 1", rc == 1 and "文件不存在" in txt,
      "rc=%d" % rc)

# ============ render_pages ============
rc, txt = run([RP, good, "--out-dir", os.path.join(PNG, "render1")])
check("render_pages 正例", rc == 0 and "共" in txt, txt.strip()[:120])
if rc == 0:
    produced = [f for f in os.listdir(os.path.join(PNG, "render1")) if f.endswith(".png")]
    check("render_pages 出图数 > 0", len(produced) > 0, "产出 %d 张" % len(produced))

    rc2, txt2 = run([RP, good, "--out-dir", os.path.join(PNG, "render-p1"), "--pages", "1"])
    if rc2 == 0:
        p1 = [f for f in os.listdir(os.path.join(PNG, "render-p1")) if f.endswith(".png")]
        check("render_pages --pages 1 只出一张", len(p1) == 1, "产出 %d 张" % len(p1))
    else:
        check("render_pages --pages 1", False, txt2.strip()[:120])

rc, txt = run([RP, os.path.join(OUT, "no-such.xlsx")])
check("render_pages 文件不存在 → 非 0", rc != 0, "rc=%d" % rc)

# ============ 规范库质量 ============
STD = os.path.join(SKILL, "references", "standards", "standards.jsonl")
rows = []
parse_err = 0
with open(STD, encoding="utf-8") as f:
    for i, line in enumerate(f, 1):
        line = line.strip()
        if not line:
            continue
        try:
            rows.append(json.loads(line))
        except json.JSONDecodeError:
            parse_err += 1

check("规范库全部可解析", parse_err == 0, "解析失败 %d 行" % parse_err)
check("规范库条数 395", len(rows) == 395, "实得 %d" % len(rows))

REQ = ("id", "category", "name", "specs", "confidence", "domain", "profile")
missing = [(r.get("id"), [k for k in REQ if not r.get(k)]) for r in rows]
missing = [m for m in missing if m[1]]
check("字段全部非空", not missing, "缺字段 %s" % missing[:3])

ids = [r["id"] for r in rows]
check("id 唯一", len(ids) == len(set(ids)), "重复 %d 个" % (len(ids) - len(set(ids))))

bad_conf = [r["id"] for r in rows if r["confidence"] not in ("high", "medium", "low")]
check("confidence 取值合法", not bad_conf, "非法 %s" % bad_conf[:3])

over = [(r["id"], len(str(r["specs"]))) for r in rows if len(str(r["specs"])) > 260]
check("specs 长度 ≤260", not over, "超限 %s" % over[:3])

valid_profiles = {"business-table", "academic-table", "finance-table", "stats-table",
                  "gov-table", "quality-table", "hr-table", "data-list"}
bad_prof = sorted({r["profile"] for r in rows} - valid_profiles)
check("profile 取值均为 8 类之一", not bad_prof, "越界 %s" % bad_prof)

check("无 source 字段残留", all("source" not in r for r in rows),
      "含 source %d 条" % sum(1 for r in rows if "source" in r))

# 检索六种方式
checks = [
    ("--list-domains", [QS, "--list-domains"], "领域"),
    ("--list-profiles", [QS, "--list-profiles"], "按表格类型"),
    ("--list-categories", [QS, "--list-categories"], "子类"),
    ("--domain", [QS, "--domain", "finance"], "命中"),
    ("--profile", [QS, "--profile", "academic-table"], "命中"),
    ("--keyword", [QS, "--keyword", "三线表"], "命中"),
    ("--category", [QS, "--category", "表头"], None),
    ("--min-conf", [QS, "--domain", "gov", "--min-conf", "high"], None),
    ("--full", [QS, "--keyword", "续表", "--full"], None),
    ("--json", [QS, "--keyword", "三线表", "--json"], "["),
    ("--limit", [QS, "--profile", "stats-table", "--limit", "2"], None),
]
for label, args, kw in checks:
    rc, txt = run(args)
    ok = rc == 0 and (kw is None or kw in txt)
    check("query_standards %s" % label, ok, "rc=%d %s" % (rc, txt.strip()[:70]))

# 不存在的领域 → 不崩
rc, txt = run([QS, "--domain", "no-such-domain"])
check("query 未知领域不崩", rc == 0 and "未命中" in txt, "rc=%d %s" % (rc, txt.strip()[:70]))

print("=" * 84)
print("check_delivery / render_pages / 规范库质量")
print("=" * 84)
npass = 0
for ok, name, detail in results:
    if ok:
        npass += 1
    print("[%s] %-40s %s" % ("OK  " if ok else "FAIL", name, detail.replace("\n", " ")[:95]))
print("-" * 84)
print("TOOLS/LIB: %d/%d PASS" % (npass, len(results)))
sys.exit(0 if npass == len(results) else 1)
