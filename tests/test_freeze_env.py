#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""冻结规则矩阵 + 环境变量优先级测试。

冻结矩阵：数据行数 {1,19,20,21} × spec.freeze.enabled {auto,always,never} × --no-freeze
  期望：auto 且 >=min_rows → 冻结；always → 总冻结；never → 从不；
        --no-freeze 覆盖一切。
环境变量：TABLE_FORMAT_PROFILE 与 TABLE_FORMAT_SPEC 的优先级。
"""
import os
import subprocess
import sys
import tempfile

import yaml
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _env import BASE, DATA, MT, OUT, SKILL  # noqa: E402

results = []


def check(name, cond, detail=""):
    results.append((cond, name, detail))


def run(args, env_extra=None, clear_env=True):
    env = dict(os.environ)
    if clear_env:
        env.pop("TABLE_FORMAT_PROFILE", None)
        env.pop("TABLE_FORMAT_SPEC", None)
    if env_extra:
        env.update(env_extra)
    r = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env=env, cwd=SKILL)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def make_src(nrows, header_offset=0):
    p = os.path.join(DATA, "freeze-%d-%d.csv" % (nrows, header_offset))
    with open(p, "w", encoding="utf-8", newline="") as f:
        f.write("序号,值\n")
        for i in range(1, nrows + 1):
            f.write("%d,%d\n" % (i, i * 10))
    return p


def gen(profile, src, tag, extra=None):
    out = os.path.join(OUT, "freeze-%s.xlsx" % tag)
    rc, txt = run([MT, "--profile", profile, "--src", src, "--out", out] + (extra or []))
    return rc, txt, out


# ---- 真实 profile 的行为矩阵 ----
# business-table: auto, min_rows=20
# academic-table: never
# 用真实 profile 验证规格确实生效
for nrows in (1, 19, 20, 21):
    src = make_src(nrows)
    rc, txt, out = gen("business-table", src, "bt-%d" % nrows)
    if rc != 0:
        check("business-table %d 行生成" % nrows, False, txt.strip()[:120])
        continue
    got = load_workbook(out).active.freeze_panes
    want = "A2" if nrows >= 20 else None
    check("business-table %d 行冻结=%s" % (nrows, want), (got or None) == want,
          "实得 %s" % got)

    # --no-freeze 必须压过 auto
    rc, txt, out2 = gen("business-table", src, "bt-nf-%d" % nrows, ["--no-freeze"])
    got2 = load_workbook(out2).active.freeze_panes if rc == 0 else "GEN_FAIL"
    check("business-table %d 行 --no-freeze 不冻" % nrows, got2 is None, "实得 %s" % got2)

# academic-table: enabled=never，25 行也不该冻
src25 = make_src(25)
rc, txt, out = gen("academic-table", src25, "ac-25")
got = load_workbook(out).active.freeze_panes if rc == 0 else "GEN_FAIL"
check("academic-table 25 行不冻结（never）", got is None, "实得 %s" % got)

# data-list 也是 auto，25 行应冻
rc, txt, out = gen("data-list", src25, "dl-25")
got = load_workbook(out).active.freeze_panes if rc == 0 else "GEN_FAIL"
check("data-list 25 行冻结 A2", got == "A2", "实得 %s" % got)

# ---- 自定义 spec 验证 always / 自定义 anchor / 自定义 min_rows ----
custom_dir = os.path.join(BASE, "custom")
os.makedirs(custom_dir, exist_ok=True)

base_spec = {
    "version": "test",
    "font": {"name": "黑体", "western": "Arial", "data_size": 10,
             "data_color": "262626", "note_color": "595959"},
    "header": {"fill": "1F3864", "font_color": "FFFFFF", "font_name": "黑体",
               "size": 12, "bold": True, "align": "center", "valign": "center",
               "wrap": True, "border_style": "thin", "border_color": "B4C6E7"},
    "title": {"fill": "D9E2F3", "font_color": "1F3864", "size": 16, "bold": True,
              "align": "center", "merge": True},
    "note": {"size": 10, "color": "595959", "align": "left", "merge": True},
    "body": {"align": "auto", "border_style": "thin", "border_color": "B4C6E7",
             "zebra": True, "zebra_color": "F7F9FD", "zebra_start": 2,
             "total_row_bold": False},
    "rules": {"three_line": False, "top_pt": 1.5, "header_pt": 0.5,
              "bottom_pt": 1.5, "vertical": True},
    "row_height": {"min_data": 20, "title_base": 28, "title_line": 21,
                   "note_base": 18, "head_base": 24, "line": 14.5},
    "col_width": {"min": 8, "max": 52, "pad": 3, "head_factor": 1.6, "text_cap": 60},
    "freeze": {"enabled": "auto", "min_rows": 20, "anchor": "auto"},
    "print": {"orientation": "landscape", "paper": "A4", "title_rows": "auto",
              "margin_cm": None, "gridlines": False},
    "number": {"format_auto": False, "align": "right", "zero_as_dash": False},
}

cases = [
    ("always-1row", {"enabled": "always", "min_rows": 20, "anchor": "auto"}, 1, "A2"),
    ("always-25row", {"enabled": "always", "min_rows": 20, "anchor": "auto"}, 25, "A2"),
    ("never-25row", {"enabled": "never", "min_rows": 20, "anchor": "auto"}, 25, None),
    ("min5-3row", {"enabled": "auto", "min_rows": 5, "anchor": "auto"}, 3, None),
    ("min5-5row", {"enabled": "auto", "min_rows": 5, "anchor": "auto"}, 5, "A2"),
    ("anchor-A3", {"enabled": "always", "min_rows": 20, "anchor": "A3"}, 3, "A3"),
]
for tag, fz, nrows, want in cases:
    spec = dict(base_spec)
    spec["freeze"] = fz
    spec_path = os.path.join(custom_dir, "spec-%s.yaml" % tag)
    with open(spec_path, "w", encoding="utf-8") as f:
        yaml.safe_dump(spec, f, allow_unicode=True, sort_keys=False)
    src = make_src(nrows)
    out = os.path.join(OUT, "custom-%s.xlsx" % tag)
    rc, txt = run([MT, "--profile", "business-table", "--src", src, "--out", out],
                  {"TABLE_FORMAT_SPEC": spec_path})
    if rc != 0:
        check("自定义 spec %s" % tag, False, txt.strip()[:140])
        continue
    got = load_workbook(out).active.freeze_panes
    check("自定义 spec %s 冻结=%s" % (tag, want), (got or None) == want, "实得 %s" % got)

# ---- 环境变量优先级 ----
# 1) TABLE_FORMAT_PROFILE 生效
out = os.path.join(OUT, "env-profile.xlsx")
rc, txt = run([MT, "--src", make_src(3), "--out", out],
              {"TABLE_FORMAT_PROFILE": "academic-table"})
ok = rc == 0 and "academic-table" in txt
check("TABLE_FORMAT_PROFILE 生效", ok, txt.strip()[:100])

# 2) 显式 --profile 压过环境变量
out2 = os.path.join(OUT, "env-vs-cli.xlsx")
rc, txt = run([MT, "--profile", "business-table", "--src", make_src(3), "--out", out2],
              {"TABLE_FORMAT_PROFILE": "academic-table"})
ok = rc == 0 and "business-table" in txt
check("--profile 压过 TABLE_FORMAT_PROFILE", ok, txt.strip()[:120])

# 3) TABLE_FORMAT_SPEC 压过 --profile
spec_never = os.path.join(custom_dir, "spec-never-25row.yaml")
src25b = make_src(25)
out3 = os.path.join(OUT, "env-spec-vs-cli.xlsx")
rc, txt = run([MT, "--profile", "business-table", "--src", src25b, "--out", out3],
              {"TABLE_FORMAT_SPEC": spec_never})
if rc == 0:
    got = load_workbook(out3).active.freeze_panes
    check("TABLE_FORMAT_SPEC 压过 --profile（never 生效）", got is None, "实得 %s" % got)
else:
    check("TABLE_FORMAT_SPEC 压过 --profile", False, txt.strip()[:120])

# 4) 未知环境变量 profile → 报错
rc, txt = run([MT, "--src", make_src(3), "--out", os.path.join(OUT, "env-bad.xlsx")],
              {"TABLE_FORMAT_PROFILE": "no-such-type"})
check("环境变量未知 profile → exit 2", rc == 2 and "未找到" in txt,
      "rc=%d %s" % (rc, txt.strip()[:100]))

print("=" * 84)
print("冻结规则矩阵 + 环境变量优先级")
print("=" * 84)
npass = 0
for ok, name, detail in results:
    if ok:
        npass += 1
    print("[%s] %-44s %s" % ("OK  " if ok else "FAIL", name, detail.replace("\n", " ")[:100]))
print("-" * 84)
print("FREEZE/ENV: %d/%d PASS" % (npass, len(results)))
sys.exit(0 if npass == len(results) else 1)
