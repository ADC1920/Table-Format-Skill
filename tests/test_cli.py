#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""CLI 边界测试：退出码与报错文本必须符合契约。

覆盖：无 profile / 未知 profile / --list-profiles / 缺参 / 文件不存在 /
      不支持的扩展名 / 同路径覆盖 / 帮助信息。
每个用例断言 (退出码, 输出关键字)。
"""
import io
import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _env import MT, OUT, SKILL, VT  # noqa: E402

TMP = OUT
MD = os.path.join(TMP, "cli-basic.md")
with open(MD, "w", encoding="utf-8") as f:
    f.write("| 名称 | 数量 |\n| --- | --- |\n| 甲 | 1 |\n| 乙 | 2 |\n")


def run(args, env_extra=None):
    env = dict(os.environ)
    env.pop("TABLE_FORMAT_PROFILE", None)
    env.pop("TABLE_FORMAT_SPEC", None)
    if env_extra:
        env.update(env_extra)
    r = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env=env, cwd=SKILL)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


results = []


def case(name, args, want_rc, want_kw, env=None):
    rc, out = run(args, env)
    ok_rc = (rc == want_rc)
    ok_kw = all(k in out for k in want_kw) if want_kw else True
    results.append((ok_rc and ok_kw, name, rc, want_rc, out.strip()[:160]))
    return ok_rc and ok_kw


# --- 1. 无 profile 必须 exit 2 ---
case("make_table 无 profile → exit 2", [MT, "--src", MD, "--out", os.path.join(TMP, "a.xlsx")],
     2, ["未指定表格类型", "business-table"])

# --- 2. 未知 profile 必须 exit 2 ---
case("make_table 未知 profile → exit 2",
     [MT, "--profile", "not-exist-type", "--src", MD, "--out", os.path.join(TMP, "b.xlsx")],
     2, ["未找到表格类型"])

# --- 3. --list-profiles 必须 exit 0 且列 8 个 ---
rc, out = run([MT, "--list-profiles"])
names = ["business-table", "academic-table", "finance-table", "stats-table",
         "gov-table", "quality-table", "hr-table", "data-list"]
ok = rc == 0 and all(n in out for n in names) and out.count("\n") >= 8
results.append((ok, "--list-profiles 列全 8 个类型", rc, 0, out.strip()[:200]))

# --- 4. --list-profiles 不应被 profile 缺失拦住 ---
rc, out = run([VT, "--list-profiles"])
ok = rc == 0 and all(n in out for n in names)
results.append((ok, "verify --list-profiles 无 profile 也可用", rc, 0, out.strip()[:120]))

# --- 5. 缺 --src/--out → argparse error (exit 2) ---
case("make_table 缺参数 → exit 2", [MT, "--profile", "business-table"], 2, ["--src"])

# --- 6. 源文件不存在 ---
case("源文件不存在", [MT, "--profile", "business-table", "--src", os.path.join(TMP, "nope.md"),
                 "--out", os.path.join(TMP, "c.xlsx")], 1, [], None)

# --- 7. 不支持的扩展名 ---
bad_ext = os.path.join(TMP, "x.txt")
with open(bad_ext, "w", encoding="utf-8") as f:
    f.write("a,b\n1,2\n")
case("不支持的扩展名", [MT, "--profile", "business-table", "--src", bad_ext,
                  "--out", os.path.join(TMP, "d.xlsx")], 1, ["不支持的输入扩展名"])

# --- 8. xlsx 同路径覆盖必须拒绝（exit 2）---
same = os.path.join(TMP, "same.xlsx")
case("先正常生成", [MT, "--profile", "business-table", "--src", MD, "--out", same], 0, ["输出:"])
case("xlsx 同路径 → exit 2", [MT, "--profile", "business-table", "--src", same, "--out", same],
     2, ["同路径", "拒绝执行"])

# --- 9. 校验不存在的文件 ---
case("verify 不存在的文件 → exit 1",
     [VT, "--profile", "business-table", os.path.join(TMP, "ghost.xlsx")], 1, ["文件不存在"])

# --- 10. 校验时无 profile → exit 2 ---
case("verify 无 profile → exit 2", [VT, same], 2, ["未指定表格类型"])

# --- 11. 校验时未知 profile → exit 2 ---
case("verify 未知 profile → exit 2", [VT, "--profile", "bogus-type", same], 2, ["未找到表格类型"])

# --- 12. 空 md（无表格块）---
empty_md = os.path.join(TMP, "empty.md")
with open(empty_md, "w", encoding="utf-8") as f:
    f.write("# 只有标题\n\n没有任何表格。\n")
case("md 无表格块 → 报错", [MT, "--profile", "business-table", "--src", empty_md,
                       "--out", os.path.join(TMP, "e.xlsx")], 1, ["未找到 Markdown 表格"])

# --- 13. 空 CSV ---
empty_csv = os.path.join(TMP, "empty.csv")
open(empty_csv, "w", encoding="utf-8").close()
case("空 CSV → 报错", [MT, "--profile", "business-table", "--src", empty_csv,
                   "--out", os.path.join(TMP, "f.xlsx")], 1, ["为空"])

# --- 14. 工作表名含 Excel 禁用字符：应清洗并成功 ---
bad_sheet = os.path.join(TMP, "badsheet.xlsx")
case("工作表名含禁用字符 → 清洗后成功",
     [MT, "--profile", "business-table", "--src", MD, "--out", bad_sheet,
      "--sheet", "含[非法]:*?字符"], 0, ["已替换为下划线"])

# --- 15. 输出目录不存在 → exit 1 且不抛栈 ---
case("输出目录不存在 → exit 1",
     [MT, "--profile", "business-table", "--src", MD,
      "--out", os.path.join(TMP, "nodir", "sub", "x.xlsx")], 1, ["输出目录不存在"])

# --- 16. 文件不存在时不得出现 Python traceback ---
_rc, _out = run([MT, "--profile", "business-table", "--src", os.path.join(TMP, "ghost.md"),
                 "--out", os.path.join(TMP, "g.xlsx")])
results.append((_rc == 1 and "Traceback" not in _out and "源文件不存在" in _out,
                "源文件不存在 → 友好报错无堆栈", _rc, 1, _out.strip()[:100]))

print("=" * 78)
print("CLI 边界测试")
print("=" * 78)
npass = 0
for ok, name, rc, want_rc, detail in results:
    mark = "OK  " if ok else "FAIL"
    if ok:
        npass += 1
    print("[%s] %-40s rc=%s(期望%s) %s" % (mark, name, rc, want_rc, detail.replace("\n", " ")))
print("-" * 78)
print("CLI: %d/%d PASS" % (npass, len(results)))
sys.exit(0 if npass == len(results) else 1)
