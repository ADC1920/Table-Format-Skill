# -*- coding: utf-8 -*-
"""run_all_tests.py — table-format 全面测试总入口。

跑六套独立测试（共 301 项断言）并汇总结果：
    test_cli.py          CLI 契约与退出码（17 项）
    test_formats.py      输入格式与异常数据形态（26 项）
    test_styles.py       八 profile 样式真值核验（182 项）
    test_freeze_env.py   冻结矩阵与环境变量优先级（20 项）
    test_negative.py     逐类破坏产物验证门禁有效性（26 项）
    test_tools_lib.py    check_delivery / render_pages / 规范库质量（30 项）

另跑四套技能自带自检作为基线：make_table / verify_table 的 --selftest、
run_profile_regression、check_skill_consistency。

用法：
    python tests/run_all_tests.py            # 全部
    python tests/run_all_tests.py --fast     # 跳过渲染相关的慢项
    python tests/run_all_tests.py --only styles

依赖：openpyxl、PyYAML；渲染项另需 LibreOffice 与 pymupdf（缺失时该项会失败，
不影响其余套件）。

临时产物落在系统临时目录下的 table-format-tests/，不污染技能目录。
"""
import argparse
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
SCRIPTS = os.path.join(SKILL, "scripts")

SUITES = [
    ("cli", "test_cli.py", "CLI 契约与退出码"),
    ("formats", "test_formats.py", "输入格式与异常数据形态"),
    ("styles", "test_styles.py", "八 profile 样式真值核验"),
    ("freeze", "test_freeze_env.py", "冻结矩阵与环境变量优先级"),
    ("negative", "test_negative.py", "门禁有效性（负向）"),
    ("tools", "test_tools_lib.py", "工具与规范库质量"),
    ("encoding", "test_encoding_source.py", "编码与源文件形态边界"),
    ("engineering", "test_engineering.py", "工程属性（幂等/差异/性能/并发）"),
    ("round3", "test_round3.py", "xlsm / 外部链接 / 大数据量"),
]

BASELINE = [
    ("make_table 自检", [os.path.join(SCRIPTS, "make_table.py"), "--selftest"]),
    ("verify_table 自检", [os.path.join(SCRIPTS, "verify_table.py"), "--selftest"]),
    ("全 profile 回归", [os.path.join(SCRIPTS, "run_profile_regression.py")]),
    ("一致性自检", [os.path.join(SCRIPTS, "check_skill_consistency.py")]),
]


def run(cmd, env=None):
    e = dict(os.environ)
    if env:
        e.update(env)
    r = subprocess.run([sys.executable] + cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env=e, cwd=SKILL)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def main():
    ap = argparse.ArgumentParser(description="table-format 全面测试")
    ap.add_argument("--only", help="只跑指定套件（cli/formats/styles/freeze/negative/tools）")
    ap.add_argument("--no-baseline", action="store_true", help="跳过技能自带自检")
    a = ap.parse_args()

    workdir = os.path.join(tempfile.gettempdir(), "table-format-tests")
    os.makedirs(workdir, exist_ok=True)

    rows = []
    if not a.no_baseline:
        print("=" * 78)
        print("基线：技能自带自检")
        print("=" * 78)
        for name, cmd in BASELINE:
            rc, out = run(cmd)
            ok = rc == 0
            rows.append((ok, name, "rc=%d" % rc))
            print("[%s] %-24s rc=%d" % ("OK  " if ok else "FAIL", name, rc))

    print()
    print("=" * 78)
    print("独立测试套件（临时目录：%s）" % workdir)
    print("=" * 78)
    for key, fname, desc in SUITES:
        if a.only and a.only != key:
            continue
        path = os.path.join(HERE, fname)
        if not os.path.isfile(path):
            rows.append((False, desc, "脚本缺失"))
            print("[FAIL] %-30s 脚本缺失 %s" % (desc, fname))
            continue
        rc, out = run([path], {"TF_TEST_WORKDIR": workdir})
        # 抓套件自己的汇总行
        summary = ""
        for line in out.strip().splitlines():
            if "PASS" in line and (":" in line):
                summary = line.strip()
        ok = rc == 0
        rows.append((ok, desc, summary[-60:] if summary else "rc=%d" % rc))
        print("[%s] %-30s %s" % ("OK  " if ok else "FAIL", desc, summary[-70:] if summary else "rc=%d" % rc))
        if not ok:
            print("      失败明细：")
            for line in out.splitlines():
                if "FAIL" in line:
                    print("      " + line.strip()[:120])

    print()
    print("=" * 78)
    npass = sum(1 for ok, _, _ in rows if ok)
    print("总计：%d/%d 套件通过" % (npass, len(rows)))
    print("ALL TESTS: %s" % ("PASS" if npass == len(rows) else "FAIL"))
    return 0 if npass == len(rows) else 1


if __name__ == "__main__":
    sys.exit(main())
