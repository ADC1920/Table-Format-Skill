# -*- coding: utf-8 -*-
"""run_profile_regression.py — 全 profile 端到端回归。

对 profiles/ 下每个表格类型跑「生成 → 校验」闭环，断言退出码全 0。
再跑一组多负载形状（长标题、超宽表、大量空值、超长单元格、单列单行、
含前后空格与换行符的文本），确保规格切换不会在某类数据上崩。

用法：
    python run_profile_regression.py            # 全部 profile
    python run_profile_regression.py --profile academic-table
    python run_profile_regression.py --check-missing-profile   # 只验无 profile 报错
"""

import argparse
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from spec_loader import (ProfileNotSpecified, list_profiles, load_spec,
                         profile_meta, resolve_profile)

import make_table as mt
import verify_table as vt


def _shape_cases(tmp):
    """多负载形状：每种返回 (名称, headers, rows, title, note)。"""
    cases = []
    # 普通小表
    cases.append(("小表", ["名称", "数量"], [["甲", 1], ["乙", 2]], None, None))
    # 超过冻结阈值的大表
    cases.append(("大表25行", ["编码", "价格"],
                  [["QT-%03d" % i, i * 100] for i in range(1, 26)], None, None))
    # 长标题 + 长说明（折行与行高）
    cases.append(("长标题长说明", ["名称", "值"], [["甲", 1]],
                  "这是一条用于验证全表自动换行与行高自适应的超长标题" * 3,
                  "这是一段很长的说明文字，用于验证说明行在超过整表总宽后自动折行并相应增高。" * 2))
    # 超宽表（12 列）
    cases.append(("超宽12列", ["列%d" % i for i in range(1, 13)],
                  [[i * 10 + j for j in range(12)] for i in range(1, 6)], None, None))
    # 大量空值
    cases.append(("含空值", ["A", "B", "C"],
                  [["甲", None, ""], [None, "乙", 3], ["", None, None]], None, None))
    # 超长单元格（单格 200 字）
    cases.append(("超长单元格", ["说明"], [["长" * 200], ["短"]], None, None))
    # 单列单行
    cases.append(("单列单行", ["唯一列"], [["唯一值"]], None, None))
    # 含换行与空格的文本
    cases.append(("含换行空格", ["文本", "值"], [["第一行\n第二行", 1], ["  前后空格  ", 2]],
                  None, None))
    # 前导零与小数混合
    cases.append(("混合数字形态", ["编码", "金额"],
                  [["009", 12.5], ["A-001", 0.0], ["100", -3.25]], None, None))
    return cases


def run_one_profile(name, tmp, shapes=True, verbose=False):
    """跑一个 profile 的全形状回归，返回 (通过数, 失败项列表)。"""
    passed, failures = 0, []
    spec = load_spec(profile=name)
    meta = profile_meta(name)
    cases = _shape_cases(tmp) if shapes else [("小表", ["名称", "数量"],
                                            [["甲", 1], ["乙", 2]], None, None)]
    for label, headers, rows, title, note in cases:
        safe = label.replace("/", "_")
        out = os.path.join(tmp, "%s_%s.xlsx" % (name, safe))
        sheet = label[:20]
        try:
            info = mt.build_workbook(headers, rows, out, sheet, title, note, False, spec)
        except Exception as exc:
            failures.append("%s/%s 生成异常：%r" % (name, label, exc))
            continue
        if not info.get("freeze_ok"):
            failures.append("%s/%s 冻结回读不符（期望 %s 得 %s）" % (
                name, label, info.get("freeze"), info.get("freeze_readback")))
        try:
            rc = vt.verify(out, None, vt.NO_FREEZE, 0, name)
        except Exception as exc:
            failures.append("%s/%s 校验异常：%r" % (name, label, exc))
            continue
        if rc != 0:
            failures.append("%s/%s 校验 FAIL（exit=%d）" % (name, label, rc))
        else:
            passed += 1
        if verbose:
            print("    [%s] %s（%s）" % ("OK" if rc == 0 else "FAIL", label, meta.get("title", "")))
    return passed, failures


def main():
    ap = argparse.ArgumentParser(description="表格技能全 profile 回归")
    ap.add_argument("--profile", help="只跑指定 profile")
    ap.add_argument("--no-shapes", action="store_true", help="只跑基础形状（快）")
    ap.add_argument("--check-missing-profile", action="store_true",
                    help="只验证「不指定 profile 必须报错」")
    ap.add_argument("--verbose", action="store_true")
    a = ap.parse_args()

    if a.check_missing_profile:
        fails = []
        try:
            load_spec(profile=None)
            fails.append("load_spec 未抛 ProfileNotSpecified")
        except ProfileNotSpecified:
            print("  [OK] 无 profile 时 load_spec 抛 ProfileNotSpecified")
        try:
            resolve_profile(None)
            fails.append("resolve_profile 未抛 ProfileNotSpecified")
        except ProfileNotSpecified:
            print("  [OK] 无 profile 时 resolve_profile 抛 ProfileNotSpecified")
        try:
            mt.run(None, None, profile=None)
            fails.append("make_table.run 未抛 ProfileNotSpecified")
        except ProfileNotSpecified:
            print("  [OK] 无 profile 时 make_table.run 抛 ProfileNotSpecified")
        except Exception:
            pass  # src/out 为 None 的其它异常不算本项
        print("CHECK: %s" % ("PASS" if not fails else "FAIL %s" % fails))
        return 0 if not fails else 1

    names = [a.profile] if a.profile else list_profiles()
    if not names:
        print("ERROR: profiles/ 下没有任何类型", file=sys.stderr)
        return 2

    tmp = tempfile.mkdtemp(prefix="table_format_regression_")
    total_pass, all_fail = 0, []
    shapes = not a.no_shapes
    for name in names:
        meta = profile_meta(name)
        print("\n== %s（%s）==" % (name, meta.get("title", "")))
        try:
            passed, failures = run_one_profile(name, tmp, shapes=shapes, verbose=a.verbose)
        except Exception as exc:
            all_fail.append("%s 整体异常：%r" % (name, exc))
            print("  FAIL 整体异常 %r" % (exc,))
            continue
        total_pass += passed
        if failures:
            all_fail.extend(failures)
            for f in failures:
                print("  FAIL %s" % f)
        else:
            print("  OK %d/%d 形状全部通过" % (passed, passed))

    print("\n合计通过 %d 个（profile × 形状）" % total_pass)
    print("REGRESSION: %s" % ("PASS" if not all_fail else "FAIL %d 项" % len(all_fail)))
    print("样例目录: %s" % tmp)
    return 0 if not all_fail else 1


if __name__ == "__main__":
    sys.exit(main())
