# -*- coding: utf-8 -*-
"""query_standards.py — 表格规范库检索。

规范库位置：`references/standards/standards.jsonl`（每行一条），由
`build_standards_catalog.py` 从调研 jsonl 合并生成，index.md 与
domain-<slug>.md 也由它生成。本脚本只读，不改库。

用法：
    python query_standards.py --list-domains                    # 领域与条数
    python query_standards.py --list-profiles                   # 按 profile 统计
    python query_standards.py --domain finance                  # 某领域全部
    python query_standards.py --keyword 三线表                   # 关键词（名称+规格）
    python query_standards.py --profile academic-table          # 某表格类型关联条目
    python query_standards.py --category 续表                    # 按子类
    python query_standards.py --domain finance --min-conf high
    python query_standards.py --keyword 续表 --full              # 输出完整 specs
"""

import argparse
import collections
import json
import os
import sys

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CATALOG = os.path.join(SKILL_ROOT, "references", "standards", "standards.jsonl")

CONF_ORDER = {"high": 0, "medium": 1, "low": 2}
CONF_MARK = {"high": "高", "medium": "中", "low": "低"}


def load_rows(path=CATALOG):
    if not os.path.isfile(path):
        print("ERROR: 规范库不存在 %s" % path, file=sys.stderr)
        print("先跑 build_standards_catalog.py 生成。", file=sys.stderr)
        raise SystemExit(2)
    rows = []
    with open(path, encoding="utf-8") as f:
        for ln, line in enumerate(f, 1):
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError as exc:
                print("WARN: 第 %d 行解析失败（%s），已跳过" % (ln, exc), file=sys.stderr)
    return rows


def match(row, keyword):
    kw = keyword.lower()
    for field in ("name", "specs", "category", "applies_to"):
        if kw in str(row.get(field) or "").lower():
            return True
    return False


def print_rows(rows, full=False):
    for r in rows:
        conf = r.get("confidence", "?")
        mark = CONF_MARK.get(conf, conf)
        print("[%s][%s] %s" % (r.get("id", "?"), mark, r.get("name", "")))
        print("    分类: %s | 关联 profile: %s" % (r.get("category", "-"),
                                              r.get("profile", "-")))
        specs = str(r.get("specs") or "")
        if not full and len(specs) > 140:
            specs = specs[:140] + "…（--full 看全文）"
        print("    规格: %s" % specs)
        if r.get("applies_to"):
            print("    适用: %s" % r["applies_to"])
        print()


def main():
    ap = argparse.ArgumentParser(description="表格规范库检索")
    ap.add_argument("--list-domains", action="store_true", help="列出领域与条数")
    ap.add_argument("--list-profiles", action="store_true", help="按表格类型统计条数")
    ap.add_argument("--list-categories", action="store_true", help="列出子类与条数")
    ap.add_argument("--domain", help="按领域过滤（如 finance / academic）")
    ap.add_argument("--profile", help="按表格类型过滤（如 academic-table）")
    ap.add_argument("--category", help="按子类过滤（模糊匹配）")
    ap.add_argument("--keyword", help="关键词（匹配名称/规格/分类/适用范围）")
    ap.add_argument("--min-conf", choices=["high", "medium", "low"],
                    help="置信度下限")
    ap.add_argument("--full", action="store_true", help="输出完整 specs")
    ap.add_argument("--limit", type=int, default=0, help="最多输出条数（0=不限）")
    ap.add_argument("--json", action="store_true", help="以 JSON 输出")
    a = ap.parse_args()

    rows = load_rows()

    if a.list_domains:
        cnt = collections.Counter(r.get("domain") or "-" for r in rows)
        print("规范库共 %d 条，%d 个领域：" % (len(rows), len(cnt)))
        for name, n in sorted(cnt.items(), key=lambda kv: (-kv[1], kv[0])):
            print("  %-14s %3d 条" % (name, n))
        return 0

    if a.list_profiles:
        cnt = collections.Counter(r.get("profile") or "-" for r in rows)
        print("按表格类型统计：")
        for name, n in sorted(cnt.items(), key=lambda kv: (-kv[1], kv[0])):
            print("  %-18s %3d 条" % (name, n))
        return 0

    if a.list_categories:
        cnt = collections.Counter(r.get("category") or "-" for r in rows)
        print("子类共 %d 个：" % len(cnt))
        for name, n in sorted(cnt.items(), key=lambda kv: (-kv[1], kv[0])):
            print("  %-24s %3d 条" % (name, n))
        return 0

    sel = rows
    if a.domain:
        d = a.domain.lower()
        sel = [r for r in sel if d in str(r.get("domain") or "").lower()]
    if a.profile:
        p = a.profile.lower()
        sel = [r for r in sel if p in str(r.get("profile") or "").lower()]
    if a.category:
        c = a.category.lower()
        sel = [r for r in sel if c in str(r.get("category") or "").lower()]
    if a.keyword:
        sel = [r for r in sel if match(r, a.keyword)]
    if a.min_conf:
        floor = CONF_ORDER[a.min_conf]
        sel = [r for r in sel if CONF_ORDER.get(r.get("confidence"), 9) <= floor]

    # 按置信度排序（高的在前），同档保持原序
    sel = sorted(sel, key=lambda r: CONF_ORDER.get(r.get("confidence"), 9))
    if a.limit and a.limit > 0:
        sel = sel[:a.limit]

    if a.json:
        print(json.dumps(sel, ensure_ascii=False, indent=2))
        return 0

    if not sel:
        print("未命中任何条目。试 --list-domains 或 --list-profiles 看可用范围。")
        return 0
    print("命中 %d 条：\n" % len(sel))
    print_rows(sel, full=a.full)
    return 0


if __name__ == "__main__":
    sys.exit(main())
