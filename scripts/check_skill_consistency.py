# -*- coding: utf-8 -*-
"""check_skill_consistency.py — 技能文档与实物一致性自检。

断言项：
    1. profiles/ 下 yaml 均可解析，且 profile.name 与文件名一致
    2. 基线 yaml（table-format.yaml）可解析，version 与 spec_loader 默认一致
    3. SKILL.md 的 profile 速选表覆盖 profiles/ 全部类型名
    4. 规范库存在时：standards.jsonl 行数 == index.md 声明的条数
    5. 各 profile 的 domain/profile 字段在规范库中确有对应条目
    6. spec_loader.DEFAULT_SPEC 的键与基线 yaml 的键集合一致（防两处漂移）

用法：
    python check_skill_consistency.py
    python check_skill_consistency.py --verbose
"""

import argparse
import collections
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yaml

from spec_loader import (DEFAULT_SPEC, list_profiles, load_spec, profile_meta,
                         resolve_profile)

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL_MD = os.path.join(SKILL_ROOT, "SKILL.md")
BASE_YAML = os.path.join(SKILL_ROOT, "table-format.yaml")
STANDARDS = os.path.join(SKILL_ROOT, "references", "standards", "standards.jsonl")
INDEX_MD = os.path.join(SKILL_ROOT, "references", "standards", "index.md")


def main():
    ap = argparse.ArgumentParser(description="表格技能一致性自检")
    ap.add_argument("--verbose", action="store_true")
    a = ap.parse_args()

    oks, fails = [], []

    def check(cond, name, detail=""):
        (oks if cond else fails).append(name + (" " + detail if detail else ""))
        if a.verbose or not cond:
            print("  [%s] %s %s" % ("OK" if cond else "FAIL", name, detail))

    # 1) profiles 可解析且 name 一致
    names = list_profiles()
    bad = []
    for n in names:
        try:
            meta = profile_meta(n)
        except Exception as exc:
            bad.append("%s 解析失败 %r" % (n, exc))
            continue
        if (meta.get("name") or "") != n:
            bad.append("%s 的 profile.name=%r 与文件名不一致" % (n, meta.get("name")))
    check(names and not bad, "profiles 解析与 name 一致性",
          "共 %d 个；%s" % (len(names), "; ".join(bad) if bad else "全部一致"))

    # 2) 基线 yaml version 与 DEFAULT_SPEC 一致
    try:
        with open(BASE_YAML, encoding="utf-8") as f:
            base = yaml.safe_load(f) or {}
        same_ver = str(base.get("version")) == str(DEFAULT_SPEC.get("version"))
        check(same_ver, "基线 yaml 与 DEFAULT_SPEC 版本一致",
              "yaml=%s default=%s" % (base.get("version"), DEFAULT_SPEC.get("version")))
    except Exception as exc:
        base = {}
        check(False, "基线 yaml 可解析", repr(exc))

    # 3) SKILL.md 覆盖全部 profile 名
    try:
        with open(SKILL_MD, encoding="utf-8") as f:
            md = f.read()
        missing = [n for n in names if n not in md]
        check(not missing, "SKILL.md 覆盖全部 profile",
              "缺 %s" % missing if missing else "共 %d 个全部出现" % len(names))
    except Exception as exc:
        check(False, "SKILL.md 可读", repr(exc))

    # 4) 规范库行数与 index.md 声明一致
    if os.path.isfile(STANDARDS) and os.path.isfile(INDEX_MD):
        rows = []
        with open(STANDARDS, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    rows.append(json.loads(line))
        with open(INDEX_MD, encoding="utf-8") as f:
            head = f.read(500)
        m = re.search(r"共 (\d+) 条", head)
        declared = int(m.group(1)) if m else None
        check(declared == len(rows), "规范库条数与 index.md 一致",
              "jsonl=%d index=%s" % (len(rows), declared))

        # 5) profile 字段在库中有条目
        cnt = collections.Counter(r.get("profile") for r in rows)
        no_entry = [n for n in names if cnt.get(n, 0) == 0]
        check(not no_entry, "每个 profile 在规范库中有条目",
              "无条目：%s" % no_entry if no_entry else
              "最少 %d 条" % min(cnt.get(n, 0) for n in names))
    else:
        print("  [SKIP] 规范库未生成（standards.jsonl / index.md）")

    # 6) DEFAULT_SPEC 与基线 yaml 键集合一致
    if base:
        def keyset(d):
            out = set()
            for k, v in (d or {}).items():
                out.add(k)
                if isinstance(v, dict):
                    out |= {"%s.%s" % (k, kk) for kk in keyset(v)}
            return out
        only_spec = keyset(DEFAULT_SPEC) - keyset(base)
        only_yaml = keyset(base) - keyset(DEFAULT_SPEC)
        check(not only_spec and not only_yaml, "DEFAULT_SPEC 与基线 yaml 键集合一致",
              "spec 独有 %s；yaml 独有 %s" % (sorted(only_spec), sorted(only_yaml)))

    # 7) 每个 profile 都能加载出完整 spec（必填顶层键齐全）
    need = ["font", "header", "title", "note", "body", "rules", "row_height",
            "col_width", "freeze", "print"]
    lack = []
    for n in names:
        s = load_spec(profile=n)
        miss = [k for k in need if k not in s]
        if miss:
            lack.append("%s 缺 %s" % (n, miss))
    check(not lack, "全部 profile 加载出完整规格", "; ".join(lack) if lack else "")

    print("\n一致性自检：%d 项通过，%d 项失败" % (len(oks), len(fails)))
    print("CONSISTENCY: %s" % ("PASS" if not fails else "FAIL"))
    return 0 if not fails else 1


if __name__ == "__main__":
    sys.exit(main())
