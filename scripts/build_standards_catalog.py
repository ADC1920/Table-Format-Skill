# -*- coding: utf-8 -*-
"""build_standards_catalog.py — 调研 jsonl 合并为规范库。

输入：若干调研 jsonl（每条一个 JSON 对象，字段 id/category/name/specs/
applies_to/confidence/domain/profile），生成三份产物到
`references/standards/`：

    standards.jsonl      全部条目（机器可查，query_standards.py 读这个）
    index.md             一行式索引（按领域分组，含条数与置信度分布）
    domain-<slug>.md     分域详表（含完整 specs，供人读）

**source 字段一律丢弃**：调研阶段的 URL 只用于本地追溯，公开仓与技能内
不带检索痕迹（与 article-format 同一口径）。

用法：
    python build_standards_catalog.py --src a.jsonl b.jsonl c.jsonl
    python build_standards_catalog.py --src-dir 调研目录
    python build_standards_catalog.py --src-dir 目录 --check   # 只校验不写
"""

import argparse
import collections
import glob
import json
import os
import sys

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(SKILL_ROOT, "references", "standards")

REQUIRED = ("id", "category", "name", "specs", "confidence", "domain", "profile")
VALID_CONF = ("high", "medium", "low")
# specs 长度上限：原定 200 字（article-format 口径），但表格规范里有几类内容
# 天然压不进 200——抽样检验的字码对照表、FMEA 的完整列清单、报表的要素枚举。
# 截断这些会直接毁掉条目的可用性（字码表砍一半就没法用），故放宽到 260，
# 只拦明显冗长的复述，不拦参数表。
SPECS_MAX = 260

DOMAIN_TITLE = {
    "stats": "统计与国标",
    "finance": "财务与披露",
    "academic": "学术与出版",
    "gov": "公文与政务",
    "quality": "质量与工程",
    "hr-project": "人事与项目运营",
}

# domain → 默认 profile 映射。调研时各路的 profile 字段是统一填的（t1 路全部填
# stats-table），但实际有跨域内容——t1 里「学术表格规范」「三线表与表线」两个
# 子类应归 academic-table。合并时按 domain 校正，避免 --profile academic-table
# 查不到自己领域的条目。调研文件里已是正确值的（同域）不受影响。
DOMAIN_PROFILE = {
    "stats": "stats-table",
    "finance": "finance-table",
    "academic": "academic-table",
    "gov": "gov-table",
    "quality": "quality-table",
    "hr-project": "business-table",
}

# 子类 → profile 覆盖：优先级高于 DOMAIN_PROFILE。
# 用于同域内跨 profile 的内容：统计路里混着学术表格条目，人事项目路里混着
# 人事表单（hr-table）与数据整备（data-list）两类——调研时都填了 business-table。
CATEGORY_PROFILE = {
    "学术表格规范": "academic-table",
    "三线表与表线": "academic-table",
    "标准编写规则": "stats-table",
    "人事档案表": "hr-table",
    "考勤绩效表": "hr-table",
    "招聘培训与会议表": "hr-table",
    "数据分析表": "data-list",
    "交叉表透视表": "data-list",
}


def load_files(paths):
    rows, errors = [], []
    seen = set()
    for p in paths:
        if not os.path.isfile(p):
            errors.append("文件不存在：%s" % p)
            continue
        with open(p, encoding="utf-8") as f:
            for ln, line in enumerate(f, 1):
                line = line.strip()
                if not line:
                    continue
                try:
                    row = json.loads(line)
                except json.JSONDecodeError as exc:
                    errors.append("%s:%d JSON 解析失败（%s）" % (
                        os.path.basename(p), ln, exc))
                    continue
                miss = [k for k in REQUIRED if not row.get(k)]
                if miss:
                    errors.append("%s:%d 缺字段 %s" % (os.path.basename(p), ln,
                                                    ",".join(miss)))
                    continue
                if row["confidence"] not in VALID_CONF:
                    errors.append("%s:%d confidence 非法值 %r" % (
                        os.path.basename(p), ln, row["confidence"]))
                    continue
                if len(str(row["specs"])) > SPECS_MAX:
                    errors.append("%s:%d specs 超 %d 字（%d）" % (
                        os.path.basename(p), ln, SPECS_MAX, len(str(row["specs"]))))
                    continue
                rid = row["id"]
                if rid in seen:
                    errors.append("%s:%d id 重复 %s" % (os.path.basename(p), ln, rid))
                    continue
                seen.add(rid)
                # source 一律丢弃
                row.pop("source", None)
                row.pop("verify_evidence", None)
                row.pop("verified", None)
                # profile 按 domain/category 校正（调研文件里的统一值可能是跨域的）
                cat = row.get("category")
                if cat in CATEGORY_PROFILE:
                    row["profile"] = CATEGORY_PROFILE[cat]
                elif row.get("domain") in DOMAIN_PROFILE:
                    row["profile"] = DOMAIN_PROFILE[row["domain"]]
                rows.append(row)
    return rows, errors


def write_catalog(rows, out_dir=OUT_DIR):
    os.makedirs(out_dir, exist_ok=True)
    rows = sorted(rows, key=lambda r: r["id"])

    # 1) standards.jsonl
    jl = os.path.join(out_dir, "standards.jsonl")
    with open(jl, "w", encoding="utf-8", newline="\n") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

    # 2) 分域详表
    by_domain = collections.defaultdict(list)
    for r in rows:
        by_domain[r["domain"]].append(r)
    for dom, items in by_domain.items():
        path = os.path.join(out_dir, "domain-%s.md" % dom)
        title = DOMAIN_TITLE.get(dom, dom)
        conf = collections.Counter(i["confidence"] for i in items)
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write("# %s（%s）规范详表\n\n" % (title, dom))
            f.write("共 %d 条｜高 %d / 中 %d / 低 %d\n\n" % (
                len(items), conf.get("high", 0), conf.get("medium", 0),
                conf.get("low", 0)))
            f.write("> 由 build_standards_catalog.py 生成，勿手改。\n\n")
            cur_cat = None
            for it in sorted(items, key=lambda x: (x["category"], x["id"])):
                if it["category"] != cur_cat:
                    cur_cat = it["category"]
                    f.write("\n## %s\n\n" % cur_cat)
                f.write("### %s｜%s\n\n" % (it["id"], it["name"]))
                f.write("- **规格**：%s\n" % it["specs"])
                f.write("- **适用**：%s\n" % it.get("applies_to", "-"))
                f.write("- **置信度**：%s｜**关联类型**：%s\n\n" % (
                    it["confidence"], it["profile"]))

    # 3) index.md
    idx = os.path.join(out_dir, "index.md")
    conf_all = collections.Counter(r["confidence"] for r in rows)
    with open(idx, "w", encoding="utf-8", newline="\n") as f:
        f.write("# 表格规范库索引\n\n")
        f.write("共 %d 条｜高 %d / 中 %d / 低 %d｜%d 个领域\n\n" % (
            len(rows), conf_all.get("high", 0), conf_all.get("medium", 0),
            conf_all.get("low", 0), len(by_domain)))
        f.write("> 由 build_standards_catalog.py 生成，勿手改。"
                "检索用 query_standards.py，详表见 domain-<领域>.md。\n\n")
        f.write("| 领域 | 条数 | 详表 |\n| --- | --- | --- |\n")
        for dom in sorted(by_domain, key=lambda d: -len(by_domain[d])):
            f.write("| %s（%s） | %d | domain-%s.md |\n" % (
                DOMAIN_TITLE.get(dom, dom), dom, len(by_domain[dom]), dom))
        f.write("\n")
        for dom in sorted(by_domain):
            items = sorted(by_domain[dom], key=lambda x: (x["category"], x["id"]))
            f.write("\n## %s（%s）\n\n" % (DOMAIN_TITLE.get(dom, dom), dom))
            for it in items:
                f.write("- `%s` [%s] **%s** — %s\n" % (
                    it["id"], it["confidence"], it["name"],
                    str(it["specs"])[:70] + ("…" if len(str(it["specs"])) > 70 else "")))
    return {"jsonl": jl, "index": idx, "domains": sorted(by_domain),
            "total": len(rows)}


def main():
    ap = argparse.ArgumentParser(description="调研 jsonl → 表格规范库")
    ap.add_argument("--src", nargs="*", help="调研 jsonl 路径（可多个）")
    ap.add_argument("--src-dir", help="调研 jsonl 所在目录（取 *.jsonl）")
    ap.add_argument("--out-dir", default=OUT_DIR, help="输出目录")
    ap.add_argument("--check", action="store_true", help="只校验，不写文件")
    a = ap.parse_args()

    paths = list(a.src or [])
    if a.src_dir:
        paths += sorted(glob.glob(os.path.join(a.src_dir, "*.jsonl")))
    if not paths:
        ap.error("需要 --src 或 --src-dir")

    rows, errors = load_files(paths)
    if errors:
        print("校验失败 %d 处：" % len(errors), file=sys.stderr)
        for e in errors[:50]:
            print("  - %s" % e, file=sys.stderr)
        if len(errors) > 50:
            print("  …共 %d 处" % len(errors), file=sys.stderr)
        return 1

    conf = collections.Counter(r["confidence"] for r in rows)
    doms = collections.Counter(r["domain"] for r in rows)
    print("校验通过：%d 条｜高 %d / 中 %d / 低 %d｜%d 个领域" % (
        len(rows), conf.get("high", 0), conf.get("medium", 0),
        conf.get("low", 0), len(doms)))
    for d, n in sorted(doms.items(), key=lambda kv: -kv[1]):
        print("  %-14s %3d" % (d, n))

    if a.check:
        print("（--check 模式，未写文件）")
        return 0

    info = write_catalog(rows, a.out_dir)
    print("已写出：")
    print("  %s" % info["jsonl"])
    print("  %s" % info["index"])
    for d in info["domains"]:
        print("  %s" % os.path.join(a.out_dir, "domain-%s.md" % d))
    return 0


if __name__ == "__main__":
    sys.exit(main())
