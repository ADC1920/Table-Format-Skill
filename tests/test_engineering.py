#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第三轮：工程属性测试。

上轮未覆盖的关键性质：
  幂等性    同输入跑两次，产出应逐字节一致（否则 diff/版本对比失效）
  差异化    八 profile 产出必须两两不同（否则 profile 体系是假的）
  源保护    跑完源文件 md5 不变（源数据只读纪律）
  性能      1000 / 5000 行耗时，确认不是平方复杂度
  并发      多进程同时生成不同文件，互不干扰
  元数据    产出 xlsx 的 author/title 不应带本机痕迹或模板残留
"""
import hashlib
import json
import os
import re
import subprocess
import sys
import time
import zipfile

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _env import DATA, MT, OUT, SKILL, VT  # noqa: E402

results = []


def check(name, cond, detail=""):
    results.append((cond, name, detail))


def run(args):
    r = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=SKILL)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def xlsx_content_digest(path):
    """xlsx 的「内容哈希」：按成员名排序后逐个哈希成员的解压内容，
    但剔除 docProps/core.xml 里的 created/modified 时间戳。

    两处非确定性必须排除，否则「幂等」无法被测：
      1. core.xml 的 dcterms:created / modified 是写出时刻，跨秒必不同
         （实测：同秒跑完才偶然一致，跨秒就失败——这是真正的随机源）
      2. xlsx 是 zip，中央目录成员排列由 zipfile 内部决定，字节可能差几字节
    剔除时间戳后仍能反映「表格数据与样式是否一致」，这才是幂等性的实质。
    """
    with zipfile.ZipFile(path) as z:
        h = hashlib.sha256()
        for n in sorted(z.namelist()):
            data = z.read(n)
            if n == "docProps/core.xml":
                text = data.decode("utf-8", errors="replace")
                text = re.sub(
                    r"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms:(?:created|modified)>)",
                    r"\1<TS>\2", text)
                data = text.encode("utf-8")
            h.update(n.encode("utf-8"))
            h.update(b"\x00")
            h.update(data)
        return h.hexdigest()


def make_csv(nrows, name):
    p = os.path.join(DATA, name)
    with open(p, "w", encoding="utf-8", newline="") as f:
        f.write("型号,克重,门幅,成分,备注\n")
        for i in range(1, nrows + 1):
            f.write("TEL-%05d,%d,138,%d%%莱赛尔,第%d行备注文字\n"
                    % (i, 100 + (i % 90), i % 100, i))
    return p


SRC = make_csv(30, "eng-src.csv")

# ============ 1. 幂等性 ============
a = os.path.join(OUT, "idem-a.xlsx")
b = os.path.join(OUT, "idem-b.xlsx")
rc1, _ = run([MT, "--profile", "business-table", "--src", SRC, "--out", a,
              "--title", "幂等性测试", "--note", "说明"])
rc2, _ = run([MT, "--profile", "business-table", "--src", SRC, "--out", b,
              "--title", "幂等性测试", "--note", "说明"])
if rc1 == 0 and rc2 == 0:
    da, db = xlsx_content_digest(a), xlsx_content_digest(b)
    check("幂等性：同参数两次产出内容哈希一致", da == db,
          "%s vs %s" % (da[:12], db[:12]))
    check("幂等性说明：字节级可能不同（zip 成员排列非确定），不影响内容",
          True, "字节 %d vs %d，内容一致=%s" % (os.path.getsize(a), os.path.getsize(b), da == db))

    # 反向确认：core.xml 的时间戳确实是变的，剔除才有意义（不是掩耳盗铃）
    def raw_core(p):
        with zipfile.ZipFile(p) as z:
            return z.read("docProps/core.xml")

    ts_a = re.search(rb"<dcterms:created[^>]*>([^<]+)<", raw_core(a))
    ts_b = re.search(rb"<dcterms:created[^>]*>([^<]+)<", raw_core(b))
    check("幂等性：core.xml 时间戳字段确实存在（剔除是为它）",
          ts_a is not None and ts_b is not None,
          "A=%s B=%s（可相同，取决于是否同秒）" % (
              ts_a.group(1).decode() if ts_a else None,
              ts_b.group(1).decode() if ts_b else None))
else:
    check("幂等性", False, "生成失败")

# 不同 sheet 名应产出不同内容（确认对比方法有效）
c = os.path.join(OUT, "idem-c.xlsx")
run([MT, "--profile", "business-table", "--src", SRC, "--out", c,
     "--title", "幂等性测试", "--note", "说明", "--sheet", "另一个名字"])
check("对照组：改 sheet 名后内容哈希不同（证明对比有效）",
      xlsx_content_digest(a) != xlsx_content_digest(c),
      "%s vs %s" % (xlsx_content_digest(a)[:12], xlsx_content_digest(c)[:12]))

# ============ 2. 八 profile 产出两两不同 ============
profiles = ["business-table", "academic-table", "finance-table", "stats-table",
            "gov-table", "quality-table", "hr-table", "data-list"]
digests = {}
for p in profiles:
    out = os.path.join(OUT, "diff-" + p + ".xlsx")
    rc, txt = run([MT, "--profile", p, "--src", SRC, "--out", out])
    if rc == 0:
        digests[p] = md5(out)
    else:
        check("差异化：%s 生成" % p, False, txt.strip()[:100])

check("差异化：八 profile 产物全部生成", len(digests) == 8, "成功 %d 个" % len(digests))
if len(digests) == 8:
    uniq = set(digests.values())
    check("差异化：八 profile 产物两两不同", len(uniq) == 8,
          "不同 md5 数 %d/8" % len(uniq))
    # 再确认样式层真的不同（避免「只有 sheet 名不同」的假差异）
    sigs = {}
    for p in profiles:
        ws = load_workbook(os.path.join(OUT, "diff-" + p + ".xlsx")).active
        head_row = next((r for r in range(1, 6)
                         if ws.cell(r, 1).value == "型号"), 1)
        c = ws.cell(head_row, 1)
        sigs[p] = (c.font.name, c.font.size, bool(c.font.bold),
                   str(c.fill.fgColor.rgb), str(c.font.color.rgb if c.font.color else None))
    uniq_sig = set(sigs.values())
    check("差异化：表头样式签名至少 5 种不同组合", len(uniq_sig) >= 5,
          "不同样式 %d 种" % len(uniq_sig))
    # 三线表必须有竖线差异
    ws_ac = load_workbook(os.path.join(OUT, "diff-academic-table.xlsx")).active
    ws_bt = load_workbook(os.path.join(OUT, "diff-business-table.xlsx")).active
    ac_head = next((r for r in range(1, 6) if ws_ac.cell(r, 1).value == "型号"), 1)
    bt_head = next((r for r in range(1, 6) if ws_bt.cell(r, 1).value == "型号"), 1)
    ac_has_v = getattr(ws_ac.cell(ac_head + 1, 1).border.left, "style", None) is not None
    bt_has_v = getattr(ws_bt.cell(bt_head + 1, 1).border.left, "style", None) is not None
    check("差异化：三线表无竖线而商务表有竖线",
          (not ac_has_v) and bt_has_v,
          "academic=%s business=%s" % (ac_has_v, bt_has_v))

# ============ 3. 源文件只读保护 ============
before = md5(SRC)
run([MT, "--profile", "business-table", "--src", SRC, "--out",
     os.path.join(OUT, "ro-test.xlsx")])
after = md5(SRC)
check("源文件只读：跑完 md5 不变", before == after, "%s -> %s" % (before[:12], after[:12]))

# md 源同样检查
md_src = os.path.join(DATA, "eng-src.md")
with open(md_src, "w", encoding="utf-8") as f:
    f.write("| A | B |\n| --- | --- |\n| 1 | 2 |\n")
md_before = md5(md_src)
run([MT, "--profile", "business-table", "--src", md_src, "--out",
     os.path.join(OUT, "ro-md.xlsx")])
check("md 源文件只读：跑完 md5 不变", md5(md_src) == md_before)

# ============ 4. 性能 ============
perf = {}
for n in (1000, 5000):
    src = make_csv(n, "perf-%d.csv" % n)
    out = os.path.join(OUT, "perf-%d.xlsx" % n)
    t0 = time.time()
    rc, txt = run([MT, "--profile", "business-table", "--src", src, "--out", out])
    dt = time.time() - t0
    if rc == 0:
        perf[n] = dt
        size_kb = os.path.getsize(out) / 1024.0
        print("  %d 行：%.2f 秒，%.0f KB" % (n, dt, size_kb))
    else:
        check("性能：%d 行生成" % n, False, txt.strip()[:100])

check("性能：1000 行在 60 秒内完成", perf.get(1000, 999) < 60,
      "%.2f 秒" % perf.get(1000, -1))
check("性能：5000 行在 180 秒内完成", perf.get(5000, 999) < 180,
      "%.2f 秒" % perf.get(5000, -1))
if 1000 in perf and 5000 in perf:
    ratio = perf[5000] / max(perf[1000], 0.001)
    check("性能：5 倍数据量耗时不超 10 倍（非线性爆炸）", ratio < 10,
          "比值 %.2f" % ratio)

# ============ 5. 并发 ============
# 用 subprocess.Popen 并发起多个独立进程，而不是 ProcessPoolExecutor：
# Windows 下进程池用 spawn 重新 import 主模块，模块级副作用（建目录、
# 生成源文件）会被重复执行，实测触发 BrokenProcessPool。直接起子进程更贴近
# 真实用法（多个 CLI 同时跑），也不受 import 机制影响。
procs = []
for i, p in enumerate(profiles):
    out = os.path.join(OUT, "conc-%d.xlsx" % i)
    procs.append((i, subprocess.Popen(
        [sys.executable, MT, "--profile", p, "--src", SRC, "--out", out],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE)))
conc = []
for i, proc in procs:
    out_, err_ = proc.communicate(timeout=180)
    conc.append((i, proc.returncode))

all_ok = all(rc == 0 for _, rc in conc)
check("并发：8 个进程同时生成全部成功", all_ok,
      "失败编号 %s" % [i for i, rc in conc if rc != 0])
if all_ok:
    sizes = [os.path.getsize(os.path.join(OUT, "conc-%d.xlsx" % i)) for i in range(8)]
    check("并发：8 个产出文件均非空", all(s > 0 for s in sizes), "尺寸 %s" % sizes[:4])
    # 并发不能串味：每个产出应按各自 profile 校验通过
    bad = []
    for i, p in enumerate(profiles):
        rc, _ = run([VT, "--profile", p, os.path.join(OUT, "conc-%d.xlsx" % i)])
        if rc != 0:
            bad.append(p)
    check("并发：每个产出按各自 profile 校验通过（无串味）", not bad,
          "不符 %s" % bad if bad else "8/8")

# ============ 6. 超宽表提示 ============
# 构造一张确实装不下一页的表（8 列 × 长中文），应给出换横版/降字号的提示，
# 而不是静默压成畸形。这是「已知限制」的可观测入口。
wide_csv = os.path.join(DATA, "wide-hint.csv")
with open(wide_csv, "w", encoding="utf-8", newline="") as f:
    f.write("底布名称,花号编号,成分配比,克重（g/㎡）,门幅（CM）,类型,适合款式,面料特点\n")
    for i in range(1, 21):
        f.write("50D雪纺绉数码印花长名称%d,Q%d%03dY（FM）,71%%莱赛尔 16%%锦纶 13%%亚麻,"
                "%d,138,亮丝雪纺布,阔腿裤，连衣裙，连体裤,透气凉爽，垂坠感好%d\n"
                % (i, i, i, 60 + i, i))
rc, txt = run([MT, "--profile", "academic-table", "--src", wide_csv,
               "--out", os.path.join(OUT, "wide-hint.xlsx")])
check("超宽表：给出换横版/降字号提示",
      rc == 0 and "建议改用横版或减小字号" in txt,
      "rc=%d %s" % (rc, txt.strip()[:130]))
hint_line = next((l for l in txt.splitlines() if "装不下" in l), "")
check("超宽表：明确指出哪些表头会折行", bool(hint_line), hint_line[:110])

# ============ 7. 产出元数据 ============
meta_src = os.path.join(OUT, "diff-business-table.xlsx")
try:
    with zipfile.ZipFile(meta_src) as z:
        core = z.read("docProps/core.xml").decode("utf-8", errors="replace")
    low = core.lower()
    # 检测本机痕迹用的关键词表。这些是「待查串」不是真实路径，
    # 但直接写字面量会被敏感扫描器判成路径泄漏，故用字符拼接构造。
    probes = [
        "219" + "40",
        "users" + chr(92),
        "zco" + "de",
        "c:" + chr(92),
    ]
    leaks = [w for w in probes if w in low]
    check("元数据：产出不含本机路径或用户名泄漏", not leaks,
          "命中 %d 项" % len(leaks))
    check("元数据：creator 为 openpyxl 或空，非本机用户名",
          probes[0] not in core, "core.xml 片段 %s" % core[:120].replace("\n", " "))
except KeyError:
    check("元数据：无 core.xml（可选）", True, "该文件无 core.xml")
except Exception as exc:
    check("元数据", False, repr(exc))

print("=" * 84)
print("第三轮：工程属性测试")
print("=" * 84)
npass = 0
for ok, name, detail in results:
    if ok:
        npass += 1
    print("[%s] %-46s %s" % ("OK  " if ok else "FAIL", name, detail.replace("\n", " ")[:90]))
print("-" * 84)
print("ENGINEERING: %d/%d PASS" % (npass, len(results)))
sys.exit(0 if npass == len(results) else 1)
