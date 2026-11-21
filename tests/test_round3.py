#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第三轮补测：xlsm 宏保留、外部链接工作簿、超大数据量。

前两轮明确列为「未覆盖」的三项，这里补齐并固化。

关于 xlsm：本技能只**读**源表取值，不重新保存源文件，因此不存在
「openpyxl 保存时丢宏」的问题（那是编辑原文件才需要 keep_vba 的场景）。
本套件验证的是：①能正常读取带宏的 xlsm ②源文件不被改动
③产出是 .xlsx，宏不随之传递——这是设计决定，需明确告知而非静默。

关于外部链接：同样是「只读」场景，读源表时不解析外部链接、不重新保存，
故不会破坏链接。验证能读出缓存值且源文件不变。

关于大数据量：实测耗时与线性度，并把数字记录下来（当前实现是
openpyxl 逐格样式去重的固有开销，非本技能逻辑问题）。
"""
import os
import sys
import time
import zipfile

from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _env import DATA, MT, OUT, SKILL, VT, run  # noqa: E402

results = []


def check(name, cond, detail=""):
    results.append((cond, name, detail))


def md5(path):
    import hashlib
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def gen(src, tag, profile="business-table", extra=None):
    out = os.path.join(OUT, "t3-" + tag + ".xlsx")
    rc, txt = run([MT, "--profile", profile, "--src", src, "--out", out] + (extra or []))
    return rc, txt, out


# ============ 1. xlsm 带宏 ============
plain = os.path.join(DATA, "t3-plain.xlsx")
wb = Workbook()
ws = wb.active
ws.append(["型号", "数量", "备注"])
for i in range(1, 6):
    ws.append(["TEL-%03d" % i, i * 10, "备注%d" % i])
wb.save(plain)

xlsm = os.path.join(DATA, "t3-macro.xlsm")
with zipfile.ZipFile(plain) as zin, zipfile.ZipFile(xlsm, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        zout.writestr(item, zin.read(item.filename))
    zout.writestr("xl/vbaProject.bin", b"\x00" * 1024)

before = md5(xlsm)
rc, txt, out = gen(xlsm, "from-macro")
check("xlsm 带宏：能正常读取", rc == 0, "rc=%d %s" % (rc, txt.strip()[:110]))
check("xlsm 带宏：源文件未被改动", md5(xlsm) == before,
      "%s -> %s" % (before[:10], md5(xlsm)[:10]))
if rc == 0:
    ws_o = load_workbook(out).active
    check("xlsm 带宏：数据完整（5 行 3 列）",
          ws_o.max_row == 6 and ws_o.max_column == 3,
          "%d 行 × %d 列" % (ws_o.max_row, ws_o.max_column))
    check("xlsm 带宏：首行数据正确", ws_o.cell(2, 1).value == "TEL-001",
          "A2=%r" % (ws_o.cell(2, 1).value,))
    # 产出是 .xlsx，宏不随之传递（设计如此，不是缺陷）
    with zipfile.ZipFile(out) as z:
        has_vba = any("vbaProject" in n for n in z.namelist())
    check("xlsm 带宏：产出为 .xlsx 且不含 vbaProject（设计决定）", not has_vba,
          "产出含 vbaProject" if has_vba else "宏未传递，符合预期")

# ============ 2. 外部链接工作簿 ============
# 在真实产物上注入 externalLink 结构
ext = os.path.join(DATA, "t3-extlink.xlsx")
with zipfile.ZipFile(plain) as z:
    parts = {n: z.read(n) for n in z.namelist()}

wbxml = parts["xl/workbook.xml"].decode("utf-8")
wbxml = wbxml.replace(
    "</workbook>",
    '<externalReferences>'
    '<externalReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'r:id="rIdExt1"/></externalReferences></workbook>')
parts["xl/workbook.xml"] = wbxml.encode("utf-8")

rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
rels = rels.replace(
    "</Relationships>",
    '<Relationship Id="rIdExt1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" '
    'Target="externalLinks/externalLink1.xml"/></Relationships>')
parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

ct = parts["[Content_Types].xml"].decode("utf-8")
ct = ct.replace(
    "</Types>",
    '<Override PartName="/xl/externalLinks/externalLink1.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>')
parts["[Content_Types].xml"] = ct.encode("utf-8")

parts["xl/externalLinks/externalLink1.xml"] = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">'
    '<sheetNames><sheetName val="Sheet1"/></sheetNames>'
    '<sheetDataSet><sheetData sheetId="0" refreshError="1"/></sheetDataSet>'
    '</externalBook></externalLink>').encode("utf-8")
parts["xl/externalLinks/_rels/externalLink1.xml.rels"] = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" '
    'Target="other-file.xlsx" TargetMode="External"/></Relationships>').encode("utf-8")

with zipfile.ZipFile(ext, "w", zipfile.ZIP_DEFLATED) as z:
    for n, data in parts.items():
        z.writestr(n, data)

# 先确认构造成功（openpyxl 能认到 external_links）
wb_chk = load_workbook(ext, data_only=True)
check("外部链接工作簿：构造成功且被 openpyxl 识别",
      len(wb_chk._external_links) == 1, "external_links=%d" % len(wb_chk._external_links))

before = md5(ext)
rc, txt, out = gen(ext, "from-extlink")
check("外部链接工作簿：能正常读取", rc == 0, "rc=%d %s" % (rc, txt.strip()[:110]))
check("外部链接工作簿：源文件未被改动（链接不被破坏）", md5(ext) == before,
      "%s -> %s" % (before[:10], md5(ext)[:10]))
if rc == 0:
    ws_o = load_workbook(out).active
    check("外部链接工作簿：数据完整", ws_o.max_row == 6,
          "max_row=%d" % ws_o.max_row)
    # 产出不应带外部链接（本技能只排版取值，不搬运链接关系）
    with zipfile.ZipFile(out) as z:
        has_ext = any("externalLink" in n for n in z.namelist())
    check("外部链接工作簿：产出不含 externalLink（只取值不搬链接）", not has_ext,
          "产出含 externalLink" if has_ext else "链接未传递，符合预期")

# ============ 3. 大数据量 ============
perf = {}
for n in (10000, 50000):
    src = os.path.join(DATA, "t3-big-%d.csv" % n)
    with open(src, "w", encoding="utf-8", newline="") as f:
        f.write("序号,型号,克重,门幅,成分,备注\n")
        for i in range(1, n + 1):
            f.write("%d,TEL-%06d,%d,138,%d%%莱赛尔,第%d行备注\n"
                    % (i, i, 100 + (i % 90), i % 100, i))
    out = os.path.join(OUT, "t3-big-%d.xlsx" % n)
    t0 = time.time()
    rc, txt = run([MT, "--profile", "business-table", "--src", src, "--out", out])
    dt = time.time() - t0
    if rc == 0:
        perf[n] = dt
        print("  %d 行: %.1f 秒, %.1f MB" % (n, dt, os.path.getsize(out) / 1024 / 1024))
    else:
        check("大数据量 %d 行生成" % n, False, txt.strip()[:100])

check("大数据量：1 万行在 30 秒内", perf.get(10000, 999) < 30,
      "%.1f 秒" % perf.get(10000, -1))
check("大数据量：5 万行在 120 秒内", perf.get(50000, 999) < 120,
      "%.1f 秒" % perf.get(50000, -1))
if 10000 in perf and 50000 in perf:
    ratio = perf[50000] / max(perf[10000], 0.001)
    check("大数据量：5 倍数据耗时不超 10 倍（线性）", ratio < 10, "比值 %.2f" % ratio)

# 大数据量产物也要能通过门禁
big_out = os.path.join(OUT, "t3-big-10000.xlsx")
if os.path.exists(big_out):
    rc, txt = run([VT, "--profile", "business-table", big_out])
    check("大数据量：1 万行产物 verify PASS", rc == 0,
          txt.strip().splitlines()[-1] if txt.strip() else "")

print("=" * 84)
print("第三轮补测：xlsm / 外部链接 / 大数据量")
print("=" * 84)
npass = 0
for ok, name, detail in results:
    if ok:
        npass += 1
    print("[%s] %-46s %s" % ("OK  " if ok else "FAIL", name, detail.replace("\n", " ")[:88]))
print("-" * 84)
print("ROUND3: %d/%d PASS" % (npass, len(results)))
sys.exit(0 if npass == len(results) else 1)
