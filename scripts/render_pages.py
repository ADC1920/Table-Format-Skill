# -*- coding: utf-8 -*-
"""render_pages.py — table-format 渲染目检通道（xlsx → PNG）。

先经 soffice 转 PDF，再用 pymupdf 渲染每页为 PNG，供交付前目检
（查截断、错位、配色，SKILL.md「规格条文」最后一条）。本脚本只出图与清单，
不做判定；判定由使用者目检或项目门禁派 visual-judge 承担。

用法：
    python render_pages.py 成品.xlsx [--out-dir DIR] [--dpi 130] [--pages 1-3,5]
    python render_pages.py --selftest

输出默认落在 %TEMP%\\table_render\\<文件名>\\，不污染成品目录。
"""

import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile

SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def find_soffice(explicit=None):
    if explicit:
        return explicit if os.path.isfile(explicit) else None
    w = shutil.which("soffice")
    if w:
        return w
    for p in SOFFICE_CANDIDATES:
        if os.path.isfile(p):
            return p
    return None


def to_pdf(xlsx, workdir, soffice):
    """soffice --headless 转 PDF。独立 UserInstallation，避免与运行中的实例冲突。"""
    profile = os.path.join(workdir, "soffice_profile")
    os.makedirs(profile, exist_ok=True)
    profile_uri = "file:///" + profile.replace("\\", "/")
    cmd = [soffice, "--headless", "--norestore",
           "-env:UserInstallation=" + profile_uri,
           "--convert-to", "pdf", "--outdir", workdir, xlsx]
    # 中文 Windows 下 soffice 控制台输出为 GBK 字节，text=True(UTF-8) 会在读取线程抛
    # UnicodeDecodeError；显式按 GBK 容错解码
    proc = subprocess.run(cmd, capture_output=True, encoding="gbk", errors="replace",
                          timeout=180)
    stem = os.path.splitext(os.path.basename(xlsx))[0]
    pdf = os.path.join(workdir, stem + ".pdf")
    if proc.returncode != 0 or not os.path.isfile(pdf):
        raise SystemExit("ERROR: soffice 转换失败（exit=%s）\nstdout: %s\nstderr: %s"
                         % (proc.returncode, proc.stdout.strip()[:500], proc.stderr.strip()[:500]))
    return pdf


def parse_pages(spec, total):
    """'1-3,5' → [1,2,3,5]；None → 全部页。1-based。"""
    if not spec:
        return list(range(1, total + 1))
    out = []
    for part in spec.split(","):
        part = part.strip()
        m = re.fullmatch(r"(\d+)-(\d+)", part)
        if m:
            out.extend(range(int(m.group(1)), int(m.group(2)) + 1))
        elif part.isdigit():
            out.append(int(part))
        else:
            raise SystemExit("ERROR: --pages 片段无法解析: %r" % part)
    return [p for p in out if 1 <= p <= total]


def render(xlsx, out_dir=None, dpi=130, pages=None, soffice=None):
    if not os.path.isfile(xlsx):
        raise SystemExit("ERROR: 文件不存在 %s" % xlsx)
    exe = find_soffice(soffice)
    if not exe:
        raise SystemExit("ERROR: 未找到 soffice（LibreOffice）。用 --soffice 指定完整路径。")

    stem = os.path.splitext(os.path.basename(xlsx))[0]
    if not out_dir:
        out_dir = os.path.join(tempfile.gettempdir(), "table_render", stem)
    os.makedirs(out_dir, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="table_render_") as workdir:
        pdf = to_pdf(xlsx, workdir, exe)
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        doc = fitz.open(pdf)
        sel = parse_pages(pages, doc.page_count)
        if not sel:
            raise SystemExit("ERROR: --pages 未命中任何页（共 %d 页）" % doc.page_count)
        zoom = dpi / 72.0
        written = []
        for pno in sel:
            page = doc[pno - 1]
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            fp = os.path.join(out_dir, "%s_p%02d.png" % (stem, pno))
            pix.save(fp)
            written.append(fp)
        total = doc.page_count
        doc.close()

    print("源文件: %s（共 %d 页，输出 %d 张）" % (xlsx, total, len(written)))
    print("输出目录: %s" % out_dir)
    for fp in written[:10]:
        print("  %s" % fp)
    if len(written) > 10:
        print("  ...（其余 %d 张）" % (len(written) - 10))
    print("目检清单：①文字截断 ②排版错位 ③配色是否符合规格（表头 #1F3864 / 隔行 #F7F9FD）")
    return 0


def selftest():
    import make_table as mt  # 兄弟脚本
    if not find_soffice():
        print("SELFTEST: SKIP（未找到 soffice）")
        return 0
    tmp = tempfile.mkdtemp(prefix="table_format_render_selftest_")
    xlsx = os.path.join(tmp, "render_sample.xlsx")
    headers = ["编码", "名称", "价格"]
    rows = [["QT-%03d" % i, "作品%d" % i, i * 100] for i in range(1, 8)]
    mt.build_workbook(headers, rows, xlsx, "表", title="渲染自检")
    out_dir = os.path.join(tmp, "out")
    rc = render(xlsx, out_dir=out_dir, dpi=100)
    pngs = glob.glob(os.path.join(out_dir, "*.png"))
    ok = rc == 0 and len(pngs) >= 1 and os.path.getsize(pngs[0]) > 1000
    print("SELFTEST:", "PASS" if ok else "FAIL", "| png=%d size=%s"
          % (len(pngs), os.path.getsize(pngs[0]) if pngs else 0))
    return 0 if ok else 1


def main():
    ap = argparse.ArgumentParser(description="table-format 渲染目检通道（xlsx → PNG）")
    ap.add_argument("xlsx", nargs="?", help="待渲染的 xlsx")
    ap.add_argument("--out-dir", help="PNG 输出目录（默认 %%TEMP%%\\table_render\\<文件名>）")
    ap.add_argument("--dpi", type=int, default=130)
    ap.add_argument("--pages", help="页范围，如 1-3,5（默认全部）")
    ap.add_argument("--soffice", help="soffice 可执行文件路径")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()
    if a.selftest:
        sys.exit(selftest())
    if not a.xlsx:
        ap.error("需要 xlsx 路径（或使用 --selftest）")
    sys.exit(render(a.xlsx, a.out_dir, a.dpi, a.pages, a.soffice))


if __name__ == "__main__":
    main()
