# -*- coding: utf-8 -*-
"""tests/_env.py — 测试脚本的公共路径与工具。

技能版本不写死本机绝对路径：技能根目录由本文件位置推导，
临时产物目录读环境变量 TF_TEST_WORKDIR，缺省落到系统临时目录。
"""
import os
import subprocess
import sys
import tempfile

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(TESTS_DIR)
SCRIPTS = os.path.join(SKILL, "scripts")
PROFILES = os.path.join(SKILL, "profiles")

# 临时目录兜底用系统 temp。早先用 sys.executable 所在目录推导，
# 便携版/Program Files 安装下会落到只读位置并抛 PermissionError
# （2026-11-21 第二轮测试实测：单独跑某个套件即复现，总入口因传了
#  TF_TEST_WORKDIR 而掩盖了它）。
BASE = os.environ.get("TF_TEST_WORKDIR") or os.path.join(
    tempfile.gettempdir(), "table-format-tests")
DATA = os.path.join(BASE, "data")
OUT = os.path.join(BASE, "out")
PNG = os.path.join(BASE, "png")

for d in (BASE, DATA, OUT, PNG):
    os.makedirs(d, exist_ok=True)

MT = os.path.join(SCRIPTS, "make_table.py")
VT = os.path.join(SCRIPTS, "verify_table.py")
CD = os.path.join(SCRIPTS, "check_delivery.py")
RP = os.path.join(SCRIPTS, "render_pages.py")
QS = os.path.join(SCRIPTS, "query_standards.py")


def run(args, env_extra=None, clear_profile_env=True, cwd=None):
    """跑子进程。默认清掉 TABLE_FORMAT_* 避免外部环境影响测试。"""
    env = dict(os.environ)
    if clear_profile_env:
        env.pop("TABLE_FORMAT_PROFILE", None)
        env.pop("TABLE_FORMAT_SPEC", None)
    if env_extra:
        env.update(env_extra)
    r = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env=env, cwd=cwd or SKILL)
    return r.returncode, (r.stdout or "") + (r.stderr or "")
