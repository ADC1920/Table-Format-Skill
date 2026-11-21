# -*- coding: utf-8 -*-
"""表格排版规格配置加载器（table-format 技能内部模块）。

规格唯一权威源：技能根目录 `table-format.yaml`（内部基线，字段继承来源）
+ `profiles/<类型>.yaml`（各界面的差异字段）。

profile 体系（v2.0 起）
--------------------------------
表格没有「通用默认样式」——统计表、财务报表、学术三线表、政务附表的
版式互相冲突（三线表不画竖线，财务报表不要隔行底纹，政务附表用仿宋）。
所有入口必须显式指定表格类型，未指定时抛 ProfileNotSpecified。

选择 profile 的优先级（高→低）：
    1. 环境变量 TABLE_FORMAT_SPEC（绝对路径，临时套用任意规格文件）
    2. 参数 profile=... 或环境变量 TABLE_FORMAT_PROFILE（类型名，如 finance-table）
    3. 无 → 抛 ProfileNotSpecified（不再静默套基线）

用法：
    from spec_loader import load_spec
    spec = load_spec(profile="finance-table")
    spec["header"]["fill"]        # '1F3864'
"""
import os

SPEC_FILENAME = "table-format.yaml"
PROFILE_DIRNAME = "profiles"
DEFAULT_PROFILE = "default"
ENV_PROFILE = "TABLE_FORMAT_PROFILE"
ENV_SPEC_PATH = "TABLE_FORMAT_SPEC"

# 内置基线：必须与 table-format.yaml 保持一致（配置文件缺失时的兜底）。
# 这是 profile 的字段继承来源，不是可选表格类型。
DEFAULT_SPEC = {
    "version": "2.0.0",
    "font": {
        "name": "黑体",
        "western": "Arial",
        "data_size": 10,
        "data_color": "262626",
        "note_color": "595959",
    },
    "header": {
        "fill": "1F3864",
        "font_color": "FFFFFF",
        "font_name": "黑体",
        "size": 12,
        "bold": True,
        "align": "center",
        "valign": "center",
        "wrap": True,
        "border_style": "thin",
        "border_color": "B4C6E7",
    },
    "title": {
        "fill": "D9E2F3",
        "font_color": "1F3864",
        "size": 16,
        "bold": True,
        "align": "center",
        "merge": True,
        "border": False,
    },
    "note": {
        "size": 10,
        "color": "595959",
        "align": "left",
        "merge": True,
    },
    "body": {
        "align": "auto",          # auto=数字右对齐文本左对齐；center/left/right 强制
        "border_style": "thin",
        "border_color": "B4C6E7",
        "zebra": True,
        "zebra_color": "F7F9FD",
        "zebra_start": 2,         # 自第 N 个数据行起着色（1 基）
        "total_row_bold": False,  # 末行（合计行）加粗
    },
    # 三线表专用：academic-table 等置 true
    "rules": {
        "three_line": False,
        "top_pt": 1.5,
        "header_pt": 0.5,
        "bottom_pt": 1.5,
        "vertical": True,         # 是否画竖线
    },
    "row_height": {
        "min_data": 20,
        "title_base": 28,
        "title_line": 21,
        "note_base": 18,
        "head_base": 24,
        "line": 14.5,
    },
    "col_width": {
        "min": 8,
        "max": 52,
        "pad": 3,
        "head_factor": 1.6,
        "text_cap": 60,
    },
    "freeze": {
        "enabled": "auto",        # auto=数据行 >= min_rows 才冻；always；never
        "min_rows": 20,
        "anchor": "auto",         # auto=表头末行下一行首列；也可写 A3 / B3
    },
    "print": {
        "orientation": "landscape",
        "paper": "A4",
        "title_rows": "auto",
        "margin_cm": None,
        "gridlines": False,
        "page_numbers": None,
    },
    "number": {
        "format_auto": False,     # 财务/统计表开启：整数千分位、负数括号
        "align": "right",
        "zero_as_dash": False,
    },
}


def _skill_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _profile_dir():
    return os.path.join(_skill_root(), PROFILE_DIRNAME)


def list_profiles():
    """列出可用表格类型 profile 名（按字母序）。不含内部基线 default。"""
    names = []
    pdir = _profile_dir()
    if os.path.isdir(pdir):
        for fn in sorted(os.listdir(pdir)):
            if fn.endswith((".yaml", ".yml")):
                names.append(os.path.splitext(fn)[0])
    return names


class ProfileNotSpecified(Exception):
    """未指定表格类型时抛出——本技能不使用默认样式，必须按表格类型选。"""


class ProfileNotFound(Exception):
    """指定了 profile 但找不到对应 yaml（类型名写错）时抛出。"""


def _profile_path(profile):
    if profile in (None, "", DEFAULT_PROFILE):
        return os.path.join(_skill_root(), SPEC_FILENAME)
    for ext in (".yaml", ".yml"):
        candidate = os.path.join(_profile_dir(), profile + ext)
        if os.path.exists(candidate):
            return candidate
    return os.path.join(_profile_dir(), profile + ".yaml")


def resolve_profile(profile=None, require=True):
    """返回本次实际生效的 profile 名（用于日志与交付说明）。"""
    if os.environ.get(ENV_SPEC_PATH):
        return "custom:" + os.environ[ENV_SPEC_PATH]
    chosen = profile or os.environ.get(ENV_PROFILE)
    if not chosen:
        if require:
            raise ProfileNotSpecified(
                "未指定表格类型 profile。本技能不使用默认样式，必须按表格类型选择：\n  "
                + "\n  ".join(list_profiles())
                + "\n用法：--profile <类型名>，或设环境变量 TABLE_FORMAT_PROFILE。")
        return DEFAULT_PROFILE
    return chosen


def _spec_path(profile=None, require=True):
    override = os.environ.get(ENV_SPEC_PATH)
    if override:
        return override
    chosen = resolve_profile(profile, require=require)
    if chosen == DEFAULT_PROFILE:
        return os.path.join(_skill_root(), SPEC_FILENAME)
    return _profile_path(chosen)


def _deep_merge(base, override):
    """递归合并：override 里出现的键覆盖 base，未出现的键保留 base 默认值。

    注意 None 的处理：键**存在**且值为 null 表示「显式清空」（如三线表把
    header.fill 置 null 以关掉底色），必须覆盖生效；只有键**不存在**才继承。
    早先写成 `elif value is not None` 会让 profile 里的 null 静默失效——
    三线表照旧带深蓝表头，且验证脚本按同一个 spec 断言也会一起误判为通过。
    """
    result = dict(base)
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_spec(verbose=False, profile=None, require=True):
    """加载指定表格类型的完整规格（与 DEFAULT_SPEC 深合并）。

    profile=None 时读环境变量 TABLE_FORMAT_PROFILE；两者都没有且 require=True
    时抛 ProfileNotSpecified。require=False 仅供内部工具读基线字段。
    """
    path = _spec_path(profile, require=require)

    try:
        import yaml
    except ImportError:
        if verbose:
            print("[提示] 未安装 PyYAML，使用内置基线规格（%s 未加载）" % path)
        return dict(DEFAULT_SPEC)

    if not os.path.exists(path):
        chosen = resolve_profile(profile, require=False)
        if require and chosen not in (None, "", DEFAULT_PROFILE):
            raise ProfileNotFound(
                "未找到表格类型 profile「%s」：%s\n可用：\n  " % (chosen, path)
                + "\n  ".join(list_profiles()))
        if verbose:
            print("[提示] 未找到 %s，使用内置基线规格" % path)
        return dict(DEFAULT_SPEC)

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
    except Exception as exc:
        if verbose:
            print("[警告] 读取 %s 失败（%s），使用内置基线规格" % (path, exc))
        return dict(DEFAULT_SPEC)

    return _deep_merge(DEFAULT_SPEC, data)


def profile_meta(profile=None):
    """读 profile 自带的说明元信息（name/title/standard/note），用于交付说明。"""
    path = _spec_path(profile, require=False)
    try:
        import yaml
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        return data.get("profile") or {}
    except Exception:
        return {}


if __name__ == "__main__":
    import json
    import sys

    if "--list" in sys.argv:
        for name in list_profiles():
            meta = profile_meta(name)
            print("%-16s %s" % (name, meta.get("title", "")))
        raise SystemExit(0)

    prof = None
    if "--profile" in sys.argv:
        idx = sys.argv.index("--profile")
        if idx + 1 < len(sys.argv):
            prof = sys.argv[idx + 1]
    spec = load_spec(verbose=True, profile=prof)
    print("profile: %s" % resolve_profile(prof))
    print(json.dumps(spec, ensure_ascii=False, indent=2))
