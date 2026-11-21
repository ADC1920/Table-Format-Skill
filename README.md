# Table Format Skill

中文表格排版技能：把 xlsx / xls / md / CSV 表格按**指定的表格类型**重排为规范 xlsx 工作簿。

适用于「做表格、美化表格、排版表格、重新排版表格、表格规范」等任何表格产出请求。
文件操作与数据加工仍走 xlsx 技能，本技能只负责规格与交付标准。

## 核心特性

**v2.0 起无默认样式**：必须按表格类型选定 profile，不指定则脚本报错退出并列出可选类型。
这样避免了「判错类型仍静默套一套规格」的问题——把学位论文的表按商务表深蓝表头排，
比多问一句的代价大得多。

### 8 种表格类型

| profile | 适用 | 关键特征 |
|---|---|---|
| `business-table` | 内部管理表、台账、运营采购项目表 | 深蓝表头 #1F3864 + 隔行底纹 #F7F9FD + 黑体（v1.x 原样式，选它可复现历史交付） |
| `academic-table` | 学位论文、期刊论文、科技报告 | 三线表，无竖线无底色无底纹，宋体 10.5，不冻结，竖版 A4 |
| `finance-table` | 四大报表、附注、管理报表、预算 | 浅蓝表头黑字，黑线框，无底纹，数字右对齐千分位，负数括号，零值「-」，末行合计加粗 |
| `stats-table` | 统计公报、年报、普查、调查汇总 | 浅灰表头，三段式表头，合计行加粗，横版 |
| `gov-table` | 公文附表、政务公开表、预算决算 | 仿宋正文，浅灰表头，黑线框，受 GB/T 9704 版心约束 |
| `quality-table` | 检验记录、抽样、FMEA、实验室报告 | 浅灰表头，黑线框，行高 22，打印网格线 |
| `hr-table` | 花名册、考勤、绩效、薪酬、招聘 | 浅蓝表头，浅灰隔行底纹，敏感字段标密级与保管期限 |
| `data-list` | 数据字典、字段清单、主数据、BOM | 左对齐表头，浅灰细框，无底纹，等线/Consolas，禁数据区合并 |

### 规范库

`references/standards/` 收录六域 395 条表格规范，来自国标、部委规章与行业规范调研：

| 领域 | 条数 | 内容 |
|---|---|---|
| stats（统计与国标） | 120 | GB/T 1.1-2020 表格编排、统计表式、GB/T 15835 数字用法、续表接排、三线表、CY/T 170 |
| finance（财务与披露） | 55 | 企业会计准则第 30 号、证监会年报准则、四大报表格式、附注、数字格式 |
| academic（学术与出版） | 55 | GB/T 7713 系列、三线表、表号表题、量与单位、英文期刊体例 |
| gov（公文与政务） | 55 | GB/T 9704-2012 公文附表、政府统计报表、预算决算公开表、办公表单 |
| quality（质量与工程） | 55 | ISO 9001 / IATF 16949、GB/T 2828.1 抽样、CNAS-CL01、GB 50300 验收 |
| hr-project（人事与项目运营） | 55 | 人事档案表、考勤绩效、项目管理、销售运营、采购供应链、数据整备 |

三层形态：`index.md`（一行式索引）、`domain-<领域>.md`（分域详表）、`standards.jsonl`（机器可查）。

## 安装

把本目录放到技能目录下（ZCode / Claude Code 的 `skills/`），或作为插件内容。依赖：

```bash
pip install openpyxl pyyaml
```

渲染目检另需 LibreOffice（`soffice`）与 pymupdf；不做渲染时可跳过。

## 用法

```bash
# 1. 生成
python scripts/make_table.py --profile finance-table --src 数据.csv --out 成品.xlsx \
    --title "2026 年管理费用明细" --note "单位：万元"

# 2. 校验（退出码 0 才可交付）
python scripts/verify_table.py --profile finance-table 成品.xlsx

# 3. 交付形态门禁
python scripts/check_delivery.py 成品.xlsx

# 4. 渲染目检（可选）
python scripts/render_pages.py 成品.xlsx --pages 1
```

列出全部类型：

```bash
python scripts/make_table.py --list-profiles
```

### 规范库检索

```bash
python scripts/query_standards.py --list-domains              # 领域与条数
python scripts/query_standards.py --domain academic           # 某领域全部
python scripts/query_standards.py --keyword 三线表            # 关键词
python scripts/query_standards.py --profile academic-table    # 某类型关联条目
python scripts/query_standards.py --domain quality --min-conf high --full
```

置信度三档：`high` 已核到官方原文或标准正文；`medium` 行业通行做法或权威二手资料；
`low` 仅见转述、口径存疑。**low 不作为排版依据。**

## 自检与回归

```bash
python scripts/make_table.py --selftest            # 10 项：样式、三线表、无 profile 报错
python scripts/verify_table.py --selftest          # 8 项：正例 + 6 反例
python scripts/run_profile_regression.py           # 8 profile × 9 负载形状 = 72 项
python scripts/check_skill_consistency.py          # 7 项：文档实物一致性
```

## 已知限制

- 表格内的公式计算不由本技能负责：`make_table.py` 只排版，不写公式也不重算。
  需要公式与重算走 xlsx 技能的 `recalc.py`，之后再用本技能校验版式。
- 图表、条件格式、数据验证、透视表不在本技能范围。
- 多工作表工作簿只处理一个表（默认第一个，可用 `--src-sheet` 指定）。
- `data-list` 的「一维结构、禁多值单元格、禁空行」是规格要求，脚本不自动拆分多值单元格。
- `gov-table` 的版心与公文页码位置只在打印设置层面近似，发文机关标志、印章不生成。
- 三线表的磅值（1.5 / 0.75 / 1.5）是 CY/T 170-2019「表框线用粗线、其他用细线」的通行落地值，
  **不是国标强制数值**，须以目标期刊或学校模板为准。
- 规范库 `low` 置信度条目来自二手转述，引用前建议 `--min-conf high` 过滤。

## 版本

当前 2.0.0。变更记录见 [CHANGELOG.md](CHANGELOG.md)。
