---
name: alarm-statistics-analysis
description: "告警数据清洗 + 多维度统计 + 交互式 HTML 报告生成工具"
version: 1.4.2
author: Hermes Agent (from user project)
license: MIT
dependencies: [python3, pandas, openpyxl, matplotlib, tqdm]
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [alerting, statistics, HTML, report, data-cleaning, excel, monitoring, multi-file]
    related_skills: []
---

# 告警统计分析工具

对原始告警 Excel 数据进行清洗和多维度统计分析，生成**单文件交互式 HTML 报告**。支持单文件和多文件对比分析。报告无需后端服务，浏览器直接打开即可使用。

## 触发条件

- 用户需要对告警 Excel 数据做统计分析
- 用户说"告警报表"、"告警分析"、"告警统计"、"告警清洗"
- 用户有监控告警数据需要生成可视化报告
- 用户需要对比多个月份/来源的告警数据

## 功能概览

| 模块 | 说明 |
|------|------|
| **摘要卡片** | 告警记录总数、未超时处理率、覆盖业务系统数（随筛选联动） |
| **多选筛选器** | 支持系统、周期、告警级别、**文件来源**四维度组合筛选 |
| **动态图表** | 树状图（单周期）、系统×周期小多组柱状图、堆叠柱状图（按报警级别分层） |
| **未超时处理率** | 各系统未超时处理率柱状图 |
| **TOP 告警** | 按处理时长降序 TOP N（默认 20），支持折叠和自定义字段 CSV 导出 |
| **明细表格** | 分页排序、搜索、筛选联动 |
| **自定义导出** | 弹出面板勾选字段，导出筛选后的 CSV |
| **🆕 多文件对比** | 同时分析多个月份/来源的告警数据，统一报告内对比 |

## 工作流程

### 标准流程（一键执行 — 🆕 推荐）

```bash
cd <项目目录>

# 单文件分析
python run_local.py --file 7月份全量告警明细0801.xlsx

# 多文件对比分析
python run_local.py --file 7月.xlsx --file 9月.xlsx --labels "7月告警,9月告警"
```

自动完成四步骤：**数据清洗 → 生成 HTML 报告 → 清理中间产物**。

报告输出到 `statistics/`：完整版 `告警统计分析报告_<ts>.html` + 精简版 `告警统计分析报告_<ts>_lite.html`。

### 自定义统计周期

```bash
python run_local.py --file data.xlsx --periods "2025-07-01 - 2025-07-31,2025-08-01 - 2025-08-31"
```

### 自定义输出文件名

```bash
python run_local.py --file data.xlsx --output my_report.html
```

### 旧版入口（兼容）

```bash
python main.py                          # 单文件（需手动配置 config.py INPUT_FILE）
python main.py --clean-only             # 仅清洗数据
python main.py --skip-clean             # 跳过清洗，仅生成报告
```

## 环境要求

```bash
pip install -r requirements.txt
```

依赖：`numpy`, `pandas`, `matplotlib`, `openpyxl`, `tqdm`

## 配置要点（config.py）

### 系统名称映射
```python
ATTR_CHINESE_MAPPING = {
    'crm': 'CRM系统',
    'bossv8': 'BOSS系统',
    # ... 新增系统只需在此添加一行
}
```

### 告警级别映射
```python
LEVEL_CHINESE_MAPPING = {
    'serious': '严重告警', 'critical': '严重告警',
    'important': '重要告警', 'major': '重要告警',
    # 自动将英文级别转为中文显示
}
```

### 切换数据文件

修改 `config.py` 中的 `INPUT_FILE`，指向你的 Excel 文件即可：
```python
INPUT_FILE = "your_alert_data.xlsx"   # 替换为实际文件名
```
程序启动时会自动检查文件是否存在。

### 自定义导出字段
```python
EXPORT_FIELDS = [
    {'key': 'attr', 'label': '所属系统'},
    {'key': 'alarm_content', 'label': '告警内容'},
    # 修改此列表即可调整前端导出弹窗选项
]
```

## 数据格式要求

原始 Excel 需包含以下关键列：

| 列名 | 说明 |
|------|------|
| `attr` | 所属系统标识 |
| `GRADE` | 告警级别（自动转大写后映射中文） |
| `alarm_first_time` | 首次告警时间（清洗时以此为周期分组） |
| `alarm_content` | 告警内容 |
| `deal_status` | 处理状态 |
| `deal_time` | 处理时长 |
| `RESOURCE_NAME` | 资源对象 |
| `ALARM_COUNT` | 告警次数（可选，默认每条为 1） |

## 项目结构

```
alarm_alter/
├── run_local.py                       # 🆕 主入口（独立执行，支持多文件）
├── main.py                            # 旧版入口（兼容，单文件模式）
├── config.py                          # 全局配置 + 字段映射
├── data_cleaner.py                    # 数据清洗（支持 output_prefix）
├── report_generator.py                # 🆕 统计合并 + HTML 报告（支持多文件对比）
├── requirements.txt                   # 依赖
├── <原始告警数据>.xlsx                 # 输入数据
├── statistics/                        # 输出目录（运行后仅保留 HTML）
│   └── 告警统计分析报告_<timestamp>.html
├── logs/                              # 日志（自动保留最新 5 个）
└── temp/                              # 临时目录（自动清理）
```

## 注意事项

1. HTML 报告通过 CDN 加载 Chart.js，浏览器需联网打开
2. 报告内嵌全量原始数据为 JSON（约 50-60 MB），首次加载可能较慢
3. `main.py` 运行完毕自动清理中间产物（`cleaned_data.xlsx`、`temp/`、合并 Excel）
4. `config.py` 启动时检查 `INPUT_FILE` 是否存在
5. 新增系统/修改导出字段只需编辑 `config.py`，无需改其他文件
6. **修改 `report_generator.py` 时慎用 `patch` 工具**：该文件包含两套近乎相同的 HTML/JS 模板，fuzzy 匹配可能误判。删除/修改大段 HTML 时用 `sed` 代替。详见 `references/editing-report-generator.md`。

## 首次执行注意事项

首次在新环境或新数据文件上运行本工具时，请逐项检查以下配置：

### 1. 确认数据列的实际情况（最重要）

不同来源的告警 Excel 列名和值格式可能不同。**不要假设默认配置能直接匹配你的数据**。

**检查方法**：
```bash
python3 -c "
import pandas as pd
df = pd.read_excel('你的文件.xlsx')
print('=== 列名 ==='); print(df.columns.tolist())
print('=== attr 唯一值 ==='); print(df['attr'].dropna().unique()[:20])
print('=== GRADE 唯一值 (前10) ==='); print(df['GRADE'].dropna().unique()[:10])
print('=== deal_status 唯一值 ==='); print(df['deal_status'].dropna().unique())
"
```

### 2. 调整 config.py 的映射表

根据第 1 步的检查结果调整映射：

| 映射 | 场景 | 处理方式 |
|------|------|---------|
| **ATTR_CHINESE_MAPPING** | 数据中有未出现的系统代码 | 添加新条目 |
| **ATTR_CHINESE_MAPPING** | 数据中代码与映射不匹配 | 修正 key 以匹配实际值 |
| **LEVEL_CHINESE_MAPPING** | `GRADE` 列是中文（严重/重要/一般） | 映射表不起作用（中文直接保留），无需修改 |
| **LEVEL_CHINESE_MAPPING** | `GRADE` 列是英文（serious/important/…） | 确保所有出现的英文值都有对应映射 |

> **注意**：`GRADE` 列值如果是中文（如 `严重`、`重要`、`一般`），`LEVEL_CHINESE_MAPPING` 不会生效（中文值直接原样保留），这是正常的。

### 3. 检查必需列是否存在

程序依赖以下列名（区分大小写）：

| 必需列 | 缺失时的后果 |
|--------|-------------|
| `alarm_first_time` | ⚠️ **报错退出** — 清洗阶段会丢弃无时间记录的行，若整列为空则无数据可处理 |
| `attr` | 所有系统显示为原值（不会报错，但无中文名） |
| `GRADE` | 告警级别为空，筛选器无效 |
| `deal_status` | 已处理/未超时率统计失效（默认显示"未处理"） |
| `ALARM_COUNT` | 每条记录计为 1（程序自动兜底） |

### 4. Python 环境准备

```bash
# 推荐：创建虚拟环境
python3 -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt
```

### 5. pandas 3.x 兼容性

如果你使用的 pandas 版本 ≥ 3.0，首次运行可能遇到 `TypeError: object of type 'float' has no len()`。

→ 快速修复：运行 `references/pandas-compat-fix.md` 中的一键替换命令。

### 6. 首次运行建议分步执行

```bash
# 第 1 步：仅清洗（验证数据能正常处理）
python main.py --clean-only

# 检查 statistics/cleaned_data.xlsx 是否生成成功、内容正确

# 第 2 步：生成报告
python main.py --skip-clean
```

分步执行能快速定位问题出在清洗阶段还是报告生成阶段。

### 7. HTML 报告依赖 CDN

报告的图表通过 Chart.js CDN 加载 (`cdn.jsdelivr.net`)，浏览器打开报告时需要联网。内网环境需提前下载 Chart.js 放到本地并修改 HTML 模板中的 `<script>` 标签。

### 8. 大文件处理

报告将全量清洗后数据内嵌为 JSON（约 50–60 MB），首次在浏览器打开可能耗时 10–30 秒，属于正常现象。

### 9. 自定义统计周期

如果数据跨越多个自然周但只想统计特定区间：
```bash
python main.py --custom-periods "2025-07-01 - 2025-07-13,2025-08-01 - 2025-08-07"
```

### 10. 快速排错清单

| 现象 | 优先检查 |
|------|---------|
| `FileNotFoundError` | `config.py` 中 `INPUT_FILE` 是否正确；文件是否在项目目录下 |
| 清洗后无数据 | `alarm_first_time` 列是否为空；自定义周期是否覆盖了数据时间范围 |
| 系统名显示英文缩写 | `ATTR_CHINESE_MAPPING` 中缺少对应 key |
| 告警级别筛选为空 | `GRADE` 列值格式与预期不符（中文/英文/混合） |
| `map(len)` 报错 | pandas 版本 ≥ 3.0，按第 5 条修复 |
| HTML 图表不显示 | 浏览器是否联网；Chart.js CDN 是否可达 |

### pandas 3.x 兼容性问题

项目原始依赖 `pandas==1.5.3`。如果环境安装了 pandas 3.x，`.astype(str).map(len).max()` 模式会因浮点 NaN 值抛出 `TypeError: object of type 'float' has no len()`。

**修复方法**：全局替换为 `.fillna('').astype(str).str.len().max()`。

> **v1.3.0 注意**：旧版 6 个独立脚本已全部合并到 `report_generator.py`。如果该文件仍含 `map(len)` 模式，对 `report_generator.py` 执行同样的替换即可。

如果安装后运行报 `map(len)` 相关错误，按上述模式替换。详见 `references/pandas-compat-fix.md`。

## 源文件

完整源码位于 `references/` 目录：

- `references/run_local.py` — 🆕 主入口（独立执行，支持 `--file` 多文件）
- `references/main.py` — 旧版入口（兼容）
- `references/config.py` — 全局配置模版
- `references/data_cleaner.py` — 数据清洗逻辑（支持 output_prefix）
- `references/report_generator.py` — 🆕 统计 + HTML 报告生成器（支持多文件对比）
- `references/requirements.txt` — Python 依赖清单
- `references/pandas-compat-fix.md` — pandas 3.x 兼容性修复指南
- `references/editing-report-generator.md` — 修改 report_generator.py 时的注意事项
