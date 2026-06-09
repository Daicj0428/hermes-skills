---
name: alarm-statistics-analysis
description: "告警数据清洗 + 多维度统计 + 交互式 HTML 报告生成工具"
version: 1.0.0
author: Hermes Agent (from user project)
license: MIT
dependencies: [python3, pandas, openpyxl, matplotlib, tqdm]
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [alerting, statistics, HTML, report, data-cleaning, excel, monitoring]
    related_skills: []
---

# 告警统计分析工具

对原始告警 Excel 数据进行清洗和多维度统计分析，生成**单文件交互式 HTML 报告**。报告无需后端服务，浏览器直接打开即可使用。

## 触发条件

- 用户需要对告警 Excel 数据做统计分析
- 用户说"告警报表"、"告警分析"、"告警统计"、"告警清洗"
- 用户有监控告警数据需要生成可视化报告

## 功能概览

| 模块 | 说明 |
|------|------|
| **摘要卡片** | 告警记录总数、未超时处理率、覆盖业务系统数（随筛选联动） |
| **多选筛选器** | 支持系统、周期、告警级别三维度组合筛选 |
| **动态图表** | 树状图（单周期）、系统×周期小多组柱状图、堆叠柱状图（按报警级别分层） |
| **未超时处理率** | 各系统未超时处理率柱状图 |
| **TOP 告警** | 按处理时长降序 TOP N（默认 20），支持折叠和自定义字段 CSV 导出 |
| **明细表格** | 分页排序、搜索、筛选联动 |
| **自定义导出** | 弹出面板勾选字段，导出筛选后的 CSV |

## 工作流程

### 标准流程（一键执行）

```bash
cd <项目目录>
python main.py
```

自动完成三步骤：**数据清洗 → 生成 HTML 报告 → 清理中间产物**

报告输出到 `statistics/告警统计分析报告_<时间戳>.html`。

### 分步执行

```bash
python main.py --clean-only                          # 仅清洗数据
python main.py --skip-clean                          # 跳过清洗，仅生成报告
```

### 自定义统计周期

```bash
python main.py --custom-periods "2025-07-01 - 2025-07-06,2025-07-07 - 2025-07-13"
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
├── main.py                              # 主流程入口
├── config.py                            # 全局配置 + 字段映射
├── data_cleaner.py                      # 数据清洗（按周分 sheet）
├── create_combined_weekly_statistics_with_chinese_names.py  # 统计 + HTML 报告
├── requirements.txt                     # 依赖
├── <原始告警数据>.xlsx                   # 输入数据（文件名在 config.py 配置）
├── statistics/                          # 输出目录（运行后仅保留 HTML）
│   └── 告警统计分析报告_<timestamp>.html
└── logs/                                # 日志（自动保留最新 3 个）
```

## 注意事项

1. HTML 报告通过 CDN 加载 Chart.js，浏览器需联网打开
2. 报告内嵌全量原始数据为 JSON（约 50-60 MB），首次加载可能较慢
3. `main.py` 运行完毕自动清理中间产物（`cleaned_data.xlsx`、`temp/`、合并 Excel）
4. `config.py` 启动时检查 `INPUT_FILE` 是否存在
5. 新增系统/修改导出字段只需编辑 `config.py`，无需改其他文件

## 源文件

完整源码位于 `references/` 目录：

- `references/main.py` — 主流程（清洗 → 报告 → 清理）
- `references/config.py` — 全局配置模版
- `references/data_cleaner.py` — 数据清洗逻辑
- `references/create_combined_weekly_statistics_with_chinese_names.py` — 统计 + HTML 报告生成器
- `references/requirements.txt` — Python 依赖清单
