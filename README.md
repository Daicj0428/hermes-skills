# Hermes Skills

Skills and scripts for the Hermes Agent knowledge base.

## Skills

### convert-image-links

检查和修复 Obsidian 知识库中的图片链接。将 Wiki 图片链接转为标准 MD 格式，修复损坏链接，URL 编码。

**工作流程**：先检查 → 确认 → 修复 → 再次检查 → 自动 git commit

详见 [convert-image-links/SKILL.md](convert-image-links/SKILL.md)

### alarm-statistics-analysis `v1.4.2`

告警数据清洗 + 多维度统计 + 交互式 HTML 报告生成工具。

对原始告警 Excel 数据清洗后，生成包含摘要卡片、四维筛选（系统/周期/级别/文件来源）、动态图表（树状图/柱状图/堆叠图）、TOP 告警排行、明细表格、CSV 导出的 HTML 报告（Chart.js，浏览器直接打开）。**支持单文件和多文件对比分析**。

**v1.4.0 新特性**：默认输出**双文件**——完整版（含自定义导出 + 原始数据）和精简版（仅保留多维度筛选/图表/CSV 导出，体积缩减 99.8%）。

**v1.4.2 修复**：`_strip_lite()` 改为按函数名精确删除，修复精简版误删 `getFilteredRawData` 导致多维筛选失效的问题；移除 `run_local.py` 自动打开浏览器功能。

**工作流程**：`python run_local.py --file data.xlsx` 单文件，或 `--file A.xlsx --file B.xlsx --labels "A,B"` 多文件对比。

详见 [alarm-statistics-analysis/SKILL.md](alarm-statistics-analysis/SKILL.md)
