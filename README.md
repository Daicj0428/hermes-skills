# Hermes Skills

Skills and scripts for the Hermes Agent knowledge base.

## Skills

### convert-image-links

检查和修复 Obsidian 知识库中的图片链接。将 Wiki 图片链接转为标准 MD 格式，修复损坏链接，URL 编码。

**工作流程**：先检查 → 确认 → 修复 → 再次检查 → 自动 git commit

详见 [convert-image-links/SKILL.md](convert-image-links/SKILL.md)

### alarm-statistics-analysis

告警数据清洗 + 多维度统计 + 交互式 HTML 报告生成工具。

对原始告警 Excel 数据清洗后，生成包含摘要卡片、多维筛选、动态图表（树状图/柱状图/堆叠图）、TOP 告警排行、明细表格、自定义 CSV 导出的单文件 HTML 报告（Chart.js，浏览器直接打开）。

**工作流程**：`python main.py` → 清洗 → 报告 → 清理，一步完成。

详见 [alarm-statistics-analysis/SKILL.md](alarm-statistics-analysis/SKILL.md)
