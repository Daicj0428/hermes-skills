# Hermes Skills

Skills and scripts for the Hermes Agent knowledge base.

## Skills

### alarm-statistics-analysis `v1.4.2`

告警数据清洗 + 多维度统计 + 交互式 HTML 报告生成工具。

对原始告警 Excel 数据清洗后，生成包含摘要卡片、四维筛选（系统/周期/级别/文件来源）、动态图表（树状图/柱状图/堆叠图）、TOP 告警排行、明细表格、CSV 导出的 HTML 报告（Chart.js，浏览器直接打开）。**支持单文件和多文件对比分析**。

**v1.4.2 新特性**：

| 特性 | 说明 |
|------|------|
| 双报告输出 | 完整版（~50MB，含全量 81K+ 原始数据 JSON + 自定义导出） + 精简版（~100KB，保留全部筛选/图表/CSV 导出） |
| 多文件对比 | 同时分析多个月份/来源，筛选器新增「文件来源」维度 |
| 四维筛选联动 | 系统 × 周期 × 告警级别 × 文件来源组合过滤，摘要卡片/图表/表格实时联动 |
| `_strip_lite()` | 按元素 ID 精确删除生成精简版，筛选功能完整无损 |

**工作流程**：

```bash
# 单文件
python run_local.py --file data.xlsx

# 多文件对比
python run_local.py --file 7月.xlsx --file 9月.xlsx --labels "7月,9月"
```

> 📖 详细文档：[Skill 分享文档（15章）](https://github.com/Daicj0428/hermes-skills/blob/main/alarm-statistics-analysis/SKILL.md)

### convert-image-links

检查和修复 Obsidian 知识库中的图片链接。将 Wiki 图片链接转为标准 MD 格式，修复损坏链接，URL 编码。

**工作流程**：先检查 → 确认 → 修复 → 再次检查 → 自动 git commit

详见 [convert-image-links/SKILL.md](convert-image-links/SKILL.md)
