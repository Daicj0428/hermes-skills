---
name: convert-image-links
description: "转换 Obsidian 笔记中的图片链接格式：Wiki→MD、URL 编码、修复损坏链接。先检查后修复，自动 git commit。"
version: 1.0.0
author: Daicj0428 & Hermes Agent
license: MIT
platforms: [linux, windows]
metadata:
  hermes:
    tags: [obsidian, image-links, wiki-conversion, knowledge-base]
    category: productivity
---

# 转换图片链接格式 (convert-image-links)

检查和修复 Obsidian 知识库中的图片链接。将 Wiki 图片链接转为标准 MD 格式，修复损坏链接，对路径做 URL 编码。

## 触发条件

当用户说以下任一语句时触发：
- 「转换图片链接格式」
- 「检查图片链接」
- 「修复图片链接」
- 「图片链接转换」

## 核心原则

1. **先检查，后修复** — 永远先运行 `final_verify.py` 查看问题
2. **修改脚本变量** — 运行前必须根据用户指定的目标修改 `MODE`、`TARGET_FILE`、`VAULT_ROOT`
3. **确认后修复** — verify 结果出来后，向用户确认再执行 fix
4. **修复后再检查** — fix 完成后再次运行 verify 确认效果
5. **自动 git 提交** — 修复完成后自动 commit（push 由 Obsidian Git 插件处理）

## 包含脚本

| 文件 | 功能 | 是否修改文件 |
|------|------|:---:|
| `scripts/final_verify.py` | 扫描笔记，报告所有损坏的图片链接 | ❌ 只读 |
| `scripts/final_fix.py` | 修复损坏的图片链接 + Wiki 链接转 MD 链接 | ✅ 会改写 |

## 配置说明

两个脚本均在文件顶部通过变量配置：

```python
MODE = "dir"                 # "single" = 单文件 | "dir" = 扫描整个目录
VAULT_ROOT = r"/path/to/vault"  # 目录模式下的扫描根目录
# TARGET_FILE = r"/path/to/file.md"  # 单文件模式
```

## 注意事项

- `final_fix.py` 会将所有 `![[image.png]]` 转为 MD 链接
- 路径中的中文和空格会被 URL 编码
- 外部链接（`https://...`）不会被修改
- WSL 环境下使用 `/mnt/c/...` 路径
- git push 由 Obsidian Git 插件自动处理

## 许可证

MIT License — 详见 LICENSE 文件
