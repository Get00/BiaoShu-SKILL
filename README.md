# 📝 BiaoShu-Writer-Pro v2.3（标书智能生成专家 · 响应文件格式保全版）

> 智能标书生成专家 —— 在原作品基础上强化了**招标文件响应格式的忠实保留**，并补齐**离线自足能力**。

[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.8%2B-green.svg)](https://www.python.org/)

**一句话**：发送招标文件，自动生成符合行业标准的专业投标书；**当招标文件附有响应文件格式时，把它当填空题——题干一字不改，只填横线。**

---

## ⚠️ 来源与授权声明（请先阅读）

本作品是以下开源项目的**修改版本（Derivative Work）**：

| 项目 | 内容 |
|---|---|
| 原作品 | **BiaoShu-SKILL**（内部名 `bid-writer-pro` / `BiaoShu-writer-pro`） |
| 作者 | **Get00 / TreaSkill** |
| 仓库 | https://github.com/Get00/BiaoShu-SKILL |
| 许可 | **Apache License 2.0** |

依据 Apache License 2.0 分发。原作品的版权与署名声明已完整保留。

- 许可证全文见 [`LICENSE`](LICENSE)
- **本版全部改动清单见 [`NOTICE`](NOTICE)**（Apache 2.0 第 4(b) 条要求被修改文件带显著声明，该文件逐项列出）
- 修改过的主要文件（`SKILL.md`、`scripts/format_docx.py`、`scripts/bid_writer_pipeline.py`、`templates/00_标书制作指南.md`、`requirements.txt`、`_meta.json`）已在文件头或 NOTICE 中标注

> 备注：原仓库 README 标注为 MIT，但其 `LICENSE` 文件正文为 **Apache License 2.0**。
> 本版以 `LICENSE` 文件为准，按 Apache 2.0 处理。

---

## 🆕 本版（v2.3）相对原作品增加了什么

### 1. 响应文件格式 = 填空题原则（核心改动）

原版本会把你自己的"政府标书标准"（仿宋 四号 28磅）套到成品上。但如果招标文件**已经给了响应文件格式**，那样做等于**改动了题干**——正文里明确列过：招标文件提供格式的，必须**优先套用并保持文字内容、字体字号、排版不变**，否则可能构成"对格式文件的实质性修改"而被否决。

本版据此新增：

- **SKILL.md 顶部新增最高优先级章节**，立下六条铁律（逐字保留题干、保留下划线、照抄原字体字号、底板填空不重建、必须读真实排版、违规后果）
- **执行流程新增"第0步：检查响应文件格式"**，成为不可跳过的前置步骤
- **`format_docx.py` 新增 `tender` 标准（保留原格式）并设为默认**，取代原来默认套用政府标准的行为
- **`bid_writer_pipeline.py` 新增步骤 1.6**：自动检测招标文件是否附有响应文件格式，命中即**强制切换为 tender** 并告警

### 2. 三个全新工具

| 工具 | 作用 |
|---|---|
| `scripts/extract_format.py` | **抓格式** —— 把招标文件的"投标/响应文件格式"章节整段抽出来，原样不改一字 |
| `scripts/apply_bid_template.py` | **底板填空** —— 以该章节为底板，只往下划线 run 和空白单元格里填内容 |
| `scripts/check_format_drift.py` | **格式漂移检查** —— 逐段比对成品与招标文件，检出题干删改与字体/字号/下划线/行距漂移 |

### 3. 离线自足能力（不依赖 MCP）

原版本依赖宿主环境的文件与文档工具；本版改为**自带纯 Python 工具**，换台电脑装好 pip 依赖即可运行：

| 工具 | 作用 |
|---|---|
| `scripts/fsutil.py` | 文件读写：读/写文本（自动识别 UTF-8/**GBK**）、列目录、通配符查找、正则搜索、复制、删除 |
| `scripts/doc_io.py` | 文档读写：**Word / Excel / PPT / PDF** 四类格式的读与写，含 `.doc`/`.xls` 老格式转换 |
| `scripts/office/` | 结构化文档工具包（按选区提取/改写 docx/xlsx/pptx/pdf，源文件只读） |
| `scripts/verify_env.py` | 环境自检 + 一键补依赖 |

---

## 🚀 安装

```bash
# 1) 环境自检（列出缺什么并给出修复命令）
python scripts/verify_env.py

# 2) 自动补齐依赖
python scripts/verify_env.py --install
#    或
pip install -r requirements.txt

# 3) 再次自检确认
python scripts/verify_env.py
```

**依赖（全部纯 Python，pip 可装）**：python-docx、openpyxl、python-pptx、pdfplumber、pypdf、reportlab、jieba、pandas；Windows 上另有可选的 pywin32（用于 `.doc`/`.xls` 转换）。

---

## 📖 使用

### 常规：从招标文件生成技术标

```bash
python scripts/bid_writer_pipeline.py "招标文件.pdf" --output-dir ./输出
```

> 若招标文件含响应文件格式，步骤 1.6 会自动把格式标准切为 `tender` 并提示。

### 格式保全三件套

```bash
# 1) 抓格式：把招标文件里的格式章节原样抽出来
python scripts/extract_format.py "招标文件.docx" --output "格式.docx"

# 2) 看看底板上有哪些空要填
python scripts/apply_bid_template.py "招标文件.docx" --section "响应文件格式" --list-blanks

# 3) 按 fields.json 填空（题干、字体、字号、下划线一律不动）
python scripts/apply_bid_template.py "招标文件.docx" \
    --section "响应文件格式" --fields fields.json --output "响应文件.docx"

# 4) 复核：成品是否忠实保留了招标文件格式
python scripts/check_format_drift.py "响应文件.docx" --source "招标文件.docx"
```

`fields.json` 示例：

```json
{
  "blanks": ["XX省电力有限公司（采购人）", "XX技术有限公司"],
  "replacements": [
    {"find": "致：采购人", "replace": "致：XX省电力有限公司（采购人）"}
  ],
  "cells": [
    {"table": 0, "row": 1, "col": 2, "value": "158.00"}
  ]
}
```

### 文档与文件工具

```bash
python scripts/doc_io.py read "报价清单.xlsx" --json      # 读表格
python scripts/doc_io.py write-docx "输出.docx" --md 内容.md --table 表格.json
python scripts/doc_io.py convert "老文件.doc" "新文件.docx"
python scripts/fsutil.py grep . --pattern "否决" --glob "*.md"
```

---

## 🧭 定位格式章节的三条实战经验

写这些脚本时踩过的坑，一并记在这里：

1. **章号不只叫"章"** —— 有的写成"第X卷""第X篇""第X部分"，正则要一并支持。
2. **章标题会出现两次** —— 目录页一次、正文一次，正文里还有"见第五章响应文件格式"这类引用；**整章标题要取最后一次匹配**才落得到正文真标题。
3. **结束边界必须"同级"** —— 起点是"第六章"就只认下一个"第X章"，**不能**把内部的"第一部分 商务文件"当成边界。

---

## ⚠️ 已知限制

- **`.doc` / `.xls` 老格式转换**需要本机装有 Microsoft Office（走 COM）或 LibreOffice；两者都没有时会明确报错。
- **命令行传 JSON 易被 shell 吃掉引号**（尤其 PowerShell）。`--table` / `--json` 同时接受"JSON 字符串"和"JSON 文件路径"，**建议用文件**。
- `scripts/office/` 工具包取自 Cherry Studio 随附的 `office-transform` 技能，非本作品原创，再分发前请自行核实其许可。

---

## 🙏 致谢

- 原作品作者 **Get00 / TreaSkill** —— [BiaoShu-SKILL](https://github.com/Get00/BiaoShu-SKILL)
- [pdfplumber](https://github.com/jsvine/pdfplumber)、[python-docx](https://github.com/python-openxml/python-docx)、[python-pptx](https://github.com/scanny/python-pptx)、[openpyxl](https://foss.heptapod.net/openpyxl/openpyxl)、[reportlab](https://www.reportlab.com/)、[jieba](https://github.com/fxsjy/jieba)
- [Trea](https://www.trae.ai/) —— AI 辅助开发平台（原作品使用）

---

## 📄 许可证

[Apache License 2.0](LICENSE) · 改动清单见 [NOTICE](NOTICE)
