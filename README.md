# 📝 BiaoShu-Writer-Pro

> 智能标书生成专家 - 通用行业投标文件自动生成工具

[![Version](https://img.shields.io/badge/version-2.2.0-blue.svg)](https://github.com/your-username/bid-writer-pro)
[![Python](https://img.shields.io/badge/python-3.8+-green.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

**一句话介绍**：发送招标文件，自动生成符合行业标准的专业投标书。

---

## ✨ 核心特性

- 🔄 **全自动流程** - 从招标文件解析到标书生成，14个步骤全自动执行
- 🏭 **多行业支持** - 支持IT/信息化、建筑工程、医疗健康、教育服务、制造业、物流运输、咨询服务等8大行业
- 🤖 **智能行业检测** - 自动识别招标文件所属行业，生成符合行业特点的专业内容
- 📈 **自我进化** - 遇到新行业时自动创建行业指南，逐步完善知识库
- 📄 **多格式支持** - 支持PDF、DOCX、DOC、TXT、XLSX格式的招标文件
- 📊 **Markdown转换** - 自动将招标文件转换为Markdown格式，方便AI阅读和处理
- 🏷️ **智能占位符** - 自动标记图片、表格、资质等非文字内容的位置
- 📁 **独立输出** - 每次运行生成带时间戳的独立文件夹，避免文件混乱

---

## 🎯 适用场景

| 场景 | 说明 |
|------|------|
| 企业投标部门 | 快速生成技术标，将编写时间从3-5天缩短到几小时 |
| 投标代理公司 | 批量处理多个招标文件，提高工作效率 |
| 跨行业投标 | 自动适配不同行业特点，无需配备各行业专业人员 |
| 标书模板积累 | 自动创建行业指南，逐步形成企业知识资产 |
| 标书质量检查 | 对照标准模板，检查标书完整性和规范性 |

---

## 🏭 支持行业

| 行业 | 关键词 | 指南文件 |
|------|--------|----------|
| 🖥️ IT/信息化 | 软件、系统集成、云计算、大数据 | `00_标书制作指南.md` |
| 🏗️ 建筑工程 | 建筑、施工、装修、市政 | `industry_建筑工程.md` |
| 🏥 医疗健康 | 医疗、医院、设备、药品 | `industry_医疗健康.md` |
| 📚 教育服务 | 教育、学校、教学、培训 | `industry_教育服务.md` |
| 🏭 制造业 | 制造、设备、自动化、智能制造 | `industry_制造业.md` |
| 🚚 物流运输 | 物流、运输、仓储、配送 | `industry_物流运输.md` |
| 💼 咨询服务 | 咨询、管理、审计、法律 | `industry_咨询服务.md` |
| 📁 通用行业 | 其他行业（自动创建指南） | 自动生成 |

---

## 🚀 快速开始

### 环境要求

- Python 3.8+
- 操作系统：Windows / macOS / Linux

### 安装

```bash
# 克隆项目
git clone https://github.com/your-username/bid-writer-pro.git
cd bid-writer-pro

# 安装依赖
pip install -r requirements.txt

# 或运行安装脚本
python install.py --auto
```

### 基本使用

```bash
# 最简单的用法（自动检测行业）
python scripts/bid_writer_pipeline.py "招标文件.pdf"

# 指定输出目录
python scripts/bid_writer_pipeline.py "招标文件.pdf" --output-dir ./标书输出

# 指定行业
python scripts/bid_writer_pipeline.py "招标文件.pdf" --industry construction

# 列出支持的行业
python scripts/bid_writer_pipeline.py --list-industries

# 检查环境
python scripts/check_environment.py
```

---

## 📁 输出结构

```
招标文件所在目录/
├── 招标文件.pdf                        # 原始招标文件
└── XX项目技术标.docx                   # ⭐ 最终标书（与招标文件同目录）

output/
└── {时间戳}_{招标文件名}/              # 示例：20260522_143025_XX项目招标/
    └── debug/                          # 所有中间文件
        ├── template_knowledge.md       # 模板知识
        ├── bid_document.md             # 招标文件Markdown版
        ├── industry_detection.json     # 行业检测结果
        ├── scoring_criteria.json       # 评分标准
        ├── key_requirements.json       # 关键要求
        ├── outline.json                # 标书大纲
        ├── word_count_check.json       # 字数检查
        └── chapters/                   # 章节内容
            ├── 00_封面.md
            ├── 00_目录.md
            ├── 01_项目概述.md
            ├── 02_技术方案.md
            └── ...
```

---

## 🔍 工作流程

```text
招标文件 → Markdown转换 → 行业检测 → 评分标准提取
    ↓
大纲生成 → 内容生成 → 字数检查 → AI去痕
    ↓
占位符插入 → 封面生成 → 目录生成 → Word生成 → 格式规范化
    ↓
最终标书（与招标文件同目录）
```

### 详细步骤

| 步骤 | 说明 | 输出文件 |
|------|------|----------|
| 0 | 通读模板知识 | `template_knowledge.md` |
| 1 | 解析招标文件 | `bid_document.md` |
| 1.5 | 检测行业类型 | `industry_detection.json` |
| 2 | 提取评分标准 | `scoring_criteria.json` |
| 3 | 提取关键要求 | `key_requirements.json` |
| 4 | 提取标书名称 | - |
| 5 | 生成大纲 | `outline.json` |
| 6 | 生成内容 | `chapters/*.md` |
| 7 | 字数检查 | `word_count_check.json` |
| 8 | AI去痕 | `humanized_*.md` |
| 9 | 插入占位符 | - |
| 10 | 生成封面 | `00_封面.md` |
| 11 | 生成目录 | `00_目录.md` |
| 12 | 生成Word | `{项目名称}技术标.docx` |
| 13 | 格式规范化 | 最终标书 |

---

## 📂 项目结构

```
bid-writer-pro/
├── SKILL.md                    # SKILL主文档（Trea会读取）
├── _meta.json                  # 元数据配置
├── README.md                   # 项目说明文档
├── requirements.txt            # 依赖库列表
├── install.py                  # 安装脚本
├── scripts/                    # 脚本目录
│   ├── bid_writer_pipeline.py  # 主流程脚本
│   ├── industry_detector.py    # 行业检测模块
│   ├── industry_content_generator.py  # 行业内容生成器
│   ├── parse_bid_file.py       # 招标文件解析
│   ├── convert_to_md.py        # Markdown转换
│   ├── extract_scoring.py      # 评分标准提取
│   ├── extract_requirements.py # 关键要求提取
│   ├── extract_bid_name.py     # 标书名称提取
│   ├── generate_outline.py     # 大纲生成
│   ├── generate_content.py     # 内容生成
│   ├── check_word_count.py     # 字数检查
│   ├── humanizer.py            # AI去痕
│   ├── convert_to_docx.py      # Word转换
│   ├── format_docx.py          # 格式规范化
│   ├── generate_placeholder.py # 占位符生成
│   ├── generate_cover.py       # 封面生成
│   ├── generate_toc.py         # 目录生成
│   └── check_environment.py    # 环境检查
├── templates/                  # 模板知识库
│   ├── 00_标书制作指南.md      # 通用指南
│   ├── industry_建筑工程.md    # 建筑工程行业指南
│   ├── industry_医疗健康.md    # 医疗健康行业指南
│   ├── industry_教育服务.md    # 教育服务行业指南
│   ├── industry_制造业.md      # 制造业行业指南
│   ├── industry_物流运输.md    # 物流运输行业指南
│   ├── industry_咨询服务.md    # 咨询服务行业指南
│   └── ...                     # 更多行业指南（自动创建）
└── references/                 # 参考资料
    └── industry_configs.json   # 行业配置文件
```

---

## ⚙️ 配置选项

### 命令行参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `bid_file` | 招标文件路径（必填） | - |
| `--output-dir, -o` | 输出目录 | `./output` |
| `--format-standard, -f` | 格式标准 | `government` |
| `--industry, -i` | 行业代码 | 自动检测 |
| `--company-profile, -c` | 企业资质文件 | - |
| `--total-pages, -p` | 预计总页数 | `40` |
| `--list-industries, -l` | 列出支持的行业 | - |

### 格式标准

| 标准 | 说明 |
|------|------|
| `government` | 政府采购标准（默认） |
| `enterprise` | 企业招标标准 |
| `highway` | 高速公路工程标准 |

---

## 🔧 常见问题

### Q: 如何添加新的行业支持？

A: SKILL会自动创建新行业的指南文件。当检测到未支持的行业时，系统会：
1. 自动分析招标文件内容
2. 生成行业指南文件
3. 保存到 `templates/` 文件夹
4. 告知用户SKILL已完善

### Q: 如何提高标书质量？

A: 建议：
1. 确保 `templates/` 文件夹中有足够的行业知识文档
2. 在生成后人工审核和调整内容
3. 根据实际情况修改占位符内容

### Q: 支持扫描版PDF吗？

A: 支持，但需要安装OCR依赖：
```bash
pip install pytesseract
# 还需要安装Tesseract OCR引擎
```

### Q: 生成的标书字数不够怎么办？

A: 系统会自动检查字数并提示，可以通过以下方式增加字数：
1. 增加 `--total-pages` 参数值
2. 在 `templates/` 中添加更多行业知识文档
3. 人工补充详细内容
---
## 📊 版本历史

| 版本 | 日期 | 更新内容 |
|------|------|---------|
| v2.2 | 2026-05-22 | 修复输出目录结构，明确时间戳文件夹说明 |
| v2.1 | 2026-05-22 | 添加行业指南自动创建功能 |
| v2.0 | 2026-05-22 | 改造为通用行业版，支持8个行业 |
| v1.9 | 2026-05-21 | 优化内容生成逻辑，中间文件保存到debug文件夹 |
| v1.8 | 2026-05-21 | 修复PDF转换编码问题 |
| v1.5 | 2026-05-21 | 添加Markdown转换功能 |
---

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

---

## 🙏 致谢

- [pdfplumber](https://github.com/jsvine/pdfplumber) - PDF解析库
- [python-docx](https://github.com/python-openxml/python-docx) - Word文档处理库
- [Trea](https://www.trae.ai/) - AI辅助开发平台

---
<img width="1448" height="1086" alt="ad0bb0c2-2043-4fd8-b2d4-10b28c1ca63b" src="https://github.com/user-attachments/assets/112788d6-8d19-4ecc-8763-7ffa1e285fe1" />

<p align="center">Made with ❤️ by BiaoShu-Writer-Pro And Trea</p>

## Stargazers over time
[![Stargazers over time](https://starchart.cc/Get00/BiaoShu-SKILL.svg?variant=adaptive)](https://starchart.cc/Get00/BiaoShu-SKILL)
