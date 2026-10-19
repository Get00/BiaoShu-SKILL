---
name: bid-writer-pro
description: 标书智能生成专家 v2.3 - 通用行业投标文件自动生成工具。支持IT/信息化、建筑工程、医疗健康、教育服务、制造业、物流运输、咨询服务等多个行业。自动检测招标文件所属行业，智能判断是否需要创建新的行业指南。**招标文件附有投标/响应文件格式时，按"填空题"原则保留原格式（题干、字体、字号、下划线一字不改），并提供抓格式、底板填空、格式漂移检查三个工具**。每次运行生成带时间戳的独立文件夹，避免文件混乱。适用场景：各行业投标文件编写、标书自动生成、标书格式规范化、响应文件格式套用。
---

> **【修改声明】** 本文件基于开源项目 [BiaoShu-SKILL](https://github.com/Get00/BiaoShu-SKILL)
> （Apache License 2.0）修改而来。修改日期：2026-10-13 ｜ 改动内容见 `NOTICE` 文件。


# 标书智能生成专家 v2.3 - 通用行业版

投标技术标文档自动生成工具。发送招标文件 → 自动转换Markdown → 生成符合行业标准格式的Word标书文档。

---

## ⚠️ 最高优先级：响应文件格式 = 填空题原则

**招标文件附有"响应文件格式"（通常在第五章，或末尾几个附件里）时，该格式是填空题，不是参考模板。**

动手前先确认：招标文件有没有给响应文件格式？
- **给了 → 严格执行本节规则。本节规则优先于后文所有"格式标准"。**
- **没给 → 才可使用后文 `government` / `enterprise` / `highway` 格式标准。**

### 填空题的六条铁律

1. **逐字保留格式中的全部文字，一个字不得删改。** 包括：
   - 括号内的说明文字（如"（如谈判过程中响应供应商需出具本委托书，否则无需出具。）"）
   - 平台操作提示行（如"（上传投标工具路径：商务文件-选择分标-按钮“编辑”-“法定代表人授权委托书批量管理”-按钮“文件上传”）"）
   - 表头、注释、"注："之后的全部文字、声明语、承诺语
   - 不得"精简""合并""改写成更通顺的表述"
2. **保留填空下划线。** 招标文件里的 `________` 是带下划线的 run，答案要填在下划线上，**禁止删掉下划线改成纯文字**。
3. **字体、字号、加粗、行距、首行缩进、对齐方式一律照抄招标文件原格式。** 招标文件用五号就用五号，小节标题用黑体小四就用黑体小四。**严禁套用任何"通用标书格式"。**
4. **用底板填空，不要从零重建。** 以招标文件中格式所在页为底板，在底板内填空；**禁止脱离底板重新排版同类文档**——这是最常见的错误来源。
5. **读取格式时必须读真实排版。** 用 `python-docx` 逐 run 读 `rPr`（字体/字号/下划线/加粗/行距），**不得只依据 `read_docx_text`、`convert_to_md` 等纯文本转写结果**——纯文本会把字体、字号、下划线全部剥离，导致你"看不见"原格式。
6. **违规后果严重。** 采购文件普遍规定"对格式文件进行实质性修改或未按格式进行签署的，视为无效文件/否决响应"。删题干、改字号都可能直接导致废标。

### 配套工具

```bash
# 抓格式：把招标文件的投标/响应文件格式整段抽出来（原样，不改一字）
python scripts/extract_format.py "招标文件.docx" --output "格式.docx"

# 底板填空：保留原格式，只填下划线与空白单元格
python scripts/apply_bid_template.py "招标文件.docx" \
  --section "响应文件格式" --fields fields.json --output "响应文件.docx"

# 核对成品与招标文件原格式是否一致
python scripts/check_format_drift.py "响应文件.docx" --source "招标文件.docx"
```

### 定位格式章节的实战经验（踩过的坑）

- **章号不只叫"章"**：有的招标文件写"第X卷""第X篇""第X部分"，正则要一并支持。
- **章标题会出现两次**：目录页一次、正文一次，且正文中还有"见第五章响应文件格式"这类引用——**整章标题取最后一次匹配**才能落到正文真标题。
- **结束边界必须"同级"**：起点是"第六章"就只认下一个"第X章"，**不能**把内部的"第一部分 商务文件"当成边界。

---

## 功能概述


### 核心能力
1. **招标文件解析** - 支持PDF/DOCX/DOC/TXT/XLSX格式
2. **文档转Markdown** - **自动将PDF/Word转换为Markdown格式，方便AI阅读和处理**
3. **行业自动检测** - **自动识别招标文件所属行业，生成符合行业特点的投标文件**
4. **评分标准提取** - 自动识别评分项、分值、权重
5. **关键要求识别** - 提取资质、技术、商务要求
6. **标书名称提取** - 从招标文件和内容中自动提取项目名称
7. **标书大纲生成** - 根据评分标准生成4级标题大纲
8. **标书内容生成** - 按章节并发编写内容
9. **字数检查** - 公式化字数检查，确保达标
10. **AI去痕处理** - 去除AI生成痕迹，使内容更自然
11. **占位符生成** - 标记需要插入图片/表格/资质的位置
12. **Word文档生成** - 输出符合行业标准格式的.docx文件

### 支持行业
- **IT/信息化**：信息技术、软件开发、系统集成、数据中心等
- **建筑工程**：建筑施工、工程总承包、装饰装修、市政工程等
- **医疗健康**：医疗设备、医疗服务、健康管理、医药等
- **教育服务**：教育信息化、教学设备、培训服务、校园建设等
- **制造业**：设备制造、工业自动化、智能制造、供应链等
- **物流运输**：物流服务、运输服务、仓储服务、供应链管理等
- **咨询服务**：管理咨询、技术咨询、审计服务、法律服务等
- **通用行业**：适用于未指定行业的通用模板

### Markdown转换优势
- **保留结构**：自动识别标题层级（一、二级标题等）
- **表格转换**：将Word/PDF表格转换为Markdown表格格式
- **方便AI阅读**：Markdown格式更适合AI理解和处理
- **调试友好**：转换后的文件保存在 `debug/bid_document.md`，便于查看

### 输出特点
- **文件命名**：自动从招标文件提取项目名称，格式为 `{项目名称}技术标.docx`
- **占位符规范**：非文字内容用统一格式的占位符标记
- **格式优先级**：**招标文件提供了响应文件格式的，一律保留原格式（见前文《填空题原则》）**；招标文件未提供格式的，才套用下方政府/企业/高速公路标准
- **行业适配**：根据检测到的行业自动生成专业内容

---

## 📦 自足能力包（迁移到新电脑必备）

本 skill **不依赖任何 MCP 服务**：所有文件读写与文档处理均由自带的纯 Python 工具完成。
换一台电脑后，装好 Python 依赖即可运行，无需重新配置 MCP。

### 自带工具

| 工具 | 能力 |
|------|------|
| `scripts/fsutil.py` | 文件读写：读/写文本（自动识别 UTF-8/GBK）、列目录、通配符查找、正则搜索、复制、删除 |
| `scripts/doc_io.py` | 文档读写：**Word(.docx) / Excel(.xlsx) / PowerPoint(.pptx) / PDF** 四类格式的读与写；`.doc`/`.xls` 老格式转换 |
| `scripts/office/` | 结构化文档包（源自 office-transform）：按选区提取/改写 docx/xlsx/pptx/pdf，源文件只读 |
| `scripts/office_extract.py` | 上述工具包的「提取」命令行入口 |
| `scripts/office_patch_copy.py` | 上述工具包的「改写副本」命令行入口 |
| `scripts/verify_env.py` | 环境自检：查 Python 版本、依赖库、自带脚本是否齐全，并给出修复命令 |

### 在新电脑上首次运行

```bash
# 1) 自检（列出缺什么、并打印修复命令）
python scripts/verify_env.py

# 2) 自动补齐缺失依赖
python scripts/verify_env.py --install
#    或
pip install -r requirements.txt

# 3) 再自检一次确认
python scripts/verify_env.py
```

### 依赖库（全部纯 Python，pip 可装）

| 库 | 用途 |
|---|---|
| python-docx | Word 读写（核心） |
| openpyxl | Excel 读写 |
| python-pptx | PowerPoint 读写 |
| pdfplumber / pypdf | PDF 读取 |
| reportlab | PDF 生成（内置 CJK 字体，中文无需外部字体文件） |
| jieba | 中文分词（字数统计 / AI 去痕） |
| pandas | 表格处理（可选） |
| pywin32 | `.doc`/`.xls` 老格式转换（仅 Windows，可选） |

### 没有 MCP 时怎么读文档

宿主若未提供文档类 MCP 工具，直接用自带工具即可：

```bash
python scripts/doc_io.py read "招标文件.docx"         # 读文本
python scripts/doc_io.py read "报价清单.xlsx" --json   # 读表格
python scripts/doc_io.py info "附件.pdf"               # 结构摘要
python scripts/fsutil.py grep . --pattern "否决" --glob "*.md"   # 全文检索
```

### 老格式 .doc / .xls 怎么转

`doc_io.py convert` 自动依次尝试：①Windows + 装了 Office → 走 Word/Excel COM；
②否则走 LibreOffice（`soffice`）。两条都不可用时明确报错并提示装其中一个。

```bash
python scripts/doc_io.py convert "老文件.doc" "新文件.docx"
```

### 注意：命令行传 JSON 易被 shell 吃掉引号

PowerShell 会把 `--table '[["a","b"]]'` 的引号剥掉导致解析失败。
**表格数据请写进 .json 文件后传文件路径**，`--table`/`--json` 都同时接受
「JSON 字符串」和「JSON 文件路径」两种形式。

---

## 环境安装


### 依赖库
```bash
pip install python-docx pdfplumber openpyxl PyPDF2 jieba pandas pywin32
```

### 字体要求
- **宋体（SimSun）** - 正文默认字体
- **黑体（SimHei）** - 标题字体
- **仿宋_GB2312** - 政府标书正文字体

将字体文件复制到系统字体目录：
- Windows: `C:\Windows\Fonts\`
- Mac: `~/Library/Fonts/`

---

## 快速使用

### 基础用法
```bash
# 1. 发送招标文件（PDF/DOCX）给AI
# 2. AI自动解析并生成大纲
# 3. 确认大纲后AI并发编写各章节
# 4. 字数检查 + AI去痕处理
# 5. 生成占位符标记
# 6. 转换生成Word文档
```

### 命令行用法（CLI）
```bash
# 完整流程
python scripts/bid_writer_pipeline.py "招标文件.pdf" --output-dir ./output

# 带企业资质信息
python scripts/bid_writer_pipeline.py "招标文件.docx" \
  --company-profile "企业资质.json" \
  --output-dir ./标书输出

# 指定格式标准
python scripts/bid_writer_pipeline.py "招标文件.pdf" \
  --format-standard government \
  --output-dir ./output
```

### 在AI IDE中使用
```python
# 直接导入使用
import sys
sys.path.insert(0, '/path/to/bid-writer-pro/scripts')

from bid_writer_pipeline import BidWriterPipeline

pipeline = BidWriterPipeline(
    bid_file="招标文件.pdf",
    output_dir="./output",
    format_standard="government"
)
result = pipeline.run()
```

### 环境检查与安装
```bash
# 检查环境是否满足要求
python scripts/check_environment.py

# 安装依赖库
python install.py
# 或直接使用pip
pip install -r requirements.txt
```

---

## 执行流程

```
【第0步】检查招标文件有无"响应文件格式" ──┬─ 有 → 走"底板填空"，全程保留原格式（tender）
                                        └─ 无 → 走常规流程
    ↓
通读模板知识 → 发送招标文件 → 转换Markdown → 解析评分标准 → 提取标书名称 → 生成大纲 →
确认大纲 → 并发编写章节 → 字数检查 → AI去痕 →
占位符生成 → 转换Word → 格式规范化（tender 时为空操作）→ 输出标书.docx
```

### ⚠️ 第0步（必须先做）：检查响应文件格式

打开招标文件，查找"响应文件格式""投标文件格式"章节（通常在第五章，或末尾附件）。

- **找到 → 整个生成过程改为"底板填空"模式：**
  1. 用 `python-docx` 逐 run 读出该章节的真实排版（字体/字号/下划线/行距/缩进）
  2. 复制该章节作为底板，**不新建文档、不改动任何题干文字**
  3. 只在下划线 run 和空白表格单元格里填入本企业信息
  4. 后续 `--format-standard` 一律传 `tender`
- **没找到 → 才按下文常规流程和格式标准处理。**

**这一步不能跳过。** 跳过就会默认套用"政府标书标准"（仿宋四号28磅），
把招标文件的五号字、下划线和说明文字全部改掉，构成"对格式文件的实质性修改"。

### ⚠️ 重要：开始撰写标书前必须先判断行业类型

**在开始撰写标书之前，必须先判断招标文件所属的行业类型**，这是生成高质量标书的关键步骤！

#### 🔍 行业判断流程（必须执行）

```text
第1步：读取招标文件内容
    ↓
第2步：AI分析招标文件，判断所属行业
    ↓
第3步：检查templates文件夹中是否存在该行业的指南文件
    ↓
┌─ 如果存在：industry_{行业名称}.md
│   └─ 按照行业指南执行，生成符合行业特点的标书
│
└─ 如果不存在：
    └─ AI进行互联网检索，收集该行业的：
       - 行业特点和规范
       - 投标文件结构要求
       - 关键内容要点
       - 评分标准关注点
       ↓
    └─ 自动生成 industry_{行业名称}.md 文件
       ↓
    └─ 保存到 templates 文件夹
       ↓
    └─ 告知用户：已创建新的行业指南，SKILL已完善
```

#### 📋 支持的行业列表（已有指南）

| 行业代码 | 行业名称 | 指南文件 |
|---------|---------|---------|
| it_information | IT/信息化 | 00_标书制作指南.md |
| construction | 建筑工程 | industry_建筑工程.md |
| medical | 医疗健康 | industry_医疗健康.md |
| education | 教育服务 | industry_教育服务.md |
| manufacturing | 制造业 | industry_制造业.md |
| logistics | 物流运输 | industry_物流运输.md |
| consulting | 咨询服务 | industry_咨询服务.md |
| general | 通用行业 | 00_标书制作指南.md |

#### 🎯 行业判断关键词

**IT/信息化**：信息化、软件、系统集成、数据中心、云计算、大数据、人工智能、数字化
**建筑工程**：建筑、施工、工程、装修、市政、道路、桥梁、隧道
**医疗健康**：医疗、健康、医院、设备、药品、诊断、治疗、康复
**教育服务**：教育、学校、教学、培训、校园、课程、实验室、图书馆
**制造业**：制造、设备、自动化、智能制造、工业、生产线、供应链、质量
**物流运输**：物流、运输、仓储、配送、供应链、快递、货运、冷链
**咨询服务**：咨询、管理、审计、法律、评估、规划、研究、策划

#### 💡 行业判断示例

**示例1**：招标文件中包含"智慧交通"、"系统集成"、"软件开发"
→ 判断为：IT/信息化
→ 使用指南：00_标书制作指南.md

**示例2**：招标文件中包含"医院"、"医疗设备"、"诊断"
→ 判断为：医疗健康
→ 使用指南：industry_医疗健康.md

**示例3**：招标文件中包含"建筑工程"、"施工"、"装修"
→ 判断为：建筑工程
→ 使用指南：industry_建筑工程.md

### ⚠️ 重要：开始撰写标书前必须先通读模板知识

**在判断行业类型后，必须通读 `templates` 文件夹中的相关行业指南**，包括：

1. **`00_标书制作指南.md`** - 标书制作的通用指南
2. **`industry_{行业名称}.md`** - 特定行业的投标指南（如果存在）
3. **`bid_doc_*.md`** - 从DOC/DOCX转换的标书知识文档
4. **`bid_knowledge_*.md`** - 从PDF转换的招投标知识文档

**通读目的**：
- 理解该行业的投标文件特点和要求
- 掌握行业特定的编写要点和规范
- 了解行业常见的评分标准和关注点
- 学习行业优秀的标书范本

### 详细流程说明

#### 0. 行业判断与指南检查（必须首先执行！）
- **必须首先执行**
- 读取招标文件内容，判断所属行业
- 检查 `templates` 文件夹中是否存在该行业的指南文件
- **如果存在**：按照行业指南执行
- **如果不存在**：AI进行互联网检索，创建新的行业指南
- 通读相关行业指南，为后续生成高质量标书奠定基础

#### 1. 招标文件解析与Markdown转换
- 支持格式：PDF、DOCX、DOC、TXT、XLSX
- **自动转换为Markdown格式**，方便AI阅读和处理
- 保留标题层级结构和表格格式
- 自动检测扫描版PDF并提示
- Markdown文件保存在 `debug/bid_document.md`

#### 2. 评分标准提取
- 识别评分办法章节
- 提取评分大类（技术、商务、价格）
- 提取评分细则和分值权重

#### 3. 标书名称提取
- 从文件名提取项目名称
- 从内容中提取项目名称
- 格式：`{项目名称}技术标.docx`

#### 4. 大纲生成
- 根据评分标准生成4级标题大纲
- 每个评分项对应一个章节
- 字数按分值比例分配

#### 5. 内容生成
- 按章节并发编写
- 严格按评分标准编写
- 每小节 ≥ 3个独立段落
- 每章节配表格

#### 6. 字数检查
- 公式：目标字数 = 评分分值 × (总页数 ÷ 总分) × 780
- 合格范围：目标 × 0.75 ~ 1.25
- 不达标则打回重写

#### 7. AI去痕处理
- 去除AI生成痕迹
- 使内容更自然、更像人类书写
- 保留核心信息完整性

#### 8. 占位符生成
- 标记需要插入图片的位置
- 标记需要插入表格的位置
- 标记需要插入资质的位置

#### 9. Word文档生成
- Markdown转Word
- 应用行业标准格式
- 生成封面和目录

---

## 任务决策树

在开始执行前，先根据用户需求选择路径：

```text
用户需求是什么？
│
├─ 【完整生成标书】有招标文件，想生成完整标书
│  └─ 运行完整流程：scripts/bid_writer_pipeline.py "招标文件.pdf" --output-dir ./output
│     └─ 自动执行14个步骤：通读模板 → 解析文件 → 提取评分 → 生成大纲 → 生成内容 → 字数检查 → AI去痕 → 生成Word
│
├─ 【仅解析招标文件】只想提取招标文件中的信息
│  └─ 运行 scripts/parse_bid_file.py "招标文件.pdf"
│     └─ 输出：文本内容、评分标准、关键要求
│
├─ 【转换文档格式】想把PDF/DOCX转换为Markdown
│  ├─ 转换单个文件：运行 scripts/convert_to_md.py "文件.pdf" --output "输出.md"
│  ├─ 批量转换PDF：运行 scripts/convert_pdf_to_md.py --source "PDF文件夹" --output "输出文件夹"
│  └─ 批量转换DOC/DOCX：运行 scripts/convert_doc_to_md.py --source "DOC文件夹" --output "输出文件夹"
│
├─ 【仅提取评分标准】只想分析招标文件的评分标准
│  └─ 运行 scripts/extract_scoring.py "招标文件内容.txt"
│     └─ 输出：scoring_criteria.json（评分大类、分值、权重）
│
├─ 【仅提取关键要求】只想提取资质、技术、商务要求
│  └─ 运行 scripts/extract_requirements.py "招标文件内容.txt"
│     └─ 输出：key_requirements.json（资质要求、技术要求、商务要求）
│
├─ 【仅生成大纲】已有评分标准，想生成标书大纲
│  └─ 运行 scripts/generate_outline.py --scoring "scoring_criteria.json"
│     └─ 输出：outline.json（4级标题大纲、字数分配）
│
├─ 【仅生成内容】已有大纲，想生成标书内容
│  └─ 运行 scripts/generate_content.py --outline "outline.json" --scoring "scoring_criteria.json"
│     └─ 输出：各章节Markdown文件
│
├─ 【检查字数】已有内容，想检查字数是否达标
│  └─ 运行 scripts/check_word_count.py --chapters "chapters文件夹" --scoring "scoring_criteria.json"
│     └─ 输出：字数检查报告（目标字数、实际字数、是否达标）
│
├─ 【AI去痕处理】已有内容，想去除AI生成痕迹
│  └─ 运行 scripts/humanizer.py --input "内容.md" --output "处理后.md"
│     └─ 输出：去痕后的自然语言内容
│
├─ 【生成占位符】已有内容，想标记图片/表格/资质位置
│  └─ 运行 scripts/generate_placeholder.py --chapters "chapters文件夹" --requirements "key_requirements.json"
│     └─ 输出：带占位符标记的内容
│
├─ 【Markdown转Word】已有Markdown内容，想生成Word文档
│  └─ 运行 scripts/convert_to_docx.py --input "内容.md" --output "标书.docx" --standard government
│     └─ 输出：符合行业标准格式的Word文档
│
├─ 【格式规范化】已有Word文档，想统一格式
│  └─ 运行 scripts/format_docx.py --input "标书.docx" --standard government
│     └─ 输出：格式规范化的Word文档
│
├─ 【生成封面】想为标书生成封面页
│  └─ 运行 scripts/generate_cover.py --name "项目名称" --requirements "key_requirements.json"
│     └─ 输出：封面页Markdown
│
├─ 【生成目录】想为标书生成目录页
│  └─ 运行 scripts/generate_toc.py --chapters "chapters文件夹"
│     └─ 输出：目录页Markdown
│
├─ 【检查环境】想检查运行环境是否满足要求
│  └─ 运行 scripts/check_environment.py
│     └─ 输出：环境检查报告（Python版本、依赖库、脚本文件、模板知识）
│
├─ 【构建模板知识库】想从标书范本中提取知识
│  ├─ 从DOC/DOCX提取：运行 scripts/extract_templates.py --source "范本文件夹" --output "templates文件夹"
│  └─ 从PDF提取：运行 scripts/convert_pdf_to_md.py --source "PDF文件夹" --output "templates文件夹"
│
├─ 【查看中间结果】想查看已生成的中间文件
│  ├─ 查看评分标准：打开 output/scoring_criteria.json
│  ├─ 查看关键要求：打开 output/key_requirements.json
│  ├─ 查看大纲：打开 output/outline.json
│  ├─ 查看章节内容：打开 output/chapters/*.md
│  └─ 查看调试信息：打开 output/debug/ 文件夹
│
└─ 【需求不清楚或包含多步】不确定需要什么
   └─ 运行完整流程：scripts/bid_writer_pipeline.py "招标文件.pdf" --output-dir ./output
      └─ 自动执行所有步骤，生成完整标书
```

### 决策树使用说明

#### 场景1：完整生成标书（最常见）
```bash
# 用户有招标文件，想生成完整标书
python scripts/bid_writer_pipeline.py "XX市智慧交通项目招标文件.pdf" --output-dir ./output
```
**自动执行**：通读模板 → 解析文件 → 提取评分 → 提取要求 → 提取名称 → 生成大纲 → 生成内容 → 字数检查 → AI去痕 → 占位符 → 封面 → 目录 → 生成Word → 格式规范化

#### 场景2：分步执行（需要中间调整）
```bash
# 步骤1：解析招标文件
python scripts/parse_bid_file.py "招标文件.pdf" --output "bid_content.txt"

# 步骤2：提取评分标准
python scripts/extract_scoring.py --input "bid_content.txt" --output "scoring_criteria.json"

# 步骤3：提取关键要求
python scripts/extract_requirements.py --input "bid_content.txt" --output "key_requirements.json"

# 步骤4：生成大纲（可人工审核调整）
python scripts/generate_outline.py --scoring "scoring_criteria.json" --requirements "key_requirements.json" --output "outline.json"

# 步骤5：生成内容
python scripts/generate_content.py --outline "outline.json" --scoring "scoring_criteria.json" --output-dir "./chapters"

# 步骤6：字数检查
python scripts/check_word_count.py --chapters "./chapters" --scoring "scoring_criteria.json"

# 步骤7：AI去痕
python scripts/humanizer.py --input-dir "./chapters" --output-dir "./chapters_humanized"

# 步骤8：生成Word
python scripts/convert_to_docx.py --input-dir "./chapters_humanized" --output "标书.docx" --standard government
```

#### 场景3：仅转换文档格式
```bash
# 将PDF转换为Markdown
python scripts/convert_to_md.py "招标文件.pdf" --output "招标文件.md"

# 批量转换PDF文件夹
python scripts/convert_pdf_to_md.py --source "./pdf_files" --output "./markdown_files"

# 批量转换DOC/DOCX文件夹
python scripts/convert_doc_to_md.py --source "./doc_files" --output "./markdown_files"
```

#### 场景4：格式规范化
```bash
# 已有Word文档，统一格式
python scripts/format_docx.py --input "标书.docx" --standard government --output "标书_规范化.docx"
```

#### 场景5：检查环境
```bash
# 检查运行环境是否满足要求
python scripts/check_environment.py
```

### 注意事项

1. **首次使用**：先运行 `scripts/check_environment.py` 检查环境
2. **安装依赖**：运行 `python install.py --auto` 或 `pip install -r requirements.txt`
3. **模板知识**：首次生成标书前，确保 `templates` 文件夹中有知识文档
4. **格式标准**：可选 `government`（政府标准）、`enterprise`（企业标准）、`highway`（高速公路标准）
5. **中间文件**：所有中间文件保存在 `output/debug/` 文件夹，便于调试和审核

---

## 占位符规范

### 占位符类型

#### 1. 图片占位符
```
[图片占位符]
类型：技术架构图
描述：XX系统总体架构图
建议尺寸：宽15cm × 高10cm
```

#### 2. 表格占位符
```
[表格占位符]
类型：项目业绩表
描述：近3年类似项目清单
建议列数：6列
建议内容：项目名称、合同金额、完成时间、项目规模、客户名称、验收情况
```

#### 3. 资质占位符
```
[资质占位符]
类型：企业资质证书
描述：信息系统集成及服务一级资质证书
要求：有效期内的彩色扫描件
```

### 常用占位符示例

| 类型 | 占位符内容 | 使用场景 |
|------|-----------|---------|
| 图片 | [图片占位符] 类型：技术架构图 | 技术方案章节 |
| 图片 | [图片占位符] 类型：网络拓扑图 | 网络设计章节 |
| 图片 | [图片占位符] 类型：业务流程图 | 业务流程章节 |
| 表格 | [表格占位符] 类型：功能模块清单 | 功能设计章节 |
| 表格 | [表格占位符] 类型：项目进度计划表 | 项目实施章节 |
| 表格 | [表格占位符] 类型：人员配置表 | 项目团队章节 |
| 表格 | [表格占位符] 类型：项目业绩表 | 项目经验章节 |
| 资质 | [资质占位符] 类型：企业营业执照 | 附件章节 |
| 资质 | [资质占位符] 类型：信息系统集成一级资质 | 附件章节 |
| 资质 | [资质占位符] 类型：ISO9001认证证书 | 附件章节 |
| 资质 | [资质占位符] 类型：项目经理PMP证书 | 附件章节 |

---

## 格式标准

> **⚠️ 适用前提：本节仅用于"招标文件未提供响应文件格式"的场景。**
> 招标文件提供了响应文件格式的，**不得使用本节任何标准**——必须走前文《响应文件格式 = 填空题原则》，
> 用 `--format-standard tender` 保留原格式。

### 保留招标文件原格式（tender）——招标文件提供格式时的唯一选择

| 元素 | 格式要求 |
|------|---------|
| 全部元素 | **完全保留源文档格式，不做任何修改**。字体、字号、加粗、行距、缩进、对齐、下划线全部照抄。 |
| 用途 | 招标文件附有响应文件格式时，作为底板填空，保证"不构成对格式文件的实质性修改"。 |
| 命令 | `--format-standard tender` |

### 政府标书标准（招标文件未提供格式时的缺省值）

| 元素 | 格式要求 |
|------|---------|
| 正文字体 | 仿宋_GB2312 |
| 正文字号 | 四号（14pt） |
| 行距 | 28磅 |
| 页边距 | 上下2.54cm，左右3.17cm |
| 首行缩进 | 2字符 |
| 一级标题 | 黑体、三号（16pt）、加粗 |
| 二级标题 | 黑体、小三（15pt）、加粗 |
| 三级标题 | 黑体、四号（14pt）、加粗 |
| 四级标题 | 仿宋_GB2312、四号（14pt）、加粗 |

### 企业标书标准

| 元素 | 格式要求 |
|------|---------|
| 正文字体 | 宋体 |
| 正文字号 | 小四（12pt） |
| 行距 | 1.5倍行距 |
| 页边距 | 上下2.54cm，左右2.54cm |
| 首行缩进 | 2字符 |
| 一级标题 | 黑体、小二（18pt）、加粗 |
| 二级标题 | 黑体、三号（16pt）、加粗 |
| 三级标题 | 黑体、小三（15pt）、加粗 |

### 高速公路/航道工程标准

| 元素 | 格式要求 |
|------|---------|
| 正文字体 | 宋体 |
| 正文字号 | 三号（16pt） |
| 行距 | 28磅 |
| 页边距 | 上下2.5cm，左右2.5cm |
| 首行缩进 | 2字符 |

---

## 内容规则

### 基本要求
- **严格按评分标准编写** - 不泛泛而谈，针对每个评分项详细展开
- **每小节 ≥ 3个独立段落** - 确保内容充实
- **每章节配表格** - 使用Markdown表格格式
- **禁用"我方/我们"** - 用"将/项目组"替代
- **禁止金额/预算描述** - 除非招标文件明确要求

### 字数控制
- **公式**：目标字数 = 评分分值 × (总页数 ÷ 总分) × 780
- **合格范围**：目标 × 0.75 ~ 1.25
- **每页字数**：28行 × 28字 = 780字

### 章节结构
```
# 第一章 项目概述
## 1.1 项目背景
### 1.1.1 建设背景
### 1.1.2 现状分析
## 1.2 项目目标
## 1.3 项目范围

# 第二章 技术方案
## 2.1 总体架构
### 2.1.1 架构设计原则
### 2.1.2 系统架构设计
## 2.2 功能设计
### 2.2.1 功能模块划分
### 2.2.2 功能详细说明
## 2.3 数据设计
## 2.4 安全设计

# 第三章 项目实施
## 3.1 实施计划
## 3.2 人员配置
## 3.3 质量保证
## 3.4 风险控制

# 第四章 项目业绩
## 4.1 类似项目经验
## 4.2 客户评价

# 第五章 售后服务
## 5.1 服务体系
## 5.2 响应时间
## 5.3 培训计划

# 附件
## 附件1 企业资质
## 附件2 人员证书
## 附件3 项目业绩证明
```

---

## 核心脚本

| 脚本 | 功能 | 输入 | 输出 |
|------|------|------|------|
| `bid_writer_pipeline.py` | 主流程脚本 | 招标文件 | 标书.docx |
| `parse_bid_file.py` | 招标文件解析 | PDF/DOCX/DOC/TXT/XLSX | 文本内容 |
| `convert_to_md.py` | **文档转Markdown** | PDF/DOCX/DOC/TXT | Markdown文件 |
| `extract_scoring.py` | 评分标准提取 | 招标文件文本 | scoring_criteria.json |
| `extract_requirements.py` | 关键要求提取 | 招标文件文本 | key_requirements.json |
| `extract_bid_name.py` | 标书名称提取 | 招标文件 | 标书名称.docx |
| `generate_outline.py` | 大纲生成 | 评分标准 | outline.json |
| `generate_content.py` | 内容生成 | 大纲+评分标准 | Markdown内容 |
| `check_word_count.py` | 字数检查 | Markdown内容 | 检查结果 |
| `humanizer.py` | AI去痕处理 | Markdown内容 | 处理后内容 |
| `convert_to_docx.py` | Markdown转Word | Markdown文件 | Word文档 |
| `convert_pdf_to_md.py` | **PDF转Markdown** | PDF文件夹 | Markdown文件 |
| `convert_doc_to_md.py` | **DOC/DOCX转Markdown** | DOC/DOCX文件夹 | Markdown文件 |
| `format_docx.py` | Word格式规范化 | Word文档 | 格式化后的Word文档 |
| `generate_placeholder.py` | 占位符生成 | 内容分析 | 占位符标记 |
| `generate_cover.py` | 封面生成 | 项目信息 | 封面页 |
| `generate_toc.py` | 目录生成 | 文档结构 | 目录页 |
| `check_environment.py` | 环境检查工具 | 无 | 检查报告 |
| `extract_templates.py` | 模板提取工具 | DOC/DOCX文件夹 | Markdown文件 |

---

## 文件结构

```
bid-writer-pro/
├── SKILL.md                    # 说明文档
├── _meta.json                  # 元数据配置
├── requirements.txt            # 依赖库列表
├── install.py                  # 安装脚本
├── scripts/                    # 脚本目录
│   ├── __init__.py             # 包初始化
│   ├── bid_writer_pipeline.py  # 主流程脚本
│   ├── parse_bid_file.py       # 招标文件解析
│   ├── convert_to_md.py        # 文档转Markdown
│   ├── convert_pdf_to_md.py    # PDF转Markdown工具
│   ├── convert_doc_to_md.py    # DOC/DOCX转Markdown工具
│   ├── extract_scoring.py      # 评分标准提取
│   ├── extract_requirements.py # 关键要求提取
│   ├── extract_bid_name.py     # 标书名称提取
│   ├── generate_outline.py     # 大纲生成
│   ├── generate_content.py     # 内容生成
│   ├── check_word_count.py     # 字数检查
│   ├── humanizer.py            # AI去痕处理
│   ├── convert_to_docx.py      # Markdown转Word
│   ├── format_docx.py          # Word格式规范化
│   ├── generate_placeholder.py # 占位符生成
│   ├── generate_cover.py       # 封面生成
│   ├── generate_toc.py         # 目录生成
│   ├── check_environment.py    # 环境检查工具
│   └── extract_templates.py    # 模板提取工具
└── templates/                  # ⭐ 模板知识文件夹（必读）
    ├── 00_标书制作指南.md       # 标书制作完整指南
    ├── bid_doc_*.md            # 标书知识文档（15个）
    └── bid_knowledge_*.md      # 招投标知识文档（5个）
```

## 输出文件结构

```
招标文件所在目录/
├── 招标文件.pdf                        # 原始招标文件
└── XX市智慧交通建设项目技术标.docx    # ⭐ 最终标书文件（与招标文件同目录）

output/                                 # 中间文件目录
└── {时间戳}_{招标文件名}/              # ⭐ 带时间戳的独立文件夹（避免混乱）
    │                                   # 示例：20260522_143025_XX项目招标/
    └── debug/                          # 所有中间文件保存在debug文件夹
        ├── template_knowledge.md       # 模板知识（通读内容）
        ├── bid_document.md             # 招标文件Markdown版本
        ├── raw_bid_content.txt         # 原始解析内容
        ├── industry_detection.json     # ⭐ 行业检测结果
        ├── scoring_criteria.json       # 评分标准
        ├── key_requirements.json       # 关键要求
        ├── outline.json                # 标书大纲
        ├── word_count_check.json       # 字数检查结果
        └── chapters/                   # 章节Markdown文件
            ├── 00_封面.md              # 封面
            ├── 00_目录.md              # 目录
            ├── 01_项目概述.md
            ├── 02_技术方案.md
            ├── 03_项目实施.md
            ├── 04_项目业绩.md
            └── 05_售后服务.md
```

### 时间戳文件夹说明

- **格式**：`{YYYYMMDD}_{HHMMSS}_{招标文件名}`
- **示例**：`20260522_143025_XX市智慧交通项目招标`
- **作用**：每次运行生成独立文件夹，避免多次运行导致文件混乱
- **位置**：在用户指定的 `--output-dir` 目录下自动创建

---

## 使用示例

### 示例1：基础使用

```bash
# 用户发送招标文件
python scripts/bid_writer_pipeline.py "XX市智慧交通项目招标文件.pdf"

# AI处理流程：
# 1. 解析招标文件... ✅
# 2. 提取评分标准... ✅ (技术60分，商务20分，报价20分)
# 3. 提取标书名称... ✅ (XX市智慧交通建设项目技术标.docx)
# 4. 生成大纲... ✅ (5章18节)
# 5. 确认大纲... ✅
# 6. 生成内容... ✅ (约25000字)
# 7. 字数检查... ✅
# 8. AI去痕... ✅
# 9. 生成占位符... ✅ (12个图片，8个表格，6个资质)
# 10. 生成Word... ✅

# 输出文件：XX市智慧交通建设项目技术标.docx
```

### 示例2：带企业资质

```bash
python scripts/bid_writer_pipeline.py "招标文件.docx" \
  --company-profile "company_profile.json" \
  --format-standard government \
  --output-dir ./标书输出
```

### 示例3：指定格式标准

```bash
# 使用政府标书标准
python scripts/bid_writer_pipeline.py "招标文件.pdf" --format-standard government

# 使用企业标书标准
python scripts/bid_writer_pipeline.py "招标文件.pdf" --format-standard enterprise

# 使用高速公路标准
python scripts/bid_writer_pipeline.py "招标文件.pdf" --format-standard highway
```

---

## 占位符处理指南

### 生成后的处理步骤

1. **查看占位符清单**
   - 打开生成的Word文档
   - 搜索 `[图片占位符]`、`[表格占位符]`、`[资质占位符]`

2. **准备素材**
   - 根据占位符描述准备相应的图片、表格、资质文件
   - 确保图片清晰、表格完整、资质在有效期内

3. **替换占位符**
   - 删除占位符文本
   - 插入相应的图片/表格/资质扫描件
   - 调整大小和位置

4. **最终检查**
   - 检查所有占位符是否已替换
   - 检查格式是否统一
   - 检查内容是否完整

---

## 注意事项

### 文件格式
- **输入**：支持PDF、DOCX、TXT、XLSX格式
- **输出**：Word文档（.docx格式）
- **扫描版PDF**：需要先进行OCR识别

### 内容质量
- **人工审核**：建议对生成内容进行人工审核
- **数据准确性**：技术参数、项目数据需要核实
- **格式一致性**：检查标题层级、字体字号是否统一

### 占位符处理
- **及时替换**：生成后尽快替换占位符
- **素材准备**：提前准备好图片、表格、资质文件
- **格式调整**：替换后调整格式保持一致

---

## 版本历史

| 版本 | 日期 | 更新内容 |
|------|------|---------|
| v2.2 | 2026-05-22 | 修复输出目录结构：明确时间戳文件夹说明、修复封面和目录保存路径 |
| v2.1 | 2026-05-22 | 添加行业指南自动创建功能：如果行业指南不存在，AI自动检索并创建，逐步完善SKILL |
| v2.0 | 2026-05-22 | 改造为通用行业版：支持8个行业、自动行业检测、行业特定内容生成 |
| v1.9 | 2026-05-21 | 优化内容生成逻辑（更详细丰富）、中间文件保存到debug文件夹、最终标书与招标文件同目录 |
| v1.8 | 2026-05-21 | 修复PDF转换编码问题，添加fix_encoding函数自动检测和修复乱码 |
| v1.7 | 2026-05-21 | 添加详细的任务决策树，包含18种场景和分步执行指南 |
| v1.6 | 2026-05-21 | 修复所有脚本的编码设置问题，避免I/O错误 |
| v1.5 | 2026-05-21 | 删除README.md（避免干扰Trea导入）、精简文件结构 |
| v1.4 | 2026-05-21 | 清理无用文件（删除references文件夹、__pycache__）、整理md文档、更新文档结构 |
| v1.3 | 2026-05-21 | 新增DOC/DOCX转Markdown功能、扩展templates知识库（包含45个DOC/DOCX转换的知识文档）、修复编码问题 |
| v1.2 | 2026-05-21 | 新增PDF转Markdown功能、扩展templates知识库（包含招投标知识文档）、更新模板知识读取逻辑 |
| v1.1 | 2026-05-21 | 新增Markdown转换功能、.doc支持、编码兼容性改进、环境检查工具 |
| v1.0 | 2026-05-21 | 初始版本，整合标书生成功能 |

---

## 技术支持

### 常见问题

**Q: 扫描版PDF无法解析怎么办？**
A: 使用OCR工具（如ABBYY FineReader、Adobe Acrobat）先进行文字识别，然后重新提交。

**Q: 生成的标书字数不达标怎么办？**
A: 系统会自动检查字数，如果不达标会提示需要补充的内容章节。

**Q: 如何修改格式标准？**
A: 可以通过 `--format-standard` 参数选择不同的格式标准，或修改 `scripts/format_docx.py` 中的 `FORMAT_STANDARDS` 配置自定义格式。

**Q: 占位符如何批量替换？**
A: 使用Word的查找替换功能（Ctrl+H），搜索占位符类型进行批量定位。

---

## 相关文档

- `templates/00_标书制作指南.md` - 标书制作完整指南
- `templates/bid_doc_*.md` - 标书知识文档
- `templates/bid_knowledge_*.md` - 招投标知识文档
