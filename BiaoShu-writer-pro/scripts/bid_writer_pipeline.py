#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书生成主流程脚本
整合所有功能，实现从招标文件到标书Word文档的完整流程
版本：v1.2 (2026-05-21)
"""

import argparse
import io
import json
import os
import sys
from pathlib import Path
from datetime import datetime

# 设置标准输出编码（解决Windows控制台编码问题）
import locale
try:
    # 尝试设置控制台编码
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from parse_bid_file import parse_file
from extract_scoring import ScoringExtractor
from extract_requirements import RequirementsExtractor
from extract_bid_name import BidNameExtractor
from generate_outline import OutlineGenerator
from generate_content import ContentGenerator
from check_word_count import check_all_chapters
from humanizer import Humanizer
from convert_to_docx import convert_md_to_word
from convert_to_md import convert_to_markdown
from format_docx import format_docx
from generate_placeholder import process_chapters
from generate_cover import generate_cover
from generate_toc import generate_toc
from industry_detector import IndustryDetector, detect_industry_from_text
from industry_content_generator import IndustryContentGenerator


class BidWriterPipeline:
    """标书生成主流程"""

    def __init__(self, bid_file: str, output_dir: str = "./output",
                 format_standard: str = "government",
                 company_profile: str = None,
                 total_pages: int = 40,
                 industry: str = None):
        self.bid_file = bid_file
        self.base_output_dir = Path(output_dir)
        self.format_standard = format_standard
        self.company_profile = company_profile
        self.total_pages = total_pages
        self.industry = industry  # 行业代码，如果为None则自动检测

        # 行业检测器
        self.industry_detector = IndustryDetector()
        self.industry_config = None
        self.industry_content_generator = None

        # 生成带时间戳和招标文件名称的文件夹
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        bid_file_stem = Path(bid_file).stem  # 获取招标文件名（不含扩展名）
        # 清理文件名中的特殊字符
        bid_file_stem = "".join(c for c in bid_file_stem if c.isalnum() or c in (' ', '-', '_')).strip()
        if not bid_file_stem:
            bid_file_stem = "bid_document"

        # 创建唯一的输出目录
        self.output_dir = self.base_output_dir / f"{timestamp}_{bid_file_stem}"
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # 创建debug文件夹
        self.debug_dir = self.output_dir / "debug"
        self.debug_dir.mkdir(exist_ok=True)

        # 中间结果
        self.bid_content = None
        self.scoring = None
        self.requirements = None
        self.bid_name = None
        self.outline = None
        self.chapters = None
        self.final_docx_path = None

    def run(self) -> str:
        """执行完整流程"""
        print("=" * 60)
        print("🚀 标书智能生成专家 v1.9")
        print("=" * 60)
        print(f"招标文件: {self.bid_file}")
        print(f"输出目录: {self.output_dir}")
        print(f"调试目录: {self.debug_dir}")
        print(f"格式标准: {self.format_standard}")
        print("=" * 60)

        try:
            # 步骤0：通读模板知识
            self._step0_read_templates()

            # 步骤1：解析招标文件
            self._step1_parse_bid_file()

            # 步骤1.5：检测行业类型
            self._step1_5_detect_industry()

            # 步骤2：提取评分标准
            self._step2_extract_scoring()

            # 步骤3：提取关键要求
            self._step3_extract_requirements()

            # 步骤4：提取标书名称
            self._step4_extract_bid_name()

            # 步骤5：生成大纲
            self._step5_generate_outline()

            # 步骤6：生成内容
            self._step6_generate_content()

            # 步骤7：字数检查
            self._step7_check_word_count()

            # 步骤8：AI去痕处理
            self._step8_humanize_content()

            # 步骤9：插入占位符
            self._step9_insert_placeholders()

            # 步骤10：生成封面
            self._step10_generate_cover()

            # 步骤11：生成目录
            self._step11_generate_toc()

            # 步骤12：合并生成Word文档
            self._step12_generate_docx()

            # 步骤13：格式规范化
            self._step13_format_docx()

            print("\n" + "=" * 60)
            print("✅ 标书生成完成！")
            print(f"📁 文件名: {self.bid_name}")
            print(f"📄 最终标书: {self.final_docx_path}")
            print(f"📂 调试目录: {self.debug_dir}")
            print("=" * 60)

            return str(self.final_docx_path)

        except Exception as e:
            print(f"\n❌ 生成失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def _step0_read_templates(self):
        """步骤0：通读模板知识"""
        print("\n[步骤0/14] 通读模板知识...")

        # 获取templates文件夹路径
        script_dir = Path(__file__).parent.parent
        templates_dir = script_dir / "templates"

        if not templates_dir.exists():
            print("  [WARNING] templates文件夹不存在，跳过模板知识学习")
            return

        # 读取所有md文件
        md_files = list(templates_dir.glob("*.md"))

        if not md_files:
            print("  [WARNING] templates文件夹中没有md文件，跳过模板知识学习")
            return

        print(f"  [INFO] 发现 {len(md_files)} 个模板知识文件")

        # 读取并合并所有模板知识
        template_knowledge = []
        for md_file in sorted(md_files):
            try:
                with open(md_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                    template_knowledge.append(f"=== {md_file.name} ===\n{content}")
                    print(f"  [INFO] 已读取: {md_file.name}")
            except Exception as e:
                print(f"  [WARNING] 读取失败 {md_file.name}: {str(e)}")

        # 保存模板知识到输出目录
        knowledge_path = self.debug_dir / "template_knowledge.md"
        with open(knowledge_path, 'w', encoding='utf-8') as f:
            f.write("\n\n".join(template_knowledge))

        print(f"  ✅ 模板知识已加载，共 {len(template_knowledge)} 个文件")
        print(f"  [INFO] 模板知识已保存到: {knowledge_path}")

    def _step1_parse_bid_file(self):
        """步骤1：解析招标文件并转换为Markdown格式"""
        print("\n[步骤1/14] 解析招标文件...")

        # 获取文件扩展名
        ext = os.path.splitext(self.bid_file)[1].lower()

        # 如果不是.txt格式，先转换为Markdown
        if ext != '.txt':
            print(f"  [INFO] 检测到{ext}格式，正在转换为Markdown...")
            md_path = self.debug_dir / "bid_document.md"
            md_content = convert_to_markdown(self.bid_file, str(md_path))

            if md_content.startswith("[错误]"):
                print(f"  [WARNING] Markdown转换失败，使用原始解析方式")
                self.bid_content = parse_file(self.bid_file)
            else:
                self.bid_content = md_content
                print(f"  [INFO] 已转换为Markdown格式: {md_path}")
        else:
            self.bid_content = parse_file(self.bid_file)

        if self.bid_content.startswith("[错误]"):
            raise Exception(f"招标文件解析失败: {self.bid_content}")

        # 保存原始内容
        debug_path = self.debug_dir / "raw_bid_content.txt"
        with open(debug_path, 'w', encoding='utf-8') as f:
            f.write(self.bid_content)

        # 统计信息
        lines = self.bid_content.split('\n')
        print(f"  ✅ 解析完成")
        print(f"  [INFO] 文本长度: {len(self.bid_content)} 字符")
        print(f"  [INFO] 行数: {len(lines)} 行")

    def _step1_5_detect_industry(self):
        """步骤1.5：检测行业类型并检查行业指南"""
        print("\n[步骤1.5/14] 检测行业类型并检查行业指南...")

        # 如果用户指定了行业，使用用户指定的行业
        if self.industry:
            industry_code = self.industry
            confidence = 1.0
            config = self.industry_detector.get_industry_config(industry_code)
            print(f"  [INFO] 使用用户指定的行业: {config.get('name', industry_code)}")
        else:
            # 自动检测行业
            industry_code, confidence, config = detect_industry_from_text(self.bid_content)
            print(f"  [INFO] 自动检测行业: {config.get('name', industry_code)} (置信度: {confidence:.2%})")

        # 保存行业信息
        self.industry_config = config
        self.industry_content_generator = IndustryContentGenerator(industry_code)

        # 检查行业指南是否存在
        industry_name = config.get("name", "")
        guide_exists = self._check_industry_guide(industry_code, industry_name)

        # 如果行业指南不存在，创建新的行业指南
        if not guide_exists:
            print(f"  [INFO] 未找到 {industry_name} 行业的指南文件，正在创建...")
            self._create_industry_guide(industry_code, industry_name)
        else:
            print(f"  [INFO] 已找到 {industry_name} 行业的指南文件")

        # 保存行业检测结果到debug文件夹
        industry_info = {
            "industry_code": industry_code,
            "industry_name": industry_name,
            "confidence": confidence,
            "chapter_structure": self.industry_detector.get_chapter_structure(industry_code),
            "key_sections": self.industry_detector.get_key_sections(industry_code),
            "guide_exists": guide_exists
        }

        industry_path = self.debug_dir / "industry_detection.json"
        with open(industry_path, 'w', encoding='utf-8') as f:
            json.dump(industry_info, f, ensure_ascii=False, indent=2)

        print(f"  ✅ 行业检测完成")
        print(f"  [INFO] 行业代码: {industry_code}")
        print(f"  [INFO] 行业名称: {industry_name}")
        print(f"  [INFO] 章节结构: {len(self.industry_detector.get_chapter_structure(industry_code))} 个章节")

    def _check_industry_guide(self, industry_code: str, industry_name: str) -> bool:
        """
        检查行业指南是否存在

        Args:
            industry_code: 行业代码
            industry_name: 行业名称

        Returns:
            bool: 是否存在
        """
        templates_dir = Path(__file__).parent.parent / "templates"

        # 检查行业指南文件
        guide_patterns = [
            f"industry_{industry_name}.md",
            f"industry_{industry_code}.md",
            f"{industry_name}.md",
            f"{industry_code}.md"
        ]

        for pattern in guide_patterns:
            guide_path = templates_dir / pattern
            if guide_path.exists():
                return True

        return False

    def _create_industry_guide(self, industry_code: str, industry_name: str):
        """
        创建行业指南

        Args:
            industry_code: 行业代码
            industry_name: 行业名称
        """
        templates_dir = Path(__file__).parent.parent / "templates"
        guide_path = templates_dir / f"industry_{industry_name}.md"

        # 生成行业指南内容
        guide_content = self._generate_industry_guide_content(industry_code, industry_name)

        # 保存到templates文件夹
        with open(guide_path, 'w', encoding='utf-8') as f:
            f.write(guide_content)

        print(f"  ✅ 已创建行业指南: {guide_path}")
        print(f"  [INFO] SKILL已完善，下次运行将直接使用该指南")

    def _generate_industry_guide_content(self, industry_code: str, industry_name: str) -> str:
        """
        生成行业指南内容

        Args:
            industry_code: 行业代码
            industry_name: 行业名称

        Returns:
            str: 行业指南内容
        """
        # 获取行业配置
        config = self.industry_detector.get_industry_config(industry_code)
        chapter_structure = config.get("chapter_structure", [])
        key_sections = config.get("key_sections", {})
        keywords = config.get("keywords", [])

        # 生成指南内容
        content = f"""# {industry_name}行业投标文件编写指南

## 一、行业特点

{industry_name}行业投标文件具有以下特点：

1. **专业性强**：涉及{', '.join(keywords[:3])}等专业领域
2. **规范性高**：需要符合行业标准和规范
3. **细节要求多**：技术方案、实施方案、质量保证等
4. **资质要求严**：需要具备相应的行业资质和认证

## 二、章节结构

"""

        # 生成章节结构
        for i, chapter in enumerate(chapter_structure, 1):
            content += f"### 2.{i} {chapter}\n"
            if chapter in key_sections:
                for section in key_sections[chapter]:
                    content += f"- {section}\n"
            content += "\n"

        content += """## 三、关键内容编写要点

"""

        # 生成关键内容要点
        for chapter, sections in key_sections.items():
            content += f"### 3.{list(key_sections.keys()).index(chapter) + 1} {chapter}\n"
            for section in sections:
                content += f"- {section}要详细，具有可操作性\n"
            content += "\n"

        content += """## 四、常见问题及解决方法

### 4.1 技术方案不详细
- 问题：技术方案过于简单，缺乏可操作性
- 解决：详细描述技术方案、实施步骤、质量措施

### 4.2 质量措施不具体
- 问题：质量措施泛泛而谈，缺乏针对性
- 解决：针对行业特点，制定具体的质量控制措施

### 4.3 安全措施不完善
- 问题：安全措施不全面，存在安全隐患
- 解决：全面分析安全风险，制定完善的安全措施

## 五、评分标准关注点

### 5.1 技术部分
- 技术方案的可行性
- 实施方案的合理性
- 质量保证的有效性
- 安全管理的规范性

### 5.2 商务部分
- 企业资质和业绩
- 项目团队配置
- 价格合理性
- 服务承诺

### 5.3 价格部分
- 报价的合理性
- 成本分析的科学性
- 价格组成的清晰性

## 六、注意事项

1. **资质要求**：确保具备相应的行业资质和认证
2. **人员要求**：项目团队要具备相应的专业资质和经验
3. **业绩要求**：类似项目业绩要真实、有效
4. **格式要求**：严格按照招标文件要求的格式编写
5. **时间要求**：按时提交投标文件，避免逾期

## 七、参考资料

1. 行业相关法规和标准
2. 行业优秀标书范本
3. 行业专家经验总结
"""

        return content

    def _step2_extract_scoring(self):
        """步骤2：提取评分标准"""
        print("\n[步骤2/14] 提取评分标准...")
        extractor = ScoringExtractor(self.bid_content)
        self.scoring = extractor.extract()

        if not self.scoring:
            print("  [WARNING] 未找到评分标准，使用默认评分")
            self.scoring = {
                "total_score": 100,
                "categories": [
                    {"name": "技术部分", "score": 60, "weight": 0.6},
                    {"name": "商务部分", "score": 20, "weight": 0.2},
                    {"name": "报价部分", "score": 20, "weight": 0.2}
                ]
            }

        # 保存评分标准到debug文件夹
        scoring_path = self.debug_dir / "scoring_criteria.json"
        with open(scoring_path, 'w', encoding='utf-8') as f:
            json.dump(self.scoring, f, ensure_ascii=False, indent=2)

        print(f"  ✅ 评分标准提取完成")
        print(f"  [INFO] 总分: {self.scoring.get('total_score', '未知')}")
        for cat in self.scoring.get('categories', []):
            print(f"  [INFO] {cat['name']}: {cat['score']}分")

    def _step3_extract_requirements(self):
        """步骤3：提取关键要求"""
        print("\n[步骤3/14] 提取关键要求...")
        extractor = RequirementsExtractor(self.bid_content)
        self.requirements = extractor.extract()

        # 保存关键要求到debug文件夹
        requirements_path = self.debug_dir / "key_requirements.json"
        with open(requirements_path, 'w', encoding='utf-8') as f:
            json.dump(self.requirements, f, ensure_ascii=False, indent=2)

        print(f"  ✅ 关键要求提取完成")
        print(f"  [INFO] 资质要求: {len(self.requirements.get('qualifications', []))} 项")
        print(f"  [INFO] 技术要求: {len(self.requirements.get('technical', []))} 项")
        print(f"  [INFO] 商务要求: {len(self.requirements.get('commercial', []))} 项")

    def _step4_extract_bid_name(self):
        """步骤4：提取标书名称"""
        print("\n[步骤4/14] 提取标书名称...")
        extractor = BidNameExtractor(self.bid_file, self.bid_content)
        self.bid_name = extractor.extract()

        if not self.bid_name:
            # 使用文件名作为备选
            base_name = os.path.splitext(os.path.basename(self.bid_file))[0]
            self.bid_name = f"{base_name}技术标.docx"

        print(f"  ✅ 标书名称提取完成")
        print(f"  [INFO] 文件名: {self.bid_name}")

    def _step5_generate_outline(self):
        """步骤5：生成大纲"""
        print("\n[步骤5/14] 生成大纲...")
        generator = OutlineGenerator(self.scoring, self.requirements)
        self.outline = generator.generate()

        # 保存大纲到debug文件夹
        outline_path = self.debug_dir / "outline.json"
        with open(outline_path, 'w', encoding='utf-8') as f:
            json.dump(self.outline, f, ensure_ascii=False, indent=2)

        print(f"  ✅ 大纲生成完成")
        print(f"  [INFO] 章节数: {len(self.outline.get('chapters', []))}")
        for chapter in self.outline.get('chapters', []):
            print(f"  [INFO] {chapter['title']}")

    def _step6_generate_content(self):
        """步骤6：生成内容"""
        print("\n[步骤6/14] 生成内容...")
        generator = ContentGenerator(
            self.outline,
            self.scoring,
            self.requirements,
            self.bid_content
        )
        self.chapters = generator.generate()

        # 保存各章节内容到debug文件夹
        for i, chapter in enumerate(self.chapters, 1):
            chapter_path = self.debug_dir / "chapters" / f"{i:02d}_{chapter['title']}.md"
            chapter_path.parent.mkdir(parents=True, exist_ok=True)
            with open(chapter_path, 'w', encoding='utf-8') as f:
                f.write(chapter['content'])

        print(f"  ✅ 内容生成完成")
        print(f"  [INFO] 生成章节数: {len(self.chapters)}")
        print(f"  [INFO] 章节文件保存在: {self.output_dir / 'debug' / 'chapters'}")

    def _step7_check_word_count(self):
        """步骤7：字数检查"""
        print("\n[步骤7/14] 字数检查...")
        results = check_all_chapters(self.chapters, self.scoring, self.total_pages)

        # 保存检查结果
        results_path = self.debug_dir / "word_count_check.json"
        with open(results_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)

        print(f"  ✅ 字数检查完成")
        for result in results.get('details', []):
            status = "✅" if result['passed'] else "❌"
            print(f"  {status} {result['chapter']}: {result['actual']}字 (目标: {result['target']}字)")

    def _step8_humanize_content(self):
        """步骤8：AI去痕处理"""
        print("\n[步骤8/14] AI去痕处理...")
        humanizer = Humanizer()

        for i, chapter in enumerate(self.chapters):
            chapter['content'] = humanizer.process(chapter['content'])

        # 保存去痕后的内容
        for i, chapter in enumerate(self.chapters, 1):
            chapter_path = self.debug_dir / f"humanized_{i:02d}_{chapter['title']}.md"
            with open(chapter_path, 'w', encoding='utf-8') as f:
                f.write(chapter['content'])

        print(f"  ✅ AI去痕处理完成")

    def _step9_insert_placeholders(self):
        """步骤9：插入占位符"""
        print("\n[步骤9/14] 插入占位符...")
        self.chapters = process_chapters(self.chapters, self.requirements)

        print(f"  ✅ 占位符插入完成")

    def _step10_generate_cover(self):
        """步骤10：生成封面"""
        print("\n[步骤10/14] 生成封面...")
        cover_path = self.debug_dir / "chapters" / "00_封面.md"
        cover_content = generate_cover(self.bid_name, self.requirements)

        with open(cover_path, 'w', encoding='utf-8') as f:
            f.write(cover_content)

        print(f"  ✅ 封面生成完成")
        print(f"  [INFO] 封面文件: {cover_path}")

    def _step11_generate_toc(self):
        """步骤11：生成目录"""
        print("\n[步骤11/14] 生成目录...")
        toc_path = self.debug_dir / "chapters" / "00_目录.md"
        toc_content = generate_toc(self.chapters)

        with open(toc_path, 'w', encoding='utf-8') as f:
            f.write(toc_content)

        print(f"  ✅ 目录生成完成")
        print(f"  [INFO] 目录文件: {toc_path}")

    def _step12_generate_docx(self):
        """步骤12：合并生成Word文档"""
        print("\n[步骤12/14] 生成Word文档...")

        # 合并所有章节内容
        all_content = []
        for chapter in self.chapters:
            all_content.append(chapter['content'])

        # 获取招标文件所在目录
        bid_file_dir = Path(self.bid_file).parent

        # 生成Word文档到招标文件所在目录
        docx_path = bid_file_dir / self.bid_name
        convert_md_to_word('\n\n'.join(all_content), str(docx_path), self.format_standard)

        # 保存最终文件路径
        self.final_docx_path = docx_path

        print(f"  ✅ Word文档生成完成")
        print(f"  [INFO] 文件路径: {docx_path}")
        print(f"  [INFO] 与招标文件同目录: {bid_file_dir}")

    def _step13_format_docx(self):
        """步骤13：格式规范化"""
        print("\n[步骤13/14] 格式规范化...")

        # 使用最终文件路径
        docx_path = self.final_docx_path
        format_docx(str(docx_path), self.format_standard)

        print(f"  ✅ 格式规范化完成")
        print(f"  [INFO] 最终文件: {docx_path}")


def main():
    parser = argparse.ArgumentParser(description='标书智能生成工具')
    parser.add_argument('bid_file', help='招标文件路径')
    parser.add_argument('--output-dir', '-o', default='./output', help='输出目录')
    parser.add_argument('--format-standard', '-f', default='government',
                       choices=['government', 'enterprise', 'highway'],
                       help='格式标准')
    parser.add_argument('--company-profile', '-c', help='企业资质文件路径')
    parser.add_argument('--total-pages', '-p', type=int, default=40, help='预计总页数')
    parser.add_argument('--industry', '-i', help='行业代码（可选，如不指定则自动检测）')
    parser.add_argument('--list-industries', '-l', action='store_true', help='列出所有支持的行业')

    args = parser.parse_args()

    # 如果用户要求列出行业
    if args.list_industries:
        detector = IndustryDetector()
        print("\n支持的行业列表：")
        print("=" * 60)
        for industry in detector.list_industries():
            print(f"  {industry['code']}: {industry['name']}")
            print(f"    {industry['description']}")
        return

    pipeline = BidWriterPipeline(
        bid_file=args.bid_file,
        output_dir=args.output_dir,
        format_standard=args.format_standard,
        company_profile=args.company_profile,
        total_pages=args.total_pages,
        industry=args.industry
    )

    result = pipeline.run()
    if result:
        print(f"\n🎉 标书生成成功！文件保存在: {result}")
    else:
        print(f"\n❌ 标书生成失败！")
        sys.exit(1)


if __name__ == '__main__':
    main()