#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书大纲生成脚本
根据评分标准生成标书大纲结构
版本：v1.0 (2026-05-21)
"""

import io
import json
import os
import sys
from pathlib import Path
from typing import Dict, List

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


class OutlineGenerator:
    """标书大纲生成器"""

    # 默认章节结构
    DEFAULT_CHAPTERS = [
        {"id": 1, "title": "项目概述", "sections": ["项目背景", "项目目标", "项目范围"]},
        {"id": 2, "title": "技术方案", "sections": ["总体架构", "功能设计", "数据设计", "安全设计"]},
        {"id": 3, "title": "项目实施", "sections": ["实施计划", "人员配置", "质量保证", "风险控制"]},
        {"id": 4, "title": "项目业绩", "sections": ["类似项目经验", "客户评价"]},
        {"id": 5, "title": "售后服务", "sections": ["服务体系", "响应时间", "培训计划"]},
        {"id": 6, "title": "附件", "sections": ["企业资质", "人员证书", "项目业绩证明"]}
    ]

    def __init__(self, scoring_criteria: Dict, requirements: Dict = None):
        self.scoring = scoring_criteria
        self.requirements = requirements or {}
        self.outline = []

    def generate(self) -> Dict:
        """生成标书大纲"""
        print("[INFO] 开始生成标书大纲...")

        # 获取评分结构
        scoring_structure = self.scoring.get('scoring_structure', {})
        categories = scoring_structure.get('categories', [])
        total_score = scoring_structure.get('total_score', 100)

        # 根据评分结构生成大纲
        if categories:
            self._generate_from_scoring(categories, total_score)
        else:
            self._generate_default()

        # 添加附件章节
        self._add_appendix()

        result = {
            "outline": self.outline,
            "metadata": {
                "total_chapters": len(self.outline),
                "total_sections": sum(len(ch.get('sections', [])) for ch in self.outline),
                "total_score": total_score
            }
        }

        print(f"[INFO] 大纲生成完成:")
        print(f"  - 章节数: {result['metadata']['total_chapters']}")
        print(f"  - 小节数: {result['metadata']['total_sections']}")

        return result

    def _generate_from_scoring(self, categories: List[Dict], total_score: int):
        """根据评分结构生成大纲"""
        chapter_id = 1

        for category in categories:
            cat_name = category.get('name', '')
            cat_score = category.get('score', 0)
            items = category.get('items', [])

            # 根据评分大类生成章节
            if '技术' in cat_name:
                chapter = self._generate_technical_chapter(chapter_id, cat_name, cat_score, items)
            elif '商务' in cat_name:
                chapter = self._generate_business_chapter(chapter_id, cat_name, cat_score, items)
            elif '价格' in cat_name:
                chapter = self._generate_price_chapter(chapter_id, cat_name, cat_score, items)
            else:
                chapter = self._generate_generic_chapter(chapter_id, cat_name, cat_score, items)

            self.outline.append(chapter)
            chapter_id += 1

    def _generate_technical_chapter(self, chapter_id: int, cat_name: str, cat_score: float, items: List[Dict]) -> Dict:
        """生成技术方案章节"""
        sections = []

        # 根据评分项生成小节
        for i, item in enumerate(items, 1):
            item_name = item.get('name', f'技术点{i}')
            sections.append({
                "id": f"{chapter_id}.{i}",
                "title": item_name,
                "score": item.get('score', 0),
                "keywords": item.get('keywords', [])
            })

        # 如果没有评分项，使用默认结构
        if not sections:
            sections = [
                {"id": f"{chapter_id}.1", "title": "总体架构设计", "score": cat_score * 0.3},
                {"id": f"{chapter_id}.2", "title": "功能模块设计", "score": cat_score * 0.3},
                {"id": f"{chapter_id}.3", "title": "技术方案详细说明", "score": cat_score * 0.4}
            ]

        return {
            "id": chapter_id,
            "title": "技术方案",
            "score": cat_score,
            "sections": sections
        }

    def _generate_business_chapter(self, chapter_id: int, cat_name: str, cat_score: float, items: List[Dict]) -> Dict:
        """生成商务部分章节"""
        sections = []

        for i, item in enumerate(items, 1):
            item_name = item.get('name', f'商务点{i}')
            sections.append({
                "id": f"{chapter_id}.{i}",
                "title": item_name,
                "score": item.get('score', 0),
                "keywords": item.get('keywords', [])
            })

        if not sections:
            sections = [
                {"id": f"{chapter_id}.1", "title": "项目实施方案", "score": cat_score * 0.4},
                {"id": f"{chapter_id}.2", "title": "项目管理方案", "score": cat_score * 0.3},
                {"id": f"{chapter_id}.3", "title": "质量保证措施", "score": cat_score * 0.3}
            ]

        return {
            "id": chapter_id,
            "title": "项目实施与管理",
            "score": cat_score,
            "sections": sections
        }

    def _generate_price_chapter(self, chapter_id: int, cat_name: str, cat_score: float, items: List[Dict]) -> Dict:
        """生成价格部分章节"""
        sections = [
            {"id": f"{chapter_id}.1", "title": "报价说明", "score": cat_score * 0.5},
            {"id": f"{chapter_id}.2", "title": "费用明细", "score": cat_score * 0.5}
        ]

        return {
            "id": chapter_id,
            "title": "报价方案",
            "score": cat_score,
            "sections": sections
        }

    def _generate_generic_chapter(self, chapter_id: int, cat_name: str, cat_score: float, items: List[Dict]) -> Dict:
        """生成通用章节"""
        sections = []

        for i, item in enumerate(items, 1):
            item_name = item.get('name', f'要点{i}')
            sections.append({
                "id": f"{chapter_id}.{i}",
                "title": item_name,
                "score": item.get('score', 0),
                "keywords": item.get('keywords', [])
            })

        if not sections:
            sections = [
                {"id": f"{chapter_id}.1", "title": f"{cat_name}详细说明", "score": cat_score}
            ]

        return {
            "id": chapter_id,
            "title": cat_name,
            "score": cat_score,
            "sections": sections
        }

    def _generate_default(self):
        """生成默认大纲"""
        for i, chapter in enumerate(self.DEFAULT_CHAPTERS, 1):
            sections = []
            for j, section_title in enumerate(chapter['sections'], 1):
                sections.append({
                    "id": f"{i}.{j}",
                    "title": section_title,
                    "score": 0
                })

            self.outline.append({
                "id": i,
                "title": chapter['title'],
                "score": 0,
                "sections": sections
            })

    def _add_appendix(self):
        """添加附件章节"""
        appendix_sections = [
            {"id": f"{len(self.outline) + 1}.1", "title": "企业营业执照", "score": 0, "is_placeholder": True},
            {"id": f"{len(self.outline) + 1}.2", "title": "企业资质证书", "score": 0, "is_placeholder": True},
            {"id": f"{len(self.outline) + 1}.3", "title": "项目业绩证明", "score": 0, "is_placeholder": True},
            {"id": f"{len(self.outline) + 1}.4", "title": "人员资格证书", "score": 0, "is_placeholder": True}
        ]

        self.outline.append({
            "id": len(self.outline) + 1,
            "title": "附件",
            "score": 0,
            "sections": appendix_sections
        })


def generate_outline_from_scoring(scoring_file: str, requirements_file: str = None) -> Dict:
    """从评分标准文件生成大纲"""
    # 读取评分标准
    with open(scoring_file, 'r', encoding='utf-8') as f:
        scoring = json.load(f)

    # 读取关键要求（可选）
    requirements = None
    if requirements_file and os.path.exists(requirements_file):
        with open(requirements_file, 'r', encoding='utf-8') as f:
            requirements = json.load(f)

    # 生成大纲
    generator = OutlineGenerator(scoring, requirements)
    return generator.generate()


def main():
    import argparse

    parser = argparse.ArgumentParser(description='标书大纲生成工具')
    parser.add_argument('scoring_file', help='评分标准JSON文件路径')
    parser.add_argument('--requirements', '-r', help='关键要求JSON文件路径（可选）')
    parser.add_argument('--output', '-o', default='outline.json', help='输出大纲文件路径')

    args = parser.parse_args()

    if not os.path.exists(args.scoring_file):
        print(f"[ERROR] 文件不存在: {args.scoring_file}")
        sys.exit(1)

    print(f"[INFO] 正在生成大纲...")

    result = generate_outline_from_scoring(args.scoring_file, args.requirements)

    # 保存结果
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"[INFO] 大纲已保存至: {output_path}")

    # 打印大纲结构
    print("\n大纲结构:")
    for chapter in result['outline']:
        print(f"  第{chapter['id']}章 {chapter['title']} ({chapter['score']}分)")
        for section in chapter.get('sections', []):
            print(f"    {section['id']} {section['title']}")


if __name__ == '__main__':
    main()
