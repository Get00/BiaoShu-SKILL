#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
评分标准提取脚本
从招标文件中自动提取评分标准、评分项、分值权重等信息
版本：v1.0 (2026-05-21)
"""

import io
import json
import re
import os
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

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


class ScoringExtractor:
    """评分标准提取器"""

    # 评分办法相关关键词
    SCORING_KEYWORDS = [
        '评标办法', '评分办法', '评分标准', '评分细则',
        '评标标准', '评标细则', '综合评分', '技术评分',
        '商务评分', '价格评分', '评分因素', '评分项目'
    ]

    # 评分方法关键词
    SCORING_METHODS = {
        '客观评分': ['客观', '量化', '硬性指标', '满足即得分'],
        '主观评分': ['主观', '综合评价', '评审委员会', '酌情'],
        '价格评分': ['价格', '报价', '低价优先', '基准价'],
    }

    # 评分项模式
    ITEM_PATTERNS = [
        r'(\d+)[\.\s]+([^\d\n]+?)[\.\s]+(\d+)[\s]*分',
        r'([一二三四五六七八九十])[\.\s]+([^\d\n]+?)[\.\s]+(\d+)[\s]*分',
        r'([\(（]\d+[\)）])\s*([^\d\n]+?)[\.\s]+(\d+)[\s]*分',
    ]

    def __init__(self, text: str):
        self.text = text
        self.scoring_section = ""
        self.categories = []
        self.document_info = {}

    def find_scoring_section(self) -> Tuple[int, int, str]:
        """查找评分办法章节"""
        lines = self.text.split('\n')
        start_idx = -1
        end_idx = -1

        for i, line in enumerate(lines):
            # 检查是否是评分办法章节标题
            if any(keyword in line for keyword in self.SCORING_KEYWORDS):
                if start_idx == -1:
                    start_idx = i
                    print(f"[INFO] 找到评分办法章节: {line.strip()}")

            # 如果已经找到开始，检查是否是下一个主要章节的开始
            if start_idx != -1 and i > start_idx:
                if re.match(r'第[一二三四五六七八九十\d]+章', line) or \
                   re.match(r'[一二三四五六七八九十]、[^评分]', line) or \
                   re.match(r'\d+\.\s+[\u4e00-\u9fa5]{2,}', line):
                    if not any(kw in line for kw in self.SCORING_KEYWORDS):
                        end_idx = i
                        break

        if start_idx == -1:
            return 0, 0, ""

        if end_idx == -1:
            end_idx = len(lines)

        section_content = '\n'.join(lines[start_idx:end_idx])
        return start_idx, end_idx, section_content

    def extract_document_info(self) -> Dict:
        """提取文档基本信息"""
        info = {
            "project_name": self._extract_project_name(),
            "bid_deadline": self._extract_bid_deadline(),
            "total_budget": self._extract_budget(),
        }
        self.document_info = info
        return info

    def _extract_project_name(self) -> str:
        """提取项目名称"""
        patterns = [
            r'项目名称[：:]\s*([^\n]+)',
            r'招标项目[：:]\s*([^\n]+)',
            r'(.{3,30}项目).{0,5}招标',
        ]
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1).strip()
        return "未知项目"

    def _extract_bid_deadline(self) -> str:
        """提取投标截止时间"""
        patterns = [
            r'投标截止[时间]?[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)',
            r'截止[时间]?[：:]\s*(\d{4}[\.\-/]\d{1,2}[\.\-/]\d{1,2})',
            r'(\d{4}年\d{1,2}月\d{1,2}日).{0,10}开标',
        ]
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1).strip()
        return "未知"

    def _extract_budget(self) -> str:
        """提取预算金额"""
        patterns = [
            r'预算[金额]?[：:]\s*([\d,\.]+\s*[万元]+)',
            r'最高限价[：:]\s*([\d,\.]+\s*[万元]+)',
            r'采购预算[：:]\s*([\d,\.]+)',
            r'([\d,\.]+)\s*万元',
        ]
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1).strip()
        return "未知"

    def extract_categories(self):
        """识别评分大类（技术、商务、价格等）"""
        category_patterns = [
            r'([技术商务价格]+)[部类]?[分]?[^\d]*(\d+)[\s]*分',
            r'([技术商务价格]+)评分[^\d]*(\d+)[\s]*分',
        ]

        lines = self.scoring_section.split('\n')
        total_score = 0

        for line in lines:
            line = line.strip()
            if not line:
                continue

            for pattern in category_patterns:
                match = re.search(pattern, line)
                if match:
                    name = match.group(1)
                    score = float(match.group(2))
                    total_score += score

                    category = {
                        "name": name,
                        "score": score,
                        "weight": 0,
                        "items": []
                    }
                    self.categories.append(category)
                    print(f"[INFO] 识别评分大类: {name} ({score}分)")
                    break

        # 计算权重
        if total_score > 0:
            for cat in self.categories:
                cat["weight"] = round(cat["score"] / total_score, 2)

    def extract_scoring_items(self):
        """提取具体评分项"""
        if not self.categories:
            self._extract_items_without_categories()
            return

        lines = self.scoring_section.split('\n')
        current_category_idx = -1
        item_counter = 0

        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue

            # 检查是否是大类标题行
            for idx, cat in enumerate(self.categories):
                if cat["name"] in line and str(int(cat["score"])) in line:
                    current_category_idx = idx
                    break

            # 提取评分项
            if current_category_idx >= 0:
                item = self._parse_scoring_item(line, item_counter)
                if item:
                    self.categories[current_category_idx]["items"].append(item)
                    item_counter += 1

    def _parse_scoring_item(self, line: str, counter: int) -> Optional[Dict]:
        """解析单行评分项"""
        for pattern in self.ITEM_PATTERNS:
            match = re.search(pattern, line)
            if match:
                item_id = match.group(1)
                name = match.group(2).strip()
                score = float(match.group(3))

                method = self._detect_scoring_method(line)
                keywords = self._extract_keywords(name)

                return {
                    "id": item_id,
                    "name": name,
                    "score": score,
                    "scoring_method": method,
                    "criteria": line,
                    "keywords": keywords
                }
        return None

    def _detect_scoring_method(self, text: str) -> str:
        """检测评分方法"""
        for method, keywords in self.SCORING_METHODS.items():
            if any(kw in text for kw in keywords):
                return method
        return "综合评价"

    def _extract_keywords(self, text: str) -> List[str]:
        """提取关键词"""
        try:
            import jieba
            words = jieba.lcut(text)
            keywords = [w for w in words if len(w) >= 2 and not w.isdigit()]
            return keywords[:5]
        except ImportError:
            return text.split()[:5]

    def _extract_items_without_categories(self):
        """在没有识别到大类的情况下直接提取评分项"""
        lines = self.scoring_section.split('\n')
        items = []

        for line in lines:
            item = self._parse_scoring_item(line, len(items))
            if item:
                items.append(item)

        if items:
            total_score = sum(item["score"] for item in items)
            category = {
                "name": "评分项目",
                "score": total_score,
                "weight": 1.0,
                "items": items
            }
            self.categories.append(category)

    def extract(self) -> Dict:
        """提取评分标准主方法"""
        # 1. 查找评分办法章节
        start, end, section = self.find_scoring_section()
        if not section:
            print("[WARNING] 未找到评分办法章节")
            return self._build_result()

        self.scoring_section = section

        # 2. 提取文档信息
        self.extract_document_info()

        # 3. 识别评分大类
        self.extract_categories()

        # 4. 提取各评分项
        self.extract_scoring_items()

        return self._build_result()

    def _build_result(self) -> Dict:
        """构建最终结果"""
        return {
            "document_info": self.document_info,
            "scoring_structure": {
                "total_score": sum(cat["score"] for cat in self.categories),
                "categories": self.categories
            },
            "extraction_metadata": {
                "section_length": len(self.scoring_section),
                "categories_count": len(self.categories),
                "total_items": sum(len(cat["items"]) for cat in self.categories)
            }
        }


def extract_scoring_from_file(file_path: str) -> Dict:
    """从文件提取评分标准"""
    from parse_bid_file import parse_file

    # 解析文件
    text = parse_file(file_path)
    if text.startswith("[错误]"):
        return {"error": text}

    # 提取评分标准
    extractor = ScoringExtractor(text)
    return extractor.extract()


def main():
    import argparse

    parser = argparse.ArgumentParser(description='招标文件评分标准提取工具')
    parser.add_argument('input_file', help='输入招标文件路径')
    parser.add_argument('--output', '-o', default='scoring_criteria.json', help='输出JSON文件路径')

    args = parser.parse_args()

    if not os.path.exists(args.input_file):
        print(f"[ERROR] 文件不存在: {args.input_file}")
        sys.exit(1)

    print(f"[INFO] 正在解析文件: {args.input_file}")

    result = extract_scoring_from_file(args.input_file)

    # 保存结果
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"[INFO] 结果已保存至: {output_path}")

    # 打印摘要
    if 'scoring_structure' in result:
        categories = result['scoring_structure'].get('categories', [])
        print(f"[INFO] 识别评分大类: {len(categories)} 个")
        for cat in categories:
            print(f"  - {cat['name']}: {cat['score']}分 ({len(cat['items'])}项)")


if __name__ == '__main__':
    main()
