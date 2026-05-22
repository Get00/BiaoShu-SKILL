#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
行业检测模块
自动识别招标文件所属行业，支持通用行业投标文件生成
版本：v1.0 (2026-05-22)
"""

import json
import os
import re
from pathlib import Path
from typing import Dict, List, Tuple


class IndustryDetector:
    """行业检测器"""

    def __init__(self, config_path: str = None):
        """
        初始化行业检测器

        Args:
            config_path: 行业配置文件路径
        """
        if config_path is None:
            # 默认配置文件路径
            script_dir = Path(__file__).parent
            config_path = script_dir.parent / "references" / "industry_configs.json"

        self.config_path = Path(config_path)
        self.config = self._load_config()
        self.industries = self.config.get("industries", {})
        self.detection_config = self.config.get("industry_detection", {})

    def _load_config(self) -> Dict:
        """加载行业配置文件"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"[WARNING] 加载行业配置文件失败: {e}")
            return {"industries": {}, "industry_detection": {}}

    def detect_industry(self, text: str) -> Tuple[str, float, Dict]:
        """
        检测文本所属行业

        Args:
            text: 待检测的文本

        Returns:
            Tuple[str, float, Dict]: (行业代码, 置信度, 行业配置)
        """
        if not text:
            return "general", 0.0, self.industries.get("general", {})

        # 统计每个行业的匹配分数
        industry_scores = {}

        for industry_code, industry_config in self.industries.items():
            keywords = industry_config.get("keywords", [])
            if not keywords:
                continue

            # 计算关键词匹配分数
            score = self._calculate_keyword_score(text, keywords)
            industry_scores[industry_code] = score

        # 找到最高分的行业
        if industry_scores:
            best_industry = max(industry_scores, key=industry_scores.get)
            best_score = industry_scores[best_industry]

            # 检查是否超过置信度阈值
            threshold = self.detection_config.get("confidence_threshold", 0.3)
            if best_score >= threshold:
                return best_industry, best_score, self.industries.get(best_industry, {})

        # 如果没有匹配的行业，返回通用行业
        fallback = self.detection_config.get("fallback_industry", "general")
        return fallback, 0.0, self.industries.get(fallback, {})

    def _calculate_keyword_score(self, text: str, keywords: List[str]) -> float:
        """
        计算关键词匹配分数

        Args:
            text: 待检测的文本
            keywords: 关键词列表

        Returns:
            float: 匹配分数 (0-1)
        """
        if not text or not keywords:
            return 0.0

        text_lower = text.lower()
        total_keywords = len(keywords)
        matched_keywords = 0

        for keyword in keywords:
            # 统计关键词出现次数
            count = text_lower.count(keyword.lower())
            if count > 0:
                matched_keywords += 1

        # 计算匹配比例
        if total_keywords > 0:
            return matched_keywords / total_keywords
        return 0.0

    def get_industry_config(self, industry_code: str) -> Dict:
        """
        获取行业配置

        Args:
            industry_code: 行业代码

        Returns:
            Dict: 行业配置
        """
        return self.industries.get(industry_code, self.industries.get("general", {}))

    def get_chapter_structure(self, industry_code: str) -> List[str]:
        """
        获取行业章节结构

        Args:
            industry_code: 行业代码

        Returns:
            List[str]: 章节结构列表
        """
        config = self.get_industry_config(industry_code)
        return config.get("chapter_structure", [])

    def get_key_sections(self, industry_code: str) -> Dict[str, List[str]]:
        """
        获取行业关键章节

        Args:
            industry_code: 行业代码

        Returns:
            Dict[str, List[str]]: 关键章节配置
        """
        config = self.get_industry_config(industry_code)
        return config.get("key_sections", {})

    def get_format_standards(self, industry_code: str) -> List[str]:
        """
        获取行业支持的格式标准

        Args:
            industry_code: 行业代码

        Returns:
            List[str]: 格式标准列表
        """
        config = self.get_industry_config(industry_code)
        return config.get("format_standards", ["government", "enterprise"])

    def list_industries(self) -> List[Dict]:
        """
        列出所有支持的行业

        Returns:
            List[Dict]: 行业列表
        """
        industries = []
        for code, config in self.industries.items():
            industries.append({
                "code": code,
                "name": config.get("name", ""),
                "description": config.get("description", "")
            })
        return industries


def detect_industry_from_file(file_path: str) -> Tuple[str, float, Dict]:
    """
    从文件检测行业

    Args:
        file_path: 文件路径

    Returns:
        Tuple[str, float, Dict]: (行业代码, 置信度, 行业配置)
    """
    try:
        # 读取文件内容
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 检测行业
        detector = IndustryDetector()
        return detector.detect_industry(content)
    except Exception as e:
        print(f"[WARNING] 行业检测失败: {e}")
        return "general", 0.0, {}


def detect_industry_from_text(text: str) -> Tuple[str, float, Dict]:
    """
    从文本检测行业

    Args:
        text: 文本内容

    Returns:
        Tuple[str, float, Dict]: (行业代码, 置信度, 行业配置)
    """
    detector = IndustryDetector()
    return detector.detect_industry(text)


def main():
    """命令行入口"""
    import argparse

    parser = argparse.ArgumentParser(description='行业检测工具')
    parser.add_argument('input_file', help='输入文件路径')
    parser.add_argument('--list', '-l', action='store_true', help='列出所有支持的行业')

    args = parser.parse_args()

    detector = IndustryDetector()

    if args.list:
        print("\n支持的行业列表：")
        print("=" * 60)
        for industry in detector.list_industries():
            print(f"  {industry['code']}: {industry['name']}")
            print(f"    {industry['description']}")
        return

    if not os.path.exists(args.input_file):
        print(f"[ERROR] 文件不存在: {args.input_file}")
        return

    print(f"\n检测文件: {args.input_file}")
    print("=" * 60)

    industry_code, confidence, config = detect_industry_from_file(args.input_file)

    print(f"检测结果:")
    print(f"  行业代码: {industry_code}")
    print(f"  行业名称: {config.get('name', '未知')}")
    print(f"  置信度: {confidence:.2%}")
    print(f"  行业描述: {config.get('description', '')}")

    print(f"\n章节结构:")
    for i, chapter in enumerate(detector.get_chapter_structure(industry_code), 1):
        print(f"  {i}. {chapter}")


if __name__ == '__main__':
    main()
