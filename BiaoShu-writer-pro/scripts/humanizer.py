#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI去痕处理脚本
去除文本中的AI生成痕迹，使内容更自然、更像人类书写
版本：v1.0 (2026-05-21)
"""

import io
import re
import sys
import os
from typing import List, Dict

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


class Humanizer:
    """AI去痕处理器"""

    # AI常见词汇和短语
    AI_PHRASES = [
        # 夸大性语言
        '至关重要', '举足轻重', '不可或缺', '极其重要',
        '显著提升', '大幅提高', '极大改善', '根本性变革',

        # 模糊性表达
        '在某种程度上', '从某种意义上说', '可以说',
        '毋庸置疑', '众所周知', '显而易见',

        # 连接性短语
        '首先', '其次', '再次', '最后',
        '一方面', '另一方面', '与此同时',
        '不仅如此', '更重要的是', '值得注意的是',

        # AI特征词汇
        '赋能', '助力', '驱动', '引领',
        '生态', '闭环', '抓手', '颗粒度',
        '底层逻辑', '顶层设计', '降维打击',
    ]

    # 需要替换的表达
    REPLACEMENTS = {
        '赋能': '支持',
        '助力': '帮助',
        '驱动': '推动',
        '引领': '带动',
        '生态': '体系',
        '闭环': '流程',
        '抓手': '手段',
        '颗粒度': '细致程度',
        '底层逻辑': '基本原理',
        '顶层设计': '整体规划',
    }

    # 需要删除的填充短语
    FILLER_PHRASES = [
        '首先，', '其次，', '再次，', '最后，',
        '一方面，', '另一方面，',
        '不仅如此，', '更重要的是，', '值得注意的是，',
        '众所周知，', '显而易见，', '毋庸置疑，',
    ]

    def __init__(self, text: str):
        self.text = text
        self.changes = []

    def humanize(self) -> str:
        """执行去痕处理"""
        result = self.text

        # 1. 替换AI特征词汇
        result = self._replace_ai_phrases(result)

        # 2. 删除填充短语
        result = self._remove_fillers(result)

        # 3. 简化复杂表达
        result = self._simplify_expressions(result)

        # 4. 增加变化性
        result = self._add_variation(result)

        return result

    def _replace_ai_phrases(self, text: str) -> str:
        """替换AI特征词汇"""
        result = text
        for old, new in self.REPLACEMENTS.items():
            if old in result:
                result = result.replace(old, new)
                self.changes.append(f"替换: {old} -> {new}")
        return result

    def _remove_fillers(self, text: str) -> str:
        """删除填充短语"""
        result = text
        for phrase in self.FILLER_PHRASES:
            if phrase in result:
                # 只删除段落开头的填充短语
                result = result.replace('\n' + phrase, '\n')
                # 删除文本开头的填充短语
                if result.startswith(phrase):
                    result = result[len(phrase):]
                self.changes.append(f"删除填充: {phrase}")
        return result

    def _simplify_expressions(self, text: str) -> str:
        """简化复杂表达"""
        result = text

        # 简化过长的句子
        sentences = result.split('。')
        simplified = []
        for sentence in sentences:
            if len(sentence) > 100:
                # 尝试在逗号处断句
                parts = sentence.split('，')
                if len(parts) > 2:
                    mid = len(parts) // 2
                    sentence = '，'.join(parts[:mid]) + '。\n' + '，'.join(parts[mid:])
                    self.changes.append("断句: 长句拆分")
            simplified.append(sentence)

        return '。'.join(simplified)

    def _add_variation(self, text: str) -> str:
        """增加表达变化性"""
        result = text

        # 替换重复的表达
        replacements = [
            ('采用', ['使用', '运用', '利用']),
            ('实现', ['达成', '完成', '达到']),
            ('提供', ['给予', '供应', '交付']),
            ('确保', ['保证', '保障', '维护']),
            ('提升', ['提高', '增强', '改善']),
        ]

        for word, alternatives in replacements:
            # 统计出现次数
            count = result.count(word)
            if count > 3:
                # 替换部分实例
                for i, alt in enumerate(alternatives):
                    if i < count - 2:
                        result = result.replace(word, alt, 1)
                        self.changes.append(f"变化: {word} -> {alt}")

        return result

    def get_changes(self) -> List[str]:
        """获取修改记录"""
        return self.changes


def humanize_file(input_file: str, output_file: str = None) -> str:
    """处理文件"""
    # 读取文件
    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # 去痕处理
    humanizer = Humanizer(content)
    result = humanizer.humanize()

    # 输出修改统计
    changes = humanizer.get_changes()
    print(f"[INFO] 去痕处理完成，共 {len(changes)} 处修改")

    # 保存结果
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(result)
        print(f"[INFO] 结果已保存至: {output_file}")

    return result


def humanize_text(text: str) -> tuple:
    """处理文本"""
    humanizer = Humanizer(text)
    result = humanizer.humanize()
    changes = humanizer.get_changes()
    return result, changes


def main():
    import argparse

    parser = argparse.ArgumentParser(description='AI去痕处理工具')
    parser.add_argument('input_file', help='输入文件路径')
    parser.add_argument('--output', '-o', help='输出文件路径')

    args = parser.parse_args()

    if not os.path.exists(args.input_file):
        print(f"[ERROR] 文件不存在: {args.input_file}")
        sys.exit(1)

    print(f"[INFO] 开始处理文件: {args.input_file}")

    output_file = args.output or args.input_file.replace('.md', '.humanized.md')
    humanize_file(args.input_file, output_file)


if __name__ == '__main__':
    main()
