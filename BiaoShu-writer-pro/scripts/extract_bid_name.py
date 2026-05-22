#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书名称提取脚本
从招标文件和内容中提取项目名称，生成标书文件名
版本：v1.0 (2026-05-21)

命名规则：
1. 优先从招标文件中提取项目名称
2. 结合投标方信息
3. 格式：{项目名称}技术标.docx
"""

import io
import re
import os
import sys
from pathlib import Path
from typing import Optional

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


class BidNameExtractor:
    """标书名称提取器"""

    def __init__(self, file_path: str = None, content: str = None):
        self.file_path = file_path
        self.content = content
        self.project_name = None

    def extract(self) -> str:
        """提取标书名称"""
        # 方法1：从内容中提取项目名称
        if self.content:
            self.project_name = self._extract_from_content()

        # 方法2：从文件名提取
        if not self.project_name and self.file_path:
            self.project_name = self._extract_from_filename()

        # 方法3：使用默认名称
        if not self.project_name:
            self.project_name = "投标项目"

        # 生成文件名
        return self._generate_filename()

    def _extract_from_content(self) -> Optional[str]:
        """从内容中提取项目名称"""
        patterns = [
            r'项目名称[：:]\s*([^\n]+)',
            r'招标项目[：:]\s*([^\n]+)',
            r'项目名称[：:]\s*(.+?)[\n\r]',
            r'关于\s*[“""](.+?)[”""].+?的招标',
            r'(.{3,30}项目).{0,5}招标',
            r'(.{3,30}项目).{0,5}采购',
            r'(.{3,30}系统).{0,5}建设',
            r'(.{3,30}平台).{0,5}建设',
        ]

        for pattern in patterns:
            match = re.search(pattern, self.content)
            if match:
                name = match.group(1).strip()
                # 清理名称
                name = self._clean_name(name)
                if name and len(name) >= 3:
                    return name

        return None

    def _extract_from_filename(self) -> Optional[str]:
        """从文件名提取项目名称"""
        if not self.file_path:
            return None

        filename = Path(self.file_path).stem

        # 移除常见后缀
        remove_suffixes = ['招标文件', '招标公告', '招标', '采购文件', '采购公告', '采购', '文件', '公告']
        name = filename
        for suffix in remove_suffixes:
            name = name.replace(suffix, '')

        # 清理名称
        name = self._clean_name(name)

        if name and len(name) >= 3:
            return name

        return None

    def _clean_name(self, name: str) -> str:
        """清理项目名称"""
        # 移除非法字符
        illegal_chars = '<>:"/\\|?*'
        for char in illegal_chars:
            name = name.replace(char, '')

        # 移除多余空格
        name = re.sub(r'\s+', ' ', name).strip()

        # 移除括号内容（如果太长）
        name = re.sub(r'[（(].{20,}[）)]', '', name)

        return name

    def _generate_filename(self) -> str:
        """生成文件名"""
        if self.project_name:
            # 确保名称以"项目"或"系统"结尾
            if not any(self.project_name.endswith(suffix) for suffix in ['项目', '系统', '平台', '工程']):
                filename = f"{self.project_name}项目技术标.docx"
            else:
                filename = f"{self.project_name}技术标.docx"
        else:
            filename = "投标文件技术标.docx"

        return filename


def extract_bid_name_from_file(file_path: str) -> str:
    """从文件提取标书名称"""
    from parse_bid_file import parse_file

    # 解析文件
    content = parse_file(file_path)
    if content.startswith("[错误]"):
        return "投标文件技术标.docx"

    # 提取名称
    extractor = BidNameExtractor(file_path=file_path, content=content)
    return extractor.extract()


def extract_bid_name_from_content(content: str, file_path: str = None) -> str:
    """从内容提取标书名称"""
    extractor = BidNameExtractor(file_path=file_path, content=content)
    return extractor.extract()


def main():
    import argparse

    parser = argparse.ArgumentParser(description='标书名称提取工具')
    parser.add_argument('input_file', help='输入招标文件路径')
    parser.add_argument('--output', '-o', help='输出名称到文件')

    args = parser.parse_args()

    if not os.path.exists(args.input_file):
        print(f"[ERROR] 文件不存在: {args.input_file}")
        sys.exit(1)

    print(f"[INFO] 正在解析文件: {args.input_file}")

    filename = extract_bid_name_from_file(args.input_file)

    print(f"[INFO] 提取的标书名称: {filename}")

    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(filename)
        print(f"[INFO] 名称已保存至: {args.output}")


if __name__ == '__main__':
    main()
