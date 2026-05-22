#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书智能生成专家 - 脚本包初始化
版本：v1.1 (2026-05-21)
"""

import io
import sys

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

# 版本信息
__version__ = "1.1.0"
__author__ = "TreaSkill"

# 导出主要模块
from .bid_writer_pipeline import BidWriterPipeline
from .parse_bid_file import parse_file
from .convert_to_md import convert_to_markdown
from .extract_scoring import ScoringExtractor
from .extract_requirements import RequirementsExtractor
from .extract_bid_name import BidNameExtractor
from .generate_outline import OutlineGenerator
from .generate_content import ContentGenerator
from .check_word_count import check_all_chapters
from .humanizer import Humanizer
from .convert_to_docx import convert_md_to_word
from .format_docx import format_docx
from .generate_placeholder import process_chapters
from .generate_cover import generate_cover
from .generate_toc import generate_toc

__all__ = [
    'BidWriterPipeline',
    'parse_file',
    'convert_to_markdown',
    'ScoringExtractor',
    'RequirementsExtractor',
    'BidNameExtractor',
    'OutlineGenerator',
    'ContentGenerator',
    'check_all_chapters',
    'Humanizer',
    'convert_md_to_word',
    'format_docx',
    'process_chapters',
    'generate_cover',
    'generate_toc',
]
