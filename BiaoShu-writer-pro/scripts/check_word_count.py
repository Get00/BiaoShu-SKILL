#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
字数检查脚本
根据评标分值和总页数，动态计算每个章节的合格字数范围
版本：v1.0 (2026-05-21)

公式：
    每页字数 = 28行 × 28字 = 780字
    每分页数 = 总页数 ÷ 总分
    章节目标 = 评分分值 × 每分页数 × 每页字数
    合格范围 = 目标字数 × 0.75 ~ 1.25
"""

import io
import sys
import re
import os
import json

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


def count_chinese_chars(text):
    """统计纯中文文字字数（忽略标点、英文、数字）"""
    # 移除所有空白字符
    text = re.sub(r'\s', '', text)
    # 中文字符范围
    chinese = re.findall(r'[\u4e00-\u9fff]', text)
    return len(chinese)


def count_markdown_words(md_text):
    """统计Markdown章节字数（排除标题标记、表格分隔线等）"""
    lines = md_text.split('\n')
    total_chars = 0

    for line in lines:
        stripped = line.strip()

        # 跳过空行
        if not stripped:
            continue

        # 跳过表格分隔行
        if re.match(r'^\|[-| :]+\|$', stripped):
            continue

        # 跳过占位符
        if stripped.startswith('[') and '占位符' in stripped:
            continue

        # 标题本身也计入字数（去掉#和空格）
        if stripped.startswith('#'):
            content = re.sub(r'^#+\s*', '', stripped)
            total_chars += count_chinese_chars(content)
            continue

        # 跳过代码块标记
        if stripped.startswith('```'):
            continue

        # 跳过图片/链接语法
        line_clean = re.sub(r'!\[.*?\]\(.*?\)', '', stripped)
        line_clean = re.sub(r'\[.*?\]\(.*?\)', '', line_clean)

        total_chars += count_chinese_chars(line_clean)

    return total_chars


def check_chapter_words(content, chapter_score, total_pages=40, total_score=100):
    """
    检查单个章节的字数

    Args:
        content: 章节内容
        chapter_score: 章节分值
        total_pages: 总页数
        total_score: 总分

    Returns:
        dict: 检查结果
    """
    # 计算基础参数
    chars_per_page = 28 * 28  # 780
    pages_per_score = total_pages / total_score

    # 计算目标字数和合格范围
    target = chapter_score * pages_per_score * chars_per_page
    min_chars = target * 0.75
    max_chars = target * 1.25

    # 统计字数
    actual = count_markdown_words(content)

    # 判断是否合格
    is_pass = min_chars <= actual <= max_chars

    return {
        "score": chapter_score,
        "target": target,
        "min": min_chars,
        "max": max_chars,
        "actual": actual,
        "pass": is_pass,
        "percentage": actual / target * 100 if target > 0 else 0
    }


def check_all_chapters(chapters_dir, scoring_file, total_pages=40):
    """
    检查所有章节的字数

    Args:
        chapters_dir: 章节目录
        scoring_file: 评分标准文件
        total_pages: 总页数

    Returns:
        dict: 检查结果
    """
    # 读取评分标准
    with open(scoring_file, 'r', encoding='utf-8') as f:
        scoring = json.load(f)

    scoring_structure = scoring.get('scoring_structure', {})
    categories = scoring_structure.get('categories', [])
    total_score = scoring_structure.get('total_score', 100)

    # 构建章节分值映射
    chapter_scores = {}
    for category in categories:
        for item in category.get('items', []):
            item_name = item.get('name', '')
            item_score = item.get('score', 0)
            chapter_scores[item_name] = item_score

    # 检查每个章节
    results = {}
    all_pass = True

    for filename in sorted(os.listdir(chapters_dir)):
        if not filename.endswith('.md'):
            continue

        filepath = os.path.join(chapters_dir, filename)

        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()

        # 提取章节名
        match = re.search(r'^#\s+(.+)$', content, re.MULTILINE)
        chapter_name = match.group(1).strip() if match else filename

        # 查找对应分值
        score = None
        for name, s in chapter_scores.items():
            if name in chapter_name or chapter_name in name:
                score = s
                break

        if score is None:
            # 默认分值
            score = total_score / len(os.listdir(chapters_dir))

        # 检查字数
        result = check_chapter_words(content, score, total_pages, total_score)
        results[filename] = {
            "chapter": chapter_name,
            **result
        }

        if not result['pass']:
            all_pass = False

    return {
        "all_pass": all_pass,
        "total_pages": total_pages,
        "total_score": total_score,
        "chapters": results
    }


def print_check_results(results):
    """打印检查结果"""
    print(f"=" * 60)
    print(f"📊 章节字数检查")
    print(f"=" * 60)
    print(f"总页数: {results['total_pages']} | 总分: {results['total_score']}")
    print(f"")

    for filename, chapter in results['chapters'].items():
        status = "✅" if chapter['pass'] else "❌"
        print(f"{status} {filename}")
        print(f"   章节: {chapter['chapter']}")
        print(f"   分值: {chapter['score']}分 | 目标: {chapter['target']:.0f}字")
        print(f"   合格: {chapter['min']:.0f} ~ {chapter['max']:.0f}字")
        print(f"   实际: {chapter['actual']}字 ({chapter['percentage']:.1f}%)")
        print()

    print(f"=" * 60)
    if results['all_pass']:
        print(f"✅ 全部章节字数合格")
    else:
        print(f"❌ 以下章节需要调整:")
        for filename, chapter in results['chapters'].items():
            if not chapter['pass']:
                print(f"   - {chapter['chapter']}: {chapter['actual']}字")


def main():
    import argparse

    parser = argparse.ArgumentParser(description='标书字数检查工具')
    parser.add_argument('chapters_dir', help='章节目录路径')
    parser.add_argument('scoring_file', help='评分标准JSON文件路径')
    parser.add_argument('--total-pages', '-p', type=int, default=40, help='总页数（默认40页）')
    parser.add_argument('--output', '-o', help='输出结果到文件')

    args = parser.parse_args()

    if not os.path.exists(args.chapters_dir):
        print(f"[ERROR] 章节目录不存在: {args.chapters_dir}")
        sys.exit(1)

    if not os.path.exists(args.scoring_file):
        print(f"[ERROR] 评分标准文件不存在: {args.scoring_file}")
        sys.exit(1)

    print(f"[INFO] 开始检查字数...")

    results = check_all_chapters(args.chapters_dir, args.scoring_file, args.total_pages)

    # 打印结果
    print_check_results(results)

    # 保存结果
    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)

        print(f"\n[INFO] 结果已保存至: {output_path}")


if __name__ == '__main__':
    main()
