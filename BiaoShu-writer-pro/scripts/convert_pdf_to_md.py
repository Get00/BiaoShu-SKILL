#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF转Markdown转换脚本
将PDF文件转换为Markdown格式，用于标书知识库构建
版本：v1.0 (2026-05-21)
"""

import io
import os
import sys
from pathlib import Path

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


def extract_pdf_content(file_path: str) -> str:
    """从PDF文件提取内容"""
    try:
        import pdfplumber
        
        content = []
        with pdfplumber.open(file_path) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text and text.strip():
                    # 添加页码标记
                    content.append(f"## 第 {i} 页\n\n{text.strip()}")
        
        return '\n\n'.join(content)
    except ImportError:
        return "[错误] pdfplumber未安装，请运行: pip install pdfplumber"
    except Exception as e:
        return f"[错误] 无法读取 {file_path}: {str(e)}"


def process_pdf_folder(source_dir: str, output_dir: str, prefix: str = "knowledge"):
    """处理PDF文件夹"""
    os.makedirs(output_dir, exist_ok=True)
    
    # 获取所有PDF文件
    pdf_files = []
    for root, dirs, files in os.walk(source_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    if not pdf_files:
        print(f"[WARNING] 在 {source_dir} 中没有找到PDF文件")
        return
    
    print(f"找到 {len(pdf_files)} 个PDF文件")
    print("=" * 60)
    
    processed_count = 0
    for pdf_path in pdf_files:
        pdf_name = os.path.basename(pdf_path)
        print(f"\n处理: {pdf_name}")
        
        # 提取内容
        content = extract_pdf_content(pdf_path)
        
        # 生成输出文件名
        md_name = pdf_name.rsplit('.', 1)[0] + '.md'
        md_path = os.path.join(output_dir, f"{prefix}_{md_name}")
        
        # 保存为md文件
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(f"# {pdf_name}\n\n")
            f.write(f"**来源**: {pdf_path}\n\n")
            f.write("---\n\n")
            f.write(content)
        
        print(f"  已保存: {os.path.basename(md_path)}")
        processed_count += 1
    
    print("\n" + "=" * 60)
    print(f"处理完成！成功转换 {processed_count}/{len(pdf_files)} 个PDF文件")
    print("=" * 60)


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='PDF转Markdown转换工具')
    parser.add_argument('--source', '-s',
                       required=True,
                       help='源PDF文件夹路径')
    parser.add_argument('--output', '-o',
                       default='./templates',
                       help='输出文件夹路径')
    parser.add_argument('--prefix', '-p',
                       default='knowledge',
                       help='输出文件名前缀 (默认: knowledge)')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.source):
        print(f"[ERROR] 源文件夹不存在: {args.source}")
        sys.exit(1)
    
    process_pdf_folder(args.source, args.output, args.prefix)


if __name__ == '__main__':
    main()