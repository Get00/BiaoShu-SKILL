#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
环境检查脚本
检查SKILL运行所需的依赖库和环境配置
版本：v1.0 (2026-05-21)
"""

import io
import os
import sys
import platform

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


def check_python_version():
    """检查Python版本"""
    version = sys.version_info
    print(f"[INFO] Python版本: {version.major}.{version.minor}.{version.micro}")

    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print("[WARNING] 建议使用Python 3.8或更高版本")
        return False

    print("[OK] Python版本符合要求")
    return True


def check_platform():
    """检查操作系统平台"""
    system = platform.system()
    print(f"[INFO] 操作系统: {system}")
    print(f"[INFO] 平台: {platform.platform()}")

    return system


def check_package(package_name, import_name=None):
    """检查单个包是否安装"""
    if import_name is None:
        import_name = package_name

    try:
        __import__(import_name)
        print(f"[OK] {package_name} 已安装")
        return True
    except ImportError:
        print(f"[ERROR] {package_name} 未安装")
        print(f"       安装命令: pip install {package_name}")
        return False


def check_all_dependencies():
    """检查所有依赖库"""
    print("\n" + "=" * 60)
    print("检查依赖库...")
    print("=" * 60)

    # 核心依赖
    core_packages = [
        ("python-docx", "docx"),
        ("pdfplumber", "pdfplumber"),
        ("openpyxl", "openpyxl"),
        ("PyPDF2", "PyPDF2"),
        ("jieba", "jieba"),
        ("pandas", "pandas"),
    ]

    # Windows特定依赖
    windows_packages = [
        ("pywin32", "win32com.client"),
    ]

    results = []

    # 检查核心依赖
    print("\n核心依赖库:")
    for package_name, import_name in core_packages:
        result = check_package(package_name, import_name)
        results.append((package_name, result))

    # 检查Windows特定依赖
    system = platform.system()
    if system == "Windows":
        print("\nWindows特定依赖库:")
        for package_name, import_name in windows_packages:
            result = check_package(package_name, import_name)
            results.append((package_name, result))
    else:
        print(f"\n[INFO] 当前系统为{system}，跳过Windows特定依赖检查")
        print("[INFO] .doc文件解析功能仅在Windows系统上可用")

    return results


def check_scripts():
    """检查脚本文件是否存在"""
    print("\n" + "=" * 60)
    print("检查脚本文件...")
    print("=" * 60)

    scripts_dir = os.path.dirname(os.path.abspath(__file__))
    required_scripts = [
        "bid_writer_pipeline.py",
        "parse_bid_file.py",
        "convert_to_md.py",
        "extract_scoring.py",
        "extract_requirements.py",
        "extract_bid_name.py",
        "generate_outline.py",
        "generate_content.py",
        "check_word_count.py",
        "humanizer.py",
        "convert_to_docx.py",
        "format_docx.py",
        "generate_placeholder.py",
        "generate_cover.py",
        "generate_toc.py",
    ]

    results = []
    for script in required_scripts:
        script_path = os.path.join(scripts_dir, script)
        if os.path.exists(script_path):
            print(f"[OK] {script}")
            results.append((script, True))
        else:
            print(f"[ERROR] {script} 不存在")
            results.append((script, False))

    return results


def check_templates():
    """检查模板知识文件是否存在"""
    print("\n" + "=" * 60)
    print("检查模板知识文件...")
    print("=" * 60)

    scripts_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(scripts_dir)
    templates_dir = os.path.join(parent_dir, "templates")

    if not os.path.exists(templates_dir):
        print(f"[WARNING] templates文件夹不存在")
        return []

    # 检查主要的模板文件
    required_files = [
        "00_标书制作指南.md",
    ]

    results = []
    for file_name in required_files:
        file_path = os.path.join(templates_dir, file_name)
        if os.path.exists(file_path):
            print(f"[OK] {file_name}")
            results.append((file_name, True))
        else:
            print(f"[WARNING] {file_name} 不存在")
            results.append((file_name, False))

    # 统计templates文件夹中的md文件数量
    md_files = [f for f in os.listdir(templates_dir) if f.endswith('.md')]
    print(f"[INFO] templates文件夹中共有 {len(md_files)} 个md文件")

    return results


def check_font_support():
    """检查字体支持"""
    print("\n" + "=" * 60)
    print("检查字体支持...")
    print("=" * 60)

    # 常用中文字体
    fonts = [
        ("宋体", "SimSun"),
        ("黑体", "SimHei"),
        ("仿宋", "FangSong"),
        ("楷体", "KaiTi"),
    ]

    system = platform.system()
    if system == "Windows":
        fonts_dir = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")
        for font_name, font_file in fonts:
            font_path = os.path.join(fonts_dir, f"{font_file}.ttf")
            if os.path.exists(font_path):
                print(f"[OK] {font_name} ({font_file})")
            else:
                print(f"[WARNING] {font_name} ({font_file}) 可能未安装")
    else:
        print(f"[INFO] 当前系统为{system}，跳过Windows字体检查")
        print("[INFO] 请确保系统安装了中文字体（宋体、黑体等）")


def generate_report(python_ok, platform_name, deps_results, scripts_results, templates_results):
    """生成检查报告"""
    print("\n" + "=" * 60)
    print("环境检查报告")
    print("=" * 60)

    # Python版本
    print(f"\n1. Python版本: {'✓ 通过' if python_ok else '✗ 不通过'}")

    # 操作系统
    print(f"2. 操作系统: {platform_name}")

    # 依赖库
    deps_ok = all(result for _, result in deps_results)
    failed_deps = [name for name, result in deps_results if not result]
    print(f"3. 依赖库: {'✓ 全部通过' if deps_ok else '✗ 部分缺失'}")
    if failed_deps:
        print(f"   缺失的库: {', '.join(failed_deps)}")

    # 脚本文件
    scripts_ok = all(result for _, result in scripts_results)
    failed_scripts = [name for name, result in scripts_results if not result]
    print(f"4. 脚本文件: {'✓ 全部存在' if scripts_ok else '✗ 部分缺失'}")
    if failed_scripts:
        print(f"   缺失的文件: {', '.join(failed_scripts)}")

    # 模板知识文件
    templates_ok = all(result for _, result in templates_results) if templates_results else False
    print(f"5. 模板知识: {'✓ 全部存在' if templates_ok else '⚠ 部分缺失（非关键）'}")

    # 总体结果
    print("\n" + "=" * 60)
    if python_ok and deps_ok and scripts_ok:
        print("✓ 环境检查通过！SKILL可以正常使用。")
    else:
        print("✗ 环境检查未通过，请修复上述问题后重试。")

    # 使用说明
    print("\n使用方法:")
    print("  python scripts/bid_writer_pipeline.py <招标文件路径> --output-dir ./output")

    # 依赖安装命令
    if failed_deps:
        print("\n安装缺失的依赖:")
        print(f"  pip install {' '.join(failed_deps)}")

    print("=" * 60)


def main():
    """主函数"""
    print("=" * 60)
    print("标书智能生成专家 - 环境检查工具")
    print("=" * 60)

    # 检查Python版本
    python_ok = check_python_version()

    # 检查平台
    platform_name = check_platform()

    # 检查依赖库
    deps_results = check_all_dependencies()

    # 检查脚本文件
    scripts_results = check_scripts()

    # 检查模板知识文件
    templates_results = check_templates()

    # 检查字体支持
    check_font_support()

    # 生成报告
    generate_report(python_ok, platform_name, deps_results, scripts_results, templates_results)


if __name__ == '__main__':
    main()
