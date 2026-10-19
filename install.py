#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
安装脚本
自动安装SKILL所需的依赖库
版本：v1.0 (2026-05-21)
"""

import io
import os
import sys
import subprocess
import platform

# 设置标准输出编码（解决Windows控制台编码问题）
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
except Exception:
    pass


def get_script_dir():
    """获取脚本所在目录"""
    return os.path.dirname(os.path.abspath(__file__))


def get_requirements_path():
    """获取requirements.txt路径"""
    script_dir = get_script_dir()
    return os.path.join(script_dir, "requirements.txt")


def install_package(package):
    """安装单个包"""
    print(f"[INFO] 正在安装 {package}...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"[OK] {package} 安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] {package} 安装失败: {e}")
        return False


def install_from_requirements():
    """从requirements.txt安装依赖"""
    requirements_path = get_requirements_path()

    if not os.path.exists(requirements_path):
        print(f"[ERROR] requirements.txt 文件不存在: {requirements_path}")
        return False

    print(f"[INFO] 正在从 {requirements_path} 安装依赖...")

    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "-r", requirements_path
        ])
        print("[OK] 所有依赖安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] 安装失败: {e}")
        return False


def install_core_packages():
    """安装核心依赖包"""
    print("\n" + "=" * 60)
    print("安装核心依赖库...")
    print("=" * 60)

    core_packages = [
        "python-docx>=0.8.11",
        "pdfplumber>=0.7.0",
        "openpyxl>=3.0.0",
        "PyPDF2>=3.0.0",
        "jieba>=0.42.1",
        "pandas>=1.3.0",
    ]

    results = []
    for package in core_packages:
        result = install_package(package)
        results.append(result)

    return all(results)


def install_windows_packages():
    """安装Windows特定依赖"""
    system = platform.system()

    if system != "Windows":
        print(f"\n[INFO] 当前系统为{system}，跳过Windows特定依赖安装")
        return True

    print("\n" + "=" * 60)
    print("安装Windows特定依赖库...")
    print("=" * 60)

    windows_packages = [
        "pywin32>=306",
    ]

    results = []
    for package in windows_packages:
        result = install_package(package)
        results.append(result)

    return all(results)


def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description='标书智能生成专家 - 依赖安装工具')
    parser.add_argument('--auto', '-a', action='store_true',
                       help='自动安装（从requirements.txt）')
    parser.add_argument('--core-only', '-c', action='store_true',
                       help='仅安装核心依赖')

    args = parser.parse_args()

    print("=" * 60)
    print("标书智能生成专家 - 依赖安装工具")
    print("=" * 60)
    print(f"\nPython版本: {sys.version}")
    print(f"操作系统: {platform.system()} {platform.release()}")
    print(f"安装路径: {get_script_dir()}")

    # 自动安装模式
    if args.auto:
        print("\n[INFO] 自动安装模式")
        success = install_from_requirements()
    elif args.core_only:
        print("\n[INFO] 仅安装核心依赖")
        success = install_core_packages()
    else:
        # 交互式模式
        print("\n请选择安装方式:")
        print("1. 从 requirements.txt 安装（推荐）")
        print("2. 逐个安装核心依赖")
        print("3. 退出")

        try:
            choice = input("\n请输入选择 (1/2/3): ").strip()
        except EOFError:
            # 非交互式环境，自动使用requirements.txt安装
            print("\n[INFO] 检测到非交互式环境，自动使用requirements.txt安装")
            choice = "1"

        if choice == "1":
            success = install_from_requirements()
        elif choice == "2":
            core_ok = install_core_packages()
            windows_ok = install_windows_packages()
            success = core_ok and windows_ok
        elif choice == "3":
            print("[INFO] 退出安装")
            return
        else:
            print("[ERROR] 无效的选择")
            return

    # 显示结果
    print("\n" + "=" * 60)
    if success:
        print("✓ 依赖安装完成！")
        print("\n现在可以使用以下命令检查环境:")
        print("  python scripts/check_environment.py")
        print("\n使用方法:")
        print("  python scripts/bid_writer_pipeline.py <招标文件路径> --output-dir ./output")
    else:
        print("✗ 部分依赖安装失败，请手动安装")
        print("\n手动安装命令:")
        print("  pip install -r requirements.txt")
    print("=" * 60)


if __name__ == '__main__':
    main()
