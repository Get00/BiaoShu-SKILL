#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
环境自检工具 —— 迁移到新电脑后先跑这个

检查 Python 版本、依赖库是否齐全、skill 自带脚本是否完整，
并给出可直接复制的安装命令。

用法：
    python scripts/verify_env.py            # 只检查
    python scripts/verify_env.py --install  # 检查并自动 pip 安装缺失的库
版本：v1.0 (2026-10-13)
"""

import io
import os
import sys
import platform
import subprocess

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

# (import 名, pip 包名, 用途, 是否必需)
DEPS = [
    ("docx",       "python-docx",  "Word 读写（核心）", True),
    ("openpyxl",   "openpyxl",     "Excel 读写", True),
    ("pptx",       "python-pptx",  "PowerPoint 读写", True),
    ("pdfplumber", "pdfplumber",   "PDF 读取（首选）", True),
    ("pypdf",      "pypdf",        "PDF 读取/拆分（office 工具包用）", True),
    ("reportlab",  "reportlab",    "PDF 生成（中文）", True),
    ("jieba",      "jieba",        "中文分词", False),
    ("PyPDF2",     "PyPDF2",       "PDF 读取（旧代码兼容）", False),
    ("pandas",     "pandas",       "表格处理（可选）", False),
]

WIN_DEPS = [
    ("win32com.client", "pywin32", ".doc/.xls 老格式转换（Windows 需要）", False),
]

# skill 自带的脚本（缺一不可）
SELF_FILES = [
    "scripts/bid_writer_pipeline.py",
    "scripts/extract_format.py",
    "scripts/apply_bid_template.py",
    "scripts/check_format_drift.py",
    "scripts/format_docx.py",
    "scripts/doc_io.py",
    "scripts/fsutil.py",
    "scripts/office/common.py",
    "scripts/office/docx.py",
    "scripts/office/xlsx.py",
    "scripts/office/pptx.py",
    "scripts/office/pdf.py",
    "scripts/office_extract.py",
    "scripts/office_patch_copy.py",
]


def check_import(mod):
    try:
        __import__(mod)
        return True, ""
    except Exception as e:
        return False, str(e)[:60]


def main():
    install = "--install" in sys.argv
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    print("=" * 62)
    print("标书智能生成专家 — 环境自检")
    print("=" * 62)
    print(f"Python : {sys.version.split()[0]}  ({sys.executable})")
    print(f"系统   : {platform.system()} {platform.release()}")
    print(f"Skill  : {root}")

    if sys.version_info < (3, 8):
        print("\n❌ Python 版本过低，需要 >= 3.8")
        return 1

    deps = DEPS + (WIN_DEPS if platform.system() == "Windows" else [])
    missing_required, missing_optional = [], []

    print("\n【依赖库】")
    for mod, pkg, use, required in deps:
        ok, err = check_import(mod)
        if ok:
            print(f"  ✅ {pkg:<14} {use}")
        else:
            tag = "必需" if required else "可选"
            print(f"  ❌ {pkg:<14} [{tag}] {use}   ← {err}")
            (missing_required if required else missing_optional).append(pkg)

    print("\n【自带脚本】")
    lost = []
    for rel in SELF_FILES:
        p = os.path.join(root, rel)
        if os.path.exists(p):
            print(f"  ✅ {rel}")
        else:
            print(f"  ❌ 缺失：{rel}")
            lost.append(rel)

    print("\n" + "=" * 62)
    if not missing_required and not lost:
        print("✅ 环境完整，可以直接使用。")
        if missing_optional:
            print(f"   （可选库未装：{', '.join(missing_optional)} — 不影响主要功能）")
        return 0

    if missing_required:
        print(f"❌ 缺少必需依赖：{', '.join(missing_required)}")
        cmd = f"{sys.executable} -m pip install -r {os.path.join(root, 'requirements.txt')}"
        print(f"   修复：{cmd}")
    if lost:
        print(f"❌ skill 文件缺失 {len(lost)} 个，建议重新安装本 skill")

    if install and missing_required:
        print("\n[INFO] 正在自动安装...")
        r = subprocess.run([sys.executable, "-m", "pip", "install"] + missing_required)
        print("[INFO] 安装结束，请重新运行本脚本确认。" if r.returncode == 0 else "[ERROR] 安装失败")
    return 1


if __name__ == '__main__':
    sys.exit(main())
