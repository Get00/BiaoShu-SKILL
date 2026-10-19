#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件读写工具 —— 不依赖任何 MCP / 外部服务，纯 Python 标准库实现

用途：把 skill 迁移到别的电脑后，即使宿主的 MCP 文件工具不可用，
     也能完成读文件、写文件、列目录、搜索等基本操作。

命令行：
    python fsutil.py read   <path>                      # 读文本（自动识别编码）
    python fsutil.py write  <path> --text "内容"         # 写文本
    python fsutil.py write  <path> --stdin              # 从标准输入写
    python fsutil.py ls     <dir> [--recursive]
    python fsutil.py glob   <dir> --pattern "*.docx"
    python fsutil.py grep   <dir> --pattern "正则" [--glob "*.py"]
    python fsutil.py stat   <path>
    python fsutil.py mkdir  <dir>
    python fsutil.py copy   <src> <dst>
    python fsutil.py rm     <path> [--recursive]

Python API：
    from fsutil import read_text, write_text, read_bytes, write_bytes, \\
                       list_dir, glob_files, grep_files, ensure_dir, copy_path
版本：v1.0 (2026-10-13)
"""

import io
import os
import re
import sys
import fnmatch
import shutil

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

ENCODINGS = ('utf-8-sig', 'utf-8', 'gb18030', 'gbk', 'big5', 'latin-1')


# ---------------- Python API ----------------

def read_text(path, encoding=None):
    """读文本文件。encoding 为空时依次尝试常见编码（含 GBK，中文文档常用）"""
    if encoding:
        with open(path, 'r', encoding=encoding, errors='replace') as f:
            return f.read()
    data = read_bytes(path)
    for enc in ENCODINGS:
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode('utf-8', errors='replace')


def write_text(path, text, encoding='utf-8'):
    """写文本文件（默认 utf-8 无 BOM，跨平台安全）"""
    ensure_dir(os.path.dirname(os.path.abspath(path)))
    with open(path, 'w', encoding=encoding, newline='') as f:
        f.write(text)
    return path


def read_bytes(path):
    with open(path, 'rb') as f:
        return f.read()


def write_bytes(path, data):
    ensure_dir(os.path.dirname(os.path.abspath(path)))
    with open(path, 'wb') as f:
        f.write(data)
    return path


def ensure_dir(path):
    if path:
        os.makedirs(path, exist_ok=True)
    return path


def list_dir(path, recursive=False):
    """列目录，返回 [{name, path, is_dir, size}]"""
    out = []
    if recursive:
        for root, dirs, files in os.walk(path):
            for d in dirs:
                p = os.path.join(root, d)
                out.append({"name": d, "path": p, "is_dir": True, "size": None})
            for fn in files:
                p = os.path.join(root, fn)
                out.append({"name": fn, "path": p, "is_dir": False,
                            "size": os.path.getsize(p)})
    else:
        for name in sorted(os.listdir(path)):
            p = os.path.join(path, name)
            out.append({"name": name, "path": p, "is_dir": os.path.isdir(p),
                        "size": None if os.path.isdir(p) else os.path.getsize(p)})
    return out


def glob_files(root, pattern):
    """按通配符（如 *.docx）递归查找文件"""
    hits = []
    for r, dirs, files in os.walk(root):
        for fn in files:
            if fnmatch.fnmatch(fn, pattern):
                hits.append(os.path.join(r, fn))
    return hits


TEXT_EXT = {'.txt', '.md', '.json', '.csv', '.py', '.html', '.xml', '.yml', '.yaml', '.log'}


def grep_files(root, pattern, file_glob='*', ignore_binary=True):
    """在目录中按正则搜索，返回 [(path, lineno, line)]"""
    rx = re.compile(pattern)
    hits = []
    for r, dirs, files in os.walk(root):
        for fn in files:
            if not fnmatch.fnmatch(fn, file_glob):
                continue
            p = os.path.join(r, fn)
            ext = os.path.splitext(fn)[1].lower()
            if ignore_binary and ext and ext not in TEXT_EXT:
                continue
            try:
                txt = read_text(p)
            except Exception:
                continue
            for i, line in enumerate(txt.splitlines(), 1):
                if rx.search(line):
                    hits.append((p, i, line.strip()[:200]))
    return hits


def copy_path(src, dst):
    ensure_dir(os.path.dirname(os.path.abspath(dst)))
    if os.path.isdir(src):
        shutil.copytree(src, dst, dirs_exist_ok=True)
    else:
        shutil.copy2(src, dst)
    return dst


def remove_path(path, recursive=False):
    if os.path.isdir(path):
        if recursive:
            shutil.rmtree(path)
        else:
            os.rmdir(path)
    elif os.path.exists(path):
        os.remove(path)
    return path


# ---------------- 命令行 ----------------

def main():
    import argparse
    ap = argparse.ArgumentParser(description='文件读写工具（不依赖 MCP）')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('read'); p.add_argument('path')
    p = sub.add_parser('write'); p.add_argument('path')
    p.add_argument('--text', default=None); p.add_argument('--stdin', action='store_true')
    p.add_argument('--encoding', default='utf-8')
    p = sub.add_parser('ls'); p.add_argument('path'); p.add_argument('--recursive', '-r', action='store_true')
    p = sub.add_parser('glob'); p.add_argument('path'); p.add_argument('--pattern', required=True)
    p = sub.add_parser('grep'); p.add_argument('path'); p.add_argument('--pattern', required=True)
    p.add_argument('--glob', default='*')
    p = sub.add_parser('stat'); p.add_argument('path')
    p = sub.add_parser('mkdir'); p.add_argument('path')
    p = sub.add_parser('copy'); p.add_argument('src'); p.add_argument('dst')
    p = sub.add_parser('rm'); p.add_argument('path'); p.add_argument('--recursive', '-r', action='store_true')

    a = ap.parse_args()

    if a.cmd == 'read':
        print(read_text(a.path))
    elif a.cmd == 'write':
        text = sys.stdin.read() if a.stdin else (a.text or '')
        print("已写入：", write_text(a.path, text, a.encoding))
    elif a.cmd == 'ls':
        for it in list_dir(a.path, a.recursive):
            tag = '[目录]' if it['is_dir'] else f"{it['size']:>9} B"
            print(f"  {tag}  {it['path']}")
    elif a.cmd == 'glob':
        for p_ in glob_files(a.path, a.pattern):
            print(p_)
    elif a.cmd == 'grep':
        for p_, ln, line in grep_files(a.path, a.pattern, a.glob):
            print(f"{p_}:{ln}: {line}")
    elif a.cmd == 'stat':
        st = os.stat(a.path)
        print(f"路径: {a.path}\n大小: {st.st_size} 字节\n修改: {st.st_mtime}")
    elif a.cmd == 'mkdir':
        print("已创建：", ensure_dir(a.path))
    elif a.cmd == 'copy':
        print("已复制：", copy_path(a.src, a.dst))
    elif a.cmd == 'rm':
        print("已删除：", remove_path(a.path, a.recursive))


if __name__ == '__main__':
    main()
