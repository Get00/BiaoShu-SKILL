#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
占位符生成脚本
在标书内容中插入占位符标记
版本：v1.0 (2026-05-21)
"""

import io
import re
import os
import sys
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


# 占位符类型定义
PLACEHOLDER_TYPES = {
    "image": {
        "prefix": "[图片占位符]",
        "fields": ["type", "description", "size"],
        "template": "[图片占位符]\n类型：{type}\n描述：{description}\n建议尺寸：{size}"
    },
    "table": {
        "prefix": "[表格占位符]",
        "fields": ["type", "description", "columns", "content"],
        "template": "[表格占位符]\n类型：{type}\n描述：{description}\n建议列数：{columns}\n建议内容：{content}"
    },
    "qualification": {
        "prefix": "[资质占位符]",
        "fields": ["type", "description", "requirement"],
        "template": "[资质占位符]\n类型：{type}\n描述：{description}\n要求：{requirement}"
    }
}

# 常用占位符模板
COMMON_PLACEHOLDERS = {
    "技术架构图": {
        "type": "image",
        "data": {
            "type": "技术架构图",
            "description": "系统总体架构图",
            "size": "宽15cm × 高10cm"
        }
    },
    "网络拓扑图": {
        "type": "image",
        "data": {
            "type": "网络拓扑图",
            "description": "系统网络拓扑结构图",
            "size": "宽15cm × 高10cm"
        }
    },
    "业务流程图": {
        "type": "image",
        "data": {
            "type": "业务流程图",
            "description": "核心业务流程图",
            "size": "宽15cm × 高10cm"
        }
    },
    "功能模块清单": {
        "type": "table",
        "data": {
            "type": "功能模块清单",
            "description": "系统功能模块列表",
            "columns": "5列",
            "content": "模块名称、功能描述、技术选型、优先级、预计工作量"
        }
    },
    "项目进度表": {
        "type": "table",
        "data": {
            "type": "项目进度计划表",
            "description": "项目实施进度安排",
            "columns": "6列",
            "content": "阶段名称、开始时间、结束时间、主要任务、交付物、负责人"
        }
    },
    "人员配置表": {
        "type": "table",
        "data": {
            "type": "人员配置表",
            "description": "项目团队人员配置",
            "columns": "5列",
            "content": "角色、姓名、职责、资质证书、投入时间"
        }
    },
    "项目业绩表": {
        "type": "table",
        "data": {
            "type": "项目业绩表",
            "description": "近3年类似项目清单",
            "columns": "6列",
            "content": "项目名称、合同金额、完成时间、项目规模、客户名称、验收情况"
        }
    },
    "服务响应时间表": {
        "type": "table",
        "data": {
            "type": "服务响应时间表",
            "description": "售后服务响应时间承诺",
            "columns": "4列",
            "content": "故障级别、响应时间、解决时间、联系方式"
        }
    },
    "企业营业执照": {
        "type": "qualification",
        "data": {
            "type": "企业营业执照",
            "description": "企业营业执照副本",
            "requirement": "有效期内的彩色扫描件"
        }
    },
    "信息系统集成资质": {
        "type": "qualification",
        "data": {
            "type": "信息系统集成资质",
            "description": "信息系统集成及服务资质证书",
            "requirement": "有效期内的彩色扫描件"
        }
    },
    "ISO9001认证证书": {
        "type": "qualification",
        "data": {
            "type": "ISO9001认证证书",
            "description": "ISO9001质量管理体系认证证书",
            "requirement": "有效期内的彩色扫描件"
        }
    },
    "项目经理PMP证书": {
        "type": "qualification",
        "data": {
            "type": "项目经理PMP证书",
            "description": "项目经理PMP认证证书",
            "requirement": "有效期内的彩色扫描件"
        }
    }
}


def generate_placeholder(placeholder_type: str, data: Dict) -> str:
    """生成占位符文本"""
    template_info = PLACEHOLDER_TYPES.get(placeholder_type)
    if not template_info:
        return f"[未知占位符类型: {placeholder_type}]"

    template = template_info["template"]
    return template.format(**data)


def insert_placeholder(content: str, placeholder_name: str, position: str = "end") -> str:
    """在内容中插入占位符"""
    placeholder_info = COMMON_PLACEHOLDERS.get(placeholder_name)
    if not placeholder_info:
        return content

    placeholder_type = placeholder_info["type"]
    placeholder_data = placeholder_info["data"]

    placeholder_text = generate_placeholder(placeholder_type, placeholder_data)

    if position == "end":
        return content + "\n\n" + placeholder_text
    elif position == "start":
        return placeholder_text + "\n\n" + content
    else:
        return content


def analyze_content_for_placeholders(content: str) -> List[Dict]:
    """分析内容，建议需要插入的占位符"""
    suggestions = []

    # 检测技术方案相关章节
    if any(keyword in content for keyword in ['架构', '系统设计', '技术方案']):
        suggestions.append({
            "name": "技术架构图",
            "reason": "检测到技术架构相关描述，建议插入架构图"
        })

    # 检测功能设计相关章节
    if any(keyword in content for keyword in ['功能', '模块', '子系统']):
        suggestions.append({
            "name": "功能模块清单",
            "reason": "检测到功能模块描述，建议插入功能清单表"
        })

    # 检测实施计划相关章节
    if any(keyword in content for keyword in ['实施', '进度', '计划']):
        suggestions.append({
            "name": "项目进度表",
            "reason": "检测到实施计划描述，建议插入进度表"
        })

    # 检测人员配置相关章节
    if any(keyword in content for keyword in ['人员', '团队', '配置']):
        suggestions.append({
            "name": "人员配置表",
            "reason": "检测到人员配置描述，建议插入人员表"
        })

    # 检测项目业绩相关章节
    if any(keyword in content for keyword in ['业绩', '案例', '经验']):
        suggestions.append({
            "name": "项目业绩表",
            "reason": "检测到项目业绩描述，建议插入业绩表"
        })

    # 检测资质相关章节
    if any(keyword in content for keyword in ['资质', '证书', '认证']):
        suggestions.append({
            "name": "企业营业执照",
            "reason": "检测到资质描述，建议插入资质证书"
        })

    return suggestions


def process_chapters(chapters_dir: str, output_dir: str = None):
    """处理所有章节文件，插入占位符"""
    if output_dir is None:
        output_dir = chapters_dir

    os.makedirs(output_dir, exist_ok=True)

    for filename in sorted(os.listdir(chapters_dir)):
        if not filename.endswith('.md'):
            continue

        filepath = os.path.join(chapters_dir, filename)
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()

        # 分析内容，获取建议
        suggestions = analyze_content_for_placeholders(content)

        # 插入占位符
        for suggestion in suggestions:
            content = insert_placeholder(content, suggestion["name"])

        # 保存处理后的文件
        output_path = os.path.join(output_dir, filename)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)

        print(f"[INFO] 已处理: {filename} (插入 {len(suggestions)} 个占位符)")


def main():
    import argparse

    parser = argparse.ArgumentParser(description='占位符生成工具')
    parser.add_argument('chapters_dir', help='章节目录路径')
    parser.add_argument('--output-dir', '-o', help='输出目录路径')
    parser.add_argument('--list', '-l', action='store_true', help='列出所有可用占位符')

    args = parser.parse_args()

    if args.list:
        print("可用占位符:")
        for name, info in COMMON_PLACEHOLDERS.items():
            print(f"  - {name}: {info['data']['description']}")
        return

    if not os.path.exists(args.chapters_dir):
        print(f"[ERROR] 章节目录不存在: {args.chapters_dir}")
        sys.exit(1)

    print(f"[INFO] 开始处理章节文件...")

    process_chapters(args.chapters_dir, args.output_dir)

    print(f"[INFO] 处理完成")


if __name__ == '__main__':
    main()
