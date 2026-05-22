#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
关键要求提取脚本
从招标文件中提取资质要求、技术要求、商务要求、废标条款等关键信息
版本：v1.0 (2026-05-21)
"""

import io
import json
import re
import os
import sys
from pathlib import Path
from typing import Dict, List, Optional

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


class RequirementsExtractor:
    """关键要求提取器"""

    # 资质相关关键词
    QUALIFICATION_KEYWORDS = [
        '资质', '认证', '证书', '资格', '许可',
        'ISO', 'CMMI', 'ITSS', 'PMP', '信息系统集成'
    ]

    # 技术要求关键词
    TECHNICAL_KEYWORDS = [
        '技术参数', '性能指标', '功能要求', '系统要求',
        '并发', '响应时间', '可用性', '兼容性'
    ]

    # 商务要求关键词
    COMMERCIAL_KEYWORDS = [
        '工期', '交付时间', '付款方式', '质保', '售后服务',
        '培训', '维护', '响应时间'
    ]

    # 废标条款关键词
    DISQUALIFICATION_KEYWORDS = [
        '废标', '无效投标', '不予受理', '取消资格',
        '资格性审查', '符合性审查', '实质性响应'
    ]

    def __init__(self, text: str):
        self.text = text
        self.requirements = {
            "qualification_requirements": {
                "enterprise": [],
                "certifications": [],
                "personnel": [],
                "performance": []
            },
            "technical_requirements": {
                "parameters": [],
                "functions": [],
                "performance": []
            },
            "commercial_requirements": {
                "delivery": {},
                "payment": {},
                "warranty": {},
                "training": {}
            },
            "disqualification_clauses": []
        }

    def extract_all(self) -> Dict:
        """提取所有关键要求"""
        print("[INFO] 开始提取关键要求...")

        self._extract_qualifications()
        self._extract_technical_requirements()
        self._extract_commercial_requirements()
        self._extract_disqualification_clauses()

        print(f"[INFO] 提取完成:")
        print(f"  - 资质要求: {sum(len(v) for v in self.requirements['qualification_requirements'].values())} 项")
        print(f"  - 技术要求: {sum(len(v) for v in self.requirements['technical_requirements'].values())} 项")
        print(f"  - 商务要求: {len(self.requirements['commercial_requirements'])} 项")
        print(f"  - 废标条款: {len(self.requirements['disqualification_clauses'])} 项")

        return self.requirements

    def _extract_qualifications(self):
        """提取资质要求"""
        lines = self.text.split('\n')
        in_qualification_section = False

        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue

            # 检查是否进入资质要求章节
            if any(kw in line for kw in ['资质要求', '资格要求', '投标人资格']):
                in_qualification_section = True
                continue

            # 检查是否离开资质要求章节
            if in_qualification_section and re.match(r'第[一二三四五六七八九十\d]+章', line):
                in_qualification_section = False
                continue

            if in_qualification_section:
                # 提取企业资质
                if any(kw in line for kw in ['企业资质', '公司资质', '营业执照']):
                    self.requirements["qualification_requirements"]["enterprise"].append({
                        "type": "企业资质",
                        "requirement": line,
                        "mandatory": True
                    })

                # 提取认证证书
                if any(kw in line for kw in ['ISO', 'CMMI', 'ITSS', '认证']):
                    self.requirements["qualification_requirements"]["certifications"].append({
                        "type": "认证证书",
                        "requirement": line,
                        "mandatory": True
                    })

                # 提取人员资格
                if any(kw in line for kw in ['项目经理', 'PMP', '高级工程师', '人员资格']):
                    self.requirements["qualification_requirements"]["personnel"].append({
                        "type": "人员资格",
                        "requirement": line,
                        "mandatory": True
                    })

                # 提取业绩要求
                if any(kw in line for kw in ['类似项目', '业绩', '合同金额']):
                    self.requirements["qualification_requirements"]["performance"].append({
                        "type": "业绩要求",
                        "requirement": line,
                        "mandatory": True
                    })

    def _extract_technical_requirements(self):
        """提取技术要求"""
        lines = self.text.split('\n')
        in_technical_section = False

        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue

            # 检查是否进入技术要求章节
            if any(kw in line for kw in ['技术要求', '技术参数', '技术规格', '功能需求']):
                in_technical_section = True
                continue

            # 检查是否离开技术要求章节
            if in_technical_section and re.match(r'第[一二三四五六七八九十\d]+章', line):
                in_technical_section = False
                continue

            if in_technical_section:
                # 提取性能指标
                if any(kw in line for kw in ['并发', '响应时间', '可用性', '性能']):
                    self.requirements["technical_requirements"]["performance"].append({
                        "category": "性能指标",
                        "item": line,
                        "mandatory": True
                    })

                # 提取功能要求
                if any(kw in line for kw in ['功能', '模块', '子系统']):
                    self.requirements["technical_requirements"]["functions"].append({
                        "module": line,
                        "requirement": line,
                        "mandatory": True
                    })

    def _extract_commercial_requirements(self):
        """提取商务要求"""
        # 提取工期要求
        duration_pattern = r'工期[：:]\s*(\d+)[日天月年]'
        match = re.search(duration_pattern, self.text)
        if match:
            self.requirements["commercial_requirements"]["delivery"] = {
                "total_duration": f"{match.group(1)}日历天"
            }

        # 提取质保要求
        warranty_pattern = r'质保[期务]?[：:]\s*(\d+)[年月]'
        match = re.search(warranty_pattern, self.text)
        if match:
            self.requirements["commercial_requirements"]["warranty"] = {
                "period": f"{match.group(1)}年"
            }

        # 提取付款方式
        payment_keywords = ['付款方式', '付款条件', '支付方式']
        for kw in payment_keywords:
            if kw in self.text:
                # 提取包含付款信息的段落
                pattern = f'{kw}[：:]([^。]+。)'
                match = re.search(pattern, self.text)
                if match:
                    self.requirements["commercial_requirements"]["payment"] = {
                        "terms": match.group(1).strip()
                    }
                break

    def _extract_disqualification_clauses(self):
        """提取废标条款"""
        lines = self.text.split('\n')
        in_disqualification_section = False

        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue

            # 检查是否进入废标条款章节
            if any(kw in line for kw in ['废标条款', '无效投标', '不予受理']):
                in_disqualification_section = True
                continue

            # 检查是否离开废标条款章节
            if in_disqualification_section and re.match(r'第[一二三四五六七八九十\d]+章', line):
                in_disqualification_section = False
                continue

            if in_disqualification_section:
                if any(kw in line for kw in ['未提供', '超过', '不符合', '未响应']):
                    self.requirements["disqualification_clauses"].append({
                        "type": "资格性审查",
                        "clause": line,
                        "consequence": "废标"
                    })


def extract_requirements_from_file(file_path: str) -> Dict:
    """从文件提取关键要求"""
    from parse_bid_file import parse_file

    # 解析文件
    text = parse_file(file_path)
    if text.startswith("[错误]"):
        return {"error": text}

    # 提取关键要求
    extractor = RequirementsExtractor(text)
    return extractor.extract_all()


def main():
    import argparse

    parser = argparse.ArgumentParser(description='招标文件关键要求提取工具')
    parser.add_argument('input_file', help='输入招标文件路径')
    parser.add_argument('--output', '-o', default='key_requirements.json', help='输出JSON文件路径')

    args = parser.parse_args()

    if not os.path.exists(args.input_file):
        print(f"[ERROR] 文件不存在: {args.input_file}")
        sys.exit(1)

    print(f"[INFO] 正在解析文件: {args.input_file}")

    result = extract_requirements_from_file(args.input_file)

    # 保存结果
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"[INFO] 结果已保存至: {output_path}")


if __name__ == '__main__':
    main()
