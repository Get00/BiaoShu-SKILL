#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
行业特定内容生成器
根据不同行业生成专业的投标文件内容
版本：v1.0 (2026-05-22)
"""

import json
from pathlib import Path
from typing import Dict, List, Optional

from industry_detector import IndustryDetector


class IndustryContentGenerator:
    """行业特定内容生成器"""

    def __init__(self, industry_code: str = "general"):
        """
        初始化内容生成器

        Args:
            industry_code: 行业代码
        """
        self.industry_code = industry_code
        self.detector = IndustryDetector()
        self.industry_config = self.detector.get_industry_config(industry_code)
        self.chapter_structure = self.detector.get_chapter_structure(industry_code)
        self.key_sections = self.detector.get_key_sections(industry_code)

    def generate_chapter_title(self, chapter_index: int, chapter_name: str) -> str:
        """
        生成章节标题

        Args:
            chapter_index: 章节索引
            chapter_name: 章节名称

        Returns:
            str: 格式化的章节标题
        """
        # 将数字转换为中文数字
        chinese_numbers = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
                          "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十"]

        if chapter_index <= len(chinese_numbers):
            number = chinese_numbers[chapter_index - 1]
        else:
            number = str(chapter_index)

        return f"第{number}章 {chapter_name}"

    def generate_section_title(self, chapter_index: int, section_index: int, section_name: str) -> str:
        """
        生成小节标题

        Args:
            chapter_index: 章节索引
            section_index: 小节索引
            section_name: 小节名称

        Returns:
            str: 格式化的小节标题
        """
        return f"{chapter_index}.{section_index} {section_name}"

    def get_industry_specific_content(self, section_type: str, project_name: str) -> str:
        """
        获取行业特定内容

        Args:
            section_type: 章节类型
            project_name: 项目名称

        Returns:
            str: 行业特定内容
        """
        # 根据行业和章节类型生成内容
        content_generators = {
            "it_information": self._generate_it_content,
            "construction": self._generate_construction_content,
            "medical": self._generate_medical_content,
            "education": self._generate_education_content,
            "manufacturing": self._generate_manufacturing_content,
            "logistics": self._generate_logistics_content,
            "consulting": self._generate_consulting_content,
            "general": self._generate_general_content
        }

        generator = content_generators.get(self.industry_code, self._generate_general_content)
        return generator(section_type, project_name)

    def _generate_it_content(self, section_type: str, project_name: str) -> str:
        """生成IT/信息化行业内容"""
        templates = {
            "技术方案": f"""本项目技术方案采用先进的技术架构，确保系统的高可用性、可扩展性和安全性。

1. **总体架构设计**
   - 采用微服务架构，实现服务解耦和独立部署
   - 使用容器化技术，支持快速部署和弹性伸缩
   - 实施DevOps流程，提高开发效率和系统稳定性

2. **技术选型**
   - 后端：Java/Python/Go等主流开发语言
   - 前端：Vue.js/React等现代前端框架
   - 数据库：MySQL/PostgreSQL/Redis等
   - 中间件：RabbitMQ/Kafka等消息队列

3. **安全设计**
   - 实施身份认证和权限控制
   - 数据加密传输和存储
   - 安全审计和日志记录""",

            "项目实施": f"""本项目采用敏捷开发方法，确保项目按时高质量交付。

1. **实施计划**
   - 需求分析阶段：2-4周
   - 系统设计阶段：2-3周
   - 开发实现阶段：8-12周
   - 测试验收阶段：4-6周
   - 部署上线阶段：2-4周

2. **人员配置**
   - 项目经理：1名，负责项目整体管理
   - 技术负责人：1名，负责技术方案设计
   - 开发工程师：3-5名，负责系统开发
   - 测试工程师：2名，负责系统测试

3. **质量保证**
   - 代码审查制度
   - 自动化测试
   - 持续集成/持续部署""",

            "售后服务": f"""我们提供全面的售后服务，确保系统稳定运行。

1. **服务体系**
   - 7×24小时技术支持
   - 专属客户经理
   - 定期巡检和维护

2. **响应时间**
   - 紧急故障：15分钟响应，2小时内解决
   - 严重故障：30分钟响应，4小时内解决
   - 一般故障：2小时内响应，8小时内解决

3. **培训计划**
   - 系统管理员培训
   - 业务用户培训
   - 技术人员培训"""
        }

        return templates.get(section_type, self._generate_general_content(section_type, project_name))

    def _generate_construction_content(self, section_type: str, project_name: str) -> str:
        """生成建筑工程行业内容"""
        templates = {
            "施工组织设计": f"""本工程施工组织设计科学合理，确保工程质量和安全。

1. **施工部署**
   - 施工区域划分
   - 施工顺序安排
   - 施工流水段划分

2. **施工方案**
   - 基础工程施工方案
   - 主体结构施工方案
   - 装饰装修施工方案
   - 安装工程施工方案

3. **施工进度**
   - 总工期控制
   - 里程碑节点
   - 进度保证措施""",

            "质量保证": f"""本工程实施全面质量管理，确保工程质量达到优良标准。

1. **质量目标**
   - 工程质量达到国家验收规范标准
   - 确保工程一次验收合格率100%
   - 争创优质工程

2. **质量体系**
   - 建立质量管理体系
   - 实施ISO9001质量管理体系
   - 定期质量审核和改进

3. **质量措施**
   - 材料进场检验
   - 工序质量控制
   - 隐蔽工程验收
   - 成品保护措施""",

            "安全文明施工": f"""本工程高度重视安全生产和文明施工。

1. **安全目标**
   - 杜绝重大安全事故
   - 杜绝重大机械设备事故
   - 杜绝重大火灾事故

2. **安全体系**
   - 建立安全管理体系
   - 实施安全生产责任制
   - 定期安全培训和演练

3. **安全措施**
   - 安全防护设施
   - 安全警示标志
   - 个人防护用品
   - 应急预案"""
        }

        return templates.get(section_type, self._generate_general_content(section_type, project_name))

    def _generate_medical_content(self, section_type: str, project_name: str) -> str:
        """生成医疗健康行业内容"""
        templates = {
            "技术方案": f"""本项目技术方案符合医疗行业标准和规范。

1. **产品配置**
   - 设备选型符合国家医疗器械标准
   - 产品性能满足临床需求
   - 产品兼容性好，易于集成

2. **技术参数**
   - 设备技术参数先进
   - 检测精度高
   - 稳定性好

3. **功能说明**
   - 功能完善，满足临床需求
   - 操作简便，易于使用
   - 数据安全，符合隐私保护要求""",

            "供货方案": f"""本项目供货方案确保设备及时、安全到达。

1. **供货计划**
   - 供货周期：合同签订后XX天内
   - 供货方式：厂家直供
   - 供货地点：用户指定地点

2. **物流方案**
   - 专业物流公司运输
   - 设备包装符合运输要求
   - 运输保险全覆盖

3. **安装调试**
   - 专业工程师现场安装
   - 设备调试和验收
   - 操作培训""",

            "售后服务": f"""我们提供全面的医疗设备售后服务。

1. **服务体系**
   - 厂家直接服务
   - 全国服务网点
   - 专业服务团队

2. **响应时间**
   - 电话支持：7×24小时
   - 远程支持：2小时内响应
   - 现场支持：24小时内到达

3. **维修保养**
   - 定期巡检和维护
   - 备品备件供应
   - 软件升级服务"""
        }

        return templates.get(section_type, self._generate_general_content(section_type, project_name))

    def _generate_education_content(self, section_type: str, project_name: str) -> str:
        """生成教育服务行业内容"""
        templates = {
            "技术方案": f"""本项目技术方案符合教育信息化发展趋势。

1. **系统架构**
   - 采用云架构，支持多校区部署
   - 模块化设计，易于扩展
   - 移动端支持，随时随地访问

2. **功能模块**
   - 教学管理模块
   - 学生管理模块
   - 资源管理模块
   - 数据分析模块

3. **技术参数**
   - 系统响应时间：≤3秒
   - 并发用户数：≥1000
   - 数据安全性：符合等保要求""",

            "培训计划": f"""我们提供全面的教育信息化培训服务。

1. **教师培训**
   - 系统操作培训
   - 教学资源制作培训
   - 信息化教学方法培训

2. **学生培训**
   - 系统使用培训
   - 在线学习培训

3. **管理员培训**
   - 系统管理培训
   - 数据维护培训
   - 故障处理培训""",

            "售后服务": f"""我们提供全面的教育信息化售后服务。

1. **服务体系**
   - 专属客户经理
   - 技术支持团队
   - 定期巡检维护

2. **响应时间**
   - 电话支持：7×24小时
   - 远程支持：2小时内响应
   - 现场支持：24小时内到达

3. **技术支持**
   - 系统升级服务
   - 数据迁移服务
   - 定制开发服务"""
        }

        return templates.get(section_type, self._generate_general_content(section_type, project_name))

    def _generate_manufacturing_content(self, section_type: str, project_name: str) -> str:
        """生成制造业行业内容"""
        templates = {
            "技术方案": f"""本项目技术方案采用先进制造技术，确保产品质量。

1. **产品设计**
   - 采用CAD/CAM/CAE技术
   - 有限元分析优化
   - 可靠性设计

2. **材料选型**
   - 材料符合国家标准
   - 材料性能满足要求
   - 材料供应商资质齐全

3. **工艺流程**
   - 工艺路线合理
   - 工艺参数优化
   - 工艺文件齐全""",

            "质量保证": f"""本项目实施全面质量管理，确保产品质量。

1. **质量目标**
   - 产品合格率≥99%
   - 客户满意度≥95%
   - 质量投诉率≤1%

2. **质量体系**
   - ISO9001质量管理体系
   - ISO14001环境管理体系
   - ISO45001职业健康安全管理体系

3. **质量措施**
   - 来料检验
   - 过程检验
   - 成品检验
   - 出厂检验""",

            "售后服务": f"""我们提供全面的制造业售后服务。

1. **服务体系**
   - 厂家直接服务
   - 全国服务网点
   - 专业服务团队

2. **响应时间**
   - 电话支持：7×24小时
   - 远程支持：2小时内响应
   - 现场支持：24小时内到达

3. **维修保养**
   - 定期巡检和维护
   - 备品备件供应
   - 技术培训服务"""
        }

        return templates.get(section_type, self._generate_general_content(section_type, project_name))

    def _generate_logistics_content(self, section_type: str, project_name: str) -> str:
        """生成物流运输行业内容"""
        templates = {
            "服务方案": f"""本项目服务方案确保物流服务高效、安全。

1. **服务内容**
   - 运输服务
   - 仓储服务
   - 配送服务
   - 增值服务

2. **运营计划**
   - 运力配置
   - 线路规划
   - 时间安排

3. **服务标准**
   - 准时率≥98%
   - 货损率≤0.1%
   - 客户满意度≥95%""",

            "资源配置": f"""本项目资源配置充足，确保服务能力和质量。

1. **车辆配置**
   - 车辆类型齐全
   - 车辆数量充足
   - 车辆状况良好

2. **人员配置**
   - 驾驶员资质齐全
   - 人员培训到位
   - 人员稳定性高

3. **仓储设施**
   - 仓库面积充足
   - 设施设备完善
   - 管理系统先进""",

            "安全管理": f"""本项目高度重视安全管理。

1. **安全目标**
   - 杜绝重大安全事故
   - 杜绝重大货损事故
   - 杜绝重大交通事故

2. **安全体系**
   - 建立安全管理体系
   - 实施安全责任制
   - 定期安全培训

3. **安全措施**
   - 车辆GPS监控
   - 驾驶员行为监控
   - 应急预案演练"""
        }

        return templates.get(section_type, self._generate_general_content(section_type, project_name))

    def _generate_consulting_content(self, section_type: str, project_name: str) -> str:
        """生成咨询服务行业内容"""
        templates = {
            "服务方案": f"""本项目服务方案专业、系统、可操作。

1. **服务内容**
   - 现状调研和分析
   - 方案设计和规划
   - 实施指导和培训
   - 效果评估和改进

2. **方法论**
   - 采用国际先进的方法论
   - 结合国内实际情况
   - 注重可操作性

3. **工具技术**
   - 专业分析工具
   - 数据分析技术
   - 知识管理平台""",

            "项目团队": f"""本项目团队专业、经验丰富。

1. **团队配置**
   - 项目总监：1名，15年以上经验
   - 项目经理：1名，10年以上经验
   - 咨询顾问：3-5名，5年以上经验

2. **人员资质**
   - 相关专业学历
   - 行业认证资质
   - 项目经验证明

3. **职责分工**
   - 明确的职责分工
   - 高效的协作机制
   - 透明的沟通机制""",

            "质量保证": f"""本项目实施全面质量管理。

1. **质量目标**
   - 客户满意度≥95%
   - 方案可行性≥90%
   - 问题解决率≥85%

2. **质量体系**
   - 项目管理体系
   - 质量控制体系
   - 知识管理体系

3. **质量措施**
   - 阶段性评审
   - 客户反馈机制
   - 持续改进机制"""
        }

        return templates.get(section_type, self._generate_general_content(section_type, project_name))

    def _generate_general_content(self, section_type: str, project_name: str) -> str:
        """生成通用行业内容"""
        templates = {
            "技术方案": f"""本项目技术方案先进、可靠、安全。

1. **总体方案**
   - 采用成熟的技术方案
   - 确保系统的可靠性
   - 注重系统的可扩展性

2. **详细设计**
   - 功能设计完善
   - 性能设计合理
   - 安全设计可靠

3. **技术参数**
   - 技术指标先进
   - 性能指标达标
   - 安全指标合规""",

            "项目实施": f"""本项目实施科学、规范、高效。

1. **实施计划**
   - 计划制定科学
   - 进度安排合理
   - 资源配置充足

2. **人员配置**
   - 团队专业
   - 经验丰富
   - 责任明确

3. **质量保证**
   - 质量体系完善
   - 质量措施到位
   - 质量控制严格""",

            "售后服务": f"""我们提供全面、专业的售后服务。

1. **服务体系**
   - 服务体系完善
   - 服务团队专业
   - 服务流程规范

2. **响应时间**
   - 响应及时
   - 处理高效
   - 反馈及时

3. **服务承诺**
   - 服务态度好
   - 服务质量高
   - 客户满意度高"""
        }

        return templates.get(section_type, f"本章节内容将根据{project_name}的具体需求进行详细编写。")

    def get_chapter_names(self) -> List[str]:
        """获取章节名称列表"""
        return self.chapter_structure

    def get_industry_name(self) -> str:
        """获取行业名称"""
        return self.industry_config.get("name", "通用行业")


def main():
    """测试函数"""
    # 测试不同行业
    industries = ["it_information", "construction", "medical", "education", "manufacturing", "logistics", "consulting", "general"]

    for industry_code in industries:
        print(f"\n{'='*60}")
        print(f"行业: {industry_code}")
        print(f"{'='*60}")

        generator = IndustryContentGenerator(industry_code)
        print(f"行业名称: {generator.get_industry_name()}")
        print(f"章节结构:")
        for i, chapter in enumerate(generator.get_chapter_names(), 1):
            print(f"  {i}. {chapter}")


if __name__ == '__main__':
    main()
