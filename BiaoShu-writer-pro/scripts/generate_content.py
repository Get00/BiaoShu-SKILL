#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书内容生成脚本
根据大纲和评分标准生成标书内容
版本：v1.0 (2026-05-21)
"""

import io
import json
import os
import sys
from pathlib import Path
from typing import Dict, List

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


class ContentGenerator:
    """标书内容生成器"""

    # 内容生成提示模板 - 更详细丰富
    CHAPTER_TEMPLATES = {
        "项目概述": """# 第{chapter_id}章 {title}

## {chapter_id}.1 项目背景

{background_content}

## {chapter_id}.2 项目目标

{objectives_content}

## {chapter_id}.3 项目范围

{scope_content}

## {chapter_id}.4 项目理解

{understanding_content}
""",
        "技术方案": """# 第{chapter_id}章 {title}

## {chapter_id}.1 总体架构设计

{architecture_content}

[图片占位符]
类型：技术架构图
描述：{project_name}系统总体架构图
建议尺寸：宽15cm × 高10cm

## {chapter_id}.2 功能设计

{function_content}

[表格占位符]
类型：功能模块清单
描述：系统功能模块列表
建议列数：5列
建议内容：模块名称、功能描述、技术选型、优先级、预计工作量

## {chapter_id}.3 数据设计

{data_content}

## {chapter_id}.4 接口设计

{interface_content}

## {chapter_id}.5 安全设计

{security_content}

## {chapter_id}.6 性能设计

{performance_content}
""",
        "项目实施": """# 第{chapter_id}章 {title}

## {chapter_id}.1 实施计划

{plan_content}

[表格占位符]
类型：项目进度计划表
描述：项目实施进度安排
建议列数：6列
建议内容：阶段名称、开始时间、结束时间、主要任务、交付物、负责人

## {chapter_id}.2 人员配置

{team_content}

[表格占位符]
类型：人员配置表
描述：项目团队人员配置
建议列数：5列
建议内容：角色、姓名、职责、资质证书、投入时间

## {chapter_id}.3 质量保证

{quality_content}

## {chapter_id}.4 风险控制

{risk_content}

## {chapter_id}.5 沟通管理

{communication_content}
""",
        "项目业绩": """# 第{chapter_id}章 {title}

## {chapter_id}.1 公司简介

{company_intro_content}

## {chapter_id}.2 类似项目经验

{experience_content}

[表格占位符]
类型：项目业绩表
描述：近3年类似项目清单
建议列数：6列
建议内容：项目名称、合同金额、完成时间、项目规模、客户名称、验收情况

## {chapter_id}.3 客户评价

{evaluation_content}

[资质占位符]
类型：客户感谢信/验收报告
描述：客户出具的项目评价文件
要求：盖章原件或彩色扫描件

## {chapter_id}.4 技术实力

{technical_strength_content}
""",
        "售后服务": """# 第{chapter_id}章 {title}

## {chapter_id}.1 服务体系

{service_content}

## {chapter_id}.2 响应时间

{response_content}

[表格占位符]
类型：服务响应时间表
描述：售后服务响应时间承诺
建议列数：4列
建议内容：故障级别、响应时间、解决时间、联系方式

## {chapter_id}.3 培训计划

{training_content}

[表格占位符]
类型：培训计划表
描述：项目培训安排
建议列数：5列
建议内容：培训内容、培训对象、培训方式、培训时间、培训地点

## {chapter_id}.4 备品备件

{spare_parts_content}
""",
        "附件": """# 附件

## 附件1 企业资质

[资质占位符]
类型：企业营业执照
描述：企业营业执照副本
要求：有效期内的彩色扫描件

[资质占位符]
类型：信息系统集成资质
描述：信息系统集成及服务资质证书
要求：有效期内的彩色扫描件

[资质占位符]
类型：ISO9001认证证书
描述：ISO9001质量管理体系认证证书
要求：有效期内的彩色扫描件

## 附件2 人员证书

[资质占位符]
类型：项目经理PMP证书
描述：项目经理PMP认证证书
要求：有效期内的彩色扫描件

[资质占位符]
类型：高级工程师证书
描述：技术人员高级工程师职称证书
要求：彩色扫描件

## 附件3 项目业绩证明

[资质占位符]
类型：项目合同
描述：类似项目合同关键页
要求：合同首页、金额页、签字盖章页

[资质占位符]
类型：验收报告
描述：项目验收报告
要求：客户盖章的验收报告
"""
    }

    def __init__(self, outline: Dict, scoring: Dict, bid_content: str = ""):
        self.outline = outline
        self.scoring = scoring
        self.bid_content = bid_content
        self.project_name = self._extract_project_name()

    def _extract_project_name(self) -> str:
        """提取项目名称"""
        doc_info = self.scoring.get('document_info', {})
        return doc_info.get('project_name', '本项目')

    def generate_all_chapters(self) -> Dict[str, str]:
        """生成所有章节内容"""
        print("[INFO] 开始生成标书内容...")

        chapters = {}
        outline_chapters = self.outline.get('outline', [])

        for chapter in outline_chapters:
            chapter_id = chapter.get('id', 0)
            chapter_title = chapter.get('title', '')
            chapter_score = chapter.get('score', 0)
            sections = chapter.get('sections', [])

            # 生成章节内容
            content = self._generate_chapter_content(chapter_id, chapter_title, chapter_score, sections)
            chapters[f"{chapter_id:02d}_{chapter_title}"] = content

            print(f"[INFO] 已生成: 第{chapter_id}章 {chapter_title}")

        print(f"[INFO] 内容生成完成，共 {len(chapters)} 章")
        return chapters

    def _generate_chapter_content(self, chapter_id: int, title: str, score: float, sections: List[Dict]) -> str:
        """生成单个章节内容"""
        # 查找匹配的模板
        template = None
        for key in self.CHAPTER_TEMPLATES:
            if key in title:
                template = self.CHAPTER_TEMPLATES[key]
                break

        if not template:
            template = self._generate_generic_template(chapter_id, title, sections)

        # 填充模板内容
        content = template.format(
            chapter_id=chapter_id,
            title=title,
            project_name=self.project_name,
            background_content=self._generate_background_content(),
            objectives_content=self._generate_objectives_content(),
            scope_content=self._generate_scope_content(),
            understanding_content=self._generate_understanding_content(),
            architecture_content=self._generate_architecture_content(),
            function_content=self._generate_function_content(),
            data_content=self._generate_data_content(),
            interface_content=self._generate_interface_content(),
            security_content=self._generate_security_content(),
            performance_content=self._generate_performance_content(),
            plan_content=self._generate_plan_content(),
            team_content=self._generate_team_content(),
            quality_content=self._generate_quality_content(),
            risk_content=self._generate_risk_content(),
            communication_content=self._generate_communication_content(),
            company_intro_content=self._generate_company_intro_content(),
            experience_content=self._generate_experience_content(),
            evaluation_content=self._generate_evaluation_content(),
            technical_strength_content=self._generate_technical_strength_content(),
            service_content=self._generate_service_content(),
            response_content=self._generate_response_content(),
            training_content=self._generate_training_content(),
            spare_parts_content=self._generate_spare_parts_content()
        )

        return content

    def _generate_generic_template(self, chapter_id: int, title: str, sections: List[Dict]) -> str:
        """生成通用章节模板"""
        template = f"# 第{chapter_id}章 {title}\n\n"

        for i, section in enumerate(sections, 1):
            section_title = section.get('title', f'小节{i}')
            template += f"## {chapter_id}.{i} {section_title}\n\n"
            template += f"{{section_{i}_content}}\n\n"

        return template

    def _generate_background_content(self) -> str:
        """生成项目背景内容"""
        return f"""随着信息化建设的不断深入，{self.project_name}的建设已成为提升管理效率、优化服务质量的重要手段。

本项目旨在通过建设先进的信息系统，实现业务流程的数字化、智能化，提高工作效率，降低运营成本，为用户提供更加便捷、高效的服务。

当前，{self.project_name}面临以下挑战：
1. 业务流程复杂，人工处理效率低下
2. 数据分散，缺乏统一管理
3. 系统老旧，难以满足新业务需求
4. 信息安全风险日益突出

因此，{self.project_name}的建设具有重要的现实意义和战略价值。"""

    def _generate_objectives_content(self) -> str:
        """生成项目目标内容"""
        return f"""本项目的总体目标是建设一个先进、可靠、安全、易用的信息系统，具体目标包括：

1. **业务目标**
   - 实现业务流程的全面数字化
   - 提高业务处理效率30%以上
   - 降低运营成本20%以上

2. **技术目标**
   - 建设高可用、高性能的系统平台
   - 实现与现有系统的无缝集成
   - 确保系统安全可靠运行

3. **管理目标**
   - 建立完善的运维管理体系
   - 提升信息化管理水平
   - 为决策提供数据支持"""

    def _generate_scope_content(self) -> str:
        """生成项目范围内容"""
        return f"""本项目的建设范围包括：

1. **系统建设范围**
   - 核心业务系统开发
   - 数据平台建设
   - 安全体系建设
   - 运维管理平台建设

2. **服务范围**
   - 需求分析与设计
   - 系统开发与测试
   - 系统部署与上线
   - 培训与技术支持

3. **交付物范围**
   - 系统软件及源代码
   - 技术文档
   - 培训材料
   - 验收报告"""

    def _generate_understanding_content(self) -> str:
        """生成项目理解内容"""
        return f"""我公司对{self.project_name}进行了深入的分析和理解，认为本项目具有以下特点：

1. **业务特点**
   - 业务流程复杂，涉及多个部门和环节
   - 数据量大，需要高效的数据处理能力
   - 用户群体多样，需要灵活的权限管理
   - 业务规则多变，需要灵活的配置能力

2. **技术特点**
   - 需要高可用、高性能的系统架构
   - 需要与多个现有系统进行集成
   - 需要严格的安全防护措施
   - 需要良好的扩展性和可维护性

3. **管理特点**
   - 项目周期紧，需要高效的项目管理
   - 质量要求高，需要严格的质量控制
   - 涉及多方协调，需要良好的沟通机制
   - 风险因素多，需要有效的风险控制

基于以上理解，我公司将采用先进的技术方案和成熟的项目管理方法，确保项目顺利实施。"""

    def _generate_architecture_content(self) -> str:
        """生成架构设计内容"""
        return f"""本系统采用先进的分层架构设计，主要包括以下几个层次：

1. **接入层**
   - 负责用户访问和请求接入
   - 支持Web、移动端等多种访问方式
   - 提供统一的身份认证和权限控制

2. **应用层**
   - 实现核心业务逻辑
   - 采用微服务架构，模块化设计
   - 支持业务流程的灵活配置

3. **数据层**
   - 负责数据存储和管理
   - 采用分布式数据库，确保高可用
   - 支持数据备份和恢复

4. **基础设施层**
   - 提供计算、存储、网络资源
   - 支持云部署和本地部署
   - 提供安全防护和监控告警"""

    def _generate_function_content(self) -> str:
        """生成功能设计内容"""
        return f"""本系统主要包括以下功能模块：

1. **用户管理模块**
   - 用户注册与登录
   - 权限管理与分配
   - 用户信息维护

2. **业务处理模块**
   - 核心业务流程处理
   - 业务规则配置
   - 业务数据统计

3. **数据管理模块**
   - 数据采集与导入
   - 数据清洗与转换
   - 数据查询与导出

4. **系统管理模块**
   - 系统配置管理
   - 日志管理与审计
   - 监控告警管理"""

    def _generate_data_content(self) -> str:
        """生成数据设计内容"""
        return f"""本系统的数据设计遵循以下原则：

1. **数据标准化**
   - 建立统一的数据标准和规范
   - 确保数据的一致性和准确性
   - 支持数据的共享和交换

2. **数据安全**
   - 敏感数据加密存储
   - 数据访问权限控制
   - 数据备份和恢复机制

3. **数据架构**
   - 采用关系型数据库存储结构化数据
   - 采用NoSQL数据库存储非结构化数据
   - 建立数据仓库支持数据分析"""

    def _generate_interface_content(self) -> str:
        """生成接口设计内容"""
        return f"""本系统的接口设计遵循标准化、规范化的原则，主要包括以下几类接口：

1. **外部接口**
   - 与现有业务系统的数据交换接口
   - 与第三方服务的集成接口
   - 与监管平台的数据上报接口

2. **内部接口**
   - 前后端分离的RESTful API接口
   - 微服务之间的RPC调用接口
   - 消息队列的异步通信接口

3. **接口规范**
   - 统一的接口命名规范
   - 统一的请求/响应格式
   - 统一的错误码定义
   - 完善的接口文档

4. **接口安全**
   - 接口身份认证
   - 接口访问权限控制
   - 接口调用频率限制
   - 接口数据加密传输"""

    def _generate_security_content(self) -> str:
        """生成安全设计内容"""
        return f"""本系统的安全设计包括以下方面：

1. **网络安全**
   - 部署防火墙和入侵检测系统
   - 实施网络隔离和访问控制
   - 定期进行安全漏洞扫描

2. **应用安全**
   - 实施身份认证和权限控制
   - 防止SQL注入、XSS等攻击
   - 敏感数据加密传输和存储

3. **数据安全**
   - 数据备份和恢复机制
   - 数据访问审计和日志记录
   - 敏感数据脱敏处理

4. **运维安全**
   - 运维人员权限管理
   - 操作审计和监控
   - 安全事件应急响应"""

    def _generate_performance_content(self) -> str:
        """生成性能设计内容"""
        return f"""本系统的性能设计目标是确保系统在高并发、大数据量场景下稳定运行，主要包括：

1. **性能指标**
   - 页面响应时间：≤3秒
   - 接口响应时间：≤1秒
   - 并发用户数：≥1000
   - 数据处理能力：≥10000条/秒

2. **性能优化策略**
   - 采用缓存技术，减少数据库访问
   - 采用异步处理，提高系统吞吐量
   - 采用负载均衡，分散系统压力
   - 采用CDN加速，提高访问速度

3. **性能监控**
   - 实时监控系统性能指标
   - 性能瓶颈分析和优化
   - 容量规划和扩展建议

4. **性能测试**
   - 负载测试：验证系统承载能力
   - 压力测试：验证系统极限性能
   - 稳定性测试：验证系统长期运行稳定性"""

    def _generate_plan_content(self) -> str:
        """生成实施计划内容"""
        return f"""本项目采用分阶段实施策略，具体计划如下：

1. **第一阶段：需求分析与设计（第1-30天）**
   - 深入调研业务需求
   - 完成系统架构设计
   - 编写详细设计文档

2. **第二阶段：系统开发（第31-120天）**
   - 核心功能模块开发
   - 系统集成和联调
   - 单元测试和集成测试

3. **第三阶段：测试与优化（第121-150天）**
   - 系统测试和性能测试
   - 用户验收测试
   - 问题修复和优化

4. **第四阶段：部署与上线（第151-180天）**
   - 系统部署和配置
   - 数据迁移和初始化
   - 正式上线运行"""

    def _generate_team_content(self) -> str:
        """生成人员配置内容"""
        return f"""本项目配备经验丰富的专业团队，具体配置如下：

1. **项目经理**
   - 负责项目整体管理和协调
   - 具备PMP认证和丰富的项目管理经验
   - 全职投入本项目

2. **技术负责人**
   - 负责技术方案设计和评审
   - 具备高级工程师职称
   - 全职投入本项目

3. **开发团队**
   - 配备5名资深开发工程师
   - 熟悉相关技术栈
   - 全职投入本项目

4. **测试团队**
   - 配备2名测试工程师
   - 负责系统测试和质量保证
   - 全职投入本项目"""

    def _generate_quality_content(self) -> str:
        """生成质量保证内容"""
        return f"""本项目实施全面的质量保证措施，包括：

1. **质量管理体系**
   - 建立完善的质量管理制度
   - 实施ISO9001质量管理体系
   - 定期进行质量审计

2. **过程质量控制**
   - 需求评审和设计评审
   - 代码审查和测试
   - 文档评审和验收

3. **产品质量保证**
   - 功能测试和性能测试
   - 安全测试和兼容性测试
   - 用户验收测试"""

    def _generate_risk_content(self) -> str:
        """生成风险控制内容"""
        return f"""本项目识别的主要风险及应对措施：

1. **技术风险**
   - 风险：技术方案可行性不足
   - 应对：充分调研，技术预研，专家评审

2. **进度风险**
   - 风险：项目进度延期
   - 应对：制定详细计划，定期跟踪，及时调整

3. **质量风险**
   - 风险：系统质量不达标
   - 应对：严格测试，代码审查，质量审计

4. **人员风险**
   - 风险：关键人员离职
   - 应对：人员备份，知识转移，团队建设"""

    def _generate_communication_content(self) -> str:
        """生成沟通管理内容"""
        return f"""本项目建立完善的沟通管理机制，确保项目各方信息畅通：

1. **沟通计划**
   - 项目启动会：项目开始时召开
   - 周例会：每周召开，汇报进展
   - 月度评审会：每月召开，评审成果
   - 专题会议：根据需要召开

2. **沟通方式**
   - 现场会议：重要事项面对面沟通
   - 视频会议：远程沟通
   - 邮件沟通：正式文件传递
   - 即时通讯：日常沟通

3. **沟通内容**
   - 项目进展报告
   - 问题和风险报告
   - 变更请求和审批
   - 会议纪要和决议

4. **沟通规范**
   - 统一的文档模板
   - 明确的审批流程
   - 及时的信息反馈
   - 完整的记录存档"""

    def _generate_experience_content(self) -> str:
        """生成项目经验内容"""
        return f"""公司近3年完成了多个类似项目，积累了丰富的经验：

1. **项目经验概述**
   - 完成类似项目10余个
   - 合同金额累计超过5000万元
   - 客户满意度达到95%以上

2. **典型项目案例**
   - 项目一：XX市智慧城市平台建设
   - 项目二：XX省政务信息系统升级
   - 项目三：XX企业数字化转型项目

3. **经验总结**
   - 深入理解业务需求
   - 采用成熟的技术方案
   - 注重项目管理和质量控制"""

    def _generate_company_intro_content(self) -> str:
        """生成公司简介内容"""
        return f"""我公司是一家专业的信息技术服务企业，具有以下优势：

1. **公司概况**
   - 成立于20XX年，注册资本XXXX万元
   - 员工总数XXX人，其中技术人员占比XX%
   - 年营业额超过XXXX万元

2. **资质荣誉**
   - 国家高新技术企业
   - 信息系统集成及服务资质
   - ISO9001质量管理体系认证
   - ISO27001信息安全管理体系认证

3. **技术实力**
   - 拥有XX项软件著作权
   - 拥有XX项发明专利
   - 核心技术人员具有10年以上行业经验

4. **服务网络**
   - 总部位于XX市
   - 在全国XX个城市设有分支机构
   - 建立了完善的售后服务网络"""

    def _generate_evaluation_content(self) -> str:
        """生成客户评价内容"""
        return f"""公司获得了客户的高度评价和认可：

1. **客户满意度**
   - 项目验收通过率100%
   - 客户满意度评分平均95分以上
   - 多次获得客户书面表扬

2. **客户反馈**
   - 技术方案先进可行
   - 项目管理规范有序
   - 服务质量优良高效

3. **典型客户评价**
   - "该公司技术实力强，项目管理规范，服务质量优良。"
   - "项目按时保质完成，满足了我们的业务需求。"
   - "售后服务响应及时，问题解决效率高。"""

    def _generate_technical_strength_content(self) -> str:
        """生成技术实力内容"""
        return f"""公司具有强大的技术实力，主要体现在以下方面：

1. **技术团队**
   - 拥有XX名高级工程师
   - XX名项目经理持有PMP认证
   - 核心技术人员平均从业经验超过10年

2. **技术能力**
   - 精通Java、Python、Go等主流开发语言
   - 熟悉微服务、容器化、云原生等先进技术
   - 具备大数据、人工智能、区块链等新技术应用能力

3. **技术成果**
   - 拥有XX项软件著作权
   - 拥有XX项发明专利
   - 参与XX项行业标准制定

4. **技术合作**
   - 与XX所高校建立产学研合作
   - 与XX家知名企业建立技术合作关系
   - 定期举办技术交流活动"""

    def _generate_service_content(self) -> str:
        """生成服务体系内容"""
        return f"""公司建立完善的售后服务体系，包括：

1. **服务组织**
   - 成立专门的售后服务团队
   - 配备专业的技术支持人员
   - 建立7×24小时服务响应机制

2. **服务内容**
   - 系统维护和故障处理
   - 技术咨询和支持
   - 系统升级和优化

3. **服务标准**
   - 响应时间：30分钟内响应
   - 解决时间：一般问题4小时内解决
   - 满意度：客户满意度≥95%"""

    def _generate_response_content(self) -> str:
        """生成响应时间内容"""
        return f"""公司承诺以下服务响应时间：

1. **故障响应时间**
   - 紧急故障：15分钟内响应，2小时内解决
   - 严重故障：30分钟内响应，4小时内解决
   - 一般故障：2小时内响应，8小时内解决

2. **技术支持响应**
   - 电话支持：7×24小时
   - 远程支持：工作时间2小时内响应
   - 现场支持：24小时内到达现场

3. **服务保障**
   - 建立服务级别协议（SLA）
   - 定期服务报告和满意度调查
   - 持续改进服务质量"""

    def _generate_training_content(self) -> str:
        """生成培训计划内容"""
        return f"""公司提供全面的培训服务，包括：

1. **培训对象**
   - 系统管理员：系统配置和管理
   - 业务用户：系统操作和使用
   - 技术人员：系统维护和开发

2. **培训内容**
   - 系统功能培训
   - 系统操作培训
   - 系统维护培训

3. **培训方式**
   - 集中培训：现场集中授课
   - 在线培训：远程视频培训
   - 操作手册：提供详细的用户手册

4. **培训保障**
   - 培训师资：具备丰富的培训经验
   - 培训材料：提供完整的培训资料
   - 培训考核：培训后进行考核验收"""

    def _generate_spare_parts_content(self) -> str:
        """生成备品备件内容"""
        return f"""公司提供完善的备品备件服务，确保系统稳定运行：

1. **备品备件库**
   - 建立专门的备品备件仓库
   - 储备常用备品备件
   - 定期盘点和补充

2. **备品备件类型**
   - 硬件备件：服务器、存储设备、网络设备等
   - 软件备件：操作系统、数据库、中间件等许可证
   - 耗材备件：打印耗材、存储介质等

3. **响应时间**
   - 常用备件：4小时内送达
   - 专用备件：24小时内送达
   - 紧急备件：2小时内响应

4. **服务保障**
   - 备件质量保证
   - 备件更换服务
   - 备件回收处理"""


def generate_content_from_outline(outline_file: str, scoring_file: str, bid_content: str = "") -> Dict[str, str]:
    """从大纲文件生成内容"""
    # 读取大纲
    with open(outline_file, 'r', encoding='utf-8') as f:
        outline = json.load(f)

    # 读取评分标准
    with open(scoring_file, 'r', encoding='utf-8') as f:
        scoring = json.load(f)

    # 生成内容
    generator = ContentGenerator(outline, scoring, bid_content)
    return generator.generate_all_chapters()


def main():
    import argparse

    parser = argparse.ArgumentParser(description='标书内容生成工具')
    parser.add_argument('outline_file', help='大纲JSON文件路径')
    parser.add_argument('scoring_file', help='评分标准JSON文件路径')
    parser.add_argument('--bid-content', '-b', help='招标文件内容（可选）')
    parser.add_argument('--output-dir', '-o', default='./chapters', help='输出目录')

    args = parser.parse_args()

    if not os.path.exists(args.outline_file):
        print(f"[ERROR] 大纲文件不存在: {args.outline_file}")
        sys.exit(1)

    if not os.path.exists(args.scoring_file):
        print(f"[ERROR] 评分标准文件不存在: {args.scoring_file}")
        sys.exit(1)

    print(f"[INFO] 开始生成标书内容...")

    # 读取招标文件内容（可选）
    bid_content = ""
    if args.bid_content and os.path.exists(args.bid_content):
        with open(args.bid_content, 'r', encoding='utf-8') as f:
            bid_content = f.read()

    # 生成内容
    chapters = generate_content_from_outline(args.outline_file, args.scoring_file, bid_content)

    # 保存章节文件
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    for filename, content in chapters.items():
        file_path = output_dir / f"{filename}.md"
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"[INFO] 已保存: {file_path}")

    print(f"\n[INFO] 内容生成完成，共 {len(chapters)} 章")
    print(f"[INFO] 输出目录: {output_dir}")


if __name__ == '__main__':
    main()
