#!/usr/bin/env python3
import os
import sys
import re
import csv
from pathlib import Path
from typing import List, Tuple, Dict
from collections import Counter

# 尝试导入可选依赖
try:
    import docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    print("警告: python-docx 未安装，无法处理.docx文件")

try:
    import jieba
    import jieba.analyse
    HAS_JIEBA = True
except ImportError:
    HAS_JIEBA = False
    print("警告: jieba 未安装，将使用简单的关键词提取")

class DocumentTagger:
    def __init__(self):
        # 文档属性分类定义
        self.document_types = {
            '需求类文档': [
                '需求', '功能需求', '业务需求', '产品需求', 'PRD', '需求分析', 
                'requirement', 'feature', '用户需求', '功能规格', '需求说明'
            ],
            '技术类文档': [
                '技术', '开发', '编程', '代码', '实现', '架构', '设计', 'API',
                'development', 'coding', '技术方案', '系统设计', '接口文档',
                '数据库', '算法', '框架', '技术规范'
            ],
            '测试类文档': [
                '测试', '质量', 'QA', '测试用例', '测试计划', 'testing', 
                '测试报告', '缺陷', 'bug', '质量保证', '验收测试', '性能测试'
            ],
            '运维类文档': [
                '运维', '部署', '监控', '服务器', '系统运维', 'devops', 
                'deploy', '运维手册', '故障处理', '备份', '安全', '配置管理'
            ],
            '知识类文档': [
                '知识', '教程', '培训', '说明', '指南', '手册', '介绍',
                '学习', '分享', '总结', '经验', '最佳实践', '规范', '标准'
            ],
            '管理类文档': [
                '管理', '项目管理', '计划', '流程', '规范', 'management', 
                'process', '会议', '决策', '报告', '总结', '制度', '政策'
            ]
        }
        
        # 项目关键词映射（用于识别所属项目）
        self.project_keywords = {
            # 业务项目
            '超品中心项目': ['超品', '超级品牌', '品牌中心', '品牌运营'],
            '直播业务项目': ['直播', '主播', '直播间', '直播运营', '直播平台'],
            '电商系统项目': ['订单', '商品', '购物车', '支付', '物流', 'OMS'],
            '财务管理项目': ['财务', '会计', '成本', '预算', '收入', '支出', '结算'],
            '营销推广项目': ['营销', '推广', '广告', '活动', '用户增长', '转化'],
            '内容运营项目': ['内容', '创作', '视频', '图片', '文章', '创意'],
            '用户运营项目': ['用户运营', '私域', '客服', '用户管理', '客户服务'],
            # 技术项目  
            '技术平台项目': ['平台', '系统', '架构', '技术', '开发', '产研'],
            '数据分析项目': ['数据', '分析', '统计', '指标', '报表', '数据库'],
            '移动端项目': ['移动', 'APP', '小程序', '移动端', 'iOS', 'Android'],
            '运维保障项目': ['运维', '部署', '监控', '服务器', '运维保障'],
            # 管理项目
            '人力资源项目': ['人力', 'HR', '招聘', '培训', '绩效', '薪酬', '员工'],
            '质量管控项目': ['质量', '品控', '品质管理', '质量保证', '测试'],
            '流程优化项目': ['流程', '优化', '规范', '标准化', '制度', '管理']
        }

    def extract_text_from_file(self, file_path: str) -> Tuple[str, str]:
        """从文件中提取文本和标题"""
        file_path = Path(file_path)
        title = file_path.stem
        text = ""
        
        if file_path.suffix.lower() == '.docx':
            if not HAS_DOCX:
                print("错误: 需要安装 python-docx 来处理 .docx 文件")
                return title, ""
            try:
                doc = docx.Document(file_path)
                text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
            except Exception as e:
                print(f"读取Word文档失败: {e}")
                return title, ""
        elif file_path.suffix.lower() == '.txt':
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    text = f.read()
            except Exception as e:
                print(f"读取文本文件失败: {e}")
                return title, ""
        else:
            print(f"不支持的文件格式: {file_path.suffix}")
            return title, ""
        
        return title, text

    def identify_project(self, text: str, title: str = "") -> str:
        """智能识别所属项目，优先从文档标题和内容中直接提取项目名称"""
        text_lower = text.lower()
        title_lower = title.lower()
        combined_text = (title + " " + text).lower()
        
        # 强化标题优先级 - 明确项目名称优先于通用词汇
        # 第一优先级：明确的项目名称
        if any(word in title_lower for word in ['超品中心', '超品']):
            return "超品中心项目"
        elif any(word in title_lower for word in ['财务', '结算', '业绩', '毛利']):
            return "财务管理项目"
        elif any(word in title_lower for word in ['新人', '入职', '培训']):
            return "人力资源项目"  
        elif any(word in title_lower for word in ['部门', '组织', '团队', '职责']):
            return "组织管理项目"
        
        # 第二优先级：项目相关业务词汇
        elif any(word in title_lower for word in ['商家', '星选', '遥望星选']):
            return "超品中心项目"
        
        # 第三优先级：通用词汇（优先级最低）
        elif any(word in title_lower for word in ['文档', '分类', '收集', '知识']):
            return "知识管理项目"
        
        # 第一步：直接从标题和内容中提取明确的项目名称
        extracted_project = self._extract_explicit_project_name(title, text)
        if extracted_project:
            return extracted_project
        
        # 第二步：使用关键词匹配（降低权重，更灵活）
        project_scores = {}
        for project, keywords in self.project_keywords.items():
            score = 0
            for keyword in keywords:
                keyword_lower = keyword.lower()
                # 文档内容匹配
                content_count = combined_text.count(keyword_lower)
                score += content_count
                
                # 标题匹配给予更高权重
                if keyword_lower in title_lower:
                    score += 3  # 降低权重，避免过度依赖预设分类
                    
            if score > 0:
                project_scores[project] = score
        
        # 第三步：如果有预设匹配但分数不高，优先使用智能推断
        if project_scores:
            max_score = max(project_scores.values())
            if max_score >= 3:  # 只有较高匹配度才使用预设分类
                best_project = max(project_scores, key=project_scores.get)
                return best_project
        
        # 第四步：智能推断项目名称
        return self._infer_project_from_content(combined_text, title)
    
    def _extract_explicit_project_name(self, title: str, text: str) -> str:
        """从标题和内容中直接提取明确的项目名称"""
        import re
        
        # 合并标题和文档开头部分用于项目名称提取
        search_text = title + " " + text[:500]  # 只搜索前500字符，提高效率
        
        # 1. 从标题中提取项目名称模式
        title_patterns = [
            r'(\w*中心)\s*[vV]?\d*\.?\d*',  # 超品中心V1.0
            r'(\w+系统)\s*[vV]?\d*\.?\d*',   # 财务系统V1.0
            r'(\w+平台)\s*[vV]?\d*\.?\d*',   # 营销平台V1.0
            r'(\w+项目)\s*[vV]?\d*\.?\d*',   # 电商项目V1.0
            r'(\w+业务)\s*[vV]?\d*\.?\d*',   # 直播业务V1.0
        ]
        
        for pattern in title_patterns:
            match = re.search(pattern, title)
            if match:
                project_core = match.group(1)
                # 将识别到的项目核心词转换为标准项目名
                return self._standardize_project_name(project_core)
        
        # 2. 从内容中提取项目关键信息
        content_patterns = [
            r'项目[:：]?\s*([^\s\n,，。]{2,10})',
            r'所属项目[:：]?\s*([^\s\n,，。]{2,10})',
            r'([^\s]{2,8})\s*PRD',  # PRD前的项目名
            r'([^\s\n]{2,10})\s*需求文档',
        ]
        
        for pattern in content_patterns:
            matches = re.findall(pattern, search_text)
            if matches:
                # 选择最可能的项目名，过滤掉无效匹配
                for match in matches:
                    match = match.strip()
                    if (len(match) >= 2 and not match.isdigit() and 
                        not match.startswith('：') and 
                        '解决' not in match):  # 过滤掉"：解决现阶段财务"这样的误匹配
                        standardized = self._standardize_project_name(match)
                        if standardized != f"{match}项目":  # 如果有明确的标准化映射才返回
                            return standardized
        
        return ""
    
    def _standardize_project_name(self, project_core: str) -> str:
        """将提取的项目核心词标准化为项目名称"""
        # 清理版本号和特殊字符
        import re
        project_core = re.sub(r'[vV]?\d+\.?\d*', '', project_core).strip()
        
        # 标准化映射
        standardization_map = {
            '超品中心': '超品中心项目',
            '超品': '超品中心项目', 
            '财务': '财务管理项目',
            '财务系统': '财务管理项目',
            '财务专项': '财务管理项目',
            '新人': '人力资源项目',
            '入职': '人力资源项目',
            '部门': '组织管理项目',
            '组织': '组织管理项目',
            '文档分类': '知识管理项目',
            '文档': '知识管理项目',
            '知识': '知识管理项目',
            '商家端': '超品中心项目',
            '遥望星选': '超品中心项目',
            '背景': '技术平台项目',  # 临时解决背景误识别问题
        }
        
        # 查找匹配的标准化名称
        for key, standard_name in standardization_map.items():
            if key in project_core:
                return standard_name
        
        # 如果没有直接匹配，根据关键词构造项目名
        if any(word in project_core for word in ['中心', '平台', '系统']):
            return f"{project_core}项目"
        elif '管理' in project_core:
            return f"{project_core}项目"
        else:
            return f"{project_core}项目"
    
    def _infer_project_from_content(self, text: str, title: str = "") -> str:
        """基于内容和标题智能推断项目类型"""
        # 提取高频词汇用于分析
        if HAS_JIEBA:
            keywords = jieba.analyse.extract_tags(text, topK=15, withWeight=False)
        else:
            # 简单分词
            words = []
            for word in text.replace('，', ' ').replace('。', ' ').split():
                word = word.strip(' \t\n.,!?;:()[]{}"\'-')
                if 2 <= len(word) <= 8:
                    words.append(word)
            from collections import Counter
            word_count = Counter(words)
            keywords = [word for word, count in word_count.most_common(15)]
        
        # 优先检查标题中的项目线索
        title_lower = title.lower()
        
        # 强化标题优先级 - 明确项目名称优先于通用词汇
        # 第一优先级：明确的项目名称
        if any(word in title_lower for word in ['超品中心', '超品']):
            return "超品中心项目"
        elif any(word in title_lower for word in ['财务', '结算', '业绩', '毛利']):
            return "财务管理项目"
        elif any(word in title_lower for word in ['新人', '入职', '培训', 'onboard']):
            return "人力资源项目"  
        elif any(word in title_lower for word in ['部门', '组织', '团队', '职责']):
            return "组织管理项目"
        # 第二优先级：项目相关业务词汇
        elif any(word in title_lower for word in ['商家', '星选', '遥望星选']):
            return "超品中心项目"
        # 第三优先级：通用词汇（优先级最低）
        elif any(word in title_lower for word in ['文档', '分类', '收集', '知识']):
            return "知识管理项目"
        
        # 基于内容关键词进行更细致的分类
        # 定义更精确的领域关键词
        financial_words = ['财务', '结算', '业绩', '毛利', '成本', '收入', '支出', '预算', '会计']
        hr_words = ['人力', '招聘', '培训', '员工', '入职', '薪酬', '绩效', '考核']
        product_words = ['产品', '需求', 'PRD', '功能', '用户', '体验', '设计']
        tech_words = ['技术', '开发', '系统', '平台', '架构', '数据库', '接口', 'API']
        operation_words = ['运营', '营销', '推广', '活动', '转化', '渠道', '客户']
        management_words = ['管理', '流程', '制度', '规范', '优化', '团队', '组织']
        knowledge_words = ['文档', '知识', '分类', '收集', '整理', '归档', '指南']
        
        # 计算各领域词汇的匹配度
        domain_scores = {
            '财务管理项目': sum(1 for kw in keywords if any(fw in kw for fw in financial_words)),
            '人力资源项目': sum(1 for kw in keywords if any(hw in kw for hw in hr_words)),
            '产品研发项目': sum(1 for kw in keywords if any(pw in kw for pw in product_words)),
            '技术平台项目': sum(1 for kw in keywords if any(tw in kw for tw in tech_words)),
            '运营推广项目': sum(1 for kw in keywords if any(ow in kw for ow in operation_words)),
            '组织管理项目': sum(1 for kw in keywords if any(mw in kw for mw in management_words)),
            '知识管理项目': sum(1 for kw in keywords if any(kw_word in kw for kw_word in knowledge_words)),
        }
        
        # 返回得分最高的项目类型
        if domain_scores:
            max_score = max(domain_scores.values())
            if max_score > 0:
                best_domain = max(domain_scores, key=domain_scores.get)
                return best_domain
        
        # 如果没有明确匹配，使用通用分类逻辑
        business_count = sum(1 for kw in keywords if any(bw in kw for bw in ['用户', '产品', '业务', '运营', '营销', '商业']))
        tech_count = sum(1 for kw in keywords if any(tw in kw for tw in ['系统', '平台', '技术', '开发', '数据', '功能']))
        mgmt_count = sum(1 for kw in keywords if any(mw in kw for mw in ['管理', '流程', '制度', '培训', '团队']))
        
        if tech_count >= business_count and tech_count >= mgmt_count:
            return "技术平台项目"
        elif business_count >= mgmt_count:
            return "业务运营项目"
        else:
            return "组织管理项目"

    def classify_document_type(self, text: str, title: str = "") -> str:
        """分类文档属性：需求类、技术类、测试类、运维类、知识类、管理类"""
        text_lower = text.lower()
        title_lower = title.lower()
        combined_text = (title + " " + text).lower()
        
        type_scores = {}
        
        # 计算每种文档类型的匹配分数
        for doc_type, keywords in self.document_types.items():
            score = 0
            for keyword in keywords:
                keyword_lower = keyword.lower()
                # 文档内容匹配
                content_count = combined_text.count(keyword_lower)
                score += content_count
                
                # 标题匹配给予更高权重
                if keyword_lower in title_lower:
                    score += 5
                    
            if score > 0:
                type_scores[doc_type] = score
        
        # 返回得分最高的文档类型
        if type_scores:
            best_type = max(type_scores, key=type_scores.get)
            return best_type
        else:
            return "知识类文档"  # 默认类型

    def _get_keywords_count(self, text_length: int) -> int:
        """根据文档长度动态确定关键词数量，不设固定上限"""
        if text_length < 200:
            return 2
        elif text_length < 500:
            return 3
        elif text_length < 1000:
            return 5
        elif text_length < 2000:
            return 8
        elif text_length < 3000:
            return 10
        elif text_length < 5000:
            return 12
        elif text_length < 8000:
            return 15
        elif text_length < 12000:
            return 18
        elif text_length < 20000:
            return 22
        else:
            # 对于超长文档，按每1000字符约1个关键词的比例
            return min(int(text_length / 1000) + 3, 30)  # 最多30个，避免过多

    def extract_keywords(self, text: str) -> List[str]:
        """只从文档原文中提取确实存在的关键词"""
        if not text.strip():
            return []
        
        # 根据文档长度确定关键词数量
        num_keywords = self._get_keywords_count(len(text))
        
        # 直接从原文提取关键词，不依赖jieba可能产生的虚假词汇
        return self._extract_verified_keywords_from_text(text, num_keywords)
    
    def _extract_verified_keywords_from_text(self, text: str, num_keywords: int) -> List[str]:
        """从原文中提取并验证关键词，完全避免jieba产生虚假词汇"""
        import re
        import string
        from collections import Counter
        
        # 只使用直接文本分割，绝对不使用jieba，避免虚假词汇
        # 方法1：按中文标点符号和空格分割
        text_clean = re.sub(r'[，。！？；：""''（）【】\[\]\{\}<>《》、\s]+', ' ', text)
        words1 = text_clean.split()
        
        # 方法2：按英文标点分割提取更多可能的词汇
        text_clean2 = re.sub(r'[.,!?;:\"\'\(\)\[\]\{\}<>\s]+', ' ', text)
        words2 = text_clean2.split()
        
        # 方法3：提取连续的中文字符组合
        chinese_words = re.findall(r'[\u4e00-\u9fff]{2,8}', text)
        
        # 合并所有候选词汇，但只使用直接分割的结果
        candidate_words = words1 + words2 + chinese_words
        
        # 严格过滤和验证
        word_count = Counter()
        
        for word in candidate_words:
            word = word.strip(string.punctuation + ' \t\n')
            
            # 基本过滤条件
            if (2 <= len(word) <= 8 and 
                not word.isdigit() and 
                word.strip() and
                not word.isspace() and
                not word.lower() in ['html', 'http', 'https', 'www']):  # 排除网页相关词汇
                
                # 绝对严格验证：该词汇必须完整存在于原文中
                if self._absolute_strict_verify_word_in_text(word, text):
                    word_count[word] += 1
        
        # 按词频排序选择关键词
        # 排除过于常见的词汇
        common_words = {'的', '了', '在', '和', '与', '为', '是', '有', '及', '等', '可以', '进行', '通过', '或者', '如果', '但是', '因为', '所以', '这个', '那个', '我们', '他们', '她们', '它们', '之后', '之前', '什么', '怎么', '哪里', '什么时候', '为什么', '一个', '一种', '一些', '所有', '每个', '所以', '因此', '然后', '现在', '已经', '仍然', '只是', '也是', '还是', '或是', '就是', '不是', '没有', '这些', '那些'}
        
        # 选择最有意义的关键词
        result_keywords = []
        for word, count in word_count.most_common():
            if (word not in common_words and 
                count >= 1 and 
                len(result_keywords) < num_keywords):
                result_keywords.append(word)
        
        return result_keywords
    
    def _absolute_strict_verify_word_in_text(self, word: str, text: str) -> bool:
        """绝对严格验证词汇确实在原文中存在"""
        # 必须完全匹配存在于原文中，不能有任何差异
        return word in text
    
    def _strict_verify_word_in_text(self, word: str, text: str) -> bool:
        """严格验证词汇确实在原文中存在"""
        # 必须完全匹配存在于原文中
        return word in text
    
    def _verify_keyword_in_text(self, keyword: str, text: str) -> bool:
        """严格验证关键词确实在文档中存在"""
        import re
        
        # 对于中文关键词，直接检查是否在文本中
        if re.search(r'[\u4e00-\u9fff]', keyword):
            return keyword in text
        
        # 对于英文关键词，检查完整词匹配
        pattern = r'\b' + re.escape(keyword) + r'\b'
        if re.search(pattern, text, re.IGNORECASE):
            return True
            
        # 额外检查：是否作为中英混合词的一部分存在
        return keyword.lower() in text.lower()
    
    def _extract_simple_keywords_from_text(self, text: str, num_keywords: int) -> List[str]:
        """从文档文本中直接提取简单关键词"""
        # 分词处理
        import re
        import string
        
        # 移除标点符号并分词
        words = []
        # 按中文标点和空格分割
        text_clean = re.sub(r'[，。！？；：""''（）【】\s]+', ' ', text)
        word_candidates = text_clean.split()
        
        # 统计词频
        word_count = {}
        for word in word_candidates:
            word = word.strip(string.punctuation + ' \t\n')
            if (2 <= len(word) <= 8 and 
                not word.isdigit() and 
                word.strip()):
                word_count[word] = word_count.get(word, 0) + 1
        
        # 按频率排序
        sorted_words = sorted(word_count.items(), key=lambda x: x[1], reverse=True)
        
        # 取前N个高频词
        result = []
        for word, count in sorted_words:
            if count >= 1 and len(result) < num_keywords:  # 至少出现1次
                result.append(word)
        
        return result
    
    def _simple_keyword_extraction(self, text: str, num_keywords: int = 5) -> List[str]:
        """简单的关键词提取方法（当jieba不可用时）"""
        # 扩大常见关键词库
        common_keywords = [
            # 业务类
            '用户', '产品', '功能', '系统', '平台', '服务', '管理', '开发', '设计', '测试',
            '优化', '体验', '需求', '方案', '项目', '业务', '数据', '分析', '运营', '营销',
            # 技术类
            '技术', '接口', '数据库', '前端', '后端', '移动端', '网站', '应用', '软件', '架构',
            # 流程类
            '流程', '规范', '标准', '制度', '政策', '培训', '考核', '绩效', '质量', '安全',
            # 财务类
            '财务', '成本', '预算', '收入', '支出', '利润', '投资', '风险', '合规', '审计',
            # 运营类
            '推广', '活动', '渠道', '客户', '市场', '品牌', '内容', '社群', '转化', '留存'
        ]
        
        # 统计关键词频率
        keyword_count = {}
        text_lower = text.lower()
        
        for keyword in common_keywords:
            count = text_lower.count(keyword)
            if count > 0:
                keyword_count[keyword] = count
        
        # 按频率排序
        sorted_keywords = sorted(keyword_count.items(), key=lambda x: x[1], reverse=True)
        result = [kw[0] for kw in sorted_keywords[:num_keywords]]
        
        # 如果关键词数量不足，使用简单的字频统计补充
        if len(result) < num_keywords:
            # 简单分词（按标点和空格）
            import string
            words = []
            for word in text.replace('，', ' ').replace('。', ' ').replace('、', ' ').split():
                word = word.strip(string.punctuation + ' \t\n')
                if 2 <= len(word) <= 6 and not word.isdigit():
                    words.append(word)
            
            # 统计词频
            word_count = Counter(words)
            additional_words = [word for word, count in word_count.most_common(num_keywords*2) 
                             if word not in result]
            
            result.extend(additional_words[:num_keywords - len(result)])
        
        return result[:num_keywords]
    
    def generate_content_summary(self, text: str, title: str = "") -> str:
        """生成详细的文档内容概述，不限制字数"""
        if not text.strip():
            return "文档内容为空"
        
        # 按段落分割文本
        paragraphs = [p.strip() for p in text.split('\n') if p.strip()]
        
        # 提取和组织内容的不同部分
        summary_parts = []
        
        if HAS_JIEBA:
            # 提取文档结构化信息
            background_info = []
            objective_info = []
            content_info = []
            solution_info = []
            result_info = []
            
            # 关键词匹配模式
            background_patterns = ['背景', '现状', '问题', '挑战', '原因', '情况']
            objective_patterns = ['目标', '目的', '期望', '预期', '计划', '要求']
            content_patterns = ['内容', '功能', '特点', '特性', '包括', '包含', '具体']
            solution_patterns = ['方案', '解决', '实现', '实施', '执行', '操作', '步骤', '流程']
            result_patterns = ['结果', '效果', '收益', '价值', '成果', '总结', '结论']
            
            # 分析每个段落
            for para in paragraphs:
                para_lower = para.lower()
                
                # 检查背景信息
                if any(pattern in para_lower for pattern in background_patterns):
                    if len(para) > 15:
                        background_info.append(para)
                
                # 检查目标信息
                elif any(pattern in para_lower for pattern in objective_patterns):
                    if len(para) > 15:
                        objective_info.append(para)
                
                # 检查解决方案
                elif any(pattern in para_lower for pattern in solution_patterns):
                    if len(para) > 15:
                        solution_info.append(para)
                
                # 检查结果效果
                elif any(pattern in para_lower for pattern in result_patterns):
                    if len(para) > 15:
                        result_info.append(para)
                
                # 检查内容描述
                elif any(pattern in para_lower for pattern in content_patterns):
                    if len(para) > 15:
                        content_info.append(para)
            
            # 如果没有结构化内容，提取关键段落
            if not any([background_info, objective_info, content_info, solution_info, result_info]):
                # 提取包含关键词的重要段落
                important_keywords = [
                    '项目', '系统', '平台', '功能', '需求', '设计', '开发', '测试', '运维',
                    '管理', '流程', '规范', '优化', '升级', '改进', '实现', '支持', '提供',
                    '用户', '客户', '业务', '服务', '产品', '方案', '计划', '目标', '效果'
                ]
                
                scored_paragraphs = []
                for para in paragraphs[:15]:  # 处理前15个段落
                    if len(para) > 30:  # 过滤太短的段落
                        score = 0
                        para_lower = para.lower()
                        for keyword in important_keywords:
                            score += para_lower.count(keyword)
                        
                        if score > 0:
                            scored_paragraphs.append((para, score))
                
                # 按分数排序，取重要段落
                scored_paragraphs.sort(key=lambda x: x[1], reverse=True)
                content_info = [para[0] for para in scored_paragraphs[:8]]  # 取前8个重要段落
            
            # 组织概述内容
            if background_info:
                summary_parts.append(f"【背景情况】{' '.join(background_info[:2])}")
            
            if objective_info:
                summary_parts.append(f"【项目目标】{' '.join(objective_info[:2])}")
            
            if content_info:
                summary_parts.append(f"【主要内容】{' '.join(content_info[:4])}")
            
            if solution_info:
                summary_parts.append(f"【实施方案】{' '.join(solution_info[:3])}")
            
            if result_info:
                summary_parts.append(f"【预期效果】{' '.join(result_info[:2])}")
            
        else:
            # 简单方法：按段落提取
            important_paragraphs = []
            for para in paragraphs[:10]:
                if len(para) > 30:
                    important_paragraphs.append(para)
            
            if important_paragraphs:
                summary_parts.append(f"【文档内容】{' '.join(important_paragraphs[:5])}")
        
        # 如果没有提取到结构化内容，使用全文概括
        if not summary_parts:
            # 使用文档开头部分
            content_preview = []
            for para in paragraphs[:6]:
                if len(para) > 20:
                    content_preview.append(para)
            
            if content_preview:
                summary_parts.append(f"【文档概述】本文档主要阐述了{title}的相关内容，包括：{' '.join(content_preview)}")
            else:
                summary_parts.append(f"【文档概述】本文档描述了{title}相关内容：{text[:300]}...")
        
        # 生成最终概述
        final_summary = " ".join(summary_parts)
        
        return final_summary
    

    def format_output(self, title: str, project: str, doc_type: str, keywords: List[str], summary: str) -> Dict[str, str]:
        """格式化输出结果为新的CSV格式"""
        return {
            '文档标题': title,
            '所属项目': project,
            '文档关键词': f"{doc_type}; " + ('; '.join(keywords) if keywords else ''),
            '文档内容概述': summary
        }

    def process_document(self, file_path: str) -> Dict[str, str]:
        """处理单个文档"""
        print(f"正在处理文档: {file_path}")
        
        # 提取文本和标题
        title, text = self.extract_text_from_file(file_path)
        
        print(f"文档标题: {title}")
        print(f"文档内容长度: {len(text)} 字符")
        
        # 处理空文档
        if not text.strip():
            print("⚠️ 文档内容为空")
            return self.format_output(title, '未知项目', '知识类文档', [], '文档内容为空')
        
        # 新的分类体系
        project = self.identify_project(text, title)
        doc_type = self.classify_document_type(text, title)
        keywords = self.extract_keywords(text)
        summary = self.generate_content_summary(text, title)
        
        print(f"所属项目: {project}")
        print(f"文档属性: {doc_type}")
        print(f"关键词数量: {len(keywords)}")
        print(f"关键词: {keywords}")
        print(f"内容概述长度: {len(summary)} 字符")
        
        # 格式化输出
        result = self.format_output(title, project, doc_type, keywords, summary)
        return result
    
    def process_directory(self, directory_path: str) -> List[Dict[str, str]]:
        """批量处理目录中的文档"""
        directory = Path(directory_path)
        
        if not directory.exists():
            print(f"错误: 目录不存在 - {directory_path}")
            return []
        
        if not directory.is_dir():
            print(f"错误: 路径不是目录 - {directory_path}")
            return []
        
        # 支持的文件格式
        supported_extensions = ['.txt', '.docx']
        
        # 获取所有支持的文件
        files_to_process = []
        for ext in supported_extensions:
            files_to_process.extend(directory.glob(f'*{ext}'))
        
        if not files_to_process:
            print(f"在目录 {directory_path} 中未找到支持的文件格式: {supported_extensions}")
            return []
        
        # 按文件名排序
        files_to_process.sort(key=lambda x: x.name)
        
        print(f"找到 {len(files_to_process)} 个文件待处理")
        print("支持的格式:", supported_extensions)
        print()
        
        results = []
        for i, file_path in enumerate(files_to_process, 1):
            print(f"\n{'='*60}")
            print(f"处理第 {i}/{len(files_to_process)} 个文件")
            print(f"{'='*60}")
            
            try:
                result = self.process_document(str(file_path))
                results.append(result)
                
                print(f"\n✅ 文件 {file_path.name} 处理完成")
                
            except Exception as e:
                error_result = {
                    '文档标题': file_path.stem,
                    '所属项目': '处理失败',
                    '文档关键词': '错误',
                    '文档内容概述': str(e)
                }
                results.append(error_result)
                print(f"\n❌ 文件 {file_path.name} 处理失败: {e}")
        
        return results

def save_results_to_csv(results: List[Dict[str, str]], output_file_path: str = "document_tags.csv"):
    """保存结果到CSV文件，避免重复"""
    output_file = Path(output_file_path)
    
    # 新的CSV列名
    fieldnames = ['文档标题', '所属项目', '文档关键词', '文档内容概述']
    
    # 读取现有内容，避免重复
    existing_results = set()
    if output_file.exists():
        try:
            with open(output_file, "r", encoding="utf-8", newline='') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # 创建唯一标识符
                    identifier = f"{row.get('文档标题', '')}-{row.get('所属项目', '')}-{row.get('文档关键词', '')}"
                    existing_results.add(identifier)
        except Exception as e:
            print(f"读取现有CSV文件失败: {e}")
    
    # 过滤出新结果
    new_results = []
    for result in results:
        identifier = f"{result['文档标题']}-{result['所属项目']}-{result['文档关键词']}"
        if identifier not in existing_results:
            new_results.append(result)
    
    # 检查是否需要写入表头
    write_header = False
    if not output_file.exists():
        write_header = True
    else:
        # 检查文件是否有表头
        try:
            with open(output_file, "r", encoding="utf-8", newline='') as f:
                first_line = f.readline().strip()
                if not first_line or first_line != ','.join(fieldnames):
                    write_header = True
        except:
            write_header = True
    
    # 写入结果
    if new_results or write_header:
        if write_header and output_file.exists():
            # 需要重写文件以添加表头
            # 先读取所有现有数据
            existing_data = []
            try:
                with open(output_file, "r", encoding="utf-8", newline='') as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if row:  # 跳过空行
                            existing_data.append({
                                '文档标题': row[0] if len(row) > 0 else '',
                                '所属项目': row[1] if len(row) > 1 else '',
                                '文档关键词': row[2] if len(row) > 2 else '',
                                '文档内容概述': row[3] if len(row) > 3 else ''
                            })
            except:
                existing_data = []
            
            # 重写文件，包含表头和所有数据
            with open(output_file, "w", encoding="utf-8", newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for row in existing_data:
                    writer.writerow(row)
                for result in new_results:
                    writer.writerow(result)
        else:
            # 正常追加模式
            file_mode = 'a' if output_file.exists() and not write_header else 'w'
            with open(output_file, file_mode, encoding="utf-8", newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                
                # 如果是新文件，写入表头
                if write_header:
                    writer.writeheader()
                
                # 写入数据
                for result in new_results:
                    writer.writerow(result)
        
        if new_results:
            print(f"\n✅ {len(new_results)} 个新结果已保存到: {output_file}")
        else:
            print(f"\n✅ CSV文件已创建: {output_file}")
    else:
        print(f"\n⚠️ 所有结果都已存在于: {output_file}，跳过重复保存")
    
    print(f"\n📁 总共处理: {len(results)} 个文件")
    print(f"📁 新增结果: {len(new_results)} 个")
    print(f"📁 重复跳过: {len(results) - len(new_results)} 个")

def main():
    if len(sys.argv) != 2:
        print("使用方法: python document_tagger.py <文档路径或目录路径>")
        print("支持格式: .docx, .txt")
        print("输出格式: CSV表格 (document_tags.csv)")
        print("示例:")
        print("  单个文件: python document_tagger.py /path/to/document.docx")
        print("  批量处理: python document_tagger.py /path/to/documents/")
        print("  批量处理当前目录下的documents文件夹: python document_tagger.py documents")
        sys.exit(1)
    
    input_path = sys.argv[1]
    
    if not os.path.exists(input_path):
        print(f"错误: 路径不存在 - {input_path}")
        sys.exit(1)
    
    # 创建标签器实例
    tagger = DocumentTagger()
    
    try:
        if os.path.isdir(input_path):
            # 处理目录
            print(f"🔄 开始批量处理目录: {input_path}")
            print("="*60)
            
            results = tagger.process_directory(input_path)
            
            if results:
                print(f"\n{'='*60}")
                print("批量处理完成 - CSV表格格式结果汇总:")
                print("="*60)
                print(f"{'序号':<4} {'文档标题':<20} {'所属项目':<15} {'文档类型':<12} {'关键词数量':<10}")
                print("-" * 80)
                for i, result in enumerate(results, 1):
                    keywords = result['文档关键词'].split('; ')[1:] if result['文档关键词'] else []
                    keyword_count = len(keywords)
                    doc_type = result['文档关键词'].split('; ')[0] if result['文档关键词'] else ''
                    print(f"{i:<4} {result['文档标题']:<20} {result['所属项目']:<15} {doc_type:<12} {keyword_count:<10}")
                print("="*60)
                
                # 保存所有结果到CSV
                save_results_to_csv(results)
            else:
                print("没有文件被处理")
                
        else:
            # 处理单个文件
            result = tagger.process_document(input_path)
            
            print("\n" + "="*60)
            print("文档标签分类结果 (CSV格式):")
            print("="*60)
            print(f"文档标题: {result['文档标题']}")
            print(f"所属项目: {result['所属项目']}")
            print(f"文档关键词: {result['文档关键词']}")
            print(f"内容概述: {result['文档内容概述'][:100]}..." if len(result['文档内容概述']) > 100 else f"内容概述: {result['文档内容概述']}")
            print("="*60)
            
            # 保存结果到CSV
            save_results_to_csv([result])
        
    except Exception as e:
        print(f"处理时出错: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()