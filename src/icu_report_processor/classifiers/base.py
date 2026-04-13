# -*- coding: utf-8 -*-
"""
分类器基类模块 - 定义报告分类的标准接口
"""

from abc import ABC, abstractmethod
from typing import Optional, Dict, Any, List, Tuple


class BaseClassifier(ABC):
    """
    报告分类器基类
    
    所有具体的分类器实现都应继承此类，并实现 classify 方法
    """
    
    def __init__(self, config: Dict[str, Any]):
        """
        初始化分类器
        
        Args:
            config: 分类器配置字典
        """
        self.config = config
    
    @abstractmethod
    def classify(self, text: str) -> Optional[str]:
        """
        对报告文本进行分类
        
        Args:
            text: PDF 提取的文本内容
            
        Returns:
            报告类型标识，无法分类返回 None
        """
        pass
    
    @abstractmethod
    def get_confidence(self, text: str, report_type: str) -> float:
        """
        获取分类置信度
        
        Args:
            text: PDF 提取的文本内容
            report_type: 报告类型标识
            
        Returns:
            置信度分数（0.0 - 1.0）
        """
        pass


class KeywordClassifier(BaseClassifier):
    """
    基于关键词的分类器
    
    通过匹配配置文件中的关键词来分类报告
    """
    
    def __init__(self, report_types_config: Dict[str, Dict[str, Any]]):
        """
        初始化关键词分类器
        
        Args:
            report_types_config: 报告类型配置字典
        """
        super().__init__(report_types_config)
        self.report_types = report_types_config
    
    def classify(self, text: str) -> Optional[str]:
        """
        基于关键词匹配进行分类
        
        策略：
        1. 首先检查关键词匹配
        2. 如果有多个匹配，选择匹配关键词最多的类型
        3. 如果关键词数量相同，选择有特征指标匹配的类型
        
        Args:
            text: PDF 提取的文本内容
            
        Returns:
            报告类型标识，无法分类返回 None
        """
        if not text:
            return None
        
        candidates = []
        
        for report_id, config in self.report_types.items():
            score = self._calculate_score(text, config)
            if score > 0:
                candidates.append((report_id, score))
        
        if not candidates:
            return None
        
        # 按分数排序，返回最高分
        candidates.sort(key=lambda x: x[1], reverse=True)
        return candidates[0][0]
    
    def _calculate_score(self, text: str, config: Dict[str, Any]) -> int:
        """
        计算文本与报告类型的匹配分数
        
        Args:
            text: PDF 文本
            config: 报告类型配置
            
        Returns:
            匹配分数
        """
        score = 0
        
        # 关键词匹配（每个匹配 +10 分）
        keywords = config.get('classification_keywords', [])
        for keyword in keywords:
            if keyword in text:
                score += 10
        
        # 特征指标匹配（每个匹配 +5 分，用于区分相似类型）
        indicators = config.get('indicator_fields', [])
        for indicator in indicators:
            if indicator in text:
                score += 5
        
        return score
    
    def get_confidence(self, text: str, report_type: str) -> float:
        """
        获取分类置信度
        
        基于匹配分数计算相对置信度
        
        Args:
            text: PDF 提取的文本内容
            report_type: 报告类型标识
            
        Returns:
            置信度分数（0.0 - 1.0）
        """
        if report_type not in self.report_types:
            return 0.0
        
        config = self.report_types[report_type]
        score = self._calculate_score(text, config)
        
        # 计算最大可能分数
        max_score = len(config.get('classification_keywords', [])) * 10 + \
                   len(config.get('indicator_fields', [])) * 5
        
        if max_score == 0:
            return 0.0
        
        # 归一化到 0-1
        confidence = min(score / max_score, 1.0)
        return confidence
    
    def get_all_matches(self, text: str) -> List[Tuple[str, int]]:
        """
        获取所有匹配的报告类型及其分数
        
        Args:
            text: PDF 提取的文本内容
            
        Returns:
            列表，每项为 (report_type, score)
        """
        matches = []
        for report_id, config in self.report_types.items():
            score = self._calculate_score(text, config)
            if score > 0:
                matches.append((report_id, score))
        
        matches.sort(key=lambda x: x[1], reverse=True)
        return matches


class CompositeClassifier(BaseClassifier):
    """
    组合分类器
    
    可以组合多个分类器，按优先级或投票机制进行分类
    """
    
    def __init__(self, classifiers: List[Tuple[BaseClassifier, float]]):
        """
        初始化组合分类器
        
        Args:
            classifiers: 分类器列表，每项为 (classifier, weight)
        """
        super().__init__({})
        self.classifiers = classifiers
    
    def classify(self, text: str) -> Optional[str]:
        """
        使用多个分类器进行投票分类
        
        Args:
            text: PDF 提取的文本内容
            
        Returns:
            得票最多的报告类型
        """
        votes: Dict[str, float] = {}
        
        for classifier, weight in self.classifiers:
            result = classifier.classify(text)
            if result:
                confidence = classifier.get_confidence(text, result)
                votes[result] = votes.get(result, 0) + confidence * weight
        
        if not votes:
            return None
        
        # 返回加权得分最高的类型
        return max(votes.items(), key=lambda x: x[1])[0]
    
    def get_confidence(self, text: str, report_type: str) -> float:
        """
        获取组合分类置信度
        
        Args:
            text: PDF 提取的文本内容
            report_type: 报告类型标识
            
        Returns:
            加权平均置信度
        """
        total_confidence = 0.0
        total_weight = 0.0
        
        for classifier, weight in self.classifiers:
            confidence = classifier.get_confidence(text, report_type)
            total_confidence += confidence * weight
            total_weight += weight
        
        if total_weight == 0:
            return 0.0
        
        return total_confidence / total_weight
