# -*- coding: utf-8 -*-
"""
解析器基类模块 - 定义报告解析的标准接口
"""

import re
from abc import ABC, abstractmethod
from typing import Dict, Any, Optional, List, Tuple
from datetime import datetime


class ParseResult:
    """
    解析结果数据类
    
    封装解析后的所有数据，便于统一处理和传递
    """
    
    def __init__(self):
        # 文件信息
        self.filename: str = ""
        self.report_type: str = ""
        
        # 时间信息
        self.collection_time: Optional[str] = None
        self.receive_time: Optional[str] = None
        self.report_time: Optional[str] = None
        self.main_time: Optional[str] = None
        
        # 患者信息
        self.patient_name: Optional[str] = None
        self.patient_id: Optional[str] = None
        self.gender: Optional[str] = None
        self.age: Optional[str] = None
        self.department: Optional[str] = None
        self.bed_number: Optional[str] = None
        self.sample_type: Optional[str] = None
        self.diagnosis: Optional[str] = None
        
        # 检测项目数据 {项目名称: 结果值}
        self.test_items: Dict[str, str] = {}
        
        # 参考区间 {项目名称: 参考区间}
        self.reference_ranges: Dict[str, str] = {}
    
    def to_dict(self) -> Dict[str, Any]:
        """
        转换为字典格式
        
        Returns:
            包含所有字段的字典
        """
        result = {
            '文件名': self.filename,
            '报告类型': self.report_type,
            '采集时间': self.collection_time,
            '接收时间': self.receive_time,
            '报告时间': self.report_time,
            '主时间': self.main_time,
            '姓名': self.patient_name,
            '病历号': self.patient_id,
            '性别': self.gender,
            '年龄': self.age,
            '科室': self.department,
            '床号': self.bed_number,
            '样本种类': self.sample_type,
            '临床诊断': self.diagnosis,
        }
        # 添加检测项目
        result.update(self.test_items)
        return result
    
    def get_reference_ranges(self) -> Dict[str, str]:
        """
        获取参考区间字典
        
        Returns:
            项目名称到参考区间的映射
        """
        return self.reference_ranges.copy()


class BaseParser(ABC):
    """
    报告解析器基类
    
    所有具体的解析器实现都应继承此类，并实现相关方法
    """
    
    def __init__(self, config: Dict[str, Any]):
        """
        初始化解析器
        
        Args:
            config: 解析器配置字典，包含正则表达式等
        """
        self.config = config
        self.patient_patterns = config.get('patient_info_patterns', {})
        self.time_patterns = config.get('time_patterns', {})
        self.table_config = config.get('table_extraction', {})
        self.excel_config = config.get('excel_config', {})
    
    @abstractmethod
    def parse(self, text: str, filename: str, report_type: str) -> ParseResult:
        """
        解析报告文本
        
        Args:
            text: PDF 提取的文本内容
            filename: 文件名
            report_type: 报告类型标识
            
        Returns:
            ParseResult 解析结果对象
        """
        pass
    
    def extract_patient_info(self, text: str) -> Dict[str, Optional[str]]:
        """
        提取患者基本信息
        
        Args:
            text: PDF 文本
            
        Returns:
            患者信息字典
        """
        info = {}
        
        for field, pattern in self.patient_patterns.items():
            match = re.search(pattern, text)
            if match:
                # 字段名映射
                field_map = {
                    'patient_name': 'patient_name',
                    'patient_id': 'patient_id',
                    'gender': 'gender',
                    'age': 'age',
                    'department': 'department',
                    'bed_number': 'bed_number',
                    'sample_type': 'sample_type',
                    'diagnosis': 'diagnosis',
                }
                key = field_map.get(field, field)
                info[key] = match.group(1).strip()
            else:
                info[field] = None
        
        return info
    
    def extract_time_info(self, text: str) -> Dict[str, Optional[str]]:
        """
        提取时间信息
        
        Args:
            text: PDF 文本
            
        Returns:
            时间信息字典，包含主时间
        """
        time_info = {}
        
        for field, pattern in self.time_patterns.items():
            match = re.search(pattern, text)
            if match:
                time_info[field] = match.group(1)
            else:
                time_info[field] = None
        
        # 确定主时间
        time_info['main_time'] = self._determine_main_time(time_info)
        
        return time_info
    
    def _determine_main_time(self, time_info: Dict[str, Optional[str]]) -> Optional[str]:
        """
        根据优先级确定主时间
        
        Args:
            time_info: 时间信息字典
            
        Returns:
            主时间字符串
        """
        priority = self.excel_config.get('time_priority', ['receive_time', 'collection_time', 'report_time'])
        
        for field in priority:
            if time_info.get(field):
                return time_info[field]
        
        return None
    
    def extract_table_data(self, text: str) -> Tuple[Dict[str, str], Dict[str, str]]:
        """
        提取表格数据（检测项目和参考区间）
        
        Args:
            text: PDF 文本
            
        Returns:
            (检测项目字典, 参考区间字典)
        """
        data = {}
        references = {}
        
        row_pattern = self.table_config.get('row_pattern', '')
        reference_patterns = self.table_config.get('reference_patterns', [])
        name_cleanup = self.table_config.get('name_cleanup', [])
        
        if not row_pattern:
            return data, references
        
        lines = text.split('\n')
        
        for line in lines:
            match = re.match(row_pattern, line)
            if match:
                item_name = match.group(1).strip()
                result = match.group(2).strip()
                trend = match.group(3).strip() if len(match.groups()) > 2 else ''
                reference = match.group(4).strip() if len(match.groups()) > 3 else ''
                
                # 清理项目名称
                for old, new in name_cleanup:
                    item_name = re.sub(old, new, item_name)
                item_name = item_name.strip()
                
                # 添加趋势标记到结果
                if trend:
                    result = f"{result} {trend}"
                
                # 过滤掉过长的项目名（避免匹配到非项目行）
                if item_name and result and len(item_name) < 30:
                    data[item_name] = result
                    
                    # 判断是否为参考区间
                    if self._is_reference_range(reference, reference_patterns):
                        references[item_name] = reference
        
        return data, references
    
    def _is_reference_range(self, value: str, patterns: List[str]) -> bool:
        """
        判断值是否为参考区间
        
        Args:
            value: 待判断的值
            patterns: 参考区间匹配模式列表
            
        Returns:
            是否为参考区间
        """
        if not value:
            return False
        
        for pattern in patterns:
            if re.match(pattern, value):
                return True
        
        return False


class DefaultParser(BaseParser):
    """
    默认解析器
    
    适用于当前 ICU 报告格式的标准解析器
    """
    
    def parse(self, text: str, filename: str, report_type: str) -> ParseResult:
        """
        解析报告文本
        
        Args:
            text: PDF 提取的文本内容
            filename: 文件名
            report_type: 报告类型标识
            
        Returns:
            ParseResult 解析结果对象
        """
        result = ParseResult()
        result.filename = filename
        result.report_type = report_type
        
        # 提取患者信息
        patient_info = self.extract_patient_info(text)
        result.patient_name = patient_info.get('patient_name')
        result.patient_id = patient_info.get('patient_id')
        result.gender = patient_info.get('gender')
        result.age = patient_info.get('age')
        result.department = patient_info.get('department')
        result.bed_number = patient_info.get('bed_number')
        result.sample_type = patient_info.get('sample_type')
        result.diagnosis = patient_info.get('diagnosis')
        
        # 提取时间信息
        time_info = self.extract_time_info(text)
        result.collection_time = time_info.get('collection_time')
        result.receive_time = time_info.get('receive_time')
        result.report_time = time_info.get('report_time')
        result.main_time = time_info.get('main_time')
        
        # 提取表格数据
        test_items, references = self.extract_table_data(text)
        result.test_items = test_items
        result.reference_ranges = references
        
        return result
