# -*- coding: utf-8 -*-
"""
配置模块 - 负责加载和管理医院配置文件
"""

import os
import yaml
from pathlib import Path
from typing import Dict, Any, Optional, List


class Config:
    """配置管理类"""
    
    def __init__(self, config_path: Optional[str] = None):
        """
        初始化配置
        
        Args:
            config_path: 配置文件路径，默认为模块目录下的 hospital_config.yaml
        """
        if config_path is None:
            config_path = Path(__file__).parent / "hospital_config.yaml"
        
        self.config_path = Path(config_path)
        self._config_data: Dict[str, Any] = {}
        self._load_config()
    
    def _load_config(self) -> None:
        """加载 YAML 配置文件"""
        if not self.config_path.exists():
            raise FileNotFoundError(f"配置文件不存在: {self.config_path}")
        
        with open(self.config_path, 'r', encoding='utf-8') as f:
            self._config_data = yaml.safe_load(f)
    
    def get_report_types(self) -> Dict[str, Dict[str, Any]]:
        """
        获取所有报告类型配置
        
        Returns:
            报告类型配置字典，key 为报告类型标识
        """
        return self._config_data.get('report_types', {})
    
    def get_report_type(self, report_type_id: str) -> Optional[Dict[str, Any]]:
        """
        获取特定报告类型配置
        
        Args:
            report_type_id: 报告类型标识
            
        Returns:
            报告类型配置字典，不存在返回 None
        """
        return self._config_data.get('report_types', {}).get(report_type_id)
    
    def get_hospital_config(self, hospital_id: str = "default") -> Dict[str, Any]:
        """
        获取医院特定配置
        
        Args:
            hospital_id: 医院标识，默认为 default
            
        Returns:
            医院配置字典
        """
        return self._config_data.get('hospitals', {}).get(hospital_id, {})
    
    def get_patient_info_patterns(self, hospital_id: str = "default") -> Dict[str, str]:
        """
        获取患者信息提取正则表达式
        
        Args:
            hospital_id: 医院标识
            
        Returns:
            字段名到正则表达式的映射
        """
        hospital_config = self.get_hospital_config(hospital_id)
        return hospital_config.get('patient_info_patterns', {})
    
    def get_time_patterns(self, hospital_id: str = "default") -> Dict[str, str]:
        """
        获取时间信息提取正则表达式
        
        Args:
            hospital_id: 医院标识
            
        Returns:
            时间字段到正则表达式的映射
        """
        hospital_config = self.get_hospital_config(hospital_id)
        return hospital_config.get('time_patterns', {})
    
    def get_table_extraction_config(self, hospital_id: str = "default") -> Dict[str, Any]:
        """
        获取表格提取配置
        
        Args:
            hospital_id: 医院标识
            
        Returns:
            表格提取配置
        """
        hospital_config = self.get_hospital_config(hospital_id)
        return hospital_config.get('table_extraction', {})
    
    def get_excel_config(self, hospital_id: str = "default") -> Dict[str, Any]:
        """
        获取 Excel 输出配置
        
        Args:
            hospital_id: 医院标识
            
        Returns:
            Excel 配置
        """
        hospital_config = self.get_hospital_config(hospital_id)
        return hospital_config.get('excel_config', {})
    
    def get_system_config(self) -> Dict[str, Any]:
        """
        获取系统配置
        
        Returns:
            系统配置字典
        """
        return self._config_data.get('system', {})
    
    def get_folder_mappings(self) -> Dict[str, str]:
        """
        获取报告类型到文件夹名称的映射
        
        Returns:
            报告类型 ID 到文件夹名称的映射
        """
        mappings = {}
        for report_id, config in self.get_report_types().items():
            mappings[report_id] = config.get('folder_name', report_id)
        return mappings
    
    def get_classification_keywords(self, report_type_id: str) -> List[str]:
        """
        获取报告类型的分类关键词
        
        Args:
            report_type_id: 报告类型标识
            
        Returns:
            关键词列表
        """
        report_config = self.get_report_type(report_type_id)
        if report_config:
            return report_config.get('classification_keywords', [])
        return []
    
    def get_indicator_fields(self, report_type_id: str) -> List[str]:
        """
        获取报告类型的特征指标字段
        
        Args:
            report_type_id: 报告类型标识
            
        Returns:
            指标字段列表
        """
        report_config = self.get_report_type(report_type_id)
        if report_config:
            return report_config.get('indicator_fields', [])
        return []


# 全局配置实例
_config_instance: Optional[Config] = None


def get_config(config_path: Optional[str] = None) -> Config:
    """
    获取全局配置实例（单例模式）
    
    Args:
        config_path: 配置文件路径
        
    Returns:
        Config 实例
    """
    global _config_instance
    if _config_instance is None or config_path is not None:
        _config_instance = Config(config_path)
    return _config_instance


def reload_config(config_path: Optional[str] = None) -> Config:
    """
    重新加载配置
    
    Args:
        config_path: 配置文件路径
        
    Returns:
        新的 Config 实例
    """
    global _config_instance
    _config_instance = Config(config_path)
    return _config_instance
