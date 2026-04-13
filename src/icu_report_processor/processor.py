# -*- coding: utf-8 -*-
"""
主处理器模块 - 整合分类、解析、导出功能
"""

import shutil
import send2trash
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from collections import defaultdict
from datetime import datetime

from .config import get_config, Config
from .classifiers import KeywordClassifier
from .parsers import DefaultParser, ParseResult
from .exporters import ExcelExporter
from .pdf_utils import (
    get_file_md5, extract_text_from_pdf, find_duplicate_files,
    get_all_pdf_files, sanitize_filename, extract_timestamp_from_text
)


class ReportProcessor:
    """
    报告处理器
    
    整合 PDF 分类、数据解析、Excel 导出全流程
    """
    
    def __init__(self, config: Optional[Config] = None, hospital_id: str = "default"):
        """
        初始化处理器
        
        Args:
            config: 配置对象，为 None 时使用默认配置
            hospital_id: 医院标识
        """
        self.config = config or get_config()
        self.hospital_id = hospital_id
        
        # 初始化分类器
        report_types = self.config.get_report_types()
        self.classifier = KeywordClassifier(report_types)
        
        # 初始化解析器
        hospital_config = self.config.get_hospital_config(hospital_id)
        self.parser = DefaultParser(hospital_config)
        
        # 初始化导出器
        excel_config = self.config.get_excel_config(hospital_id)
        self.exporter = ExcelExporter(excel_config)
        
        # 获取文件夹映射
        self.folder_mappings = self.config.get_folder_mappings()
    
    def process_directory(self, input_dir: Path, output_dir: Path,
                         organized_dir: Optional[Path] = None,
                         patient_info: Optional[Dict[str, str]] = None) -> Dict[str, List[ParseResult]]:
        """
        处理目录中的所有 PDF 报告
        
        Args:
            input_dir: 输入目录（原始 PDF 文件）
            output_dir: 输出目录（Excel 文件）
            organized_dir: 可选，分类后的 PDF 存储目录
            patient_info: 可选，患者基本信息
            
        Returns:
            按报告类型分组的解析结果
        """
        print("=" * 60)
        print("ICU 检测报告处理系统")
        print("=" * 60)
        
        # 确保输出目录存在
        output_dir.mkdir(parents=True, exist_ok=True)
        if organized_dir:
            organized_dir.mkdir(parents=True, exist_ok=True)
            # 创建各报告类型子目录
            for folder_name in self.folder_mappings.values():
                (organized_dir / folder_name).mkdir(exist_ok=True)
        
        # 获取所有 PDF 文件
        pdf_files = get_all_pdf_files(input_dir)
        print(f"\n扫描到 {len(pdf_files)} 个 PDF 文件")
        
        # 去重
        print("\n[1/4] 检测并处理重复文件...")
        pdf_files = self._remove_duplicates(pdf_files)
        print(f"      去重后剩余 {len(pdf_files)} 个文件")
        
        # 分类
        print("\n[2/4] 分类报告...")
        classified = self._classify_reports(pdf_files)
        for report_type, files in classified.items():
            print(f"      {report_type}: {len(files)} 个文件")
        
        # 解析
        print("\n[3/4] 解析报告数据...")
        parsed_data = self._parse_reports(classified)
        total_records = sum(len(records) for records in parsed_data.values())
        print(f"      成功解析 {total_records} 条记录")
        
        # 如果需要，复制文件到分类目录
        if organized_dir:
            print("\n[4/4] 整理文件...")
            self._organize_files(classified, organized_dir)
        
        # 导出 Excel
        print("\n生成 Excel 报告...")
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = output_dir / f"检测报告汇总_{timestamp}.xlsx"
        self.exporter.export(parsed_data, output_path, patient_info)
        
        print("\n" + "=" * 60)
        print("处理完成！")
        print(f"输出文件: {output_path}")
        print("=" * 60)
        
        return parsed_data
    
    def _remove_duplicates(self, pdf_files: List[Path]) -> List[Path]:
        """
        删除重复文件，只保留一份
        
        Args:
            pdf_files: PDF 文件路径列表
            
        Returns:
            去重后的文件列表
        """
        # 计算 MD5 并分组
        hash_map = defaultdict(list)
        for pdf_path in pdf_files:
            md5 = get_file_md5(pdf_path)
            if md5:
                hash_map[md5].append(pdf_path)
        
        unique_files = []
        duplicates_count = 0
        
        for md5, files in hash_map.items():
            # 保留第一个，删除其余
            files.sort(key=lambda x: x.name)
            unique_files.append(files[0])
            
            for dup_file in files[1:]:
                try:
                    # 移动到回收站而不是永久删除
                    send2trash.send2trash(str(dup_file))
                    duplicates_count += 1
                    print(f"      [移到回收站] {dup_file.name}")
                except Exception as e:
                    print(f"      [移动失败] {dup_file.name}: {e}")
        
        if duplicates_count > 0:
            print(f"      共移动 {duplicates_count} 个重复文件到回收站")
        
        return unique_files
    
    def _classify_reports(self, pdf_files: List[Path]) -> Dict[str, List[Path]]:
        """
        对 PDF 文件进行分类
        
        Args:
            pdf_files: PDF 文件路径列表
            
        Returns:
            按报告类型分组的文件字典
        """
        classified = defaultdict(list)
        unclassified = []
        
        for pdf_path in pdf_files:
            text = extract_text_from_pdf(pdf_path)
            if not text:
                print(f"      [无法读取] {pdf_path.name}")
                continue
            
            report_type = self.classifier.classify(text)
            if report_type:
                # 使用中文文件夹名作为 key
                folder_name = self.folder_mappings.get(report_type, report_type)
                classified[folder_name].append(pdf_path)
            else:
                unclassified.append(pdf_path)
        
        if unclassified:
            print(f"      [未分类] {len(unclassified)} 个文件")
        
        return dict(classified)
    
    def _parse_reports(self, classified: Dict[str, List[Path]]) -> Dict[str, List[ParseResult]]:
        """
        解析分类后的报告
        
        Args:
            classified: 按报告类型分组的文件字典
            
        Returns:
            按报告类型分组的解析结果
        """
        parsed_data = {}
        
        for report_type, files in classified.items():
            results = []
            for pdf_path in files:
                text = extract_text_from_pdf(pdf_path)
                if text:
                    result = self.parser.parse(text, pdf_path.name, report_type)
                    results.append(result)
            
            if results:
                parsed_data[report_type] = results
        
        return parsed_data
    
    def _organize_files(self, classified: Dict[str, List[Path]], 
                       organized_dir: Path) -> None:
        """
        将文件整理到分类目录
        
        Args:
            classified: 按报告类型分组的文件字典
            organized_dir: 目标整理目录
        """
        # 创建反向映射（文件夹名 -> report_type id）
        folder_to_id = {v: k for k, v in self.folder_mappings.items()}
        
        for folder_name, files in classified.items():
            target_dir = organized_dir / folder_name
            target_dir.mkdir(exist_ok=True)
            
            for pdf_path in files:
                # 提取时间戳用于命名
                text = extract_text_from_pdf(pdf_path)
                timestamp = extract_timestamp_from_text(text) if text else None
                
                if timestamp:
                    safe_time = sanitize_filename(timestamp)
                else:
                    safe_time = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                # 构建新文件名
                report_id = folder_to_id.get(folder_name, folder_name)
                new_name = f"{report_id}_{safe_time}.pdf"
                dest_path = target_dir / new_name
                
                # 处理重名
                counter = 1
                while dest_path.exists():
                    new_name = f"{report_id}_{safe_time}_{counter}.pdf"
                    dest_path = target_dir / new_name
                    counter += 1
                
                shutil.copy2(pdf_path, dest_path)
    
    def process_single_file(self, pdf_path: Path) -> Optional[ParseResult]:
        """
        处理单个 PDF 文件
        
        Args:
            pdf_path: PDF 文件路径
            
        Returns:
            解析结果，失败返回 None
        """
        text = extract_text_from_pdf(pdf_path)
        if not text:
            return None
        
        report_type = self.classifier.classify(text)
        if not report_type:
            report_type = "未知类型"
        
        folder_name = self.folder_mappings.get(report_type, report_type)
        result = self.parser.parse(text, pdf_path.name, folder_name)
        
        return result


def create_processor(config_path: Optional[str] = None, 
                    hospital_id: str = "default") -> ReportProcessor:
    """
    创建报告处理器实例
    
    Args:
        config_path: 配置文件路径
        hospital_id: 医院标识
        
    Returns:
        ReportProcessor 实例
    """
    config = get_config(config_path)
    return ReportProcessor(config, hospital_id)
