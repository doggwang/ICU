# -*- coding: utf-8 -*-
"""
PDF 工具模块 - 提供 PDF 文件处理功能
"""

import hashlib
import subprocess
from pathlib import Path
from typing import Optional, List, Dict, Tuple
from collections import defaultdict

import pdfplumber


def get_file_md5(filepath: Path) -> Optional[str]:
    """
    计算文件的 MD5 哈希值
    
    Args:
        filepath: 文件路径
        
    Returns:
        MD5 哈希字符串，失败返回 None
    """
    hash_md5 = hashlib.md5()
    try:
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception:
        return None


def extract_text_with_pdfplumber(pdf_path: Path) -> Optional[str]:
    """
    使用 pdfplumber 提取 PDF 文本
    
    Args:
        pdf_path: PDF 文件路径
        
    Returns:
        提取的文本内容，失败返回 None
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
            return text
    except Exception as e:
        print(f"pdfplumber 读取失败 {pdf_path}: {e}")
        return None


def extract_text_with_pdftotext(pdf_path: Path, timeout: int = 30) -> Optional[str]:
    """
    使用 pdftotext 命令行工具提取 PDF 文本
    
    Args:
        pdf_path: PDF 文件路径
        timeout: 超时时间（秒）
        
    Returns:
        提取的文本内容，失败返回 None
    """
    try:
        result = subprocess.run(
            ['pdftotext', '-layout', str(pdf_path), '-'],
            capture_output=True,
            timeout=timeout
        )
        try:
            return result.stdout.decode('utf-8')
        except UnicodeDecodeError:
            return result.stdout.decode('latin-1', errors='replace')
    except Exception as e:
        print(f"pdftotext 读取失败 {pdf_path}: {e}")
        return None


def extract_text_from_pdf(pdf_path: Path, method: str = 'pdfplumber') -> Optional[str]:
    """
    从 PDF 中提取文本
    
    Args:
        pdf_path: PDF 文件路径
        method: 提取方法，'pdfplumber' 或 'pdftotext'
        
    Returns:
        提取的文本内容，失败返回 None
    """
    if method == 'pdfplumber':
        return extract_text_with_pdfplumber(pdf_path)
    elif method == 'pdftotext':
        return extract_text_with_pdftotext(pdf_path)
    else:
        raise ValueError(f"不支持的提取方法: {method}")


def find_duplicate_files(pdf_files: List[Path]) -> Dict[str, List[Path]]:
    """
    查找重复文件（基于 MD5）
    
    Args:
        pdf_files: PDF 文件路径列表
        
    Returns:
        MD5 到文件路径列表的映射
    """
    hash_map = defaultdict(list)
    
    for pdf_path in pdf_files:
        md5 = get_file_md5(pdf_path)
        if md5:
            hash_map[md5].append(pdf_path)
    
    # 只返回有重复的文件组
    return {md5: files for md5, files in hash_map.items() if len(files) > 1}


def get_all_pdf_files(directory: Path, recursive: bool = True) -> List[Path]:
    """
    获取目录中的所有 PDF 文件
    
    Args:
        directory: 目标目录
        recursive: 是否递归子目录
        
    Returns:
        PDF 文件路径列表
    """
    if recursive:
        return list(directory.rglob("*.pdf"))
    else:
        return list(directory.glob("*.pdf"))


def sanitize_filename(name: str, illegal_chars_replace: Dict[str, str] = None,
                     collapse_underscores: bool = True) -> str:
    """
    清理文件名中的非法字符
    
    Args:
        name: 原始文件名
        illegal_chars_replace: 非法字符替换规则
        collapse_underscores: 是否合并连续下划线
        
    Returns:
        清理后的文件名
    """
    import re
    
    if illegal_chars_replace is None:
        illegal_chars_replace = {
            r'[<>\":/\\|?*\s]': '_',
            r'[\x00-\x1f\x7f-\xff]': ''
        }
    
    result = name
    for pattern, replacement in illegal_chars_replace.items():
        result = re.sub(pattern, replacement, result)
    
    if collapse_underscores:
        result = re.sub(r'_+', '_', result)
    
    return result


def extract_timestamp_from_text(text: str, pattern: str = r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})') -> Optional[str]:
    """
    从文本中提取时间戳
    
    Args:
        text: 文本内容
        pattern: 时间戳正则表达式
        
    Returns:
        时间戳字符串，未找到返回 None
    """
    import re
    matches = re.findall(pattern, text)
    if matches:
        return matches[0]
    return None
