# -*- coding: utf-8 -*-
"""
PDF分类和去重脚本 v3
用于ICU检验报告的自动分类和整理

功能：
1. 扫描raw文件夹及子文件夹中的所有PDF
2. 基于MD5检测重复文件并删除（保留一份）
3. 对新增/未分类的文件进行分类
4. 不移动原始文件，保持其在raw中的位置

分类目标：检测报告整理/
  - 生化检验/
  - 血气分析/
  - 血常规/
  - 尿常规/
  - 药敏试验/
  - D二聚体/
  - BNP心衰标志物/
  - PCT降钙素原/
  - 炎症因子/
  - ACT活化凝血时间/
  - TBNK免疫细胞/
  - 心脏超声/

使用方法：
1. 将新的PDF文件放入 raw/ 文件夹或其子文件夹
2. 运行脚本：python rename_pdfs.py
"""

import re
import shutil
import hashlib
import subprocess
from pathlib import Path
from collections import defaultdict

# 路径设置
BASE_DIR = Path("c:/Users/39863/Desktop/ICU")
RAW_DIR = BASE_DIR / "raw"
OUTPUT_DIR = BASE_DIR / "检测报告整理"

# 中文文件夹名称映射
FOLDER_NAMES = {
    "Biochemistry": "生化检验",
    "Blood_Routine": "血常规",
    "Blood_Gas": "血气分析",
    "Drug_Sensitivity": "药敏试验",
    "BNP": "BNP心衰标志物",
    "Urine_Routine": "尿常规",
    "D_Dimer": "D二聚体",
    "PCT": "PCT降钙素原",
    "TBNK": "TBNK免疫细胞",
    "Echocardiography": "心脏超声",
    "Cytokines": "炎症因子",
    "ACT": "ACT活化凝血时间",
}

def get_md5(filepath):
    """计算文件的MD5哈希值"""
    hash_md5 = hashlib.md5()
    try:
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception:
        return None

def extract_receive_time(text):
    """从PDF文本中提取接收时间（第一个时间戳）"""
    patterns = [
        r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})',
    ]
    for pattern in patterns:
        matches = re.findall(pattern, text)
        if matches:
            return matches[0]
    return None

def categorize(text):
    """根据PDF内容确定检查类型"""
    if "BNP" in text:
        return "BNP"
    if "TBNK" in text:
        return "TBNK"
    if "(ICU)(POCT)" in text:
        return "Blood_Gas"
    if "HCRP" in text:
        return "Blood_Routine"
    if "MIC" in text and "g/ml" in text:
        return "Drug_Sensitivity"
    if "+D2" in text:
        return "D_Dimer"
    if "ng/mL" in text:
        return "PCT"
    if "LVEF" in text or "IVC" in text:
        return "Echocardiography"
    if "IL-2" in text or "IL-4" in text or "IL-6" in text or "IL-8" in text or "IL-10" in text:
        return "Cytokines"
    if "AST:ALT" in text or "UN:CREA" in text or "eGFR-EPI" in text:
        return "Biochemistry"
    if "1.003-1.030" in text or "(-)(" in text:
        return "Urine_Routine"
    if "ICU12" in text and "g/L" in text and "2.00-4.00" in text:
        return "Urine_Routine"
    if "U/L" in text and "ICU12" in text and "mol/L" in text:
        return "ACT"
    if "mmol/L" in text and "g/L" in text and "U/L" in text:
        return "Biochemistry"
    return None

def sanitize_filename(name):
    """移除文件名中的非法字符"""
    name = re.sub(r'[<>:"/\\|?*\s]', '_', name)
    name = re.sub(r'[\x00-\x1f\x7f-\xff]', '', name)
    name = re.sub(r'_+', '_', name)
    return name

def read_pdf_text(pdf_path):
    """使用pdftotext读取PDF文本"""
    try:
        result = subprocess.run(
            ['pdftotext', '-layout', str(pdf_path), '-'],
            capture_output=True,
            timeout=30
        )
        try:
            return result.stdout.decode('utf-8')
        except:
            return result.stdout.decode('latin-1', errors='replace')
    except Exception:
        return ""

def get_all_pdf_files(directory):
    """获取目录及所有子目录中的PDF文件"""
    pdf_files = []
    for item in Path(directory).iterdir():
        if item.is_file() and item.suffix.lower() == '.pdf':
            pdf_files.append(item)
        elif item.is_dir():
            pdf_files.extend(get_all_pdf_files(item))
    return pdf_files

def get_classified_info():
    """获取已分类文件夹中的文件信息 {timestamp: (category, md5)}"""
    classified = {}  # timestamp -> (category, md5)
    if not OUTPUT_DIR.exists():
        return classified
    for item in OUTPUT_DIR.iterdir():
        if item.is_dir():
            for pdf in item.glob("*.pdf"):
                # 格式: Category_YYYY-MM-DD_HH_MM_SS[.pdf or _N.pdf]
                stem = pdf.stem
                parts = stem.split('_', 1)
                if len(parts) >= 2:
                    cat = parts[0]
                    timestamp = parts[1].rsplit('_', 1)[0]  # 去掉可能的_N后缀
                    md5 = get_md5(pdf)
                    if md5 and timestamp not in classified:
                        classified[timestamp] = (cat, md5)
    return classified

def process_pdfs():
    """处理raw目录中的所有PDF文件"""
    print("=" * 50)
    print("ICU检验报告分类脚本 v3")
    print("=" * 50)

    # 确保输出目录存在
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for cat in FOLDER_NAMES.values():
        (OUTPUT_DIR / cat).mkdir(exist_ok=True)

    # 获取所有PDF文件
    pdf_files = get_all_pdf_files(RAW_DIR)
    print(f"\n扫描到 {len(pdf_files)} 个PDF文件")

    # 获取已分类文件的信息
    classified_info = get_classified_info()
    classified_timestamps = set(classified_info.keys())
    print(f"已分类文件中包含 {len(classified_timestamps)} 个不同时间戳")

    # 按MD5分组，检测重复
    print("\n[第一步] 检测并处理重复文件...")
    hash_map = defaultdict(list)
    for pdf in pdf_files:
        md5 = get_md5(pdf)
        if md5:
            hash_map[md5].append(pdf)

    duplicates_deleted = 0
    files_to_classify = []

    for md5, files in hash_map.items():
        if len(files) > 1:
            # 有重复，按名称排序，保留第一个
            files.sort(key=lambda x: x.name)
            keep_file = files[0]
            files_to_classify.append((keep_file, md5))
            # 删除其余重复文件
            for f in files[1:]:
                try:
                    f.unlink()
                    duplicates_deleted += 1
                    print(f"  [删除重复] {f.relative_to(RAW_DIR)}")
                except Exception as e:
                    print(f"  [删除失败] {f.relative_to(RAW_DIR)}")
        else:
            files_to_classify.append((files[0], md5))

    # 第二步：分类新增文件
    print(f"\n[第二步] 检查并分类新增文件...")
    print(f"待处理文件数: {len(files_to_classify)}")

    success_count = 0
    skipped_count = 0

    for pdf_file, md5 in files_to_classify:
        text = read_pdf_text(pdf_file)
        timestamp = extract_receive_time(text)

        if not timestamp:
            print(f"  [跳过-无时间] {pdf_file.name}")
            continue

        # 检查是否已分类（通过时间戳+MD5匹配）
        if timestamp in classified_timestamps:
            stored_cat, stored_md5 = classified_info[timestamp]
            if stored_md5 == md5:
                # 完全匹配，已分类
                skipped_count += 1
                continue

        # 需要分类
        cat = categorize(text)
        if cat:
            folder_name = FOLDER_NAMES.get(cat, cat)
            cat_dir = OUTPUT_DIR / folder_name

            safe_time = sanitize_filename(timestamp)
            new_name = f"{sanitize_filename(cat)}_{safe_time}.pdf"
            dest_path = cat_dir / new_name

            # 处理重名
            counter = 1
            while dest_path.exists():
                new_name = f"{sanitize_filename(cat)}_{safe_time}_{counter}.pdf"
                dest_path = cat_dir / new_name
                counter += 1

            shutil.copy2(pdf_file, dest_path)
            success_count += 1
            print(f"  [{folder_name}] {new_name}")

            # 更新已分类信息，防止同一timestamp被多次处理
            if timestamp not in classified_info:
                classified_info[timestamp] = (cat, md5)
                classified_timestamps.add(timestamp)
        else:
            print(f"  [未分类] {pdf_file.name}")

    # 统计
    print(f"\n{'='*50}")
    print(f"完成！")
    print(f"{'='*50}")
    print(f"删除了 {duplicates_deleted} 个重复文件")
    print(f"新增分类 {success_count} 个文件")
    print(f"跳过已分类 {skipped_count} 个文件")

    # 显示各文件夹文件数
    print(f"\n各文件夹文件数:")
    for cat in sorted(FOLDER_NAMES.keys()):
        folder = FOLDER_NAMES[cat]
        count = len(list((OUTPUT_DIR / folder).glob("*.pdf")))
        if count > 0:
            print(f"  {folder}: {count} 个文件")

if __name__ == "__main__":
    process_pdfs()