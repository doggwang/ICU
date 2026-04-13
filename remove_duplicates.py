# -*- coding: utf-8 -*-
"""
删除重复文件脚本
用于删除分类后PDF中的重复文件（基于MD5校验）

使用方法：
1. 运行脚本：python remove_duplicates.py
2. 脚本会自动检测并删除同一检查类型中内容完全相同的文件
3. 每个唯一文件只保留一份
"""

import hashlib
import subprocess
import time
from pathlib import Path
from collections import defaultdict

ICU_DIR = Path("c:/Users/39863/Desktop/ICU")

# 中文文件夹名称
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
    except Exception as e:
        return None

def safe_delete(filepath):
    """安全删除文件（尝试多种方法）"""
    filepath = Path(filepath)
    if not filepath.exists():
        return True

    # 方法1: pathlib
    try:
        filepath.unlink()
        return True
    except:
        pass

    # 方法2: cmd del
    try:
        subprocess.run(['cmd', '/c', 'del', '/f', '/q', str(filepath)],
                      capture_output=True, timeout=10)
        if not filepath.exists():
            return True
    except:
        pass

    # 方法3: powershell
    try:
        subprocess.run(['powershell', '-Command', f'Remove-Item -Force "{filepath}"'],
                      capture_output=True, timeout=10)
        if not filepath.exists():
            return True
    except:
        pass

    return False

def find_and_remove_duplicates():
    """查找并删除重复的PDF文件"""
    # 跳过raw文件夹
    skip_dirs = {'raw', '.claude'}
    total_deleted = 0
    total_kept = 0

    for item in sorted(ICU_DIR.iterdir()):
        if not item.is_dir() or item.name in skip_dirs:
            continue

        print(f"\n=== 处理 {item.name} ===")

        # 按MD5哈希值分组
        hash_map = defaultdict(list)
        pdf_files = list(item.glob("*.pdf"))

        for pdf in pdf_files:
            md5 = get_md5(pdf)
            if md5:
                hash_map[md5].append(pdf)

        # 处理每组重复文件
        for md5, files in sorted(hash_map.items()):
            if len(files) > 1:
                print(f"  MD5 {md5[:8]}... 发现 {len(files)} 个相同文件")

                # 按名称排序
                files.sort(key=lambda x: x.name)

                keep_file = None
                delete_files = []

                for f in files:
                    name = f.stem
                    # 检查是否是重复文件（以_N结尾）
                    if '_' in name:
                        parts = name.split('_')
                        try:
                            int(parts[-1])
                            delete_files.append(f)
                        except ValueError:
                            if keep_file is None:
                                keep_file = f
                            else:
                                delete_files.append(f)
                    else:
                        if keep_file is None:
                            keep_file = f
                        else:
                            delete_files.append(f)

                # 如果没找到原文件，保留第一个
                if keep_file is None:
                    keep_file = files[0]
                    delete_files = files[1:]

                print(f"    保留: {keep_file.name}")

                # 删除重复文件
                for f in delete_files:
                    print(f"    删除: {f.name}")
                    if safe_delete(f):
                        print(f"      成功")
                        total_deleted += 1
                    else:
                        print(f"      失败（文件被锁定）")
                    time.sleep(0.1)

                total_kept += 1

    print(f"\n{'='*50}")
    print(f"完成！删除了 {total_deleted} 个重复文件，保留了 {total_kept} 组唯一文件")

if __name__ == "__main__":
    find_and_remove_duplicates()