# -*- coding: utf-8 -*-
"""
ICU 报告处理工具 - 主入口

功能：
  1. 只分类整理 PDF
  2. 分类整理 + 生成 Excel
  3. 从已分类的文件生成 Excel

以后可扩展：
  - 日志系统
  - GUI 界面
  - 更多导出格式
"""

import sys
import shutil
from pathlib import Path
from datetime import datetime
from typing import List

# 添加 src 到路径
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from icu_report_processor import create_processor
from icu_report_processor.pdf_utils import get_file_md5, get_all_pdf_files


def print_menu():
    """打印菜单"""
    print("\n" + "=" * 50)
    print("ICU 报告处理工具")
    print("=" * 50)
    print("\n请选择功能：")
    print("  1. 只分类整理 PDF")
    print("  2. 分类整理 + 生成 Excel")
    print("  3. 从已分类的文件生成 Excel")
    print("  0. 退出")
    print("\n" + "=" * 50)


def get_existing_files_map(organized_dir: Path) -> dict:
    """
    获取已存在文件的 MD5 映射
    
    Returns:
        dict: MD5 -> 文件路径
    """
    existing_files = {}
    if not organized_dir.exists():
        return existing_files
    
    # 获取所有已存在的 PDF 文件
    pdf_files = get_all_pdf_files(organized_dir)
    
    for pdf_path in pdf_files:
        md5 = get_file_md5(pdf_path)
        if md5:
            existing_files[md5] = pdf_path
    
    return existing_files


def archive_processed_files(processed_files: List[Path], raw_dir: Path) -> Path:
    """
    将处理过的文件归档到 raw/已处理/日期/ 目录
    
    Args:
        processed_files: 已处理的文件列表
        raw_dir: raw 文件夹路径
        
    Returns:
        归档目录路径
    """
    # 创建归档目录：raw/已处理/2025-04-13/
    today = datetime.now().strftime('%Y-%m-%d')
    archive_dir = raw_dir / "已处理" / today
    archive_dir.mkdir(parents=True, exist_ok=True)
    
    archived_count = 0
    for pdf_path in processed_files:
        try:
            # 保持原有的子目录结构
            relative_path = pdf_path.relative_to(raw_dir)
            target_path = archive_dir / relative_path
            target_path.parent.mkdir(parents=True, exist_ok=True)
            
            # 移动文件
            shutil.move(str(pdf_path), str(target_path))
            archived_count += 1
        except Exception as e:
            print(f"  [归档失败] {pdf_path.name}: {e}")
    
    print(f"\n  已归档 {archived_count} 个文件到: {archive_dir}")
    return archive_dir


def option_1_classify_only():
    """选项1：只分类整理 PDF"""
    print("\n【功能1：只分类整理 PDF】")
    
    base_dir = Path(__file__).parent
    input_dir = base_dir / "raw"
    output_dir = base_dir / "检测报告整理"
    
    if not input_dir.exists():
        print(f"错误：找不到输入文件夹 {input_dir}")
        print("请把 PDF 文件放在 raw 文件夹里")
        return False
    
    print(f"输入：{input_dir}")
    print(f"输出：{output_dir}")
    
    # 检查是否已存在分类文件夹
    mode = "full"
    if output_dir.exists() and any(output_dir.iterdir()):
        print(f"\n检测到 {output_dir.name} 文件夹已存在且不为空")
        print("请选择处理方式：")
        print("  1. 清空后重新处理")
        print("  2. 增量处理（只处理新文件）")
        print("  3. 取消")
        
        choice = input("\n请选择 (1/2/3): ").strip()
        
        if choice == "1":
            mode = "full"
            print("\n将清空现有文件夹并重新处理...")
            # 清空文件夹
            import shutil
            for item in output_dir.iterdir():
                if item.is_dir():
                    shutil.rmtree(item)
                else:
                    item.unlink()
        elif choice == "2":
            mode = "incremental"
            print("\n将使用增量模式处理...")
        else:
            print("\n已取消")
            return False
    
    try:
        processor = create_processor()
        
        # 只执行分类整理
        pdf_files = get_all_pdf_files(input_dir)
        print(f"\n找到 {len(pdf_files)} 个 PDF 文件")
        
        if len(pdf_files) == 0:
            print("没有文件需要处理")
            return False
        
        # 如果是增量模式，过滤掉已存在的文件
        if mode == "incremental":
            existing_files = get_existing_files_map(output_dir)
            new_files = []
            skipped_count = 0
            
            for pdf_path in pdf_files:
                md5 = get_file_md5(pdf_path)
                if md5 and md5 in existing_files:
                    skipped_count += 1
                else:
                    new_files.append(pdf_path)
            
            pdf_files = new_files
            print(f"  跳过 {skipped_count} 个已存在的文件")
            print(f"  实际处理 {len(pdf_files)} 个新文件")
            
            if len(pdf_files) == 0:
                print("\n没有新文件需要处理")
                return True
        
        # 分类
        classified = processor._classify_reports(pdf_files)
        
        # 整理文件
        processor._organize_files(classified, output_dir)
        
        print("\n✓ 分类完成！")
        for folder_name, files in classified.items():
            print(f"  {folder_name}: {len(files)} 个文件")
        
        if mode == "incremental":
            print(f"\n（原有文件已保留，只添加了新文件）")
        
        # 归档处理过的文件
        print("\n[归档] 移动已处理的文件到 raw/已处理/...")
        archive_processed_files(pdf_files, input_dir)
        
        return True
        
    except Exception as e:
        print(f"\n错误：{e}")
        import traceback
        traceback.print_exc()
        return False


def option_2_classify_and_export():
    """选项2：分类整理 + 生成 Excel"""
    print("\n【功能2：分类整理 + 生成 Excel】")
    
    base_dir = Path(__file__).parent
    input_dir = base_dir / "raw"
    organized_dir = base_dir / "检测报告整理"
    output_dir = base_dir
    
    if not input_dir.exists():
        print(f"错误：找不到输入文件夹 {input_dir}")
        print("请把 PDF 文件放在 raw 文件夹里")
        return False
    
    print(f"输入：{input_dir}")
    print(f"分类输出：{organized_dir}")
    print(f"Excel 输出：{output_dir}")
    
    # 检查是否已存在分类文件夹
    mode = "full"
    if organized_dir.exists() and any(organized_dir.iterdir()):
        print(f"\n检测到 {organized_dir.name} 文件夹已存在且不为空")
        print("请选择处理方式：")
        print("  1. 清空后重新处理")
        print("  2. 增量处理（只处理新文件）")
        print("  3. 取消")
        
        choice = input("\n请选择 (1/2/3): ").strip()
        
        if choice == "1":
            mode = "full"
            print("\n将清空现有文件夹并重新处理...")
            # 清空文件夹
            import shutil
            for item in organized_dir.iterdir():
                if item.is_dir():
                    shutil.rmtree(item)
                else:
                    item.unlink()
        elif choice == "2":
            mode = "incremental"
            print("\n将使用增量模式处理...")
        else:
            print("\n已取消")
            return False
    
    try:
        processor = create_processor()
        
        # 获取所有 PDF 文件
        pdf_files = get_all_pdf_files(input_dir)
        print(f"\n扫描到 {len(pdf_files)} 个 PDF 文件")
        
        # 如果是增量模式，过滤掉已存在的文件
        if mode == "incremental":
            existing_files = get_existing_files_map(organized_dir)
            new_files = []
            skipped_count = 0
            
            for pdf_path in pdf_files:
                md5 = get_file_md5(pdf_path)
                if md5 and md5 in existing_files:
                    skipped_count += 1
                else:
                    new_files.append(pdf_path)
            
            pdf_files = new_files
            print(f"  跳过 {skipped_count} 个已存在的文件")
            print(f"  实际处理 {len(pdf_files)} 个新文件")
            
            if len(pdf_files) == 0:
                print("\n没有新文件需要处理")
                # 仍然生成 Excel，包含所有数据
                print("\n将基于现有文件生成 Excel...")
        
        if len(pdf_files) > 0:
            # 去重
            print("\n[1/4] 检测并处理重复文件...")
            pdf_files = processor._remove_duplicates(pdf_files)
            print(f"      去重后剩余 {len(pdf_files)} 个文件")
            
            # 分类
            print("\n[2/4] 分类报告...")
            classified = processor._classify_reports(pdf_files)
            for report_type, files in classified.items():
                print(f"      {report_type}: {len(files)} 个文件")
            
            # 解析
            print("\n[3/4] 解析报告数据...")
            new_parsed_data = processor._parse_reports(classified)
            total_records = sum(len(records) for records in new_parsed_data.values())
            print(f"      成功解析 {total_records} 条记录")
            
            # 整理文件
            print("\n[4/4] 整理文件...")
            processor._organize_files(classified, organized_dir)
        else:
            new_parsed_data = {}
        
        # 如果是增量模式，需要合并所有数据（新旧一起）
        if mode == "incremental":
            print("\n合并新旧数据...")
            # 重新扫描 organized_dir 中的所有文件
            all_pdf_files = get_all_pdf_files(organized_dir)
            print(f"  共 {len(all_pdf_files)} 个文件")
            
            # 重新分类和解析所有文件
            all_classified = processor._classify_reports(all_pdf_files)
            all_parsed_data = processor._parse_reports(all_classified)
        else:
            all_parsed_data = new_parsed_data
        
        # 导出 Excel
        print("\n生成 Excel 报告...")
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = output_dir / f"检测报告汇总_{timestamp}.xlsx"
        processor.exporter.export(all_parsed_data, output_path, None)
        
        print("\n✓ 处理完成！")
        print(f"输出文件: {output_path}")
        
        if mode == "incremental":
            print(f"\n（原有文件已保留，Excel 包含所有数据）")
        
        # 归档处理过的文件
        if len(pdf_files) > 0:
            print("\n[归档] 移动已处理的文件到 raw/已处理/...")
            archive_processed_files(pdf_files, input_dir)
        
        return True
        
    except Exception as e:
        print(f"\n错误：{e}")
        import traceback
        traceback.print_exc()
        return False


def option_3_export_only():
    """选项3：从已分类的文件生成 Excel"""
    print("\n【功能3：从已分类的文件生成 Excel】")
    
    base_dir = Path(__file__).parent
    input_dir = base_dir / "检测报告整理"
    output_dir = base_dir
    
    if not input_dir.exists():
        print(f"错误：找不到输入文件夹 {input_dir}")
        print("请先运行功能1或功能2生成分类文件夹")
        return False
    
    print(f"输入：{input_dir}")
    print(f"输出：{output_dir}")
    
    try:
        processor = create_processor()
        
        results = processor.process_directory(
            input_dir=input_dir,
            output_dir=output_dir,
            organized_dir=None,  # 不重新整理
            patient_info=None
        )
        
        print("\n✓ Excel 生成完成！")
        return True
        
    except Exception as e:
        print(f"\n错误：{e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """主函数"""
    while True:
        print_menu()
        
        choice = input("\n请输入数字 (0/1/2/3): ").strip()
        
        if choice == "0":
            print("\n再见！")
            break
        elif choice == "1":
            option_1_classify_only()
        elif choice == "2":
            option_2_classify_and_export()
        elif choice == "3":
            option_3_export_only()
        else:
            print("\n无效的选择，请重新输入")
        
        input("\n按回车键继续...")


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n程序已退出")
        sys.exit(0)
