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
from pathlib import Path
from datetime import datetime

# 添加 src 到路径
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from icu_report_processor import create_processor


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
    
    try:
        processor = create_processor()
        
        # 只执行分类整理
        from icu_report_processor.pdf_utils import get_all_pdf_files
        pdf_files = get_all_pdf_files(input_dir)
        print(f"\n找到 {len(pdf_files)} 个 PDF 文件")
        
        if len(pdf_files) == 0:
            print("没有文件需要处理")
            return False
        
        # 分类
        classified = processor._classify_reports(pdf_files)
        
        # 整理文件
        processor._organize_files(classified, output_dir)
        
        print("\n✓ 分类完成！")
        for folder_name, files in classified.items():
            print(f"  {folder_name}: {len(files)} 个文件")
        
        return True
        
    except Exception as e:
        print(f"\n错误：{e}")
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
    
    try:
        processor = create_processor()
        
        results = processor.process_directory(
            input_dir=input_dir,
            output_dir=output_dir,
            organized_dir=organized_dir,
            patient_info=None
        )
        
        print("\n✓ 处理完成！")
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
