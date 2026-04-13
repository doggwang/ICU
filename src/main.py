# -*- coding: utf-8 -*-
"""
ICU 检测报告处理系统 - 主入口脚本

使用方法:
    python main.py --input raw --output . --organized 检测报告整理
    
或作为模块运行:
    python -m main --input raw --output .
"""

import argparse
import sys
from pathlib import Path

# 添加 src 到路径
sys.path.insert(0, str(Path(__file__).parent))

from icu_report_processor import create_processor


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='ICU 检测报告处理系统 - 自动分类、解析、导出报告',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 基本用法
  python main.py --input raw --output .
  
  # 同时整理文件到分类目录
  python main.py --input raw --output . --organized 检测报告整理
  
  # 指定患者信息
  python main.py --input raw --output . --patient-name 张三 --patient-id 12345
        """
    )
    
    parser.add_argument(
        '--input', '-i',
        type=str,
        default='raw',
        help='输入目录，包含原始 PDF 报告 (默认: raw)'
    )
    
    parser.add_argument(
        '--output', '-o',
        type=str,
        default='.',
        help='输出目录，用于保存 Excel 文件 (默认: 当前目录)'
    )
    
    parser.add_argument(
        '--organized', '-g',
        type=str,
        default=None,
        help='整理后的 PDF 存储目录，不指定则不整理文件'
    )
    
    parser.add_argument(
        '--config', '-c',
        type=str,
        default=None,
        help='配置文件路径，默认使用内置配置'
    )
    
    parser.add_argument(
        '--hospital', '-H',
        type=str,
        default='default',
        help='医院标识，用于加载对应的解析规则 (默认: default)'
    )
    
    # 患者信息参数
    patient_group = parser.add_argument_group('患者信息（可选）')
    patient_group.add_argument('--patient-name', type=str, help='患者姓名')
    patient_group.add_argument('--patient-id', type=str, help='病历号/住院号')
    patient_group.add_argument('--patient-gender', type=str, help='性别')
    patient_group.add_argument('--patient-age', type=str, help='年龄')
    patient_group.add_argument('--department', type=str, help='科室')
    patient_group.add_argument('--bed-number', type=str, help='床号')
    
    args = parser.parse_args()
    
    # 构建路径
    input_dir = Path(args.input)
    output_dir = Path(args.output)
    organized_dir = Path(args.organized) if args.organized else None
    
    # 验证输入目录
    if not input_dir.exists():
        print(f"错误: 输入目录不存在: {input_dir}")
        sys.exit(1)
    
    # 构建患者信息字典
    patient_info = {}
    if args.patient_name:
        patient_info['姓名'] = args.patient_name
    if args.patient_id:
        patient_info['病历号/住院号'] = args.patient_id
    if args.patient_gender:
        patient_info['性别'] = args.patient_gender
    if args.patient_age:
        patient_info['年龄'] = args.patient_age
    if args.department:
        patient_info['科室'] = args.department
    if args.bed_number:
        patient_info['床号'] = args.bed_number
    
    patient_info = patient_info if patient_info else None
    
    # 创建处理器并运行
    try:
        processor = create_processor(args.config, args.hospital)
        processor.process_directory(
            input_dir=input_dir,
            output_dir=output_dir,
            organized_dir=organized_dir,
            patient_info=patient_info
        )
    except Exception as e:
        print(f"\n错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
