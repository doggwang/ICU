# -*- coding: utf-8 -*-
"""
新系统示例脚本 - 使用重构后的处理器处理现有报告

这个脚本演示如何使用新的模块化系统处理 ICU 检测报告
"""

import sys
from pathlib import Path
from datetime import datetime

# 添加 src 到路径
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from icu_report_processor import create_processor


def main():
    """主函数"""
    # 设置路径
    base_dir = Path(__file__).parent
    input_dir = base_dir / "检测报告整理"  # 使用已分类的报告
    output_dir = base_dir
    
    # 患者信息（可选）
    patient_info = {
        '姓名': '王路生',
        '性别': '男',
        '年龄': '67岁',
        '住院号': '1075642',
        '床号': '15',
        '科室': 'ICU',
    }
    
    print("=" * 60)
    print("ICU 检测报告处理 - 新系统演示")
    print("=" * 60)
    
    # 创建处理器
    processor = create_processor()
    
    # 处理报告
    try:
        results = processor.process_directory(
            input_dir=input_dir,
            output_dir=output_dir,
            organized_dir=None,  # 不重新整理文件
            patient_info=patient_info
        )
        
        # 打印统计信息
        print("\n各报告类型统计:")
        for report_type, records in results.items():
            print(f"  {report_type}: {len(records)} 条记录")
        
    except Exception as e:
        print(f"\n处理失败: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == '__main__':
    sys.exit(main())
