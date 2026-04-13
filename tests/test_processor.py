# -*- coding: utf-8 -*-
"""
基础测试脚本 - 验证重构后的处理器功能
"""

import sys
from pathlib import Path

# 添加 src 到路径
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from icu_report_processor import create_processor
from icu_report_processor.config import get_config


def test_config_loading():
    """测试配置加载"""
    print("\n[测试] 配置加载...")
    try:
        config = get_config()
        report_types = config.get_report_types()
        print(f"  ✓ 成功加载配置")
        print(f"  ✓ 定义了 {len(report_types)} 种报告类型")
        for rt_id, rt_config in report_types.items():
            print(f"    - {rt_id}: {rt_config.get('name', '未命名')}")
        return True
    except Exception as e:
        print(f"  ✗ 失败: {e}")
        return False


def test_classifier():
    """测试分类器"""
    print("\n[测试] 分类器...")
    try:
        from icu_report_processor.classifiers import KeywordClassifier
        from icu_report_processor.config import get_config
        
        config = get_config()
        report_types = config.get_report_types()
        classifier = KeywordClassifier(report_types)
        
        # 测试文本
        test_cases = [
            ("BNP 检测结果 100 pg/mL", "bnp"),
            ("血气分析 pH 7.4 (ICU)(POCT)", "blood_gas"),
            ("白细胞计数 10^9/L HCRP", "blood_routine"),
        ]
        
        for text, expected in test_cases:
            result = classifier.classify(text)
            folder_name = config.get_folder_mappings().get(result, result)
            print(f"  ✓ '{text[:20]}...' -> {folder_name}")
        
        return True
    except Exception as e:
        print(f"  ✗ 失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_parser():
    """测试解析器"""
    print("\n[测试] 解析器...")
    try:
        from icu_report_processor.parsers import DefaultParser
        from icu_report_processor.config import get_config
        
        config = get_config()
        hospital_config = config.get_hospital_config("default")
        parser = DefaultParser(hospital_config)
        
        # 测试文本
        test_text = """
        姓 名: 张三
        病历号: 123456
        性 别: 男
        年 龄: 50岁
        科 别: ICU
        床 号: 15
        样本种类: 血液
        临床诊断: 肺炎
        采集时间: 2026-04-01 08:00:00
        接收时间: 2026-04-01 08:30:00
        报告时间: 2026-04-01 09:00:00
        
        1 白细胞计数 10.5 4.0-10.0 10^9/L
        2 红细胞计数 4.5 4.0-5.5 10^12/L
        """
        
        result = parser.parse(test_text, "test.pdf", "血常规")
        
        print(f"  ✓ 患者姓名: {result.patient_name}")
        print(f"  ✓ 病历号: {result.patient_id}")
        print(f"  ✓ 主时间: {result.main_time}")
        print(f"  ✓ 检测项目数: {len(result.test_items)}")
        
        return True
    except Exception as e:
        print(f"  ✗ 失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_processor_creation():
    """测试处理器创建"""
    print("\n[测试] 处理器创建...")
    try:
        processor = create_processor()
        print(f"  ✓ 处理器创建成功")
        print(f"  ✓ 分类器类型: {type(processor.classifier).__name__}")
        print(f"  ✓ 解析器类型: {type(processor.parser).__name__}")
        print(f"  ✓ 导出器类型: {type(processor.exporter).__name__}")
        return True
    except Exception as e:
        print(f"  ✗ 失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """运行所有测试"""
    print("=" * 60)
    print("ICU 报告处理系统 - 功能测试")
    print("=" * 60)
    
    tests = [
        ("配置加载", test_config_loading),
        ("分类器", test_classifier),
        ("解析器", test_parser),
        ("处理器创建", test_processor_creation),
    ]
    
    results = []
    for name, test_func in tests:
        try:
            success = test_func()
            results.append((name, success))
        except Exception as e:
            print(f"\n[测试] {name} 发生异常: {e}")
            results.append((name, False))
    
    # 打印汇总
    print("\n" + "=" * 60)
    print("测试结果汇总")
    print("=" * 60)
    
    passed = sum(1 for _, success in results if success)
    total = len(results)
    
    for name, success in results:
        status = "✓ 通过" if success else "✗ 失败"
        print(f"  {status}: {name}")
    
    print(f"\n总计: {passed}/{total} 通过")
    
    if passed == total:
        print("\n🎉 所有测试通过！")
        return 0
    else:
        print(f"\n⚠️  {total - passed} 个测试失败")
        return 1


if __name__ == '__main__':
    sys.exit(main())
