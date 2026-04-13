# -*- coding: utf-8 -*-
"""
ICU患者指标研究Excel创建脚本 - v4
完全基于extract_reports.py的架构
患者：王路生，男，67岁，住院号1075642，床号15，ICU
"""

import re
from pathlib import Path
from collections import defaultdict
from datetime import datetime

import pdfplumber
import pandas as pd
from openpyxl.styles import Alignment, Font


def extract_time_info(text):
    """从PDF文本中提取时间信息（采集时间、接收时间、报告时间）"""
    time_info = {}

    # 匹配采集时间
    collection_match = re.search(r'采集时间[:\s]*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', text)
    if collection_match:
        time_info['采集时间'] = collection_match.group(1)

    # 匹配接收时间
    receive_match = re.search(r'接收时间[:\s]*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', text)
    if receive_match:
        time_info['接收时间'] = receive_match.group(1)

    # 匹配报告时间
    report_match = re.search(r'报告时间[:\s]*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', text)
    if report_match:
        time_info['报告时间'] = report_match.group(1)

    # 获取主时间（优先级：接收时间 > 采集时间 > 报告时间）
    if time_info.get('接收时间'):
        time_info['主时间'] = time_info['接收时间']
    elif time_info.get('采集时间'):
        time_info['主时间'] = time_info['采集时间']
    elif time_info.get('报告时间'):
        time_info['主时间'] = time_info['报告时间']
    else:
        time_info['主时间'] = None

    return time_info


def extract_patient_info(text):
    """提取患者基本信息"""
    info = {}

    # 姓名
    name_match = re.search(r'姓\s*名[:\s]*([^\s\n]+)', text)
    if name_match:
        info['姓名'] = name_match.group(1).strip()

    # 病历号
    id_match = re.search(r'病\s*历\s*号[:\s]*([^\s\n]+)', text)
    if id_match:
        info['病历号'] = id_match.group(1).strip()

    # 性别
    gender_match = re.search(r'性\s*别[:\s]*([^\s\n]+)', text)
    if gender_match:
        info['性别'] = gender_match.group(1).strip()

    # 年龄
    age_match = re.search(r'年\s*龄[:\s]*([^\s\n]+)', text)
    if age_match:
        info['年龄'] = age_match.group(1).strip()

    # 科室
    dept_match = re.search(r'科\s*别[:\s]*([^\s\n]+)', text)
    if dept_match:
        info['科室'] = dept_match.group(1).strip()

    # 床号
    bed_match = re.search(r'床\s*号[:\s]*([^\s\n]+)', text)
    if bed_match:
        info['床号'] = bed_match.group(1).strip()

    # 样本种类
    sample_match = re.search(r'样本种类[:\s]*([^\s\n]+)', text)
    if sample_match:
        info['样本种类'] = sample_match.group(1).strip()

    # 临床诊断
    diagnosis_match = re.search(r'临床诊断[:\s]*([^\n]+)', text)
    if diagnosis_match:
        info['临床诊断'] = diagnosis_match.group(1).strip()

    return info


def is_reference_range(value):
    """判断一个值是否是参考区间（而不是单位）"""
    # 参考区间格式：数字-数字 或 <数字 或 >数字
    if re.match(r'^[\d.]+-[\d.]+$', value):
        return True
    if re.match(r'^[<>][\d.]+$', value):
        return True
    if re.match(r'^[\d.]+$', value):
        return True
    return False


def extract_table_data_with_reference(text):
    """从文本中提取表格数据（项目、结果和参考区间）"""
    data = {}
    reference_ranges = {}

    # 按行分割文本
    lines = text.split('\n')

    for line in lines:
        # 匹配格式：序号 项目名称 结果 [↑↓] 参考区间 单位...
        # 例如：1 *血液酸碱度 7.252 ↓ 7.35-7.45 电极法
        # 或者：12 *红细胞计数 2.78 ↓ 4.30-5.80 10^12/L 鞘流阻抗技术
        match = re.match(r'\d+\s+\*?([^\d]+?)\s+([\d<>.]+)\s*([\u2191\u2193]?)\s+([^\s]+)', line)
        if match:
            item_name = match.group(1).strip()
            result = match.group(2).strip()
            trend = match.group(3).strip()  # ↑或↓
            reference = match.group(4).strip()

            # 清理项目名称
            item_name = item_name.replace('*', '').strip()

            # 如果有趋势标记，附加到结果上
            if trend:
                result = f"{result} {trend}"

            # 过滤掉过长的项目名（避免匹配到非项目行）
            if item_name and result and len(item_name) < 30:
                data[item_name] = result
                # 保存参考区间（只保存看起来像参考区间的值，不是单位）
                if is_reference_range(reference):
                    reference_ranges[item_name] = reference

    return data, reference_ranges


def extract_text_from_pdf(pdf_path):
    """从PDF中提取所有文本内容"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
            return text
    except Exception as e:
        print(f"读取PDF失败 {pdf_path}: {e}")
        return None


def process_pdf_file(pdf_path, report_type):
    """处理单个PDF文件"""
    text = extract_text_from_pdf(str(pdf_path))
    if not text:
        return None, None

    # 提取各类信息
    time_info = extract_time_info(text)
    patient_info = extract_patient_info(text)
    table_data, reference_ranges = extract_table_data_with_reference(text)

    # 合并所有数据
    result = {
        '文件名': pdf_path.name,
        '报告类型': report_type,
        **time_info,
        **patient_info,
        **table_data
    }

    return result, reference_ranges


def process_all_reports(base_path):
    """处理所有报告"""
    base_path = Path(base_path)
    all_data = defaultdict(list)
    all_references = defaultdict(dict)  # 存储每个报告类型的参考区间

    # 定义报告类型和文件夹名称的映射
    report_folders = {
        '血气分析': '血气分析',
        '血常规': '血常规',
        'D二聚体': 'D二聚体',
        'BNP心衰标志物': 'BNP心衰标志物',
        'PCT降钙素原': 'PCT降钙素原',
        '炎症因子': '炎症因子',
        'TBNK免疫细胞': 'TBNK免疫细胞',
        'ACT活化凝血时间': 'ACT活化凝血时间',
        '生化检验': '生化检验',
        '尿常规': '尿常规',
        '药敏试验': '药敏试验',
        '心脏超声': '心脏超声'
    }

    # 遍历所有报告类型文件夹
    for report_type, folder_name in report_folders.items():
        folder_path = base_path / folder_name
        if not folder_path.exists():
            print(f"文件夹不存在: {folder_path}")
            continue

        print(f"\n正在处理: {report_type}")
        pdf_files = list(folder_path.glob('*.pdf'))
        print(f"  找到 {len(pdf_files)} 个PDF文件")

        for pdf_file in pdf_files:
            data, references = process_pdf_file(pdf_file, report_type)
            if data:
                all_data[report_type].append(data)
                # 合并参考区间（以第一个有值的为准）
                for item, ref in references.items():
                    if item not in all_references[report_type]:
                        all_references[report_type][item] = ref

    return all_data, all_references


def create_excel_with_headers(writer, df, sheet_name, reference_ranges):
    """创建带有参考区间的表头的Excel工作表"""
    # 写入数据（不包含自定义表头）
    df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)

    # 获取工作表
    worksheet = writer.sheets[sheet_name]

    # 获取当前表头
    headers = df.columns.tolist()

    # 创建新的表头行（包含参考区间）
    new_headers = []
    for col in headers:
        if col in reference_ranges and reference_ranges[col]:
            new_headers.append(f"{col}\n(参考: {reference_ranges[col]})")
        else:
            new_headers.append(col)

    # 写入新的表头
    for col_idx, header in enumerate(new_headers, 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.value = header
        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        cell.font = Font(bold=True)

    # 调整行高以适应多行表头
    worksheet.row_dimensions[1].height = 35

    # 调整列宽
    for col_idx, col in enumerate(headers, 1):
        # 计算列宽
        header_len = len(str(new_headers[col_idx-1]))
        max_data_len = 0

        # 检查该列的数据长度
        if col in df.columns:
            for value in df[col].astype(str):
                max_data_len = max(max_data_len, len(value))

        # 设置列宽（取表头和数据的最大值，限制在50以内）
        adjusted_width = min(max(header_len, max_data_len) + 2, 50)

        # 获取列字母
        if col_idx <= 26:
            col_letter = chr(64 + col_idx)
        else:
            col_letter = chr(64 + (col_idx - 1) // 26) + chr(65 + (col_idx - 1) % 26)

        worksheet.column_dimensions[col_letter].width = adjusted_width


def create_patient_info_sheet(writer):
    """创建患者信息工作表"""
    df_patient = pd.DataFrame([
        {'项目': '姓名', '值': '王路生'},
        {'项目': '性别', '值': '男'},
        {'项目': '年龄', '值': '67岁'},
        {'项目': '住院号', '值': '1075642'},
        {'项目': '床号', '值': '15'},
        {'项目': '科室', '值': 'ICU'},
        {'项目': '入院日期', '值': '2026-04-01'},
    ])
    df_patient.to_excel(writer, sheet_name='患者信息', index=False)
    worksheet = writer.sheets['患者信息']
    worksheet.column_dimensions['A'].width = 15
    worksheet.column_dimensions['B'].width = 20


def create_excel_report(all_data, all_references, output_path):
    """创建Excel报告，每种报告类型一个工作表"""
    output_path = Path(output_path)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 创建患者信息表
        create_patient_info_sheet(writer)

        # 创建一个汇总表
        summary_data = []

        for report_type, records in all_data.items():
            if not records:
                continue

            # 创建DataFrame
            df = pd.DataFrame(records)

            # 定义固定列顺序
            fixed_cols = ['主时间', '接收时间', '采集时间', '报告时间', '报告类型', '姓名', '病历号', '性别', '年龄', '科室', '床号', '样本种类', '临床诊断']

            # 获取其他列（检测项目）
            other_cols = [c for c in df.columns if c not in fixed_cols and c != '文件名']

            # 重新排列列顺序
            cols = [c for c in fixed_cols if c in df.columns] + other_cols + ['文件名']
            df = df[[c for c in cols if c in df.columns]]

            # 按主时间排序
            if '主时间' in df.columns:
                df = df.sort_values('主时间')

            # 写入工作表（带参考区间的表头）
            sheet_name = report_type[:31]  # Excel工作表名最多31个字符
            reference_ranges = all_references.get(report_type, {})

            # 使用自定义函数创建工作表
            create_excel_with_headers(writer, df, sheet_name, reference_ranges)

            # 记录汇总信息
            time_range = '未知'
            if '主时间' in df.columns:
                valid_times = df['主时间'].dropna()
                if len(valid_times) > 0:
                    time_range = f"{valid_times.iloc[0]} ~ {valid_times.iloc[-1]}"

            summary_data.append({
                '报告类型': report_type,
                '记录数': len(records),
                '时间范围': time_range
            })

            print(f"  已写入工作表: {sheet_name} ({len(records)} 条记录)")

        # 写入汇总表
        if summary_data:
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='数据汇总', index=False)
            worksheet = writer.sheets['数据汇总']
            worksheet.column_dimensions['A'].width = 20
            worksheet.column_dimensions['B'].width = 10
            worksheet.column_dimensions['C'].width = 40

    print(f"\nExcel文件已保存: {output_path}")


def main():
    """主函数"""
    # 设置路径
    base_path = r'c:\Users\39863\Desktop\ICU\检测报告整理'

    # 生成带时间戳的文件名
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = rf'c:\Users\39863\Desktop\ICU\患者指标研究_王路生_{timestamp}.xlsx'

    print("=" * 60)
    print("ICU患者指标研究Excel创建工具 (v4)")
    print("=" * 60)

    # 处理所有报告
    print("\n开始提取数据...")
    all_data, all_references = process_all_reports(base_path)

    # 统计信息
    total_records = sum(len(records) for records in all_data.values())
    print(f"\n提取完成！共处理 {total_records} 条记录")

    # 创建Excel
    print("\n正在生成Excel文件...")
    create_excel_report(all_data, all_references, output_path)

    print("\n" + "=" * 60)
    print("处理完成！")
    print(f"输出文件: {output_path}")
    print("=" * 60)


if __name__ == '__main__':
    main()