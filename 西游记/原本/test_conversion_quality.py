#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
西游记PDF转文本效果测试脚本
用于评估PDF转文本的质量，包括完整性、准确性和格式保留等
"""

import os
import sys
import re
import argparse
from typing import Dict, List, Any

def load_text_file(file_path: str) -> str:
    """
    加载文本文件内容
    
    参数:
        file_path: 文件路径
    
    返回:
        文件内容字符串
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"✗ 加载文件 {file_path} 失败: {e}")
        return ''

def count_pages(text: str) -> int:
    """
    统计文本中的页数
    
    参数:
        text: 文本内容
    
    返回:
        页数
    """
    # 匹配页码标记 [第X页]
    page_matches = re.findall(r'\[第(\d+)页\]', text)
    page_numbers = [int(num) for num in page_matches] if page_matches else []
    
    if not page_numbers:
        return 0
    
    # 返回最大页码（粗略估计）
    return max(page_numbers)

def count_annotations(text: str) -> Dict[str, int]:
    """
    统计文本中的注释
    
    参数:
        text: 文本内容
    
    返回:
        注释统计信息
    """
    # 匹配注释标记 [注释X]:
    annotation_matches = re.findall(r'\[注释(\d+)\]:', text)
    
    # 匹配独立注释模式 数字+空格+文本
    standalone_annotations = re.findall(r'^\d+\s+[^\d\s]', text, re.MULTILINE)
    
    return {
        '标记注释数': len(annotation_matches),
        '独立注释数': len(standalone_annotations),
        '总注释数': len(annotation_matches) + len(standalone_annotations)
    }

def analyze_text_structure(text: str) -> Dict[str, Any]:
    """
    分析文本结构
    
    参数:
        text: 文本内容
    
    返回:
        文本结构分析结果
    """
    lines = text.split('\n')
    non_empty_lines = [line.strip() for line in lines if line.strip()]
    
    # 计算行长度统计
    line_lengths = [len(line) for line in non_empty_lines]
    avg_line_length = sum(line_lengths) / len(line_lengths) if line_lengths else 0
    max_line_length = max(line_lengths) if line_lengths else 0
    min_line_length = min(line_lengths) if line_lengths else 0
    
    # 检查是否有明显的乱码
    # 这里使用简单的启发式方法：统计包含连续非汉字字符的行
    potential_gibberish_lines = []
    for line in non_empty_lines:
        # 匹配连续的非汉字、非常见标点符号的字符序列
        if re.search(r'[^\u4e00-\u9fa5\s，。！？；："\'（）【】《》、,.!?;:"]{10,}', line):
            potential_gibberish_lines.append(line)
    
    # 检查标题结构
    has_chapter_markers = any(re.search(r'第[一二三四五六七八九十百千]+[回卷]', line) for line in non_empty_lines)
    
    return {
        '总行数': len(lines),
        '非空行数': len(non_empty_lines),
        '平均行长度': round(avg_line_length, 2),
        '最大行长度': max_line_length,
        '最小行长度': min_line_length,
        '疑似乱码行数': len(potential_gibberish_lines),
        '包含章节标记': has_chapter_markers
    }

def evaluate_conversion_quality(text_content: str, original_pdf_pages: int = 1632) -> Dict[str, Any]:
    """
    评估转换质量
    
    参数:
        text_content: 转换后的文本内容
        original_pdf_pages: 原始PDF的页数（默认1632页）
    
    返回:
        转换质量评估结果
    """
    # 基本统计
    page_count = count_pages(text_content)
    annotation_stats = count_annotations(text_content)
    structure_stats = analyze_text_structure(text_content)
    
    # 计算页面覆盖率
    page_coverage = min(100, (page_count / original_pdf_pages) * 100) if original_pdf_pages > 0 else 0
    
    # 计算完整性得分
    completeness_score = min(100, page_coverage)
    
    # 计算质量得分
    # 基于页面覆盖率、注释提取、文本结构等综合评估
    quality_score = completeness_score
    
    # 乱码惩罚
    if structure_stats['疑似乱码行数'] > 0:
        # 每10行乱码扣1分
        quality_score = max(0, quality_score - (structure_stats['疑似乱码行数'] // 10))
    
    # 章节标记加分
    if structure_stats['包含章节标记']:
        quality_score = min(100, quality_score + 5)
    
    return {
        '页面覆盖率': f"{page_count}/{original_pdf_pages} ({page_coverage:.1f}%)",
        '页面数': page_count,
        '完整性得分': round(completeness_score, 1),
        '质量得分': round(quality_score, 1),
        '注释统计': annotation_stats,
        '文本结构': structure_stats
    }

def compare_with_original_sample(text_content: str, sample_file: str = None) -> Dict[str, Any]:
    """
    与原始样本文本比较（如果提供）
    
    参数:
        text_content: 转换后的文本内容
        sample_file: 原始样本文件路径
    
    返回:
        比较结果
    """
    if not sample_file or not os.path.exists(sample_file):
        return {
            '比较结果': '未提供原始样本或文件不存在',
            '可比较性': False
        }
    
    # 加载样本文件
    sample_content = load_text_file(sample_file)
    if not sample_content:
        return {
            '比较结果': '无法加载原始样本',
            '可比较性': False
        }
    
    # 简单比较：检查样本中的关键词是否在转换文本中存在
    # 提取样本中的一些关键词（简单版本）
    sample_words = re.findall(r'[\u4e00-\u9fa5]{2,}', sample_content)[:20]  # 取前20个关键词
    
    if not sample_words:
        return {
            '比较结果': '样本中没有可提取的关键词',
            '可比较性': False
        }
    
    # 计算匹配率
    matched_words = [word for word in sample_words if word in text_content]
    match_rate = (len(matched_words) / len(sample_words)) * 100 if sample_words else 0
    
    return {
        '样本关键词数': len(sample_words),
        '匹配关键词数': len(matched_words),
        '匹配率': round(match_rate, 1),
        '可比较性': True,
        '比较结果': f"关键词匹配率: {match_rate:.1f}% ({len(matched_words)}/{len(sample_words)})"
    }

def generate_quality_report(quality_stats: Dict[str, Any], comparison_results: Dict[str, Any]) -> str:
    """
    生成质量报告
    
    参数:
        quality_stats: 质量统计数据
        comparison_results: 比较结果
    
    返回:
        格式化的质量报告
    """
    report_lines = [
        "===== 西游记PDF转文本质量评估报告 =====\n",
        "【基本信息】",
        f"页面覆盖率: {quality_stats['页面覆盖率']}",
        f"完整性得分: {quality_stats['完整性得分']}/100",
        f"整体质量得分: {quality_stats['质量得分']}/100\n",
        
        "【注释统计】",
        f"标记注释数: {quality_stats['注释统计']['标记注释数']}",
        f"独立注释数: {quality_stats['注释统计']['独立注释数']}",
        f"总注释数: {quality_stats['注释统计']['总注释数']}\n",
        
        "【文本结构分析】",
        f"总行数: {quality_stats['文本结构']['总行数']}",
        f"非空行数: {quality_stats['文本结构']['非空行数']}",
        f"平均行长度: {quality_stats['文本结构']['平均行长度']} 字符",
        f"最大行长度: {quality_stats['文本结构']['最大行长度']} 字符",
        f"最小行长度: {quality_stats['文本结构']['最小行长度']} 字符",
        f"疑似乱码行数: {quality_stats['文本结构']['疑似乱码行数']}",
        f"包含章节标记: {'是' if quality_stats['文本结构']['包含章节标记'] else '否'}\n"
    ]
    
    # 添加比较结果
    if comparison_results['可比较性']:
        report_lines.extend([
            "【与原始样本比较】",
            f"关键词匹配率: {comparison_results['匹配率']}%",
            f"匹配关键词数: {comparison_results['匹配关键词数']}/{comparison_results['样本关键词数']}\n"
        ])
    else:
        report_lines.extend([
            "【与原始样本比较】",
            f"{comparison_results['比较结果']}\n"
        ])
    
    # 添加评估总结
    quality_score = quality_stats['质量得分']
    if quality_score >= 90:
        evaluation = "优秀: 转换质量非常好，文本完整性高，格式保留良好。"
    elif quality_score >= 80:
        evaluation = "良好: 转换质量较好，文本基本完整，可能存在少量格式问题。"
    elif quality_score >= 70:
        evaluation = "一般: 转换质量基本可用，文本大部分完整，但存在一些格式或内容问题。"
    elif quality_score >= 60:
        evaluation = "较差: 转换质量有待提高，文本可能有较多缺失或格式混乱。"
    else:
        evaluation = "很差: 转换质量不佳，文本缺失严重或格式严重混乱。"
    
    report_lines.extend([
        "【评估总结】",
        evaluation,
        "\n" + "建议："
    ])
    
    # 添加改进建议
    suggestions = []
    
    if quality_stats['完整性得分'] < 80:
        suggestions.append("✓ 增加转换的页面范围，确保覆盖完整内容")
    
    if quality_stats['文本结构']['疑似乱码行数'] > 0:
        suggestions.append("✓ 优化字符编码处理，减少乱码产生")
    
    if quality_stats['注释统计']['总注释数'] == 0:
        suggestions.append("✓ 增强注释提取功能，确保注释内容被正确识别")
    
    if not quality_stats['文本结构']['包含章节标记']:
        suggestions.append("✓ 优化章节识别，保留原文的结构层次")
    
    if not comparison_results['可比较性']:
        suggestions.append("✓ 建议使用原始样本进行更精确的比较测试")
    
    if suggestions:
        report_lines.extend(suggestions)
    else:
        report_lines.append("✓ 当前转换质量已满足基本需求，如有特殊要求可进一步优化。")
    
    return '\n'.join(report_lines)

def main():
    """
    主函数
    """
    print("===== 西游记PDF转文本质量测试工具 =====\n")
    
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='西游记PDF转文本质量测试工具')
    parser.add_argument('--text-file', type=str, required=True, help='转换后的文本文件路径')
    parser.add_argument('--sample-file', type=str, help='原始样本文本文件路径（可选）')
    parser.add_argument('--output', type=str, default='转换质量报告.txt', help='质量报告输出路径')
    parser.add_argument('--original-pages', type=int, default=1632, help='原始PDF的页数')
    
    args = parser.parse_args()
    
    # 加载文本文件
    print(f"加载文本文件: {args.text_file}")
    text_content = load_text_file(args.text_file)
    
    if not text_content:
        print("✗ 无法加载文本文件，程序终止。")
        return
    
    print("✓ 文本文件加载成功")
    
    # 评估转换质量
    print("\n正在分析转换质量...")
    quality_stats = evaluate_conversion_quality(text_content, args.original_pages)
    
    # 与原始样本比较（如果提供）
    comparison_results = compare_with_original_sample(text_content, args.sample_file)
    
    # 生成质量报告
    report = generate_quality_report(quality_stats, comparison_results)
    
    # 输出报告到控制台
    print("\n" + report)
    
    # 保存报告到文件
    try:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"\n✓ 质量报告已保存到: {args.output}")
    except Exception as e:
        print(f"✗ 保存质量报告失败: {e}")
    
    print("\n===== 质量测试完成 =====")

if __name__ == "__main__":
    main()
