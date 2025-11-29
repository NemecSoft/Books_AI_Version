#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
西游记PDF文件详细分析脚本
用于更深入地分析PDF文件的内容、结构和注释格式
"""

import os
import sys
import re
from typing import List, Dict, Any

def check_dependencies():
    """检查必要的依赖库"""
    required = ['pdfplumber', 'PyPDF2']
    missing = []
    
    for lib in required:
        try:
            __import__(lib)
        except ImportError:
            missing.append(lib)
    
    if missing:
        print(f"缺少必要的库: {', '.join(missing)}")
        print("请运行: pip install pdfplumber PyPDF2")
        return False
    return True

def extract_pages_content(pdf_path: str, pages_to_extract: List[int]) -> Dict[int, str]:
    """提取指定页码的内容"""
    print(f"\n正在提取PDF指定页面的内容...")
    
    content = {}
    try:
        import pdfplumber
        
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            
            for page_num in pages_to_extract:
                # 页码转换（用户输入是1-based，Python是0-based）
                if 1 <= page_num <= total_pages:
                    page = pdf.pages[page_num - 1]
                    text = page.extract_text()
                    if text:
                        content[page_num] = text
                        print(f"✓ 已提取第{page_num}页")
                    else:
                        print(f"⚠ 第{page_num}页没有提取到文本")
                else:
                    print(f"⚠ 页码{page_num}超出范围（总页数：{total_pages}）")
    except Exception as e:
        print(f"提取页面内容时出错: {e}")
    
    return content

def search_annotation_patterns(text: str) -> List[str]:
    """在文本中搜索可能的注释模式"""
    patterns = []
    
    # 分割文本行为行
    lines = text.split('\n')
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        
        # 模式1: 数字 + 空格 + 中文文本（可能是注释编号）
        if re.search(r'^\s*\d+\s+[\u4e00-\u9fa5]', line):
            patterns.append(f"模式1 (数字+空格+文本): 第{i+1}行: {line}")
        
        # 模式2: 行中的上标数字（如正文中标注的注释引用）
        if re.search(r'[\u4e00-\u9fa5][0-9]+(?![0-9]|[a-zA-Z])', line):
            patterns.append(f"模式2 (上标数字): 第{i+1}行: {line}")
        
        # 模式3: 括号中的注释标记
        if re.search(r'[【\[\(]注[^【\[\(]*[】\]\)]', line):
            patterns.append(f"模式3 (括号注释): 第{i+1}行: {line}")
        
        # 模式4: 特殊格式的行（如多个空格分隔的内容）
        if re.search(r'\s{4,}', line):
            patterns.append(f"模式4 (多空格): 第{i+1}行: {line}")
        
        # 模式5: 明显的标题行
        if re.search(r'^[\u4e00-\u9fa5]{2,8}\s*回[\u4e00-\u9fa5]+', line):
            patterns.append(f"模式5 (回目): 第{i+1}行: {line}")
    
    return patterns

def find_annotation_sections(content: Dict[int, str]) -> Dict[str, List[str]]:
    """查找可能的注释部分"""
    print("\n正在寻找可能的注释部分...")
    
    # 检查是否有专门的注释页
    annotation_sections = {'正文注释': [], '独立注释页': []}
    
    for page_num, text in content.items():
        lines = text.split('\n')
        
        # 检查是否有大量符合注释模式的行
        annotation_count = 0
        potential_annotation_page = False
        
        for line in lines:
            line = line.strip()
            if re.search(r'^\s*\d+\s+[\u4e00-\u9fa5]', line):
                annotation_count += 1
                annotation_sections['正文注释'].append(f"第{page_num}页: {line}")
        
        # 如果一页中有超过10个可能的注释行，标记为可能的独立注释页
        if annotation_count >= 10:
            potential_annotation_page = True
            annotation_sections['独立注释页'].append(f"第{page_num}页 (发现{annotation_count}个可能的注释条目)")
    
    return annotation_sections

def analyze_text_structure(text: str) -> Dict[str, Any]:
    """分析文本结构"""
    print("\n正在分析文本结构...")
    
    # 统计各种元素
    stats = {
        '总行数': 0,
        '空行数': 0,
        '回目数': 0,
        '可能的注释引用数': 0,
        '平均行长度': 0
    }
    
    lines = text.split('\n')
    stats['总行数'] = len(lines)
    
    total_chars = 0
    non_empty_lines = 0
    
    for line in lines:
        # 统计空行
        if not line.strip():
            stats['空行数'] += 1
        else:
            total_chars += len(line)
            non_empty_lines += 1
            
            # 统计回目
            if re.search(r'^[\u4e00-\u9fa5]{2,8}\s*回[\u4e00-\u9fa5]+', line.strip()):
                stats['回目数'] += 1
            
            # 统计可能的注释引用
            if re.search(r'[\u4e00-\u9fa5][0-9]+(?![0-9]|[a-zA-Z])', line):
                stats['可能的注释引用数'] += 1
    
    # 计算平均行长度
    if non_empty_lines > 0:
        stats['平均行长度'] = total_chars / non_empty_lines
    
    return stats

def extract_full_sample(text: str, max_lines: int = 20) -> str:
    """提取完整的文本样本"""
    lines = text.split('\n')
    return '\n'.join(lines[:max_lines])

def main():
    """主函数"""
    print("===== 西游记PDF文件详细分析工具 =====\n")
    
    # 检查依赖
    if not check_dependencies():
        return
    
    # PDF文件路径
    pdf_path = os.path.join(os.path.dirname(__file__), '西游记（上下册）--中华经典小说注释系列.pdf')
    
    if not os.path.exists(pdf_path):
        print(f"错误：PDF文件不存在: {pdf_path}")
        return
    
    # 获取PDF总页数
    total_pages = 0
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
        print(f"PDF文件包含 {total_pages} 页")
    except Exception as e:
        print(f"获取页数时出错: {e}")
        return
    
    # 提取不同位置的页面进行分析
    # 我们将提取：前几页、中间页、随机页，以全面了解PDF结构
    pages_to_extract = []
    
    # 前几页（可能包含目录）
    pages_to_extract.extend([1, 2, 3, 4, 5])
    
    # 中间页（可能包含正文和注释）
    if total_pages > 20:
        middle = total_pages // 2
        pages_to_extract.extend([middle-1, middle, middle+1])
    
    # 后面的页（可能有独立的注释部分）
    if total_pages > 10:
        pages_to_extract.extend([total_pages-5, total_pages-3, total_pages-1])
    
    # 去重并排序
    pages_to_extract = sorted(list(set(pages_to_extract)))
    
    # 提取页面内容
    content = extract_pages_content(pdf_path, pages_to_extract)
    
    if not content:
        print("\n错误：没有提取到任何页面内容！")
        return
    
    # 分析每一页的内容
    for page_num, text in content.items():
        print(f"\n{'='*60}")
        print(f"第{page_num}页内容分析")
        print(f"{'='*60}")
        
        # 显示样本内容
        print("\n页面内容样本（前10行）:")
        sample_lines = text.split('\n')[:10]
        for i, line in enumerate(sample_lines):
            print(f"[{i+1:2d}] {line}")
        
        # 搜索可能的注释模式
        patterns = search_annotation_patterns(text)
        if patterns:
            print(f"\n发现的可能注释模式 ({len(patterns)}个):")
            for pattern in patterns[:10]:  # 只显示前10个
                print(f"  {pattern}")
            if len(patterns) > 10:
                print(f"  ... 等{len(patterns)}个模式")
        else:
            print("\n未发现明显的注释模式")
        
        # 分析文本结构
        stats = analyze_text_structure(text)
        print("\n文本结构统计:")
        for key, value in stats.items():
            print(f"  {key}: {value}")
    
    # 查找可能的注释部分
    annotation_sections = find_annotation_sections(content)
    
    print("\n" + "="*60)
    print("注释分析总结")
    print("="*60)
    
    if annotation_sections['独立注释页']:
        print("\n可能的独立注释页:")
        for section in annotation_sections['独立注释页']:
            print(f"  {section}")
    else:
        print("\n未发现明显的独立注释页")
    
    if annotation_sections['正文注释']:
        print(f"\n发现的可能注释条目 ({len(annotation_sections['正文注释'])}个):")
        for anno in annotation_sections['正文注释'][:15]:  # 只显示前15个
            print(f"  {anno}")
        if len(annotation_sections['正文注释']) > 15:
            print(f"  ... 等{len(annotation_sections['正文注释'])}个注释条目")
    else:
        print("\n未发现明显的注释条目")
    
    print("\n===== 分析完成 =====")
    print("请根据分析结果确定PDF的注释格式，以便进行下一步处理。")

if __name__ == "__main__":
    main()
