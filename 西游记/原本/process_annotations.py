#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
西游记注释处理脚本
用于处理和优化从PDF中提取的注释内容
"""

import os
import sys
import re
import argparse
from typing import List, Dict, Any

def load_annotations_from_file(annotation_file: str) -> Dict[int, List[Dict[str, Any]]]:
    """
    从注释文件中加载注释内容
    
    参数:
        annotation_file: 注释文件路径
    
    返回:
        注释字典，格式为 {编号: [{页码: int, 内容: str}, ...]}
    """
    annotations = {}
    
    try:
        print(f"加载注释文件: {annotation_file}")
        
        with open(annotation_file, 'r', encoding='utf-8') as f:
            content = f.read()
            
            # 跳过文件头信息
            content = content.split('\n\n', 1)
            if len(content) > 1:
                content = content[1]  # 跳过前两行的标题信息
            else:
                content = content[0]
            
            # 使用正则表达式匹配注释块
            # 注释格式: "注释 X:\n  [第Y页] 内容"
            comment_blocks = re.findall(r'注释\s+(\d+):\s*\n\s*\[第(\d+)页\]\s*(.*?)(?=\n\s*注释\s+\d+:|$)', 
                                      content, re.DOTALL)
            
            for num, page, text in comment_blocks:
                num = int(num)
                page = int(page)
                text = text.strip()
                
                if num not in annotations:
                    annotations[num] = []
                
                annotations[num].append({
                    '页码': page,
                    '内容': text
                })
            
        print(f"✓ 成功加载 {len(annotations)} 个注释")
        return annotations
        
    except Exception as e:
        print(f"✗ 加载注释文件失败: {e}")
        return {}

def load_annotations_from_main_text(text_file: str) -> Dict[int, Dict[str, Any]]:
    """
    从主文本文件中提取注释
    
    参数:
        text_file: 主文本文件路径
    
    返回:
        注释字典，格式为 {编号: {页码: int, 内容: str}}
    """
    annotations = {}
    
    try:
        print(f"从主文本中提取注释: {text_file}")
        
        current_page = 1
        
        with open(text_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                
                # 更新当前页码
                page_match = re.search(r'\[第(\d+)页\]', line)
                if page_match:
                    current_page = int(page_match.group(1))
                
                # 提取注释行
                annotation_match = re.search(r'\[注释(\d+)\]:\s*(.*)', line)
                if annotation_match:
                    num = int(annotation_match.group(1))
                    text = annotation_match.group(2).strip()
                    
                    annotations[num] = {
                        '页码': current_page,
                        '内容': text
                    }
        
        print(f"✓ 从主文本中提取了 {len(annotations)} 个注释")
        return annotations
        
    except Exception as e:
        print(f"✗ 从主文本中提取注释失败: {e}")
        return {}

def merge_annotations(annotations1: Dict, annotations2: Dict) -> Dict[int, List[Dict[str, Any]]]:
    """
    合并两个注释字典
    
    参数:
        annotations1: 第一个注释字典
        annotations2: 第二个注释字典
    
    返回:
        合并后的注释字典
    """
    merged = {}
    
    # 先添加第一个字典的注释
    for num, entries in annotations1.items():
        if isinstance(entries, list):
            merged[num] = entries.copy()
        else:
            # 如果是单个字典，转换为列表
            merged[num] = [entries.copy()]
    
    # 添加第二个字典的注释（去重）
    for num, entries in annotations2.items():
        if num not in merged:
            if isinstance(entries, list):
                merged[num] = entries.copy()
            else:
                merged[num] = [entries.copy()]
        else:
            # 检查是否需要合并
            if isinstance(entries, list):
                for new_entry in entries:
                    # 检查是否重复
                    is_duplicate = False
                    for existing_entry in merged[num]:
                        if new_entry['内容'] == existing_entry['内容']:
                            is_duplicate = True
                            break
                    if not is_duplicate:
                        merged[num].append(new_entry)
            else:
                # 检查单个条目是否重复
                is_duplicate = False
                for existing_entry in merged[num]:
                    if entries['内容'] == existing_entry['内容']:
                        is_duplicate = True
                        break
                if not is_duplicate:
                    merged[num].append(entries.copy())
    
    print(f"✓ 合并后共有 {len(merged)} 个注释")
    return merged

def clean_and_optimize_annotations(annotations: Dict[int, List[Dict[str, Any]]]) -> Dict[int, List[Dict[str, Any]]]:
    """
    清理和优化注释内容
    
    参数:
        annotations: 原始注释字典
    
    返回:
        优化后的注释字典
    """
    optimized = {}
    
    for num, entries in annotations.items():
        optimized_entries = []
        
        for entry in entries:
            # 清理注释内容
            content = entry['内容'].strip()
            
            # 去除多余的空格
            content = re.sub(r'\s+', ' ', content)
            
            # 确保注释以句号结尾（如果需要）
            if content and content[-1] not in ['。', '！', '？', '；', '：', '”', '』', '）', '】']:
                content += '。'
            
            # 更新条目
            optimized_entry = {
                '页码': entry['页码'],
                '内容': content
            }
            
            optimized_entries.append(optimized_entry)
        
        optimized[num] = optimized_entries
    
    print(f"✓ 注释优化完成")
    return optimized

def save_annotations_to_file(annotations: Dict[int, List[Dict[str, Any]]], output_file: str, 
                           format_type: str = 'standard') -> bool:
    """
    将注释保存到文件
    
    参数:
        annotations: 注释字典
        output_file: 输出文件路径
        format_type: 格式类型 ('standard', 'compact', 'detailed')
    
    返回:
        是否保存成功
    """
    try:
        print(f"保存注释到文件: {output_file} (格式: {format_type})")
        
        with open(output_file, 'w', encoding='utf-8') as f:
            # 写入文件头
            f.write(f"西游记 注释整理\n")
            f.write(f"总注释数: {len(annotations)}\n")
            f.write(f"整理时间: {import_time().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("\n")
            
            # 按编号排序
            sorted_nums = sorted(annotations.keys())
            
            for num in sorted_nums:
                entries = annotations[num]
                
                if format_type == 'standard':
                    # 标准格式
                    f.write(f"注释 {num}:\n")
                    for entry in entries:
                        f.write(f"  [第{entry['页码']}页] {entry['内容']}\n")
                    f.write("\n")
                    
                elif format_type == 'compact':
                    # 紧凑格式
                    if len(entries) == 1:
                        f.write(f"{num}. {entries[0]['内容']} [第{entries[0]['页码']}页]\n")
                    else:
                        # 多个页码的注释
                        pages = ', '.join([f"第{e['页码']}页" for e in entries])
                        # 使用第一个条目的内容
                        f.write(f"{num}. {entries[0]['内容']} [{pages}]\n")
                        
                elif format_type == 'detailed':
                    # 详细格式
                    f.write(f"=== 注释 {num} ===\n")
                    for i, entry in enumerate(entries):
                        f.write(f"页码: 第{entry['页码']}页\n")
                        f.write(f"内容: {entry['内容']}\n")
                        if i < len(entries) - 1:
                            f.write("---\n")
                    f.write("\n")
        
        print(f"✓ 注释成功保存到 {output_file}")
        return True
        
    except Exception as e:
        print(f"✗ 保存注释文件失败: {e}")
        return False

def import_time():
    """导入time模块"""
    import time
    return time

def create_annotation_index(annotations: Dict[int, List[Dict[str, Any]]], output_file: str) -> bool:
    """
    创建注释索引文件
    
    参数:
        annotations: 注释字典
        output_file: 输出文件路径
    
    返回:
        是否创建成功
    """
    try:
        print(f"创建注释索引: {output_file}")
        
        # 收集所有注释内容，用于创建索引
        index_entries = []
        
        for num, entries in annotations.items():
            # 使用第一个条目的内容作为索引文本
            if entries:
                # 提取关键词（简单版本）
                content = entries[0]['内容']
                # 获取前30个字符作为摘要
                summary = content[:30] + ('...' if len(content) > 30 else '')
                
                # 收集页码信息
                pages = sorted(list(set([e['页码'] for e in entries])))
                page_str = ', '.join([str(p) for p in pages])
                
                index_entries.append({
                    '编号': num,
                    '摘要': summary,
                    '页码': page_str
                })
        
        # 按编号排序
        index_entries.sort(key=lambda x: x['编号'])
        
        # 保存索引
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("西游记 注释索引\n")
            f.write(f"总注释数: {len(index_entries)}\n")
            f.write(f"生成时间: {import_time().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("\n")
            
            # 写入表头
            f.write(f"{'编号':<6}{'页码':<15}{'摘要'}\n")
            f.write("-" * 60 + "\n")
            
            # 写入索引条目
            for entry in index_entries:
                f.write(f"{entry['编号']:<6}{entry['页码']:<15}{entry['摘要']}\n")
        
        print(f"✓ 注释索引成功创建")
        return True
        
    except Exception as e:
        print(f"✗ 创建注释索引失败: {e}")
        return False

def main():
    """主函数"""
    print("===== 西游记注释处理工具 =====\n")
    
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='西游记注释处理工具')
    parser.add_argument('--annotation-file', type=str, help='注释文件路径')
    parser.add_argument('--text-file', type=str, help='主文本文件路径')
    parser.add_argument('--output', type=str, default='西游记_注释_处理后.txt', help='输出文件路径')
    parser.add_argument('--format', type=str, choices=['standard', 'compact', 'detailed'], 
                       default='standard', help='输出格式')
    parser.add_argument('--create-index', action='store_true', help='创建注释索引')
    parser.add_argument('--index-file', type=str, default='西游记_注释索引.txt', help='索引文件路径')
    
    args = parser.parse_args()
    
    # 确保至少有一个输入文件
    if not args.annotation_file and not args.text_file:
        print("错误：请提供注释文件或主文本文件！")
        parser.print_help()
        return
    
    # 加载注释
    annotations = {}
    
    if args.annotation_file:
        annotations1 = load_annotations_from_file(args.annotation_file)
        annotations = merge_annotations(annotations, annotations1)
    
    if args.text_file:
        annotations2 = load_annotations_from_main_text(args.text_file)
        annotations = merge_annotations(annotations, annotations2)
    
    if not annotations:
        print("错误：没有成功加载任何注释！")
        return
    
    # 优化注释
    optimized_annotations = clean_and_optimize_annotations(annotations)
    
    # 保存处理后的注释
    save_annotations_to_file(optimized_annotations, args.output, args.format)
    
    # 创建索引（如果需要）
    if args.create_index:
        create_annotation_index(optimized_annotations, args.index_file)
    
    print("\n===== 注释处理完成 =====")

if __name__ == "__main__":
    main()
