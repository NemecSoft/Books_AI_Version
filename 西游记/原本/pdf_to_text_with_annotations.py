#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
西游记PDF转文本脚本（保留注释）
将PDF文件转换为文本格式，并保留和标记注释内容
"""

import os
import sys
import re
import argparse
from typing import List, Dict, Any

def check_dependencies():
    """检查并安装必要的依赖库"""
    required = ['pdfplumber', 'tqdm']
    missing = []
    
    for lib in required:
        try:
            __import__(lib)
        except ImportError:
            missing.append(lib)
    
    if missing:
        print(f"缺少必要的库: {', '.join(missing)}")
        print("正在尝试安装...")
        try:
            import subprocess
            for lib in missing:
                print(f"安装 {lib}...")
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', lib])
            print("✓ 依赖安装成功")
            return True
        except Exception as e:
            print(f"✗ 依赖安装失败: {e}")
            print("请手动安装依赖：pip install pdfplumber tqdm")
            return False
    return True

def extract_pdf_to_text(pdf_path: str, output_path: str, 
                       start_page: int = 1, end_page: int = None,
                       save_annotations: bool = True,
                       include_page_numbers: bool = True,
                       create_annotation_file: bool = False) -> Dict[str, Any]:
    # 导入time模块
    import time
    """
    将PDF转换为文本并保留注释
    
    参数:
        pdf_path: PDF文件路径
        output_path: 输出文本文件路径
        start_page: 开始页码（1-based）
        end_page: 结束页码（1-based），None表示到最后一页
        save_annotations: 是否在主文本中保存注释
        include_page_numbers: 是否在文本中包含页码标记
        create_annotation_file: 是否创建单独的注释文件
    
    返回:
        包含转换结果的字典
    """
    print(f"\n开始处理PDF文件: {os.path.basename(pdf_path)}")
    
    # 初始化结果统计
    results = {
        '成功': False,
        '总页数': 0,
        '处理页数': 0,
        '提取的注释数': 0,
        '错误': None
    }
    
    # 导入库
    try:
        import pdfplumber
        from tqdm import tqdm
    except ImportError:
        results['错误'] = "导入必要的库失败"
        return results
    
    # 检查文件是否存在
    if not os.path.exists(pdf_path):
        results['错误'] = f"PDF文件不存在: {pdf_path}"
        return results
    
    try:
        # 创建输出目录
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 打开PDF文件
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            results['总页数'] = total_pages
            
            # 确定处理范围
            if end_page is None or end_page > total_pages:
                end_page = total_pages
            
            if start_page < 1 or start_page > total_pages:
                start_page = 1
            
            if start_page > end_page:
                start_page, end_page = end_page, start_page
            
            pages_to_process = list(range(start_page-1, end_page))  # 转换为0-based索引
            results['处理页数'] = len(pages_to_process)
            
            print(f"\n处理范围: 第{start_page}页 - 第{end_page}页 (共{results['处理页数']}页)")
            
            # 准备注释存储
            annotations = {}
            
            # 开始提取
            with open(output_path, 'w', encoding='utf-8') as out_file:
                # 进度条
                with tqdm(total=len(pages_to_process), desc="处理进度") as pbar:
                    for page_idx in pages_to_process:
                        page_num = page_idx + 1  # 转回1-based页码
                        page = pdf.pages[page_idx]
                        
                        # 提取文本
                        text = page.extract_text()
                        
                        if text:
                            # 添加页码标记
                            if include_page_numbers:
                                out_file.write(f"\n\n[第{page_num}页]\n")
                            
                            # 处理文本行
                            lines = text.split('\n')
                            page_annotations = []
                            
                            for line in lines:
                                line = line.strip()
                                if not line:
                                    out_file.write('\n')
                                    continue
                                
                                # 检查是否是注释行（数字+空格+中文文本）
                                annotation_match = re.match(r'^(\d+)\s+([\u4e00-\u9fa5].*)$', line)
                                if annotation_match and save_annotations:
                                    # 提取注释编号和内容
                                    anno_num = int(annotation_match.group(1))
                                    anno_text = annotation_match.group(2)
                                    
                                    # 存储注释
                                    if anno_num not in annotations:
                                        annotations[anno_num] = []
                                    annotations[anno_num].append({
                                        '页码': page_num,
                                        '内容': anno_text
                                    })
                                    
                                    # 在文本中标记注释
                                    out_file.write(f"[注释{anno_num}]: {anno_text}\n")
                                    page_annotations.append(anno_num)
                                else:
                                    # 处理正文中的注释引用
                                    # 查找正文中的数字引用（如"菩萨12"）
                                    processed_line = re.sub(
                                        r'([\u4e00-\u9fa5])\s*(\d+)\s*',
                                        lambda m: f"{m.group(1)}[{m.group(2)}]",
                                        line
                                    )
                                    out_file.write(f"{processed_line}\n")
                            
                            # 更新进度和注释计数
                            results['提取的注释数'] += len(page_annotations)
                        else:
                            print(f"⚠ 第{page_num}页没有提取到文本")
                        
                        pbar.update(1)
            
            # 如果需要，创建单独的注释文件
            if create_annotation_file and annotations:
                annotation_file = os.path.splitext(output_path)[0] + "_注释.txt"
                with open(annotation_file, 'w', encoding='utf-8') as anno_file:
                    anno_file.write(f"西游记 PDF注释\n")
                    anno_file.write(f"总注释数: {len(annotations)}\n")
                    anno_file.write(f"生成时间: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                    
                    # 按注释编号排序
                    for num in sorted(annotations.keys()):
                        anno_file.write(f"\n注释 {num}:\n")
                        for anno in annotations[num]:
                            anno_file.write(f"  [第{anno['页码']}页] {anno['内容']}\n")
                
                print(f"\n✓ 注释文件已创建: {os.path.basename(annotation_file)}")
            
            results['成功'] = True
            print(f"\n✓ 文本提取完成")
            print(f"  输出文件: {os.path.basename(output_path)}")
            print(f"  提取的注释数: {results['提取的注释数']}")
            
    except Exception as e:
        results['错误'] = f"处理过程中出错: {str(e)}"
        print(f"\n✗ 处理失败: {str(e)}")
    
    return results

def main():
    """主函数，处理命令行参数"""
    print("===== 西游记PDF转文本工具 =====\n")
    print("此工具可以将PDF文件转换为文本格式，并保留注释内容")
    
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='西游记PDF转文本工具')
    parser.add_argument('--pdf', type=str, help='输入PDF文件路径', 
                       default='西游记（上下册）--中华经典小说注释系列.pdf')
    parser.add_argument('--output', type=str, help='输出文本文件路径',
                       default='西游记_转txt版.txt')
    parser.add_argument('--start', type=int, help='开始页码', default=1)
    parser.add_argument('--end', type=int, help='结束页码', default=None)
    parser.add_argument('--no-annotations', action='store_true', help='不保存注释')
    parser.add_argument('--no-page-numbers', action='store_true', help='不包含页码标记')
    parser.add_argument('--separate-annotations', action='store_true', help='创建单独的注释文件')
    
    args = parser.parse_args()
    
    # 转换为绝对路径
    pdf_path = os.path.abspath(args.pdf)
    output_path = os.path.abspath(args.output)
    
    # 检查依赖
    if not check_dependencies():
        return
    
    # 导入时间模块用于时间戳
    import time
    
    # 执行转换
    results = extract_pdf_to_text(
        pdf_path=pdf_path,
        output_path=output_path,
        start_page=args.start,
        end_page=args.end,
        save_annotations=not args.no_annotations,
        include_page_numbers=not args.no_page_numbers,
        create_annotation_file=args.separate_annotations
    )
    
    # 显示结果
    print("\n===== 转换结果摘要 =====")
    if results['成功']:
        print(f"✓ 转换成功")
        print(f"  处理页数: {results['处理页数']}/{results['总页数']}")
        print(f"  提取的注释数: {results['提取的注释数']}")
        print(f"  输出文件: {output_path}")
    else:
        print(f"✗ 转换失败")
        if results['错误']:
            print(f"  错误信息: {results['错误']}")
    
    print("\n===== 完成 =====")

if __name__ == "__main__":
    main()
