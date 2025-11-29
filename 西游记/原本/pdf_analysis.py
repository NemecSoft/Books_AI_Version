#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
西游记PDF文件分析脚本
用于分析PDF文件的结构、提取文本内容和注释信息
"""

import os
import sys
from typing import List, Dict, Any

# 确保中文显示正常
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False    # 用来正常显示负号

def check_and_install_dependencies():
    """检查并安装必要的依赖库"""
    required_libraries = [
        ('PyPDF2', 'PyPDF2'),
        ('pdfplumber', 'pdfplumber'),
        ('matplotlib', 'matplotlib')
    ]
    
    for import_name, install_name in required_libraries:
        try:
            __import__(import_name)
            print(f"✓ {import_name} 已安装")
        except ImportError:
            print(f"✗ {import_name} 未安装，正在尝试安装...")
            try:
                import subprocess
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', install_name])
                print(f"✓ {import_name} 安装成功")
            except Exception as e:
                print(f"✗ {import_name} 安装失败: {e}")
                print("请手动安装依赖：pip install PyPDF2 pdfplumber matplotlib")

# 先检查依赖
check_and_install_dependencies()

# 然后导入库
try:
    import PyPDF2
    import pdfplumber
except ImportError:
    print("错误：请先安装必要的依赖库！")
    sys.exit(1)

def analyze_pdf_structure(pdf_path: str) -> Dict[str, Any]:
    """分析PDF文件结构"""
    print(f"\n正在分析PDF文件: {pdf_path}")
    
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            # 基本信息
            info = {
                '总页数': len(reader.pages),
                '元数据': {}
            }
            
            # 提取元数据
            if reader.metadata:
                for key, value in reader.metadata.items():
                    # 清理键名
                    clean_key = str(key).replace('/','')
                    # 尝试解码值
                    try:
                        if isinstance(value, bytes):
                            clean_value = value.decode('utf-8')
                        else:
                            clean_value = str(value)
                        info['元数据'][clean_key] = clean_value
                    except:
                        info['元数据'][clean_key] = str(value)
            
            return info
    except Exception as e:
        print(f"分析PDF结构时出错: {e}")
        return {'错误': str(e)}

def extract_sample_text(pdf_path: str, num_pages: int = 3) -> List[str]:
    """提取PDF样本文本"""
    print(f"\n正在提取PDF样本文本（前{num_pages}页）...")
    
    samples = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            pages_to_extract = min(num_pages, total_pages)
            
            for i in range(pages_to_extract):
                page = pdf.pages[i]
                text = page.extract_text()
                
                if text:
                    # 只取前200个字符作为样本
                    sample_text = text[:200] + ('...' if len(text) > 200 else '')
                    samples.append(f"第{i+1}页:\n{sample_text}\n{'-'*50}")
                else:
                    samples.append(f"第{i+1}页: 无法提取文本\n{'-'*50}")
    except Exception as e:
        print(f"提取文本时出错: {e}")
    
    return samples

def detect_annotation_patterns(pdf_path: str, start_page: int = 1, num_pages: int = 10) -> List[str]:
    """检测注释模式"""
    print(f"\n正在检测注释模式（从第{start_page}页开始的{num_pages}页）...")
    
    patterns = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            
            # 确保页码有效
            if start_page > total_pages:
                print(f"警告：起始页码{start_page}超过总页数{total_pages}")
                start_page = 1
            
            # 计算要分析的页码范围
            end_page = min(start_page + num_pages - 1, total_pages)
            
            for i in range(start_page-1, end_page):
                page = pdf.pages[i]
                text = page.extract_text()
                
                if text:
                    # 分割文本行进行分析
                    lines = text.split('\n')
                    
                    # 寻找可能的注释模式
                    for line in lines:
                        line = line.strip()
                        if not line:
                            continue
                        
                        # 检测可能的注释格式
                        # 1. 检查数字+空格+文本的模式
                        import re
                        if re.search(r'^\d+\s+[\u4e00-\u9fa5]', line):
                            patterns.append(f"第{i+1}页: {line}")
                        # 2. 检查特殊标记的注释
                        elif re.search(r'[【\[\(]注[^【\[\(]*[】\]\)]', line):
                            patterns.append(f"第{i+1}页: {line}")
                        # 3. 检查行尾的上标数字
                        elif re.search(r'[\u4e00-\u9fa5][0-9]+$', line):
                            patterns.append(f"第{i+1}页: {line}")
    except Exception as e:
        print(f"检测注释模式时出错: {e}")
    
    return patterns

def main():
    """主函数"""
    print("===== 西游记PDF文件分析工具 =====\n")
    
    # PDF文件路径
    pdf_path = os.path.join(os.path.dirname(__file__), '西游记（上下册）--中华经典小说注释系列.pdf')
    
    if not os.path.exists(pdf_path):
        print(f"错误：PDF文件不存在: {pdf_path}")
        return
    
    # 1. 分析PDF结构
    structure = analyze_pdf_structure(pdf_path)
    print("\nPDF基本信息:")
    for key, value in structure.items():
        if key == '元数据':
            print(f"  {key}:")
            for k, v in value.items():
                print(f"    {k}: {v}")
        else:
            print(f"  {key}: {value}")
    
    # 2. 提取样本文本
    samples = extract_sample_text(pdf_path)
    print("\n文本样本:")
    for sample in samples:
        print(sample)
    
    # 3. 检测注释模式
    annotations = detect_annotation_patterns(pdf_path)
    print("\n检测到的可能注释模式:")
    if annotations:
        for anno in annotations[:20]:  # 只显示前20个
            print(anno)
        if len(annotations) > 20:
            print(f"... 等共{len(annotations)}个可能的注释模式")
    else:
        print("未检测到明显的注释模式，可能需要更深入的分析")
    
    print("\n===== 分析完成 =====")

if __name__ == "__main__":
    main()
