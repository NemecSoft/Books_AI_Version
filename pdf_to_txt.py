#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF文件转TXT工具
用于将PDF文件转换为文本格式
"""

import os
import sys
import time
import pdfplumber


def pdf_to_txt(pdf_path, txt_path=None, start_page=0, end_page=None):
    """
    将PDF文件转换为TXT文件
    
    参数:
        pdf_path (str): PDF文件路径
        txt_path (str): 输出TXT文件路径，如果不提供则在原目录生成同名TXT
        start_page (int): 开始页码（从0开始）
        end_page (int): 结束页码，如果不提供则转换所有页面
    
    返回:
        str: 生成的TXT文件路径
    """
    # 检查输入文件是否存在
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
    
    # 如果未提供输出路径，在原目录生成同名TXT
    if txt_path is None:
        base_name = os.path.splitext(pdf_path)[0]
        txt_path = f"{base_name}.txt"
    
    # 确保输出目录存在
    output_dir = os.path.dirname(txt_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    print(f"开始将PDF转换为TXT...")
    print(f"输入文件: {pdf_path}")
    print(f"输出文件: {txt_path}")
    
    start_time = time.time()
    total_text = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            
            # 确定要处理的页面范围
            if end_page is None or end_page > total_pages:
                end_page = total_pages
            
            if start_page >= total_pages:
                raise ValueError(f"开始页码 {start_page} 超过总页数 {total_pages}")
            
            if start_page < 0:
                start_page = 0
            
            print(f"总页数: {total_pages}")
            print(f"处理页码范围: {start_page + 1} - {end_page}")
            
            # 逐页提取文本
            for page_num in range(start_page, end_page):
                page = pdf.pages[page_num]
                page_text = page.extract_text()
                
                if page_text:
                    total_text.append(page_text)
                
                # 显示进度
                progress = (page_num - start_page + 1) / (end_page - start_page) * 100
                elapsed = time.time() - start_time
                if progress > 0:
                    remaining = elapsed / progress * (100 - progress)
                else:
                    remaining = 0
                
                print(f"进度: {progress:.1f}% | 已处理: {page_num + 1}/{end_page}页 | "
                      f"用时: {elapsed:.1f}秒 | 预计剩余: {remaining:.1f}秒", end="\r")
        
        # 将所有文本写入TXT文件
        with open(txt_path, 'w', encoding='utf-8') as txt_file:
            txt_file.write('\n\n'.join(total_text))
        
        total_time = time.time() - start_time
        print(f"\n转换完成！")
        print(f"输出文件已保存到: {txt_path}")
        print(f"总用时: {total_time:.1f}秒")
        print(f"平均每页处理时间: {total_time / (end_page - start_page):.3f}秒")
        
        return txt_path
        
    except Exception as e:
        print(f"转换过程中发生错误: {str(e)}")
        raise


if __name__ == "__main__":
    # 设置命令行输出编码为UTF-8
    sys.stdout.reconfigure(encoding='utf-8')
    
    # 示例使用
    pdf_file = r"d:\AI\books\二十四史\《二十四史全译 三国志 第二册》主编：许嘉璐.pdf"
    txt_file = r"d:\AI\books\二十四史\《二十四史全译 三国志 第二册》主编：许嘉璐.txt"
    
    try:
        pdf_to_txt(pdf_file, txt_file)
    except Exception as e:
        print(f"操作失败: {e}")
        sys.exit(1)
