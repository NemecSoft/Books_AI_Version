#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import re
from PyPDF2 import PdfReader

# 设置控制台编码为UTF-8
try:
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def 清理文本(text):
    """
    清理文本中的注释符号、多余空格等干扰内容
    
    Args:
        text: 需要清理的文本
        
    Returns:
        清理后的文本
    """
    # 去除行尾的注释数字（如：  29  ）
    text = re.sub(r'\s+\d+\s*$', '', text, flags=re.MULTILINE)
    
    # 去除行中的注释数字（如："慢慢的叙阔  29 ，"）
    text = re.sub(r'(\S)\s+(\d+)\s+', '\1 ', text)
    
    # 去除多余的空白字符（连续的空格、制表符等）
    text = re.sub(r'\s+', ' ', text)
    
    # 去除行首行尾的空白
    lines = [line.strip() for line in text.split('\n')]
    text = '\n'.join(lines)
    
    # 修复因PDF提取导致的不必要的换行
    # 中文标点后如果不是换行，保留换行；中文标点后如果是换行，可能需要合并
    # 这是一个启发式方法，可能需要根据实际文本调整
    paragraphs = []
    current_paragraph = []
    for line in text.split('\n'):
        if line:
            current_paragraph.append(line)
            # 如果行以句号、问号、感叹号、冒号、分号结尾，可能是一个段落的结束
            if line.endswith(('。', '？', '！', '：', '；', '”', '）')):
                paragraphs.append(' '.join(current_paragraph))
                current_paragraph = []
    # 处理最后一个段落
    if current_paragraph:
        paragraphs.append(' '.join(current_paragraph))
    
    # 重新组合段落，保留段落间的换行
    cleaned_text = '\n\n'.join(paragraphs)
    
    return cleaned_text


def pdf_to_txt(pdf_path, txt_path):
    """
    将PDF文件转换为TXT文件，增加文本清理功能
    
    Args:
        pdf_path: PDF文件路径
        txt_path: 输出的TXT文件路径
    """
    try:
        print(f"开始处理PDF文件: {pdf_path}")
        
        # 读取PDF文件
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            total_pages = len(reader.pages)
            print(f"PDF总页数: {total_pages}")
            
            # 提取文本
            all_text = []
            for page_num in range(total_pages):
                page = reader.pages[page_num]
                text = page.extract_text()
                if text:
                    # 对每一页的文本进行清理
                    cleaned_text = 清理文本(text)
                    all_text.append(cleaned_text)
                
                # 显示进度
                if (page_num + 1) % 10 == 0 or page_num + 1 == total_pages:
                    print(f"已处理: {page_num + 1}/{total_pages} 页")
            
            # 合并所有文本
            combined_text = '\n\n'.join(all_text)
            
            # 再次进行整体清理，确保格式一致
            final_text = 清理文本(combined_text)
            
            # 保存到TXT文件
            with open(txt_path, 'w', encoding='utf-8') as output_file:
                output_file.write(final_text)
            
            print(f"转换完成! TXT文件已保存至: {txt_path}")
            print(f"提取的文本长度: {len(final_text)} 字符")
            
    except Exception as e:
        print(f"转换过程中发生错误: {str(e)}")
        raise


if __name__ == "__main__":
    # 设置PDF和TXT文件路径
    pdf_file = "d:\\AI\\books\\西游记\\原本\\西游记（上下册）--中华经典小说注释系列.pdf"
    txt_file = "d:\\AI\\books\\西游记\\西游记_转txt版.txt"
    
    # 确保输出目录存在
    output_dir = os.path.dirname(txt_file)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 执行转换
    pdf_to_txt(pdf_file, txt_file)
