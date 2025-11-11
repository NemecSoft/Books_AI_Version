#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
使用OCR技术将扫描版PDF转换为TXT工具
需要系统上安装Tesseract OCR引擎
"""

import os
import sys
import time
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import io


def check_tesseract_installed():
    """
    检查Tesseract OCR是否已安装并可用
    
    返回:
        tuple: (是否可用, 错误消息)
    """
    try:
        # 尝试运行tesseract命令获取版本信息
        import subprocess
        subprocess.run(['tesseract', '--version'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True, None
    except FileNotFoundError:
        return False, "Tesseract OCR引擎未安装或不在系统PATH中"
    except subprocess.CalledProcessError:
        return False, "Tesseract OCR引擎已安装但无法正常运行"
    except Exception as e:
        return False, f"检查Tesseract时发生未知错误: {str(e)}"


def pdf_to_txt_ocr(pdf_path, txt_path=None, start_page=0, end_page=None, lang='chi_sim'):
    """
    使用OCR技术将PDF文件转换为TXT文件
    
    参数:
        pdf_path (str): PDF文件路径
        txt_path (str): 输出TXT文件路径，如果不提供则在原目录生成同名TXT
        start_page (int): 开始页码（从0开始）
        end_page (int): 结束页码，如果不提供则转换所有页面
        lang (str): OCR语言，默认为简体中文
    
    返回:
        str: 生成的TXT文件路径
    """
    # 检查Tesseract是否可用
    is_installed, error_msg = check_tesseract_installed()
    if not is_installed:
        raise EnvironmentError(f"OCR转换失败: {error_msg}\n请安装Tesseract OCR引擎并添加到系统PATH中")
    
    # 检查输入文件是否存在
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
    
    # 如果未提供输出路径，在原目录生成同名TXT
    if txt_path is None:
        base_name = os.path.splitext(pdf_path)[0]
        txt_path = f"{base_name}_ocr.txt"
    
    # 确保输出目录存在
    output_dir = os.path.dirname(txt_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    print(f"开始使用OCR技术将PDF转换为TXT...")
    print(f"输入文件: {pdf_path}")
    print(f"输出文件: {txt_path}")
    print(f"使用OCR语言: {lang}")
    
    start_time = time.time()
    total_text = []
    
    try:
        # 打开PDF文件
        with fitz.open(pdf_path) as doc:
            total_pages = len(doc)
            
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
                page = doc[page_num]
                
                # 尝试直接提取文本（对于非扫描PDF）
                direct_text = page.get_text()
                
                if direct_text.strip():
                    # 如果能直接提取到文本，使用直接提取的内容
                    page_text = direct_text
                else:
                    # 否则使用OCR处理
                    # 将页面转换为图像
                    pix = page.get_pixmap(dpi=300)  # 高DPI以提高OCR质量
                    img = Image.open(io.BytesIO(pix.tobytes()))
                    
                    # 使用Tesseract进行OCR
                    page_text = pytesseract.image_to_string(img, lang=lang)
                
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
        
        # 检查文件大小
        file_size = os.path.getsize(txt_path)
        print(f"生成的TXT文件大小: {file_size} 字节")
        
        return txt_path
        
    except EnvironmentError as e:
        print(f"\n{str(e)}")
        print("\nTesseract OCR安装指南:")
        print("1. Windows用户: 从 https://github.com/UB-Mannheim/tesseract/wiki 下载安装包")
        print("2. 安装时记录安装路径，如 C:\\Program Files\\Tesseract-OCR")
        print("3. 将安装目录添加到系统环境变量PATH中")
        print("4. 或在代码中设置pytesseract.pytesseract.tesseract_cmd = '安装路径\\tesseract.exe'")
        raise
    except Exception as e:
        print(f"\n转换过程中发生错误: {str(e)}")
        raise


if __name__ == "__main__":
    # 设置命令行输出编码为UTF-8
    sys.stdout.reconfigure(encoding='utf-8')
    
    # 示例使用
    pdf_file = r"d:\AI\books\二十四史\《二十四史全译 三国志 第二册》主编：许嘉璐.pdf"
    txt_file = r"d:\AI\books\二十四史\《二十四史全译 三国志 第二册》主编：许嘉璐_ocr.txt"
    
    try:
        # 如果Tesseract未安装，可以先检测并提供替代方案
        is_installed, error_msg = check_tesseract_installed()
        if not is_installed:
            print(f"警告: {error_msg}")
            print("将尝试使用PDF文本提取作为替代方案...")
            
            # 如果Tesseract不可用，使用PyMuPDF直接提取文本
            print("使用PyMuPDF直接提取文本...")
            with fitz.open(pdf_file) as doc:
                all_text = []
                for page in doc:
                    all_text.append(page.get_text())
                
            with open(txt_file, 'w', encoding='utf-8') as f:
                f.write('\n\n'.join(all_text))
            
            print(f"已将文本提取到: {txt_file}")
            print("注意：如果PDF是扫描版，您需要安装Tesseract OCR以获得更好的转换效果")
        else:
            # 如果Tesseract可用，使用OCR转换
            pdf_to_txt_ocr(pdf_file, txt_file)
    except Exception as e:
        print(f"操作失败: {e}")
        sys.exit(1)
