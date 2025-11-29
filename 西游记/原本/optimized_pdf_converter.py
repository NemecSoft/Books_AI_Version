#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
优化版西游记PDF转文本工具
功能：更高效地将PDF转换为文本，更好地保留格式和注释
"""

import os
import sys
import re
import argparse
import time
from typing import List, Dict, Tuple, Any, Optional

# 尝试导入必要的库
try:
    import fitz  # PyMuPDF
    import pdfplumber
except ImportError:
    print("正在安装必要的库...")
    os.system(f"{sys.executable} -m pip install PyMuPDF pdfplumber")
    
    try:
        import fitz
        import pdfplumber
        print("库安装成功！")
    except ImportError:
        print("✗ 库安装失败，请手动安装：pip install PyMuPDF pdfplumber")
        sys.exit(1)

class PDFConverter:
    """PDF转换器类"""
    
    def __init__(self, pdf_path: str, verbose: bool = False):
        """
        初始化PDF转换器
        
        参数:
            pdf_path: PDF文件路径
            verbose: 是否显示详细信息
        """
        self.pdf_path = pdf_path
        self.verbose = verbose
        self.total_pages = 0
        
        # 检查文件是否存在
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
        
        # 获取PDF总页数
        try:
            with fitz.open(pdf_path) as doc:
                self.total_pages = len(doc)
            print(f"✓ PDF文件加载成功，共 {self.total_pages} 页")
        except Exception as e:
            raise RuntimeError(f"加载PDF文件失败: {e}")
    
    def extract_page_with_fitz(self, page_num: int) -> str:
        """
        使用PyMuPDF提取页面文本
        
        参数:
            page_num: 页码（从0开始）
        
        返回:
            提取的文本
        """
        try:
            with fitz.open(self.pdf_path) as doc:
                if page_num < 0 or page_num >= len(doc):
                    return ""
                
                page = doc[page_num]
                text = page.get_text("text")  # 使用text模式保留基本格式
                return text
        except Exception as e:
            if self.verbose:
                print(f"使用PyMuPDF提取第{page_num + 1}页失败: {e}")
            return ""
    
    def extract_page_with_pdfplumber(self, page_num: int) -> str:
        """
        使用pdfplumber提取页面文本
        
        参数:
            page_num: 页码（从0开始）
        
        返回:
            提取的文本
        """
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                if page_num < 0 or page_num >= len(pdf.pages):
                    return ""
                
                page = pdf.pages[page_num]
                text = page.extract_text() or ""
                return text
        except Exception as e:
            if self.verbose:
                print(f"使用pdfplumber提取第{page_num + 1}页失败: {e}")
            return ""
    
    def extract_page_text(self, page_num: int) -> str:
        """
        提取页面文本，尝试两种方法以获得最佳结果
        
        参数:
            page_num: 页码（从1开始）
        
        返回:
            提取的文本
        """
        # 转换为0-based索引
        idx = page_num - 1
        
        # 首先尝试PyMuPDF
        text = self.extract_page_with_fitz(idx)
        
        # 如果PyMuPDF提取失败或结果太少，尝试pdfplumber
        if not text or len(text.strip()) < 10:
            text = self.extract_page_with_pdfplumber(idx)
        
        return text
    
    def detect_annotations(self, text: str) -> List[Tuple[int, str]]:
        """
        检测文本中的注释
        
        参数:
            text: 页面文本
        
        返回:
            注释列表 [(编号, 内容), ...]
        """
        annotations = []
        
        # 模式1: 数字 + 空格 + 非数字文本（独立注释）
        pattern1 = re.compile(r'^(\d+)\s+([^\d\s].*?)$', re.MULTILINE)
        matches = pattern1.findall(text)
        for match in matches:
            try:
                num = int(match[0])
                content = match[1].strip()
                if content and len(content) > 2:  # 确保注释内容合理
                    annotations.append((num, content))
            except ValueError:
                pass
        
        # 模式2: [注释X]: 格式
        pattern2 = re.compile(r'\[注释(\d+)\]:\s*(.*?)(?=\[注释\d+\]:|$)', re.DOTALL)
        matches = pattern2.findall(text)
        for match in matches:
            try:
                num = int(match[0])
                content = match[1].strip()
                if content:
                    annotations.append((num, content))
            except ValueError:
                pass
        
        # 模式3: 可能的注释页格式
        pattern3 = re.compile(r'^(\d{1,3})\s+[一二三四五六七八九十百千]+、\s+(.*?)$', re.MULTILINE)
        matches = pattern3.findall(text)
        for match in matches:
            try:
                num = int(match[0])
                content = match[1].strip()
                if content:
                    annotations.append((num, content))
            except ValueError:
                pass
        
        return annotations
    
    def clean_text(self, text: str) -> str:
        """
        清理和优化提取的文本
        
        参数:
            text: 原始文本
        
        返回:
            清理后的文本
        """
        # 替换多个换行符为单个
        text = re.sub(r'\n\s*\n', '\n\n', text)
        
        # 去除每行开头和结尾的空格
        lines = text.split('\n')
        cleaned_lines = [line.strip() for line in lines]
        
        # 移除空白行
        cleaned_lines = [line for line in cleaned_lines if line.strip()]
        
        # 重新组合文本
        cleaned_text = '\n'.join(cleaned_lines)
        
        # 处理可能的乱码（简单方法：移除连续的非中文字符）
        # 只保留中文字符、常见标点、数字和英文字母
        cleaned_text = re.sub(r'[^\u4e00-\u9fa5\u3000-\u303f\u2000-\u206f\u2e80-\u2eff\u31c0-\u31ef\u3200-\u32ff\u3300-\u33ff\u4dc0-\u4dff\u9fb0-\u9fff\u0020-\u007e\u00a0-\u00ff\n]+', ' ', cleaned_text)
        
        # 去除多余的空格
        cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
        
        return cleaned_text
    
    def convert_range(self, start_page: int, end_page: int, 
                     separate_annotations: bool = False) -> Tuple[str, List[Tuple[int, int, str]]]:
        """
        转换指定范围的页面
        
        参数:
            start_page: 起始页码（从1开始）
            end_page: 结束页码
            separate_annotations: 是否分离注释
        
        返回:
            (主文本, 注释列表[(页码, 编号, 内容), ...])
        """
        # 验证页码范围
        start_page = max(1, start_page)
        end_page = min(self.total_pages, end_page)
        
        if start_page > end_page:
            raise ValueError(f"无效的页码范围: {start_page}-{end_page}")
        
        print(f"\n开始转换页面 {start_page}-{end_page} / {self.total_pages}")
        
        main_text = []
        all_annotations = []
        
        for page_num in range(start_page, end_page + 1):
            # 显示进度
            if self.verbose or (page_num - start_page) % 10 == 0 or page_num == end_page:
                progress = ((page_num - start_page + 1) / (end_page - start_page + 1)) * 100
                print(f"转换进度: {page_num}/{end_page} ({progress:.1f}%)", end='\r')
            
            # 提取页面文本
            text = self.extract_page_text(page_num)
            
            # 如果文本为空，尝试用另一种方法
            if not text.strip():
                print(f"\n⚠ 页面 {page_num} 提取失败，尝试备用方法...")
                # 可以在这里尝试其他提取方法
                continue
            
            # 清理文本
            cleaned_text = self.clean_text(text)
            
            # 检测注释
            annotations = self.detect_annotations(cleaned_text)
            
            # 如果需要分离注释，从主文本中移除
            if separate_annotations and annotations:
                for num, content in annotations:
                    # 将注释信息保存
                    all_annotations.append((page_num, num, content))
                    
                    # 从主文本中移除注释行
                    pattern = re.compile(rf'^{re.escape(str(num))}\s+{re.escape(content[:20])}.*?$', re.MULTILINE)
                    cleaned_text = pattern.sub('', cleaned_text)
            
            # 添加页码标记和文本
            main_text.append(f"[第{page_num}页]")
            main_text.append(cleaned_text)
            main_text.append("")  # 页面间空行
        
        print("\n✓ 页面转换完成")
        
        # 组合主文本
        combined_text = '\n'.join(main_text)
        
        # 清理最终文本
        combined_text = self.clean_text(combined_text)
        
        print(f"✓ 提取到 {len(all_annotations)} 个注释")
        return combined_text, all_annotations
    
    def save_results(self, main_text: str, annotations: List[Tuple[int, int, str]], 
                    output_path: str, annotation_path: Optional[str] = None) -> None:
        """
        保存转换结果
        
        参数:
            main_text: 主文本内容
            annotations: 注释列表
            output_path: 主文本输出路径
            annotation_path: 注释输出路径
        """
        # 保存主文本
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(main_text)
            print(f"✓ 主文本成功保存到: {output_path}")
        except Exception as e:
            print(f"✗ 保存主文本失败: {e}")
        
        # 保存注释（如果提供了路径）
        if annotation_path and annotations:
            try:
                with open(annotation_path, 'w', encoding='utf-8') as f:
                    f.write("西游记 PDF注释\n")
                    f.write(f"总注释数: {len(annotations)}\n")
                    f.write(f"生成时间: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write("\n")
                    
                    # 按页码和编号排序注释
                    sorted_annotations = sorted(annotations, key=lambda x: (x[1], x[0]))
                    
                    for page_num, num, content in sorted_annotations:
                        f.write(f"注释 {num}:\n")
                        f.write(f"  [第{page_num}页] {content}\n")
                        f.write("\n")
                
                print(f"✓ 注释成功保存到: {annotation_path}")
            except Exception as e:
                print(f"✗ 保存注释文件失败: {e}")

def main():
    """
    主函数
    """
    print("===== 优化版西游记PDF转文本工具 =====")
    
    # 设置命令行参数
    parser = argparse.ArgumentParser(description='优化版西游记PDF转文本工具')
    parser.add_argument('--pdf', type=str, default='西游记（上下册）--中华经典小说注释系列.pdf', 
                       help='PDF文件路径')
    parser.add_argument('--start', type=int, default=1, help='起始页码')
    parser.add_argument('--end', type=int, default=20, help='结束页码')
    parser.add_argument('--output', type=str, default='西游记_优化版.txt', help='输出文件路径')
    parser.add_argument('--separate-annotations', action='store_true', 
                       help='是否将注释分离到单独文件')
    parser.add_argument('--annotation-output', type=str, default='西游记_优化版_注释.txt', 
                       help='注释输出文件路径')
    parser.add_argument('--verbose', action='store_true', help='显示详细信息')
    parser.add_argument('--all-pages', action='store_true', help='转换所有页面')
    
    args = parser.parse_args()
    
    try:
        # 确保输出目录存在
        output_dir = os.path.dirname(os.path.abspath(args.output))
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 初始化转换器
        converter = PDFConverter(args.pdf, args.verbose)
        
        # 确定页码范围
        start_page = args.start
        if args.all_pages:
            end_page = converter.total_pages
        else:
            end_page = min(args.end, converter.total_pages)
        
        # 记录开始时间
        start_time = time.time()
        
        # 转换指定范围的页面
        main_text, annotations = converter.convert_range(
            start_page, end_page, args.separate_annotations
        )
        
        # 保存结果
        annotation_path = args.annotation_output if args.separate_annotations else None
        converter.save_results(main_text, annotations, args.output, annotation_path)
        
        # 计算转换时间
        elapsed_time = time.time() - start_time
        
        # 显示转换摘要
        print(f"\n===== 转换摘要 =====")
        print(f"PDF文件: {args.pdf}")
        print(f"转换范围: 第 {start_page}-{end_page} 页 (共 {converter.total_pages} 页)")
        print(f"输出文件: {args.output}")
        if args.separate_annotations:
            print(f"注释文件: {args.annotation_output}")
        print(f"提取注释数: {len(annotations)}")
        print(f"转换耗时: {elapsed_time:.2f} 秒")
        print(f"平均速度: {((end_page - start_page + 1) / elapsed_time):.2f} 页/秒")
        print("===== 转换完成 =====")
        
    except KeyboardInterrupt:
        print("\n✗ 转换被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ 转换过程中发生错误: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
