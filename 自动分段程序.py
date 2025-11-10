#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文本自动分段程序
功能：根据标点符号智能分段，使文本更易于阅读
作者：AI助手
日期：2023-11-20
"""

import re
import os
import argparse

class 文本分段器:
    """文本分段器类，提供智能分段功能"""
    
    def __init__(self):
        # 定义结束句子的标点符号
        self.结束标点 = ['。', '！', '？', '；', '!', '?', ';', '.']
        # 定义段落最大长度
        self.最大段落长度 = 100
        # 定义段落最小长度
        self.最小段落长度 = 20
    
    def 分段(self, 文本, 最大段落长度=None, 最小段落长度=None):
        """
        对文本进行智能分段
        
        参数:
            文本: 要分段的原始文本
            最大段落长度: 段落的最大长度
            最小段落长度: 段落的最小长度
            
        返回:
            分段后的文本字符串
        """
        # 如果没有指定，使用默认值
        if 最大段落长度 is None:
            最大段落长度 = self.最大段落长度
        if 最小段落长度 is None:
            最小段落长度 = self.最小段落长度
        
        # 首先去除多余的空白字符
        文本 = re.sub(r'\s+', ' ', 文本).strip()
        
        结果 = []
        当前段落 = ""
        
        # 遍历文本，进行分段
        i = 0
        文本长度 = len(文本)
        
        while i < 文本长度:
            # 添加当前字符到当前段落
            当前段落 += 文本[i]
            i += 1
            
            # 检查是否达到结束标点
            if i < 文本长度 and 文本[i-1] in self.结束标点:
                # 检查段落长度
                if len(当前段落) >= 最小段落长度:
                    # 如果段落长度超过最大值，或者下一个字符是换行符，就结束当前段落
                    if len(当前段落) >= 最大段落长度 or (i < 文本长度 and 文本[i] == '\n'):
                        结果.append(当前段落.strip())
                        当前段落 = ""
        
        # 添加最后一个段落
        if 当前段落.strip():
            结果.append(当前段落.strip())
        
        # 合并结果，每个段落一行
        return '\n\n'.join(结果)
    
    def 处理文件(self, 输入文件路径, 输出文件路径=None, 最大段落长度=None, 最小段落长度=None):
        """
        处理文件，对文件中的文本进行分段
        
        参数:
            输入文件路径: 输入文件的路径
            输出文件路径: 输出文件的路径，如果为None则自动生成
            最大段落长度: 段落的最大长度
            最小段落长度: 段落的最小长度
        """
        # 检查输入文件是否存在
        if not os.path.exists(输入文件路径):
            print(f"错误：输入文件 '{输入文件路径}' 不存在！")
            return False
        
        # 如果没有指定输出文件路径，自动生成
        if 输出文件路径 is None:
            文件名, 文件扩展名 = os.path.splitext(输入文件路径)
            输出文件路径 = f"{文件名}_分段{文件扩展名}"
        
        try:
            # 读取输入文件
            with open(输入文件路径, 'r', encoding='utf-8') as f:
                原始文本 = f.read()
            
            # 进行分段
            分段后文本 = self.分段(原始文本, 最大段落长度, 最小段落长度)
            
            # 写入输出文件
            with open(输出文件路径, 'w', encoding='utf-8') as f:
                f.write(分段后文本)
            
            print(f"分段完成！输出文件：'{输出文件路径}'")
            return True
        except Exception as e:
            print(f"处理文件时出错：{e}")
            return False

def main():
    """主函数，处理命令行参数"""
    parser = argparse.ArgumentParser(description='文本自动分段程序')
    parser.add_argument('输入文件', help='要分段的输入文件路径')
    parser.add_argument('-o', '--输出文件', help='分段后的输出文件路径')
    parser.add_argument('-m', '--最大长度', type=int, help='段落的最大长度')
    parser.add_argument('-n', '--最小长度', type=int, help='段落的最小长度')
    
    args = parser.parse_args()
    
    分段器 = 文本分段器()
    分段器.处理文件(args.输入文件, args.输出文件, args.最大长度, args.最小长度)

if __name__ == "__main__":
    main()

# 使用示例：
# 1. 基本使用：python 自动分段程序.py input.txt
# 2. 指定输出文件：python 自动分段程序.py input.txt -o output.txt
# 3. 自定义段落长度：python 自动分段程序.py input.txt -m 150 -n 30
