#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
《水浒全传》章回分割脚本
作者：AI助手
功能：使用正则表达式按章回分割水浒全传文本文件
"""

import os
import re

def split_water_margin():
    # 设置文件路径
    source_file = "d:\\AI\\books\\水浒传\\水浒全传.txt"
    regex_file = "d:\\AI\\books\\水浒传\\正则.txt"
    output_dir = "d:\\AI\\books\\水浒传\\章回"
    
    # 读取正则表达式
    with open(regex_file, 'r', encoding='utf-8') as f:
        regex_pattern = f.read().strip()
    
    # 创建输出目录
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"创建输出目录: {output_dir}")
    
    # 读取源文件内容
    with open(source_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    print(f"开始分割《水浒全传》...")
    print(f"使用正则表达式: {regex_pattern}")
    
    # 查找所有章回标题
    chapter_matches = list(re.finditer(regex_pattern, content))
    
    print(f"找到 {len(chapter_matches)} 个章回标题")
    
    # 如果没有找到章回，尝试其他可能的格式
    if len(chapter_matches) == 0:
        print("未找到章回标题，尝试其他格式...")
        backup_regex = r"第[一二三四五六七八九十百零\d]+回"
        chapter_matches = list(re.finditer(backup_regex, content))
        print(f"使用备用正则找到 {len(chapter_matches)} 个章回标题")
    
    # 分割文件
    for i, match in enumerate(chapter_matches):
        current_title = match.group()
        current_position = match.start()
        
        # 确定章回结束位置
        if i < len(chapter_matches) - 1:
            next_position = chapter_matches[i + 1].start()
            chapter_content = content[current_position:next_position].strip()
        else:
            chapter_content = content[current_position:].strip()
        
        # 生成文件名（提取章回号）
        chapter_number_match = re.search(r"第(.{0,5}?)回", current_title)
        if chapter_number_match:
            chapter_number = chapter_number_match.group(1)
            # 尝试转换为数字
            try:
                if chapter_number.isdigit():
                    file_number = int(chapter_number)
                else:
                    # 处理中文数字
                    chinese_numbers = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}
                    file_number = i + 1
            except:
                file_number = i + 1
        else:
            file_number = i + 1
        
        # 确保文件名是3位数字格式
        filename = f"{file_number:03d}.txt"
        output_path = os.path.join(output_dir, filename)
        
        # 写入章回内容
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(chapter_content)
        
        print(f"已生成: {filename} - {current_title}")
    
    print(f"分割完成！共生成 {len(chapter_matches)} 个章回文件。")
    print(f"文件保存在: {output_dir}")

if __name__ == "__main__":
    split_water_margin()