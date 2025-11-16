#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
清理三国演义分割结果
删除只包含标题的文件，保留完整内容的文件
"""

import os
import re

def 清理分割结果():
    """清理分割结果，只保留完整内容的回目文件"""
    
    章回目录 = "d:\\AI\\books\\三国演义\\章回\\"
    
    # 获取目录中的所有文件
    文件列表 = os.listdir(章回目录)
    
    # 找出需要删除的文件（只包含数字的文件名）
    要删除的文件 = []
    要保留的文件 = []
    
    for 文件名 in 文件列表:
        if 文件名.endswith('.txt'):
            # 检查是否是纯数字文件名（如001.txt, 002.txt等）
            基础名 = 文件名[:-4]  # 去掉.txt
            if 基础名.isdigit():
                要删除的文件.append(文件名)
            else:
                要保留的文件.append(文件名)
    
    print(f"找到 {len(要删除的文件)} 个需要删除的文件（纯数字文件名）")
    print(f"找到 {len(要保留的文件)} 个需要保留的文件（包含中文的文件名）")
    
    # 删除纯数字文件名的文件
    for 文件名 in 要删除的文件:
        文件路径 = os.path.join(章回目录, 文件名)
        try:
            os.remove(文件路径)
            print(f"已删除：{文件名}")
        except Exception as e:
            print(f"删除文件 {文件名} 失败：{e}")
    
    # 重命名保留的文件，统一格式
    要保留的文件.sort()  # 按文件名排序
    
    for i, 文件名 in enumerate(要保留的文件, 1):
        旧路径 = os.path.join(章回目录, 文件名)
        新文件名 = f"{i:03d}.txt"
        新路径 = os.path.join(章回目录, 新文件名)
        
        try:
            os.rename(旧路径, 新路径)
            print(f"已重命名：{文件名} -> {新文件名}")
        except Exception as e:
            print(f"重命名文件 {文件名} 失败：{e}")
    
    print(f"清理完成！最终保留 {len(要保留的文件)} 个回目文件")

if __name__ == "__main__":
    清理分割结果()