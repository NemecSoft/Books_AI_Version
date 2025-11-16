#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
三国演义分割脚本
按照正则表达式将文本分成120回
"""

import re
import os

def 分割三国演义():
    """按照正则表达式分割三国演义文本"""
    
    # 文件路径
    输入文件 = "d:\\AI\\books\\三国演义\\三国演义.txt"
    输出目录 = "d:\\AI\\books\\三国演义\\章回\\"
    
    # 创建输出目录
    if not os.path.exists(输出目录):
        os.makedirs(输出目录)
    
    # 读取文本
    try:
        with open(输入文件, 'r', encoding='utf-8') as f:
            内容 = f.read()
        print(f"成功读取文件：{输入文件}")
    except Exception as e:
        print(f"读取文件失败：{e}")
        return
    
    # 使用正则表达式查找所有回目
    正则模式 = r'第(.{0,5}?)回'
    回目列表 = re.finditer(正则模式, 内容)
    
    回目信息 = []
    for match in 回目列表:
        回目编号 = match.group(1)
        回目位置 = match.start()
        回目信息.append({
            '编号': 回目编号,
            '位置': 回目位置,
            '完整标题': match.group(0)
        })
    
    print(f"找到 {len(回目信息)} 个回目")
    
    # 分割每一回
    for i, 回目 in enumerate(回目信息):
        回编号 = 回目['编号']
        开始位置 = 回目['位置']
        
        # 确定结束位置（下一个回目的开始位置，或文件末尾）
        if i < len(回目信息) - 1:
            结束位置 = 回目信息[i + 1]['位置']
        else:
            结束位置 = len(内容)
        
        # 提取回目内容
        回目内容 = 内容[开始位置:结束位置].strip()
        
        # 生成文件名（补零到3位）
        文件名 = f"{回编号.zfill(3)}.txt"
        文件路径 = os.path.join(输出目录, 文件名)
        
        # 保存文件
        try:
            with open(文件路径, 'w', encoding='utf-8') as f:
                f.write(回目内容)
            print(f"已保存：{文件名}")
        except Exception as e:
            print(f"保存文件 {文件名} 失败：{e}")
    
    print(f"分割完成！共生成 {len(回目信息)} 个回目文件")

if __name__ == "__main__":
    分割三国演义()