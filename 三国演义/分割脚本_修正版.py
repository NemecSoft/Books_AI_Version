#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
三国演义分割脚本 - 修正版
按照正则表达式"第(.{0,5}?)回"将三国演义分割成120回
确保只匹配正文中的回目，避免目录中的回目干扰
"""

import re
import os

def 分割三国演义():
    """分割三国演义文本为120回"""
    
    # 设置文件路径
    源文件路径 = r"d:\AI\books\三国演义\三国演义.txt"
    输出目录 = r"d:\AI\books\三国演义\章回"
    
    # 确保输出目录存在
    os.makedirs(输出目录, exist_ok=True)
    
    # 读取源文件
    try:
        with open(源文件路径, 'r', encoding='utf-8') as f:
            内容 = f.read()
    except Exception as e:
        print(f"读取文件失败: {e}")
        return
    
    # 找到正文开始位置（跳过简介和目录）
    正文开始标记 = "第一回　宴桃园豪杰三结义　斩黄巾英雄首立功　"
    正文开始位置 = 内容.find(正文开始标记)
    if 正文开始位置 == -1:
        print("未找到正文开始位置")
        return
    
    正文内容 = 内容[正文开始位置:]
    
    # 使用正则表达式查找所有回目
    # 确保匹配完整的回目格式：第XXX回　标题1　标题2
    回目模式 = r'第([一二三四五六七八九十百零]+)回[　\s]+([^　\n]+)[　\s]+([^　\n]+)'
    回目列表 = re.finditer(回目模式, 正文内容)
    
    回目信息 = []
    for match in 回目列表:
        回数 = match.group(1)
        标题1 = match.group(2)
        标题2 = match.group(3)
        开始位置 = match.start()
        回目信息.append({
            '回数': 回数,
            '标题1': 标题1,
            '标题2': 标题2,
            '位置': 开始位置
        })
    
    print(f"找到 {len(回目信息)} 个回目")
    
    # 将中文数字转换为阿拉伯数字
    def 中文数字转阿拉伯数字(中文数字):
        数字映射 = {
            '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
            '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
            '百': 100, '零': 0
        }
        
        # 处理特殊情况
        if '百' in 中文数字:
            if 中文数字 == '一百':
                return 100
            elif 中文数字 == '一百一':
                return 111
            elif 中文数字 == '一百二':
                return 112
            elif 中文数字.startswith('一百一'):
                return 110 + int(数字映射.get(中文数字[3:], 0))
            elif 中文数字.startswith('一百二'):
                return 120 + int(数字映射.get(中文数字[3:], 0))
            elif 中文数字.startswith('一百三'):
                return 130 + int(数字映射.get(中文数字[3:], 0))
            elif 中文数字.startswith('一百四'):
                return 140 + int(数字映射.get(中文数字[3:], 0))
            elif 中文数字.startswith('一百五'):
                return 150 + int(数字映射.get(中文数字[3:], 0))
            elif 中文数字.startswith('一百六'):
                return 160 + int(数字映射.get(中文数字[3:], 0))
            elif 中文数字.startswith('一百七'):
                return 170 + int(数字映射.get(中文数字[3:], 0))
            elif 中文数字.startswith('一百八'):
                return 180 + int(数字映射.get(中文数字[3:], 0))
            elif 中文数字.startswith('一百九'):
                return 190 + int(数字映射.get(中文数字[3:], 0))
        elif '十' in 中文数字:
            if 中文数字 == '十':
                return 10
            elif 中文数字.startswith('十'):
                return 10 + int(数字映射.get(中文数字[1:], 0))
            elif 中文数字.endswith('十'):
                return int(数字映射.get(中文数字[0], 0)) * 10
            else:
                # 处理"二十"、"三十"等情况
                if len(中文数字) == 2:
                    return int(数字映射.get(中文数字[0], 0)) * 10
                else:
                    return int(数字映射.get(中文数字[0], 0)) * 10 + int(数字映射.get(中文数字[2], 0))
        else:
            return int(数字映射.get(中文数字, 0))
    
    # 分割每一回的内容
    for i, 回目 in enumerate(回目信息):
        回数 = 回目['回数']
        阿拉伯回数 = 中文数字转阿拉伯数字(回数)
        
        # 格式化为三位数
        文件名 = f"{int(阿拉伯回数):03d}.txt"
        
        # 确定本回内容的开始和结束位置
        开始位置 = 回目['位置']
        
        # 下一个回目的开始位置（如果有的话）
        if i + 1 < len(回目信息):
            结束位置 = 回目信息[i + 1]['位置']
        else:
            结束位置 = len(正文内容)
        
        # 提取本回内容
        回内容 = 正文内容[开始位置:结束位置].strip()
        
        # 保存到文件
        文件路径 = os.path.join(输出目录, 文件名)
        try:
            with open(文件路径, 'w', encoding='utf-8') as f:
                f.write(回内容)
            print(f"已保存: {文件名} - 第{回数}回 {回目['标题1']} {回目['标题2']}")
        except Exception as e:
            print(f"保存文件 {文件名} 失败: {e}")
    
    print(f"分割完成！共生成 {len(回目信息)} 个回目文件")

if __name__ == "__main__":
    分割三国演义()