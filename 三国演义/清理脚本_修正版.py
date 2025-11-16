#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
清理章回目录下的所有txt文件
"""

import os
import glob

def 清理文件():
    """清理章回目录下的所有txt文件"""
    
    章回目录 = r"d:\AI\books\三国演义\章回"
    
    # 查找所有txt文件
    txt文件列表 = glob.glob(os.path.join(章回目录, "*.txt"))
    
    print(f"找到 {len(txt文件列表)} 个txt文件")
    
    # 删除所有txt文件
    for 文件路径 in txt文件列表:
        try:
            os.remove(文件路径)
            print(f"已删除: {os.path.basename(文件路径)}")
        except Exception as e:
            print(f"删除文件 {文件路径} 失败: {e}")
    
    print("清理完成！")

if __name__ == "__main__":
    清理文件()