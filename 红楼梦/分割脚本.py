#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
韩湘子全传文本分割脚本
功能：根据章节标题自动分割文本为多个小文件
"""

import os
import re
import shutil

def 分割韩湘子全传文本():
    """分割韩湘子全传文本文件，按照章节创建小文件"""
    
    # 源文件路径
    源文件路径 = r'.\红楼梦_主体.txt'
    # 输出目录
    输出目录 = r'.\章回'
    
    # 确保输出目录存在
    os.makedirs(输出目录, exist_ok=True)
    print(f"输出目录: {输出目录}")
    
    # 清理输出目录中的现有文件
    for 文件 in os.listdir(输出目录):
        文件路径 = os.path.join(输出目录, 文件)
        if os.path.isfile(文件路径):
            try:
                os.remove(文件路径)
                print(f"删除旧文件: {文件}")
            except Exception as e:
                print(f"删除文件时出错 {文件}: {e}")
    
    # 章节匹配正则表达式
    # 匹配模式：第[数字]章 或 第[中文数字]章，但排除"(第X章完)"这样的行
    章节正则 = re.compile(r'^[^（]*第([○零一二三四五六七八九十百千万\d]+)回[^）]*$')
    
    # 当前章节内容和文件
    当前章节内容 = []
    当前文件名 = None
    当前文件 = None
    章节计数 = 0
    
    try:
        # 读取源文件（尝试不同编码）
        编码方式 = ['utf-8', 'gbk', 'utf-16', 'gb2312']
        内容 = None
        
        for 编码 in 编码方式:
            try:
                with open(源文件路径, 'r', encoding=编码) as f:
                    内容 = f.readlines()
                print(f"成功以{编码}编码读取文件")
                break
            except UnicodeDecodeError:
                continue
        
        if 内容 is None:
            raise Exception("无法解码文件，请检查文件编码格式")
        
        # 遍历每一行
        for 行 in 内容:
            # 去除前后空白字符（保留换行符）
            处理行 = 行.rstrip('\r\n')
            
            # 检查是否是章节标题行
            匹配结果 = 章节正则.search(处理行)
            if 匹配结果:
                # 关闭当前打开的文件
                if 当前文件:
                    当前文件.write('\n'.join(当前章节内容))
                    当前文件.close()
                    当前章节内容 = []
                    章节计数 += 1
                
                # 提取章节号
                章节号文本 = 匹配结果.group(1)
                
                # 转换章节号为数字格式（如果是中文数字，这里简单处理为纯数字）
                try:
                    # 如果已经是数字，直接转换
                    章节号 = int(章节号文本)
                except ValueError:
                    # 如果是中文数字，暂时使用序号代替
                    章节号 = 章节计数 + 1
                
                # 格式化文件名：001_第一章 xxx.txt
                文件名前缀 = f"{章节号:03d}_"
                文件名后缀 = 处理行.replace('\t', ' ').replace('/', '_').replace('\\', '_') + '.txt'
                当前文件名 = 文件名前缀 + 文件名后缀
                
                # 完整文件路径
                文件完整路径 = os.path.join(输出目录, 当前文件名)
                
                # 打开新文件
                当前文件 = open(文件完整路径, 'w', encoding='utf-8')
                print(f"创建章节文件: {当前文件名}")
                
                # 添加章节标题到当前章节内容
                当前章节内容.append(处理行)
            else:
                # 如果当前有打开的章节文件，将内容添加进去
                if 当前文件:
                    当前章节内容.append(处理行)
        
        # 处理最后一个章节
        if 当前文件:
            当前文件.write('\n'.join(当前章节内容))
            当前文件.close()
            章节计数 += 1
        
        print(f"分割完成！共创建{章节计数}个章节文件")
        print(f"所有文件已保存到: {输出目录}")
        
    except Exception as e:
        print(f"处理过程中出现错误: {str(e)}")
        # 确保文件被关闭
        if 当前文件:
            当前文件.close()
    finally:
        # 再次确保文件被关闭
        if '当前文件' in locals() and 当前文件:
            try:
                当前文件.close()
            except:
                pass

def 主函数():
    """主函数"""
    print("=== 韩湘子全传文本分割工具 ===")
    print("功能：将韩湘子全传文本按章节分割为多个小文件")
    print("=" * 50)
    
    # 执行分割
    分割韩湘子全传文本()
    
    print("\n程序执行完毕！")

if __name__ == "__main__":
    主函数()
