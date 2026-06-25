#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import configparser
import json
import sys
import os

def load_config(config_file='config.ini'):
    """
    从配置文件读取参数
    
    支持两种格式:
    1. INI格式 (使用configparser)
    2. JSON格式 (如果文件是.json后缀)
    """
    config = {
        'filename': None,
        'max_length': 80,
        'encoding': 'utf-8',
        'output_format': 'console'  # console 或 json
    }
    
    if not os.path.exists(config_file):
        print(f"警告: 配置文件 '{config_file}' 不存在，使用默认参数")
        return config
    
    # 根据文件扩展名选择解析方式
    if config_file.endswith('.json'):
        return load_config_json(config_file)
    else:
        return load_config_ini(config_file)


def load_config_ini(config_file):
    """从INI格式配置文件读取"""
    config = {
        'filename': None,
        'max_length': 80,
        'encoding': 'utf-8',
        'output_format': 'console'
    }
    
    try:
        parser = configparser.ConfigParser()
        parser.read(config_file, encoding='utf-8')
        
        if parser.has_section('settings'):
            config['filename'] = parser.get('settings', 'filename', fallback=None)
            config['max_length'] = parser.getint('settings', 'max_length', fallback=80)
            config['encoding'] = parser.get('settings', 'encoding', fallback='utf-8')
            config['output_format'] = parser.get('settings', 'output_format', fallback='console')
        else:
            print(f"警告: 配置文件中没有 [settings] 部分")
            
    except Exception as e:
        print(f"读取配置文件出错: {e}")
    
    return config


def load_config_json(config_file):
    """从JSON格式配置文件读取"""
    config = {
        'filename': None,
        'max_length': 80,
        'encoding': 'utf-8',
        'output_format': 'console'
    }
    
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        config['filename'] = data.get('filename', None)
        config['max_length'] = data.get('max_length', 80)
        config['encoding'] = data.get('encoding', 'utf-8')
        config['output_format'] = data.get('output_format', 'console')
        
    except Exception as e:
        print(f"读取JSON配置文件出错: {e}")
    
    return config


def check_line_length(config):
    """
    根据配置检查文本文件行长度
    """
    filename = config.get('filename')
    max_length = config.get('max_length', 80)
    encoding = config.get('encoding', 'utf-8')
    output_format = config.get('output_format', 'console')
    
    if not filename:
        print("错误: 配置文件中未指定文件名")
        return []
    
    if not os.path.exists(filename):
        print(f"错误: 文件 '{filename}' 不存在")
        return []
    
    try:
        with open(filename, 'r', encoding=encoding) as file:
            lines = file.readlines()
        
        exceeded_lines = []
        total_lines = len(lines)
        
        for line_num, line in enumerate(lines, start=1):
            line_stripped = line.rstrip('\n\r')
            char_count = len(line_stripped)
            
            if char_count > max_length:
                exceeded_lines.append({
                    'line_num': line_num,
                    'count': char_count,
                    'content': line_stripped[:50] + '...' if len(line_stripped) > 50 else line_stripped,
                    'full_content': line_stripped  # 用于JSON输出
                })
        
        # 输出结果
        if output_format == 'json':
            output_json(exceeded_lines, total_lines, max_length, filename)
        else:
            output_console(exceeded_lines, total_lines, max_length, filename)
        
        return exceeded_lines
        
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return []


def output_console(exceeded_lines, total_lines, max_length, filename):
    """控制台输出"""
    print(f"正在检查文件: {filename}")
    print(f"最大字符数限制: {max_length}")
    print("-" * 60)
    
    if exceeded_lines:
        print(f"发现 {len(exceeded_lines)} 行超出字符数限制:\n")
        for item in exceeded_lines:
            print(f"第 {item['line_num']:4d} 行: {item['count']:4d} 个字符 -> {item['content']}")
    else:
        print(f"✓ 所有 {total_lines} 行都在字符数限制内！")
    
    print("-" * 60)
    print(f"总行数: {total_lines}, 超出行数: {len(exceeded_lines)}")


def output_json(exceeded_lines, total_lines, max_length, filename):
    """JSON格式输出"""
    result = {
        'filename': filename,
        'max_length': max_length,
        'total_lines': total_lines,
        'exceeded_count': len(exceeded_lines),
        'exceeded_lines': exceeded_lines
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description='从配置文件读取参数检查文本行长度')
    parser.add_argument('-c', '--config', default='config.ini',
                       help='配置文件路径 (支持 .ini 或 .json, 默认: config.ini)')
    
    args = parser.parse_args()
    
    # 从配置文件加载参数
    config = load_config(args.config)
    
    # 检查必需参数
    if not config['filename']:
        print("错误: 配置文件中未指定要检查的文件名")
        print("\n请在配置文件中设置 filename 参数")
        sys.exit(1)
    
    # 执行检查
    check_line_length(config)


if __name__ == "__main__":
    main()