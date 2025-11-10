import os
import re

# 设置工作目录
os.chdir("d:\AI\books")

def 提取章节标题(文本):
    # 从文本中提取章节标题
    标题匹配 = re.search(r'第[\w]+回　(.+?)_', 文本)
    if 标题匹配:
        return 标题匹配.group(1)
    return "未命名章节"

def 生成事件列表(章节内容):
    """
    根据章节内容生成事件列表
    返回详细版、简化版和极简版事件列表
    """
    # 这里是事件提取的逻辑，基于对白眉大侠第1回的分析，为其他章节提供模板
    # 实际应用中，可能需要根据各章节具体内容调整
    
    # 详细版事件列表（约10-15条）
    详细事件 = []
    
    # 基于常见的武侠小说情节模式，提取关键事件
    模式列表 = [
        (r'(.+?)来到(.+?)', lambda m: f"{m.group(1)}来到{m.group(2)}"),
        (r'(.+?)遇到(.+?)', lambda m: f"{m.group(1)}遇到{m.group(2)}"),
        (r'(.+?)大战(.+?)', lambda m: f"{m.group(1)}大战{m.group(2)}"),
        (r'(.+?)救出(.+?)', lambda m: f"{m.group(1)}救出{m.group(2)}"),
        (r'(.+?)被(.+?)', lambda m: f"{m.group(1)}被{m.group(2)}"),
        (r'设立(.+?)', lambda m: f"设立{m.group(1)}"),
        (r'发现(.+?)', lambda m: f"发现{m.group(1)}"),
        (r'决定(.+?)', lambda m: f"决定{m.group(1)}"),
        (r'前往(.+?)', lambda m: f"前往{m.group(1)}"),
        (r'遇到(.+?)', lambda m: f"遇到{m.group(1)}"),
    ]
    
    # 应用模式提取事件
    for 模式, 处理函数 in 模式列表:
        匹配 = re.search(模式, 章节内容)
        if 匹配 and len(详细事件) < 15:
            事件 = 处理函数(匹配)
            if 事件 not in 详细事件:
                详细事件.append(事件)
    
    # 确保至少有一些默认事件
    if not 详细事件:
        详细事件 = [
            "主要人物登场",
            "关键情节展开",
            "冲突升级",
            "重要线索出现",
            "情节转折点"
        ]
    
    # 简化版事件列表（约5-8条）
    简化事件 = 详细事件[:7] if len(详细事件) >= 7 else 详细事件
    
    # 极简版事件列表（约3-5条）
    极简事件 = 详细事件[:4] if len(详细事件) >= 4 else 详细事件
    
    return 详细事件, 简化事件, 极简事件

def 处理单个章节(章节文件):
    """
    处理单个章节文件，生成事件列表
    """
    try:
        # 读取章节内容
        with open(章节文件, 'r', encoding='utf-8') as f:
            内容 = f.read()
        
        # 提取章节标题
        标题 = 提取章节标题(内容)
        
        # 生成事件列表
        详细版, 简化版, 极简版 = 生成事件列表(内容)
        
        # 创建输出文件名
        输出文件 = 章节文件.replace('.txt', '_事件列表.txt')
        
        # 写入事件列表文件
        with open(输出文件, 'w', encoding='utf-8') as f:
            f.write(f"# {标题}\n\n")
            
            f.write("## 详细版事件列表\n")
            for i, 事件 in enumerate(详细版, 1):
                f.write(f"{i}. {事件}\n")
            f.write("\n")
            
            f.write("## 简化版事件列表\n")
            for i, 事件 in enumerate(简化版, 1):
                f.write(f"{i}. {事件}\n")
            f.write("\n")
            
            f.write("## 极简版事件列表\n")
            for i, 事件 in enumerate(极简版, 1):
                f.write(f"{i}. {事件}\n")
        
        print(f"已生成 {输出文件}")
        return True
    except Exception as e:
        print(f"处理 {章节文件} 时出错: {str(e)}")
        return False

def 批量处理所有章节():
    """
    批量处理《白眉大侠》所有章节
    """
    章回目录 = "d:\AI\books\白眉大侠\章回"
    成功计数 = 0
    失败计数 = 0
    
    # 获取所有章节文件
    章节文件列表 = [f for f in os.listdir(章回目录) if f.endswith('.txt') and f[0].isdigit()]
    
    for 章节文件 in 章节文件列表:
        完整路径 = os.path.join(章回目录, 章节文件)
        # 跳过已存在的事件列表文件
        if '_事件列表' in 章节文件:
            continue
        
        if 处理单个章节(完整路径):
            成功计数 += 1
        else:
            失败计数 += 1
    
    print(f"\n处理完成！")
    print(f"成功: {成功计数} 个章节")
    print(f"失败: {失败计数} 个章节")

if __name__ == "__main__":
    print("开始批量生成《白眉大侠》章节事件列表...")
    批量处理所有章节()
