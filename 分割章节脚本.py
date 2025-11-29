import os
import re

# 设置控制台编码为UTF-8以支持中文显示
os.system('chcp 65001')

def 按章节分割文件(输入文件路径, 输出目录):
    # 确保输出目录存在
    os.makedirs(输出目录, exist_ok=True)
    
    print(f"开始处理文件: {输入文件路径}")
    
    try:
        # 添加详细调试信息：打印文件前100行，确保能看到章节标题
        print("\n--- 文件前100行内容（用于调试） ---")
        文件内容 = []
        with open(输入文件路径, 'r', encoding='utf-8') as f:
            for i in range(100):
                行 = f.readline()
                if not 行:
                    break
                文件内容.append(行)
                行处理 = 行.rstrip('\r\n')
                print(f"第{i+1}行: '{行处理}'")
                # 特别标记可能的章节标题行
                if "CHAPTER" in 行处理.upper() and "第" in 行处理 and "章" in 行处理:
                    print(f"  ===> 疑似章节标题: '{行处理}'")
        print("--- 调试结束 ---")
        
        # 使用用户建议的简洁正则表达式
        # 匹配格式：CHAPTER 后面最多15个任意字符，然后到章字为止
        章节模式 = re.compile(r'CHAPTER(.{0,15}?)章', re.IGNORECASE)
        
        
        当前章节 = None
        当前章节内容 = []
        章节计数 = 0
        总行数 = 0
        
        # 逐行读取文件并处理
        with open(输入文件路径, 'r', encoding='utf-8') as f:
            for 行号, 行 in enumerate(f, 1):
                总行数 += 1
                原行 = 行
                行 = 行.rstrip('\r\n')
                
                # 使用正则表达式匹配章节标题
                匹配结果 = 章节模式.search(行)
                if 匹配结果:
                    print(f"找到章节标题: '{行}' (行号: {行号})")
                    
                    # 如果已经有当前章节内容，先保存
                    if 当前章节:
                        保存章节(当前章节, 当前章节内容, 输出目录)
                        章节计数 += 1
                    
                    # 从匹配结果中提取信息
                    匹配内容 = 匹配结果.group(1)
                    
                    # 提取章节号 - 从匹配的内容中提取数字
                    # 方法1：从匹配结果中提取数字
                    数字匹配 = re.search(r'(\d+)', 匹配内容)
                    # 方法2：如果方法1失败，尝试从整个行中提取第X章中的数字
                    if not 数字匹配:
                        数字匹配 = re.search(r'第(\d+)章', 行)
                    
                    中文章节号 = "00"
                    if 数字匹配:
                        中文章节号 = 数字匹配.group(1)
                    
                    # 提取章节名称 - 章后面的内容
                    章节名称匹配 = re.search(r'章\s+(.+)', 行)
                    章节名称 = "未知"
                    if 章节名称匹配:
                        章节名称 = 章节名称匹配.group(1).strip()
                    
                    # 构造章节文件名
                    当前章节 = f"{int(中文章节号):02d}_{章节名称}"
                    当前章节内容 = [原行]
                    print(f"处理后的章节信息: {当前章节}")
                elif 当前章节:
                    # 添加到当前章节内容
                    当前章节内容.append(原行)
        
        # 保存最后一个章节
        if 当前章节:
            保存章节(当前章节, 当前章节内容, 输出目录)
            章节计数 += 1
        
        print(f"文件分割完成！共分割出 {章节计数} 个章节")
        
    except Exception as e:
        print(f"处理过程中出错: {str(e)}")

def 保存章节(章节名, 章节内容, 输出目录):
    # 清理文件名，移除非法字符
    安全章节名 = 章节名.replace('/', '').replace('\\', '').replace(':', '').replace('*', '').replace('?', '').replace('"', '').replace('<', '').replace('>', '').replace('|', '')
    文件路径 = os.path.join(输出目录, f"{安全章节名}.txt")
    
    with open(文件路径, 'w', encoding='utf-8') as f:
        f.writelines(章节内容)
    
    print(f"  已保存章节: {文件路径}")

if __name__ == "__main__":
    # 输入文件路径
    输入文件 = "d:\\AI\\books\\事实 (汉斯·罗斯林  欧拉·罗斯林  安娜·罗斯林·罗朗德)\\事实_转换后.txt"
    
    # 输出目录
    输出目录 = "d:\\AI\\books\\事实 (汉斯·罗斯林  欧拉·罗斯林  安娜·罗斯林·罗朗德)\\章回"
    
    # 执行分割
    按章节分割文件(输入文件, 输出目录)
