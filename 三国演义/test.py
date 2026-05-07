# 假设你的文件名为 'data.txt'
file_path = '三国演义.txt'

try:
    with open(file_path, 'r', encoding='utf-8') as file:
        # 使用 enumerate 获取行号，从1开始
        for line_number, line in enumerate(file, start=1):
            # strip() 方法会移除字符串首尾的空白字符（包括换行符）
            clean_line = line.strip()
            
            # 检查处理后的行长度是否超过30个字符
            if len(clean_line) > 80:
                print(f"第 {line_number} 行 超长 ({len(clean_line)} 字符): {clean_line}")
                
except FileNotFoundError:
    print(f"错误：找不到文件 '{file_path}'，请确认文件路径是否正确。")
except Exception as e:
    print(f"发生了一个未知错误: {e}")