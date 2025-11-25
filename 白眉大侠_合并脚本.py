import os

# 设置编码和文件路径
input_dir = "d:\\AI\\books\\白眉大侠\\章回"
output_file = "d:\\AI\\books\\白眉大侠\\白眉大侠_合并版.txt"

# 创建输出目录（如果不存在）
output_dir = os.path.dirname(output_file)
os.makedirs(output_dir, exist_ok=True)

print(f"开始合并文件，从 {input_dir} 目录读取章回文件...")

# 打开输出文件（使用UTF-8编码）
with open(output_file, 'w', encoding='utf-8') as out_file:
    # 遍历文件（从001到145）
    for i in range(1, 146):
        # 格式化文件名，确保三位数格式
        file_name = f"{i:03d}.txt"
        file_path = os.path.join(input_dir, file_name)
        
        try:
            # 检查文件是否存在
            if not os.path.exists(file_path):
                print(f"警告：文件不存在 - {file_path}")
                continue
            
            # 读取文件内容
            with open(file_path, 'r', encoding='utf-8') as in_file:
                content = in_file.read()
                
            # 写入文件内容，并添加分隔符
            out_file.write(f"=== 第{i}回 ===\n\n")
            out_file.write(content)
            out_file.write("\n\n")  # 添加两个空行作为回目间隔
            
            print(f"已合并: {file_name}")
            
        except Exception as e:
            print(f"处理文件 {file_name} 时出错: {str(e)}")

print(f"\n合并完成！\n输出文件: {output_file}")

# 统计合并结果
try:
    # 检查输出文件大小
    file_size = os.path.getsize(output_file)
    print(f"输出文件大小: {file_size / 1024:.2f} KB")
    
    # 统计行数
    with open(output_file, 'r', encoding='utf-8') as f:
        line_count = sum(1 for line in f)
    print(f"输出文件行数: {line_count}")
    
    print("合并操作成功完成！")
    
except Exception as e:
    print(f"统计结果时出错: {str(e)}")