import json
import os

# 文件路径
file_path = "d:\\AI\\books\\白眉大侠\\章回\\事件列表\\001_事件列表.json"

# 检查文件是否存在
if not os.path.exists(file_path):
    print(f"错误: 文件 {file_path} 不存在")
    exit(1)

# 检查文件权限
print(f"文件可读: {os.access(file_path, os.R_OK)}")
print(f"文件可写: {os.access(file_path, os.W_OK)}")

# 读取文件内容
print("\n读取文件内容:")
try:
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    print("文件读取成功")
    print(f"基础信息.标题: {data['基础信息']['标题']}")
    
    # 尝试修改并保存文件
    print("\n尝试修改并保存文件...")
    # 创建一个临时备份
    backup_path = file_path + ".bak"
    with open(backup_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"备份文件已创建: {backup_path}")
    
    # 修改标题
    original_title = data['基础信息']['标题']
    data['基础信息']['标题'] = original_title + " (已测试修改)"
    
    # 保存修改
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("文件已修改并保存")
    
    # 重新读取文件确认修改
    print("\n重新读取文件确认修改:")
    with open(file_path, 'r', encoding='utf-8') as f:
        new_data = json.load(f)
    print(f"修改后的标题: {new_data['基础信息']['标题']}")
    
    # 恢复原始内容
    with open(file_path, 'w', encoding='utf-8') as f:
        with open(backup_path, 'r', encoding='utf-8') as backup_f:
            original_data = json.load(backup_f)
        json.dump(original_data, f, ensure_ascii=False, indent=2)
    print("\n文件已恢复原始内容")
    os.remove(backup_path)
    print("备份文件已删除")
    
    # 再次读取确认恢复
    with open(file_path, 'r', encoding='utf-8') as f:
        restored_data = json.load(f)
    print(f"恢复后的标题: {restored_data['基础信息']['标题']}")
    
    # 测试直接写入一个新的测试文件
    test_file_path = "d:\\AI\\books\\test_write.json"
    print(f"\n测试写入新文件: {test_file_path}")
    test_data = {"测试": "写入成功"}
    with open(test_file_path, 'w', encoding='utf-8') as f:
        json.dump(test_data, f, ensure_ascii=False, indent=2)
    print(f"新文件写入成功，文件大小: {os.path.getsize(test_file_path)} 字节")
    
    # 读取测试文件
    with open(test_file_path, 'r', encoding='utf-8') as f:
        read_test_data = json.load(f)
    print(f"读取的测试数据: {read_test_data}")
    
    # 清理测试文件
    os.remove(test_file_path)
    print(f"测试文件已删除")
    
    print("\n结论: 文件IO操作正常")
    
    # 提示可能的问题
    print("\n如果您确认修改了文件但没有生效，可能的原因:")
    print("1. 修改后没有正确保存")
    print("2. 修改的是不同路径的文件")
    print("3. 使用的编辑器没有写入权限")
    print("4. 程序可能在读取缓存的文件内容")
    
except Exception as e:
    print(f"发生错误: {str(e)}")
