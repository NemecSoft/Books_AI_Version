import os
import subprocess

def 批量分段处理():
    """
    批量处理《白眉大侠》所有章回文件，为其生成分段版本
    """
    # 设置目录和文件路径
    章回目录 = r'd:\AI\books\白眉大侠\章回'
    分段程序 = r'd:\AI\books\自动分段程序.py'
    
    # 确保目录存在
    if not os.path.exists(章回目录):
        print(f"错误：章回目录不存在 - {章回目录}")
        return
    
    # 确保分段程序存在
    if not os.path.exists(分段程序):
        print(f"错误：分段程序不存在 - {分段程序}")
        return
    
    成功计数 = 0
    失败计数 = 0
    
    # 处理001到145的章回文件
    for i in range(1, 146):
        # 格式化序号，确保三位数
        序号 = f"{i:03d}"
        源文件 = os.path.join(章回目录, f"{序号}.txt")
        
        # 检查源文件是否存在
        if not os.path.exists(源文件):
            print(f"警告：文件不存在 - {源文件}")
            失败计数 += 1
            continue
        
        try:
            print(f"正在处理第{序号}回...")
            # 调用分段程序
            结果 = subprocess.run(
                ["python", 分段程序, 源文件],
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=300  # 设置5分钟超时
            )
            
            if 结果.returncode == 0:
                print(f"第{序号}回处理成功")
                成功计数 += 1
            else:
                print(f"警告：第{序号}回处理失败")
                print(f"错误输出: {结果.stderr}")
                失败计数 += 1
                
        except Exception as e:
            print(f"错误：处理第{序号}回时出错 - {str(e)}")
            失败计数 += 1
    
    # 输出处理结果统计
    print("\n===== 处理结果统计 =====")
    print(f"成功处理: {成功计数} 个文件")
    print(f"处理失败: {失败计数} 个文件")
    print(f"总文件数: {成功计数 + 失败计数}")
    print("========================")

if __name__ == "__main__":
    print("开始批量处理《白眉大侠》章回文件...")
    批量分段处理()
    print("\n批量处理完成！")
