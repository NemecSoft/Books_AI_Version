import json
import os

# 检查JSON文件的路径
json_files = [
    "d:\\AI\\books\\白眉大侠\\章回\\事件列表\\001_事件列表.json",
    "d:\\AI\\books\\白眉大侠\\章回\\事件列表\\002_事件列表.json",
    "d:\\AI\\books\\白眉大侠\\章回\\事件列表\\003_事件列表.json"
]

# 格式规范要求的关键字段
required_fields = ["基础信息", "详细版", "简化版", "极简版", "绘图提示词"]
basic_info_fields = ["标题", "版本"]

all_valid = True

print("开始检查JSON文件格式...")
print("=" * 50)

for json_file in json_files:
    print(f"\n检查文件: {json_file}")
    try:
        # 读取JSON文件
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 检查必填字段
        valid = True
        for field in required_fields:
            if field not in data:
                print(f"❌ 缺少必填字段: {field}")
                valid = False
            elif field == "基础信息" and not isinstance(data[field], dict):
                print(f"❌ 字段 {field} 应该是对象类型")
                valid = False
            elif field not in ["基础信息", "绘图提示词"] and not isinstance(data[field], list):
                print(f"❌ 字段 {field} 应该是数组类型")
                valid = False
            elif field == "绘图提示词" and not isinstance(data[field], str):
                print(f"❌ 字段 {field} 应该是字符串类型")
                valid = False
        
        # 检查基础信息
        if "基础信息" in data:
            for field in basic_info_fields:
                if field not in data["基础信息"]:
                    print(f"❌ 基础信息中缺少必填字段: {field}")
                    valid = False
        
        # 检查数组内容
        if "详细版" in data and isinstance(data["详细版"], list):
            print(f"✅ 详细版包含 {len(data['详细版'])} 条事件")
        
        if "简化版" in data and isinstance(data["简化版"], list):
            if 3 <= len(data["简化版"]) <= 5:
                print(f"✅ 简化版包含 {len(data['简化版'])} 条事件，符合要求")
            else:
                print(f"⚠️  简化版包含 {len(data['简化版'])} 条事件，建议3-5条")
        
        if "极简版" in data and isinstance(data["极简版"], list):
            if 2 <= len(data["极简版"]) <= 4:
                print(f"✅ 极简版包含 {len(data['极简版'])} 条事件，符合要求")
            else:
                print(f"⚠️  极简版包含 {len(data['极简版'])} 条事件，建议2-4条")
        
        if valid:
            print("✅ 格式检查通过！")
        else:
            all_valid = False
            print("❌ 格式检查未通过！")
            
    except json.JSONDecodeError as e:
        all_valid = False
        print(f"❌ JSON格式错误: {e}")
    except Exception as e:
        all_valid = False
        print(f"❌ 读取文件时出错: {e}")

print("\n" + "=" * 50)
if all_valid:
    print("✅ 所有JSON文件格式检查通过！")
else:
    print("❌ 部分JSON文件格式检查未通过，请检查并修正。")
