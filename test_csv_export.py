import json
import os
import pandas as pd
import types
from datetime import datetime

class TestCSVExporter:
    def __init__(self):
        self.json_data = []
    
    def _safe_convert_to_list(self, data):
        """安全地将数据转换为列表"""
        try:
            if isinstance(data, types.GeneratorType):
                return list(data)
            elif hasattr(data, '__iter__') and not isinstance(data, (str, bytes, dict)):
                return list(data)
            else:
                return [data] if data is not None else []
        except:
            return []
    
    def mock_data(self):
        """创建模拟数据"""
        mock_data = [
            {
                "标题": "测试书籍1",
                "极简版": ["事件A", "事件B", "事件C"],
                "简化版": ["这是事件A的简化描述", "这是事件B的简化描述"],
                "详细版": ["这是事件A的详细描述，包含更多内容。", "这是事件B的详细描述，包含更多内容。"]
            },
            {
                "标题": "测试书籍2",
                "极简版": ["事件X", "事件Y"],
                "简化版": ["这是事件X的简化描述", "这是事件Y的简化描述", "这是事件Z的简化描述"],
                "详细版": ["这是事件X的详细描述，包含更多内容。"]
            }
        ]
        
        # 测试生成器防御功能
        def generator_example():
            yield "生成器事件1"
            yield "生成器事件2"
        
        mock_data[0]["极简版"] = generator_example()
        self.json_data = mock_data
        print("已创建模拟数据并包含生成器对象")
    
    def test_csv_export(self):
        """测试CSV导出功能"""
        try:
            print("开始测试CSV导出功能...")
            
            # 防御性编程：确保json_data是列表类型
            if isinstance(self.json_data, types.GeneratorType):
                self.json_data = list(self.json_data)
            elif hasattr(self.json_data, '__iter__') and not hasattr(self.json_data, '__len__') and not isinstance(self.json_data, (str, bytes, dict)):
                self.json_data = list(self.json_data)
            if not isinstance(self.json_data, list):
                self.json_data = [self.json_data] if self.json_data is not None else []
            
            if not self.json_data:
                print("错误: 没有数据可导出")
                return False
            
            # 创建数据框
            data_to_export = []
            
            for data in self.json_data:
                if not isinstance(data, dict):
                    continue
                
                # 处理各字段
                极简版_list = self._safe_convert_to_list(data.get("极简版", []))
                简化版_list = self._safe_convert_to_list(data.get("简化版", []))
                详细版_list = self._safe_convert_to_list(data.get("详细版", []))
                
                print(f"书籍: {data.get('标题')}")
                print(f"  极简版列表类型: {type(极简版_list)}, 内容: {极简版_list}")
                print(f"  简化版列表类型: {type(简化版_list)}, 内容: {简化版_list}")
                print(f"  详细版列表类型: {type(详细版_list)}, 内容: {详细版_list}")
                
                # 生成字符串内容（CSV中用;分隔多行内容）
                极简版_str = "; ".join([f"{i}. {str(event)}" for i, event in enumerate(极简版_list, 1)] if 极简版_list else [])
                简化版_str = "; ".join([f"{i}. {str(event)}" for i, event in enumerate(简化版_list, 1)] if 简化版_list else [])
                详细版_str = "; ".join([f"{i}. {str(event)}" for i, event in enumerate(详细版_list, 1)] if 详细版_list else [])
                
                data_to_export.append([
                    data.get("标题", ""),
                    极简版_str,
                    简化版_str,
                    详细版_str
                ])
            
            # 创建DataFrame
            df = pd.DataFrame(data_to_export, columns=["书籍名称", "极简版", "简化版", "详细版"])
            
            print(f"\n创建的DataFrame形状: {df.shape}")
            print(f"列名: {list(df.columns)}")
            print("\nDataFrame预览:")
            print(df)
            
            # 导出为CSV
            file_path = f"测试导出_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
            
            print(f"\nCSV文件已导出到: {file_path}")
            
            # 验证导出的文件
            if os.path.exists(file_path):
                print(f"文件大小: {os.path.getsize(file_path)} 字节")
                # 重新读取以验证内容
                df_read = pd.read_csv(file_path, encoding='utf-8-sig')
                print(f"\n重新读取的CSV形状: {df_read.shape}")
                print("重新读取的CSV预览:")
                print(df_read)
                return True
            else:
                print("错误: CSV文件未创建成功")
                return False
                
        except Exception as e:
            print(f"导出CSV失败: {str(e)}")
            return False

# 运行测试
if __name__ == "__main__":
    print("=== CSV导出功能测试 ===")
    exporter = TestCSVExporter()
    exporter.mock_data()
    success = exporter.test_csv_export()
    print(f"\n测试结果: {'成功' if success else '失败'}")
