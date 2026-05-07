import re

def txt_to_html(input_txt_path, output_html_path):
    # 读取TXT文件
    with open(input_txt_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 转义HTML特殊字符，防止解析错误
    content = content.replace('&', '&amp;')
    content = content.replace('<', '&lt;')
    content = content.replace('>', '&gt;')

    # 1. 处理人名：[]包裹，替换为蓝色加粗span
    content = re.sub(r'\[([^\]]+)\]', 
                    r'<span class="person-name">\1</span>', 
                    content)
    # 2. 处理地名：【】包裹，替换为绿色斜体span
    content = re.sub(r'【([^】]+)】', 
                    r'<span class="place-name">\1</span>', 
                    content)
    # 3. 处理对话：引号包裹（支持"" '' “”），替换为红色等宽字体span
    # 匹配单引号、双引号、中文引号内的内容
    content = re.sub(r'["“]([^"”]+)["”]', 
                    r'<span class="dialogue">\1</span>', 
                    content)
    content = re.sub(r"['‘]([^'’]+)['’]", 
                    r'<span class="dialogue">\1</span>', 
                    content)

    # 替换换行符为HTML的<br>标签
    content = content.replace('\n', '<br>')

    # 构建完整的HTML结构，包含内嵌CSS
    html_template = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>小说文本渲染</title>
    <style>
        body {{
            font-family: "Microsoft YaHei", sans-serif;
            line-height: 1.8;
            margin: 20px auto;
            max-width: 1000px;
            padding: 0 20px;
            background-color: #f5f5f5;
        }}
        /* 主体文本样式 */
        .text-body {{
            font-size: 16px;
            color: #333333;
        }}
        /* 人名样式：蓝色+加粗 */
        .person-name {{
            color: #0000FF;
            font-weight: bold;
        }}
        /* 地名样式：绿色+斜体 */
        .place-name {{
            color: #008000;
            font-style: italic;
        }}
        /* 对话样式：红色+等宽字体 */
        .dialogue {{
            color: #FF0000;
            font-family: "Courier New", monospace;
        }}
    </style>
</head>
<body>
    <div class="text-body">
        {content}
    </div>
</body>
</html>'''

    # 写入HTML文件
    with open(output_html_path, 'w', encoding='utf-8') as f:
        f.write(html_template)

if __name__ == "__main__":
    # 替换为你的TXT文件路径和输出HTML路径
    input_path = "三国演义.txt"  # 你的输入文件
    output_path = "三国演义渲染版.html"  # 输出的HTML文件
    txt_to_html(input_path, output_path)
    print(f"转换完成！已生成文件：{output_path}")