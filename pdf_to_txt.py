import os
import pdfplumber

def pdf_to_text(pdf_path, output_path):
    # 设置控制台编码为UTF-8以支持中文显示
    os.system('chcp 65001')
    
    print(f"开始处理PDF文件: {pdf_path}")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"PDF文件共有 {total_pages} 页")
            
            text_content = []
            
            # 逐页提取文本
            for i, page in enumerate(pdf.pages):
                if i % 10 == 0:
                    print(f"处理第 {i+1}/{total_pages} 页...")
                
                page_text = page.extract_text()
                if page_text:
                    text_content.append(page_text)
            
            # 合并文本并写入文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(text_content))
            
            print(f"文本提取完成！已保存到: {output_path}")
            print(f"提取的文本总长度: {len(''.join(text_content))} 字符")
            
    except Exception as e:
        print(f"处理过程中出错: {str(e)}")

if __name__ == "__main__":
    # 输入PDF路径和输出TXT路径
    pdf_file = "d:\\AI\\books\\事实 (汉斯·罗斯林  欧拉·罗斯林  安娜·罗斯林·罗朗德)\\事实 (汉斯·罗斯林  欧拉·罗斯林  安娜·罗斯林·罗朗德).pdf"
    txt_file = "d:\\AI\\books\\事实 (汉斯·罗斯林  欧拉·罗斯林  安娜·罗斯林·罗朗德)\\事实_转换后.txt"
    
    # 检查PDF文件是否存在
    if not os.path.exists(pdf_file):
        print(f"错误：找不到PDF文件 {pdf_file}")
    else:
        # 确保输出目录存在
        os.makedirs(os.path.dirname(txt_file), exist_ok=True)
        
        # 调用转换函数
        pdf_to_text(pdf_file, txt_file)
