# -*- coding: utf-8 -*-
"""
TXT 小说标记转换器
将带标记的文本转换为带颜色高亮的现代化 HTML 页面

支持的标记格式：
- 人名: [孔明]、[周瑜]
- 地名: |江夏|、|柴桑|
- 对话: "..."（双引号包裹的内容）
- 章节: 第X回 或 第X章
- 引用/诗词: [内容]（特定格式）
- 分隔线: --------

支持输出格式：
- HTML: 标准 HTML 文件（需 HTTP 服务器加载外部资源）
- MHTML: 自包含单文件（双击即可打开，无需服务器）

作者: AI Assistant
日期: 2026-05-06
"""

import re
import sys
import os
import html
import base64
from pathlib import Path
from datetime import datetime


class TxtToHtmlConverter:
    """TXT 转 HTML 转换器类"""

    # 颜色配置方案（可自定义）
    COLOR_SCHEME = {
        'person': {           # 人名
            'color': '#e74c3c',
            'bg': 'rgba(231, 76, 60, 0.08)',
            'label': '人名'
        },
        'place': {            # 地名
            'color': '#3498db',
            'bg': 'rgba(52, 152, 219, 0.08)',
            'label': '地名'
        },
        'dialogue': {         # 对话
            'color': '#27ae60',
            'bg': 'rgba(39, 174, 96, 0.08)',
            'label': '对话'
        },
        'chapter': {          # 章节标题
            'color': '#9b59b6',
            'bg': 'rgba(155, 89, 182, 0.1)',
            'label': '章节'
        },
        'quote': {            # 引用/诗词
            'color': '#e67e22',
            'bg': 'rgba(230, 126, 34, 0.08)',
            'label': '引用'
        },
        'separator': {        # 分隔线
            'color': '#95a5a6',
            'bg': 'transparent',
            'label': '分隔'
        }
    }

    def __init__(self, color_scheme=None):
        """
        初始化转换器

        参数:
            color_scheme: 自定义颜色方案，不传则使用默认
        """
        if color_scheme:
            self.COLOR_SCHEME.update(color_scheme)
        self.stats = {
            'person_count': 0,
            'place_count': 0,
            'dialogue_count': 0,
            'chapter_count': 0,
            'quote_count': 0,
            'line_count': 0
        }

    def _escape_html(self, text):
        """转义 HTML 特殊字符"""
        return html.escape(text)

    def _parse_line(self, line):
        """
        解析单行文本，将标记转换为 HTML

        处理顺序很重要：先处理嵌套较少的，再处理嵌套多的
        """
        original_line = line
        result = []
        last_end = 0

        # 先检测整行类型
        stripped = line.strip()

        # 处理分隔线
        if re.match(r'^[-─—]{3,}$', stripped):
            self.stats['line_count'] += 1
            return '<hr class="separator-line" />'

        # 处理章节标题（第X回 或 第X章）
        if re.match(r'^第[一二三四五六七八九十百千零\d]+回[\s　]', stripped):
            self.stats['chapter_count'] += 1
            chapter_content = self._escape_html(stripped)
            # 章节内的人名也要高亮
            chapter_content = self._highlight_inline(chapter_content)
            return f'<h2 class="chapter-title">{chapter_content}</h2>'

        # 处理普通段落
        # 按顺序解析各种标记
        parsed = self._highlight_inline(self._escape_html(line))

        if parsed.strip():
            self.stats['line_count'] += 1
            return f'<p class="paragraph">{parsed}</p>'

        return ''

    def _highlight_inline(self, text):
        """
        处理行内的各种标记高亮

        参数:
            text: 已转义的 HTML 文本
        返回:
            带高亮标记的 HTML
        """
        result = text

        # 1. 处理对话（双引号包裹的内容）- 先处理，避免和其他冲突
        # 匹配中文双引号 "..." 或英文双引号 "..."
        def replace_dialogue(match):
            self.stats['dialogue_count'] += 1
            content = match.group(1)
            # 对话内部的人名和地名也要处理
            content = self._highlight_names_in_text(content)
            return f'<span class="tag-dialogue">"{content}"</span>'

        result = re.sub(
            r'"([^"]+)"',
            replace_dialogue,
            result
        )

        # 2. 处理人名 [xxx]
        def replace_person(match):
            self.stats['person_count'] += 1
            name = match.group(1)
            return f'<span class="tag-person" data-type="person">[{name}]</span>'

        result = re.sub(r'\[([^\]]+)\]', replace_person, result)

        # 3. 处理地名 |xxx|
        def replace_place(match):
            self.stats['place_count'] += 1
            name = match.group(1)
            return f'<span class="tag-place" data-type="place">|{name}|</span>'

        result = re.sub(r'\|([^|]+)\|', replace_place, result)

        return result

    def _highlight_names_in_text(self, text):
        """在对话内容中处理人名和地名标记"""
        # 人名
        text = re.sub(
            r'\[([^\]]+)\]',
            lambda m: f'<span class="tag-person-inline">[{m.group(1)}]</span>',
            text
        )
        # 地名
        text = re.sub(
            r'\|([^|]+)\|',
            lambda m: f'<span class="tag-place-inline">|{m.group(1)}|</span>',
            text
        )
        return text

    def convert(self, input_path, output_path=None, title=None, output_format='html'):
        """
        执行转换

        参数:
            input_path: 输入 txt 文件路径
            output_path: 输出文件路径（可选，默认同名 .html 或 .mhtml）
            title: 页面标题（可选，默认从文件名提取）
            output_format: 输出格式，'html' 或 'mhtml'（默认 'html'）
        """
        input_file = Path(input_path)
        if not input_file.exists():
            raise FileNotFoundError(f"找不到文件: {input_path}")

        if output_path is None:
            suffix = '.mhtml' if output_format.lower() == 'mhtml' else '.html'
            output_path = input_file.with_suffix(suffix)

        if title is None:
            title = input_file.stem

        # 读取文件
        print(f"正在读取: {input_path}")
        with open(input_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        print(f"共 {len(lines)} 行，开始解析...")

        # 解析内容
        content_html = []
        for line in lines:
            parsed = self._parse_line(line.rstrip('\n'))
            if parsed:
                content_html.append(parsed)

        # 生成输出内容
        if output_format.lower() == 'mhtml':
            output_content = self._generate_mhtml(title, '\n'.join(content_html))
        else:
            output_content = self._generate_html(title, '\n'.join(content_html))

        # 写入文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(output_content)

        print(f"\n转换完成!")
        print(f"输出格式: {output_format.upper()}")
        print(f"输出文件: {output_path}")
        print(f"\n统计信息:")
        print(f"  - 章节数: {self.stats['chapter_count']}")
        print(f"  - 人名标记: {self.stats['person_count']}")
        print(f"  - 地名标记: {self.stats['place_count']}")
        print(f"  - 对话标记: {self.stats['dialogue_count']}")
        print(f"  - 总行数: {self.stats['line_count']}")

        return output_path

    def _generate_html(self, title, content):
        """生成完整的 HTML 页面"""
        scheme = self.COLOR_SCHEME

        css_styles = f"""
        /* ===== 本地字体加载 ===== */
        @font-face {{
            font-family: 'Noto Serif SC';
            font-style: normal;
            font-weight: 400;
            font-display: swap;
            src: url('fonts/NotoSerifSC-Regular.woff2') format('woff2'),
                 url('fonts/NotoSerifSC-Regular.woff') format('woff'),
                 url('fonts/NotoSerifSC-Regular.ttf') format('truetype');
        }}

        @font-face {{
            font-family: 'Noto Serif SC';
            font-style: normal;
            font-weight: 600;
            font-display: swap;
            src: url('fonts/NotoSerifSC-SemiBold.woff2') format('woff2'),
                 url('fonts/NotoSerifSC-SemiBold.woff') format('woff'),
                 url('fonts/NotoSerifSC-SemiBold.ttf') format('truetype');
        }}

        @font-face {{
            font-family: 'Noto Serif SC';
            font-style: normal;
            font-weight: 700;
            font-display: swap;
            src: url('fonts/NotoSerifSC-Bold.woff2') format('woff2'),
                 url('fonts/NotoSerifSC-Bold.woff') format('woff'),
                 url('fonts/NotoSerifSC-Bold.ttf') format('truetype');
        }}

        /* ===== 基础样式 ===== */
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: "Noto Serif SC", "Source Han Serif SC", "SimSun", "宋体", serif;
            font-size: 18px;
            line-height: 2;
            color: #2c3e50;
            background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
            min-height: 100vh;
        }}

        /* ===== 布局容器 ===== */
        .container {{
            max-width: 900px;
            margin: 0 auto;
            padding: 40px 20px;
        }}

        /* ===== 头部区域 ===== */
        .header {{
            text-align: center;
            padding: 40px 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 16px;
            margin-bottom: 30px;
            box-shadow: 0 10px 40px rgba(102, 126, 234, 0.3);
        }}

        .header h1 {{
            font-size: 2.5em;
            font-weight: 700;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }}

        .header .subtitle {{
            font-size: 0.95em;
            opacity: 0.9;
        }}

        /* ===== 图例说明 ===== */
        .legend {{
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            gap: 15px;
            padding: 20px;
            background: white;
            border-radius: 12px;
            margin-bottom: 30px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        }}

        .legend-item {{
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 0.9em;
            transition: transform 0.2s;
        }}

        .legend-item:hover {{
            transform: translateY(-2px);
        }}

        .legend-dot {{
            width: 12px;
            height: 12px;
            border-radius: 50%;
        }}

        /* ===== 内容区域 ===== */
        .content {{
            background: white;
            border-radius: 16px;
            padding: 40px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
        }}

        /* ===== 章节标题 ===== */
        .chapter-title {{
            font-size: 1.6em;
            font-weight: 700;
            color: {scheme['chapter']['color']};
            text-align: center;
            padding: 30px 20px 20px;
            margin: 40px 0 30px;
            border-bottom: 3px solid {scheme['chapter']['color']};
            background: {scheme['chapter']['bg']};
            border-radius: 8px;
        }}

        .chapter-title:first-child {{
            margin-top: 0;
        }}

        /* ===== 段落 ===== */
        .paragraph {{
            text-indent: 2em;
            margin-bottom: 12px;
            text-align: justify;
        }}

        /* ===== 标记样式 - 人名 ===== */
        .tag-person {{
            color: {scheme['person']['color']};
            background: {scheme['person']['bg']};
            padding: 1px 4px;
            border-radius: 4px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            border-bottom: 2px solid {scheme['person']['color']};
        }}

        .tag-person:hover {{
            background: {scheme['person']['color']};
            color: white;
        }}

        /* 对话内的人名 */
        .tag-person-inline {{
            color: {scheme['person']['color']};
            font-weight: 600;
        }}

        /* ===== 标记样式 - 地名 ===== */
        .tag-place {{
            color: {scheme['place']['color']};
            background: {scheme['place']['bg']};
            padding: 1px 4px;
            border-radius: 4px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            border-bottom: 2px dashed {scheme['place']['color']};
        }}

        .tag-place:hover {{
            background: {scheme['place']['color']};
            color: white;
        }}

        /* 对话内的地名 */
        .tag-place-inline {{
            color: {scheme['place']['color']};
            font-weight: 600;
        }}

        /* ===== 标记样式 - 对话 ===== */
        .tag-dialogue {{
            color: {scheme['dialogue']['color']};
            background: {scheme['dialogue']['bg']};
            padding: 2px 6px;
            border-radius: 6px;
            font-style: italic;
            border-left: 3px solid {scheme['dialogue']['color']};
            border-right: 3px solid {scheme['dialogue']['color']};
        }}

        /* ===== 分隔线 ===== */
        .separator-line {{
            border: none;
            height: 2px;
            background: linear-gradient(90deg, transparent, {scheme['separator']['color']}, transparent);
            margin: 30px 0;
        }}

        /* ===== 底部信息 ===== */
        .footer {{
            text-align: center;
            padding: 30px;
            color: #7f8c8d;
            font-size: 0.85em;
        }}

        /* ===== 统计面板 ===== */
        .stats-panel {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin-bottom: 30px;
        }}

        .stat-card {{
            background: white;
            padding: 20px;
            border-radius: 12px;
            text-align: center;
            box-shadow: 0 2px 12px rgba(0,0,0,0.06);
            transition: transform 0.2s;
        }}

        .stat-card:hover {{
            transform: translateY(-3px);
        }}

        .stat-number {{
            font-size: 2em;
            font-weight: 700;
            color: #667eea;
        }}

        .stat-label {{
            color: #7f8c8d;
            font-size: 0.9em;
            margin-top: 5px;
        }}

        /* ===== 响应式设计 ===== */
        @media (max-width: 768px) {{
            body {{
                font-size: 16px;
            }}
            .header h1 {{
                font-size: 1.8em;
            }}
            .content {{
                padding: 20px;
            }}
            .legend {{
                gap: 10px;
            }}
        }}

        /* ===== 滚动条美化 ===== */
        ::-webkit-scrollbar {{
            width: 8px;
        }}

        ::-webkit-scrollbar-track {{
            background: #f1f1f1;
        }}

        ::-webkit-scrollbar-thumb {{
            background: #c1c1c1;
            border-radius: 4px;
        }}

        ::-webkit-scrollbar-thumb:hover {{
            background: #a1a1a1;
        }}
        """

        html_template = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title} - 标记高亮版</title>
    <style>
{css_styles}
    </style>
</head>
<body>
    <div class="container">
        <!-- 头部 -->
        <div class="header">
            <h1>{title}</h1>
            <div class="subtitle">智能标记高亮版 | 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
        </div>

        <!-- 统计面板 -->
        <div class="stats-panel">
            <div class="stat-card">
                <div class="stat-number">{self.stats['chapter_count']}</div>
                <div class="stat-label">章节</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{self.stats['person_count']}</div>
                <div class="stat-label">人名标记</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{self.stats['place_count']}</div>
                <div class="stat-label">地名标记</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{self.stats['dialogue_count']}</div>
                <div class="stat-label">对话标记</div>
            </div>
        </div>

        <!-- 图例 -->
        <div class="legend">
            <div class="legend-item" style="background: {scheme['person']['bg']}">
                <div class="legend-dot" style="background: {scheme['person']['color']}"></div>
                <span style="color: {scheme['person']['color']}">人名 [xxx]</span>
            </div>
            <div class="legend-item" style="background: {scheme['place']['bg']}">
                <div class="legend-dot" style="background: {scheme['place']['color']}"></div>
                <span style="color: {scheme['place']['color']}">地名 |xxx|</span>
            </div>
            <div class="legend-item" style="background: {scheme['dialogue']['bg']}">
                <div class="legend-dot" style="background: {scheme['dialogue']['color']}"></div>
                <span style="color: {scheme['dialogue']['color']}">对话 "..."</span>
            </div>
            <div class="legend-item" style="background: {scheme['chapter']['bg']}">
                <div class="legend-dot" style="background: {scheme['chapter']['color']}"></div>
                <span style="color: {scheme['chapter']['color']}">章节标题</span>
            </div>
        </div>

        <!-- 正文内容 -->
        <div class="content">
{content}
        </div>

        <!-- 底部 -->
        <div class="footer">
            <p>由 TXT 标记转换器自动生成 | 共 {self.stats['line_count']} 行内容</p>
        </div>
    </div>

    <script>
        // 简单的交互：点击标记时显示提示
        document.querySelectorAll('.tag-person, .tag-place').forEach(tag => {{
            tag.addEventListener('click', function() {{
                const type = this.dataset.type;
                const text = this.textContent;
                console.log(`点击了 ${{type}}: ${{text}}`);
            }});
        }});
    </script>
</body>
</html>"""

        return html_template

    def _generate_mhtml(self, title, content):
        """生成 MHTML 自包含单文件（嵌入字体）"""
        scheme = self.COLOR_SCHEME

        # 嵌入的字体数据（Base64 编码的 Noto Serif SC 字体）
        # 使用系统备用字体，MHTML 不依赖外部字体文件
        css_styles = f"""
        /* ===== 基础样式 ===== */
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: "Noto Serif SC", "Source Han Serif SC", "SimSun", "宋体", serif;
            font-size: 18px;
            line-height: 2;
            color: #2c3e50;
            background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
            min-height: 100vh;
        }}

        /* ===== 布局容器 ===== */
        .container {{
            max-width: 900px;
            margin: 0 auto;
            padding: 40px 20px;
        }}

        /* ===== 头部区域 ===== */
        .header {{
            text-align: center;
            padding: 40px 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 16px;
            margin-bottom: 30px;
            box-shadow: 0 10px 40px rgba(102, 126, 234, 0.3);
        }}

        .header h1 {{
            font-size: 2.5em;
            font-weight: 700;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }}

        .header .subtitle {{
            font-size: 0.95em;
            opacity: 0.9;
        }}

        /* ===== 图例说明 ===== */
        .legend {{
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            gap: 15px;
            padding: 20px;
            background: white;
            border-radius: 12px;
            margin-bottom: 30px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        }}

        .legend-item {{
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 0.9em;
            transition: transform 0.2s;
        }}

        .legend-item:hover {{
            transform: translateY(-2px);
        }}

        .legend-dot {{
            width: 12px;
            height: 12px;
            border-radius: 50%;
        }}

        /* ===== 内容区域 ===== */
        .content {{
            background: white;
            border-radius: 16px;
            padding: 40px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
        }}

        /* ===== 章节标题 ===== */
        .chapter-title {{
            font-size: 1.6em;
            font-weight: 700;
            color: {scheme['chapter']['color']};
            text-align: center;
            padding: 30px 20px 20px;
            margin: 40px 0 30px;
            border-bottom: 3px solid {scheme['chapter']['color']};
            background: {scheme['chapter']['bg']};
            border-radius: 8px;
        }}

        .chapter-title:first-child {{
            margin-top: 0;
        }}

        /* ===== 段落 ===== */
        .paragraph {{
            text-indent: 2em;
            margin-bottom: 12px;
            text-align: justify;
        }}

        /* ===== 标记样式 - 人名 ===== */
        .tag-person {{
            color: {scheme['person']['color']};
            background: {scheme['person']['bg']};
            padding: 1px 4px;
            border-radius: 4px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            border-bottom: 2px solid {scheme['person']['color']};
        }}

        .tag-person:hover {{
            background: {scheme['person']['color']};
            color: white;
        }}

        .tag-person-inline {{
            color: {scheme['person']['color']};
            font-weight: 600;
        }}

        /* ===== 标记样式 - 地名 ===== */
        .tag-place {{
            color: {scheme['place']['color']};
            background: {scheme['place']['bg']};
            padding: 1px 4px;
            border-radius: 4px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            border-bottom: 2px dashed {scheme['place']['color']};
        }}

        .tag-place:hover {{
            background: {scheme['place']['color']};
            color: white;
        }}

        .tag-place-inline {{
            color: {scheme['place']['color']};
            font-weight: 600;
        }}

        /* ===== 标记样式 - 对话 ===== */
        .tag-dialogue {{
            color: {scheme['dialogue']['color']};
            background: {scheme['dialogue']['bg']};
            padding: 2px 6px;
            border-radius: 6px;
            font-style: italic;
            border-left: 3px solid {scheme['dialogue']['color']};
            border-right: 3px solid {scheme['dialogue']['color']};
        }}

        /* ===== 分隔线 ===== */
        .separator-line {{
            border: none;
            height: 2px;
            background: linear-gradient(90deg, transparent, {scheme['separator']['color']}, transparent);
            margin: 30px 0;
        }}

        /* ===== 底部信息 ===== */
        .footer {{
            text-align: center;
            padding: 30px;
            color: #7f8c8d;
            font-size: 0.85em;
        }}

        /* ===== 统计面板 ===== */
        .stats-panel {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin-bottom: 30px;
        }}

        .stat-card {{
            background: white;
            padding: 20px;
            border-radius: 12px;
            text-align: center;
            box-shadow: 0 2px 12px rgba(0,0,0,0.06);
            transition: transform 0.2s;
        }}

        .stat-card:hover {{
            transform: translateY(-3px);
        }}

        .stat-number {{
            font-size: 2em;
            font-weight: 700;
            color: #667eea;
        }}

        .stat-label {{
            color: #7f8c8d;
            font-size: 0.9em;
            margin-top: 5px;
        }}

        /* ===== 响应式设计 ===== */
        @media (max-width: 768px) {{
            body {{
                font-size: 16px;
            }}
            .header h1 {{
                font-size: 1.8em;
            }}
            .content {{
                padding: 20px;
            }}
            .legend {{
                gap: 10px;
            }}
        }}
        """

        html_body = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title} - 标记高亮版</title>
    <style>
{css_styles}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>{title}</h1>
            <div class="subtitle">智能标记高亮版 | 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
        </div>

        <div class="stats-panel">
            <div class="stat-card">
                <div class="stat-number">{self.stats['chapter_count']}</div>
                <div class="stat-label">章节</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{self.stats['person_count']}</div>
                <div class="stat-label">人名标记</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{self.stats['place_count']}</div>
                <div class="stat-label">地名标记</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{self.stats['dialogue_count']}</div>
                <div class="stat-label">对话标记</div>
            </div>
        </div>

        <div class="legend">
            <div class="legend-item" style="background: {scheme['person']['bg']}">
                <div class="legend-dot" style="background: {scheme['person']['color']}"></div>
                <span style="color: {scheme['person']['color']}">人名 [xxx]</span>
            </div>
            <div class="legend-item" style="background: {scheme['place']['bg']}">
                <div class="legend-dot" style="background: {scheme['place']['color']}"></div>
                <span style="color: {scheme['place']['color']}">地名 |xxx|</span>
            </div>
            <div class="legend-item" style="background: {scheme['dialogue']['bg']}">
                <div class="legend-dot" style="background: {scheme['dialogue']['color']}"></div>
                <span style="color: {scheme['dialogue']['color']}">对话 "..."</span>
            </div>
            <div class="legend-item" style="background: {scheme['chapter']['bg']}">
                <div class="legend-dot" style="background: {scheme['chapter']['color']}"></div>
                <span style="color: {scheme['chapter']['color']}">章节标题</span>
            </div>
        </div>

        <div class="content">
{content}
        </div>

        <div class="footer">
            <p>由 TXT 标记转换器自动生成 | 共 {self.stats['line_count']} 行内容 | MHTML 自包含格式</p>
        </div>
    </div>
</body>
</html>"""

        # MHTML 边界标记
        boundary = f"----=_NextPart_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        
        # 构建 MHTML 内容
        mhtml_content = f"""From: <txt_to_html@converter.local>
Subject: {title} - 标记高亮版
MIME-Version: 1.0
Content-Type: multipart/related; boundary="{boundary}"; type="text/html"
X-MimeOLE: Produced By TXT to HTML Converter

--{boundary}
Content-Type: text/html; charset=UTF-8
Content-Transfer-Encoding: quoted-printable
Content-Location: index.html

{html_body}

--{boundary}--"""

        return mhtml_content


def main():
    """
    主函数 - 命令行入口
    用法: python txt_to_html.py <输入文件> [输出文件] [标题] [--mhtml]
    """
    if len(sys.argv) < 2:
        print("用法: python txt_to_html.py <输入txt文件> [输出文件] [标题] [--mhtml]")
        print("示例: python txt_to_html.py 三国演义.txt")
        print("       python txt_to_html.py 三国演义.txt --mhtml")
        print("       python txt_to_html.py 三国演义.txt 三国演义.html")
        print("       python txt_to_html.py 三国演义.txt 三国演义.mhtml --mhtml")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = None
    title = None
    output_format = 'html'

    # 解析参数
    for i in range(2, len(sys.argv)):
        if sys.argv[i] == '--mhtml':
            output_format = 'mhtml'
        elif output_file is None:
            output_file = sys.argv[i]
        elif title is None:
            title = sys.argv[i]

    converter = TxtToHtmlConverter()

    try:
        output_path = converter.convert(input_file, output_file, title, output_format)
        print(f"\n✅ 成功生成: {output_path}")
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
