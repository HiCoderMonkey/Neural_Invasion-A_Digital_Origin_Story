#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown 转 Word 文档转换脚本
将 Markdown 文件转换为 Word (.docx) 格式
"""

import re
import sys
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


def parse_markdown_to_docx(md_file_path, output_file_path=None):
    """
    将 Markdown 文件转换为 Word 文档
    
    Args:
        md_file_path: Markdown 文件路径
        output_file_path: 输出 Word 文件路径（可选，默认与输入文件同名）
    """
    # 读取 Markdown 文件
    md_path = Path(md_file_path)
    if not md_path.exists():
        print(f"错误：文件不存在 - {md_file_path}")
        return False
    
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 创建 Word 文档
    doc = Document()
    
    # 设置中文字体
    doc.styles['Normal'].font.name = '宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    doc.styles['Normal'].font.size = Pt(12)
    
    # 设置标题样式
    heading1_style = doc.styles['Heading 1']
    heading1_style.font.name = '黑体'
    heading1_style._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    heading1_style.font.size = Pt(18)
    heading1_style.font.bold = True
    
    heading2_style = doc.styles['Heading 2']
    heading2_style.font.name = '黑体'
    heading2_style._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    heading2_style.font.size = Pt(16)
    heading2_style.font.bold = True
    
    heading3_style = doc.styles['Heading 3']
    heading3_style.font.name = '黑体'
    heading3_style._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    heading3_style.font.size = Pt(14)
    heading3_style.font.bold = True
    
    # 按行处理内容
    lines = content.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        # 空行
        if not line.strip():
            if i < len(lines) - 1:  # 不是最后一行
                doc.add_paragraph()
            i += 1
            continue
        
        # 一级标题 (# 或 ##)
        if line.startswith('# '):
            text = line[2:].strip()
            # 处理标题中的粗体
            para = doc.add_heading(level=1)
            add_formatted_text(para, text)
            i += 1
            continue
        
        # 二级标题 (##)
        if line.startswith('## '):
            text = line[3:].strip()
            para = doc.add_heading(level=2)
            add_formatted_text(para, text)
            i += 1
            continue
        
        # 三级标题 (###)
        if line.startswith('### '):
            text = line[4:].strip()
            para = doc.add_heading(level=3)
            add_formatted_text(para, text)
            i += 1
            continue
        
        # 普通段落
        para = doc.add_paragraph()
        add_formatted_text(para, line)
        i += 1
    
    # 保存文档
    if output_file_path is None:
        output_file_path = md_path.with_suffix('.docx')
    else:
        output_file_path = Path(output_file_path)
    
    doc.save(str(output_file_path))
    print(f"✓ 转换成功！")
    print(f"  输入文件：{md_path}")
    print(f"  输出文件：{output_file_path}")
    return True


def add_formatted_text(paragraph, text):
    """
    向段落添加格式化的文本（支持粗体、链接等）
    
    Args:
        paragraph: docx 段落对象
        text: 要添加的文本
    """
    # 处理链接 [文本](URL)
    link_pattern = r'\[([^\]]+)\]\(([^\)]+)\)'
    
    # 检查是否有链接
    if re.search(link_pattern, text):
        # 有链接，分段处理
        last_end = 0
        
        for match in re.finditer(link_pattern, text):
            # 添加链接前的文本
            if match.start() > last_end:
                before_text = text[last_end:match.start()]
                add_bold_text(paragraph, before_text)
            
            # 添加链接文本（只显示文本，不添加超链接）
            link_text = match.group(1)
            add_bold_text(paragraph, link_text)
            
            last_end = match.end()
        
        # 添加剩余文本
        if last_end < len(text):
            remaining_text = text[last_end:]
            add_bold_text(paragraph, remaining_text)
    else:
        # 没有链接，直接处理粗体
        add_bold_text(paragraph, text)


def add_bold_text(paragraph, text):
    """
    向段落添加文本，处理粗体格式
    
    Args:
        paragraph: docx 段落对象
        text: 要添加的文本
    """
    bold_pattern = r'\*\*([^\*]+)\*\*'
    last_end = 0
    
    for match in re.finditer(bold_pattern, text):
        # 添加粗体前的普通文本
        if match.start() > last_end:
            paragraph.add_run(text[last_end:match.start()])
        
        # 添加粗体文本
        run = paragraph.add_run(match.group(1))
        run.bold = True
        
        last_end = match.end()
    
    # 添加剩余文本
    if last_end < len(text):
        paragraph.add_run(text[last_end:])


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("使用方法：")
        print(f"  python {sys.argv[0]} <markdown文件路径> [输出文件路径]")
        print("\n示例：")
        print(f"  python {sys.argv[0]} 言情版/正文/正文.md")
        print(f"  python {sys.argv[0]} 言情版/正文/正文.md 输出.docx")
        sys.exit(1)
    
    md_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    # 检查是否安装了 python-docx
    try:
        import docx
    except ImportError:
        print("错误：未安装 python-docx 库")
        print("请运行以下命令安装：")
        print("  pip install python-docx")
        sys.exit(1)
    
    success = parse_markdown_to_docx(md_file, output_file)
    
    if success:
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == '__main__':
    main()

