#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
创建测试PPT文件
"""

from pptx import Presentation
from pptx.util import Pt
import os

def create_test_ppt():
    """创建一个测试用的PPT文件"""
    # 创建新演示文稿
    prs = Presentation()
    
    # 添加标题页
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = "测试演示文稿"
    subtitle.text = "用于文本填充测试"
    
    # 添加内容页1
    bullet_slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(bullet_slide_layout)
    title = slide.shapes.title
    title.text = "第一部分"
    
    content = slide.placeholders[1]
    content.text = "这里是第一部分的内容区域，可以被替换"
    
    # 添加内容页2
    slide = prs.slides.add_slide(bullet_slide_layout)
    title = slide.shapes.title
    title.text = "第二部分"
    
    content = slide.placeholders[1]
    content.text = "这里是第二部分的内容区域，可以被替换"
    
    # 保存文件
    test_ppt_path = "test_template.pptx"
    prs.save(test_ppt_path)
    
    print(f"✅ 测试PPT文件已创建: {os.path.abspath(test_ppt_path)}")
    return os.path.abspath(test_ppt_path)

if __name__ == "__main__":
    create_test_ppt() 