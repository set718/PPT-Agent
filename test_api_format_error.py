#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试API响应格式化错误
"""

import os
import json
from config import get_config
from utils import AIProcessor, PPTProcessor
from pptx import Presentation

def test_format_error():
    print("=== 测试API格式化错误修复 ===")
    
    config = get_config()
    ppt_path = config.default_ppt_template
    
    if not os.path.exists(ppt_path):
        print(f"PPT文件不存在: {ppt_path}")
        return
    
    # 测试文本
    test_text = "人工智能发展历程"
    
    print(f"测试文本: {test_text}")
    print()
    
    # 模拟JSON响应包含花括号的情况
    mock_json_content = '''
    {
        "assignments": [
            {
                "slide_index": 0,
                "action": "replace_placeholder",
                "placeholder": "title",
                "content": "人工智能发展历程",
                "reason": "提取主题作为标题"
            }
        ]
    }
    '''
    
    # 测试JSON解析
    try:
        parsed = json.loads(mock_json_content)
        print("JSON解析成功")
        print(f"解析结果: {parsed}")
    except Exception as e:
        print(f"JSON解析失败: {e}")
        return
    
    # 加载PPT
    presentation = Presentation(ppt_path)
    ppt_processor = PPTProcessor(presentation)
    
    # 测试assignment应用
    try:
        results = ppt_processor.apply_assignments(parsed)
        print("Assignment应用成功")
        for result in results:
            print(f"  {result}")
    except Exception as e:
        print(f"Assignment应用失败: {e}")
        # 打印完整的错误信息
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_format_error()