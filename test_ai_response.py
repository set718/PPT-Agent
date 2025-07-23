#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试AI响应的脚本
"""

import os
import json
from pptx import Presentation
from config import get_config
from utils import AIProcessor, PPTAnalyzer

def test_ai_response():
    print("测试AI文本分析响应...")
    
    # 获取API密钥
    api_key = input("请输入OpenRouter API密钥: ").strip()
    if not api_key:
        print("需要API密钥")
        return
    
    config = get_config()
    ppt_path = config.default_ppt_template
    
    if not os.path.exists(ppt_path):
        print(f"PPT文件不存在: {ppt_path}")
        return
    
    # 分析PPT结构
    presentation = Presentation(ppt_path)
    ppt_structure = PPTAnalyzer.analyze_ppt_structure(presentation)
    
    print("PPT结构中的占位符:")
    for i, slide_info in enumerate(ppt_structure['slides']):
        print(f"第{i+1}页:")
        for placeholder_name in slide_info['placeholders'].keys():
            print(f"  - {placeholder_name}")
    
    # 初始化AI处理器
    ai_processor = AIProcessor(api_key)
    
    # 测试文本
    test_text = "人工智能发展历程包括符号主义、专家系统和深度学习三个阶段。当前的大语言模型展现出前所未有的能力。"
    
    print(f"\n测试文本: {test_text}")
    print("\n正在调用AI...")
    
    try:
        # 直接调用AI分析
        assignments = ai_processor.analyze_text_for_ppt(test_text, ppt_structure)
        
        print("\nAI返回的完整响应:")
        print(json.dumps(assignments, ensure_ascii=False, indent=2))
        
        # 检查assignments结构
        if "assignments" in assignments:
            assignments_list = assignments["assignments"]
            print(f"\n解析到 {len(assignments_list)} 个分配方案:")
            
            for i, assignment in enumerate(assignments_list):
                print(f"\n分配 {i+1}:")
                print(f"  动作: {assignment.get('action', 'unknown')}")
                print(f"  页码: {assignment.get('slide_index', 0) + 1}")
                print(f"  占位符: {assignment.get('placeholder', 'unknown')}")
                print(f"  内容: '{assignment.get('content', '')}'")
                print(f"  原因: {assignment.get('reason', 'none')}")
        else:
            print("❌ AI响应中没有 'assignments' 字段")
            
    except Exception as e:
        print(f"AI调用失败: {e}")

if __name__ == "__main__":
    test_ai_response()