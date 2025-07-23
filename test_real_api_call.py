#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试真实的API调用过程，看看格式化错误具体在哪里发生
"""

import os
from config import get_config
from utils import AIProcessor, PPTAnalyzer
from pptx import Presentation

def test_real_api_call():
    print("=== 测试真实API调用过程 ===")
    
    config = get_config()
    ppt_path = config.default_ppt_template
    
    if not os.path.exists(ppt_path):
        print(f"PPT文件不存在: {ppt_path}")
        return
    
    # 测试文本 - 使用与用户相同的文本
    test_text = "人工智能发展历程"
    
    print(f"测试文本: {test_text}")
    print()
    
    # 使用无效的API密钥来触发错误
    invalid_api_key = "sk-invalid-key-for-testing"
    
    try:
        # 加载PPT结构
        presentation = Presentation(ppt_path)
        ppt_structure = PPTAnalyzer.analyze_ppt_structure(presentation)
        
        print("PPT结构分析完成")
        
        # 创建AI处理器（这里会用无效的API密钥）
        ai_processor = AIProcessor(invalid_api_key)
        
        print("开始AI文本分析...")
        
        # 调用AI分析 - 这里应该会失败并触发备用方案
        assignments = ai_processor.analyze_text_for_ppt(test_text, ppt_structure)
        
        print("AI分析完成，结果:")
        print(f"类型: {type(assignments)}")
        print(f"内容: {assignments}")
        
    except Exception as e:
        print(f"测试过程中出现错误: {e}")
        # 打印完整的错误信息
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_real_api_call()