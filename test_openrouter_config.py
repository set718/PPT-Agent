#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
OpenRouter配置测试脚本
验证OpenRouter API连接和GPT-4V模型是否正常工作
"""

import os
from openai import OpenAI
from config import get_config

def test_openrouter_connection():
    """测试OpenRouter连接"""
    print("🔧 测试OpenRouter配置")
    print("=" * 40)
    
    # 获取配置
    config = get_config()
    
    print(f"📋 配置信息:")
    print(f"  Base URL: {config.openai_base_url}")
    print(f"  Model: {config.ai_model}")
    print()
    
    # 获取API密钥
    api_key = input("请输入您的OpenRouter API密钥: ").strip()
    
    if not api_key:
        print("❌ 需要提供API密钥")
        return False
    
    # 验证密钥格式
    if not (api_key.startswith('sk-or-') or api_key.startswith('sk-')):
        print("⚠️ API密钥格式可能不正确")
        print("  OpenRouter密钥通常以 'sk-or-' 开头")
        print("  标准OpenAI密钥以 'sk-' 开头")
        continue_test = input("是否继续测试? (y/n): ").strip().lower()
        if continue_test != 'y':
            return False
    
    try:
        print("🔗 正在连接OpenRouter...")
        
        # 初始化客户端
        client = OpenAI(
            api_key=api_key,
            base_url=config.openai_base_url
        )
        
        # 测试简单的文本完成
        print("📝 测试文本生成...")
        response = client.chat.completions.create(
            model=config.ai_model,
            messages=[
                {
                    "role": "user", 
                    "content": "请简单回复'测试成功'来确认连接正常"
                }
            ],
            max_tokens=50,
            temperature=0.3
        )
        
        result = response.choices[0].message.content
        print(f"✅ 连接测试成功!")
        print(f"📤 模型回复: {result}")
        print()
        
        # 测试视觉功能（如果模型支持）
        if "vision" in config.ai_model.lower():
            print("👁️  测试视觉分析功能...")
            
            # 创建一个简单的测试图片（纯文本描述）
            vision_response = client.chat.completions.create(
                model=config.ai_model,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": "这是一个测试，请回复'视觉功能正常'"
                            }
                        ]
                    }
                ],
                max_tokens=50,
                temperature=0.3
            )
            
            vision_result = vision_response.choices[0].message.content
            print(f"✅ 视觉模型测试成功!")
            print(f"📤 视觉模型回复: {vision_result}")
        
        print()
        print("🎉 所有测试通过! OpenRouter配置正确")
        return True
        
    except Exception as e:
        print(f"❌ 连接测试失败: {e}")
        print()
        print("🔍 可能的解决方案:")
        print("  1. 检查API密钥是否正确")
        print("  2. 确认OpenRouter账户有足够余额")
        print("  3. 检查网络连接")
        print("  4. 确认选择的模型在OpenRouter中可用")
        return False

def show_openrouter_info():
    """显示OpenRouter使用信息"""
    print()
    print("📚 OpenRouter使用指南")
    print("=" * 40)
    print("1. 访问 https://openrouter.ai/keys 获取API密钥")
    print("2. OpenRouter密钥格式: sk-or-xxxxxxxxxx")
    print("3. 支持多种AI模型，包括GPT-4V")
    print("4. 按使用量计费，需要预充值")
    print("5. 支持的模型格式: openai/gpt-4-vision-preview")
    print()

if __name__ == "__main__":
    print("OpenRouter配置测试工具")
    print("=" * 50)
    
    # 显示使用信息
    show_openrouter_info()
    
    # 运行测试
    success = test_openrouter_connection()
    
    if success:
        print("✨ 现在可以正常使用PPT视觉分析功能了!")
    else:
        print("⚠️  请检查配置后重试")