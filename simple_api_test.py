#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
简单的API测试脚本 - 测试GPT-4o和其他模型
"""

from openai import OpenAI

# 在这里输入你的OpenRouter API密钥
API_KEY = "YOUR_API_KEY_HERE"  # 请替换为你的实际OpenRouter API密钥

def test_api_models():
    print("测试GPT-4o和其他可用模型...")
    
    client = OpenAI(
        api_key=API_KEY,
        base_url="https://openrouter.ai/api/v1"
    )
    
    # 按优先级测试可用的模型
    models_to_test = [
        "openai/gpt-4o",              # 首选：GPT-4o
        "openai/gpt-4",               # GPT-4
        "openai/gpt-4-turbo",         # GPT-4 Turbo
        "anthropic/claude-3-haiku"    # 备选：Claude
    ]
    
    for model in models_to_test:
        try:
            print(f"尝试模型: {model}")
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "user", "content": "请回复'模型连接成功'"}
                ],
                max_tokens=20
            )
            
            print("✅ 测试成功！")
            print(f"✅ 推荐使用模型: {model}")
            print(f"✅ AI响应: {response.choices[0].message.content}")
            print(f"✅ 此模型可用于你的PPT项目")
            return True
            
        except Exception as e:
            print(f"❌ 模型 {model} 失败: {e}")
            continue
    
    print("❌ 所有模型都测试失败！")
    print("请检查：")
    print("1. API密钥是否正确")
    print("2. 网络连接是否正常")
    print("3. 账户余额是否充足")
    return False

if __name__ == "__main__":
    if API_KEY == "YOUR_API_KEY_HERE":
        print("请先在脚本中设置你的API密钥！")
    else:
        test_api_models()