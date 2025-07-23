#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
简单的API测试脚本
"""

from openai import OpenAI

# 在这里输入你的API密钥
API_KEY = "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"  # 请替换为你的实际API密钥

def test_gpt4v():
    print("测试GPT-4 Vision Preview...")
    
    client = OpenAI(
        api_key=API_KEY,
        base_url="https://openrouter.ai/api/v1"
    )
    
    # 尝试几个可能可用的模型
    models_to_test = [
        "openai/gpt-4-vision-preview",
        "openai/gpt-4o",
        "openai/gpt-4",
        "anthropic/claude-3-haiku"
    ]
    
    for model in models_to_test:
        try:
            print(f"尝试模型: {model}")
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "user", "content": "请回复'API连接成功'"}
                ],
                max_tokens=20
            )
            
            print("✅ 成功！")
            print(f"使用模型: {model}")
            print(f"响应: {response.choices[0].message.content}")
            return True
            
        except Exception as e:
            print(f"❌ 模型 {model} 失败: {e}")
            continue
    
    print("❌ 所有模型都测试失败！")
    return False

if __name__ == "__main__":
    if API_KEY == "YOUR_API_KEY_HERE":
        print("请先在脚本中设置你的API密钥！")
    else:
        test_gpt4v()