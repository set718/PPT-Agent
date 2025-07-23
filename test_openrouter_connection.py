#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试OpenRouter连接和模型可用性
"""

import requests
from openai import OpenAI

def test_openrouter_connection(api_key):
    """测试OpenRouter连接"""
    print("=== OpenRouter连接测试 ===")
    
    if not api_key:
        print("❌ 请提供有效的API密钥")
        return False
    
    client = OpenAI(
        api_key=api_key,
        base_url="https://openrouter.ai/api/v1"
    )
    
    # 测试1: 尝试简单的文本生成（不使用vision）
    print("\n1. 测试基础文本生成...")
    try:
        response = client.chat.completions.create(
            model="openai/gpt-3.5-turbo",  # 先用便宜的模型测试
            messages=[
                {"role": "user", "content": "请回复'连接成功'"}
            ],
            max_tokens=10
        )
        print(f"✅ 基础连接成功: {response.choices[0].message.content}")
    except Exception as e:
        print(f"❌ 基础连接失败: {e}")
        return False
    
    # 测试2: 检查GPT-4 Vision模型
    print("\n2. 测试GPT-4 Vision模型...")
    try:
        response = client.chat.completions.create(
            model="openai/gpt-4o",
            messages=[
                {"role": "user", "content": "请回复'GPT-4V连接成功'"}
            ],
            max_tokens=20
        )
        print(f"✅ GPT-4 Vision连接成功: {response.choices[0].message.content}")
        return True
    except Exception as e:
        print(f"❌ GPT-4 Vision连接失败: {e}")
        
        # 尝试其他可能的GPT-4V模型名称
        alternative_models = [
            "openai/gpt-4o",
            "openai/gpt-4-turbo",
            "openai/gpt-4o-mini"
        ]
        
        print("\n尝试其他GPT-4模型...")
        for model in alternative_models:
            try:
                print(f"测试模型: {model}")
                response = client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "user", "content": "请回复'模型可用'"}
                    ],
                    max_tokens=10
                )
                print(f"✅ {model} 可用: {response.choices[0].message.content}")
                print(f"建议将配置中的模型更改为: {model}")
                return True
            except Exception as alt_e:
                print(f"❌ {model} 不可用: {alt_e}")
        
        return False

def test_account_info(api_key):
    """测试账户信息"""
    print("\n=== 账户信息测试 ===")
    
    try:
        # 使用OpenRouter的API查看账户信息
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        response = requests.get("https://openrouter.ai/api/v1/auth/key", headers=headers)
        
        if response.status_code == 200:
            data = response.json()
            print(f"✅ 账户验证成功")
            print(f"用户ID: {data.get('data', {}).get('id', 'N/A')}")
            print(f"余额: ${data.get('data', {}).get('usage', 'N/A')}")
        else:
            print(f"❌ 账户验证失败: {response.status_code} - {response.text}")
            
    except Exception as e:
        print(f"❌ 账户信息获取失败: {e}")

if __name__ == "__main__":
    print("OpenRouter连接诊断工具")
    print("=" * 50)
    print("请使用您的有效API密钥运行以下测试：")
    print()
    print("Python测试代码：")
    print("```python")
    print("from test_openrouter_connection import test_openrouter_connection, test_account_info")
    print("api_key = 'your-api-key-here'")
    print("test_account_info(api_key)")
    print("test_openrouter_connection(api_key)")
    print("```")
    print()
    print("或者，您可以手动测试以下模型是否可用：")
    print("- openai/gpt-4o (当前配置，推荐)")
    print("- openai/gpt-4-turbo")
    print("- openai/gpt-4-vision-preview (已弃用)")
    print("- openai/gpt-4o-mini")
    print()
    print("常见404错误原因：")
    print("1. 模型名称已更新或不存在")
    print("2. API密钥没有访问该模型的权限")
    print("3. 账户余额不足")
    print("4. 模型暂时不可用")
    
    # 显示当前配置
    from config import get_config
    config = get_config()
    print(f"\n当前配置的模型: {config.ai_model}")
    print(f"当前配置的API端点: {config.openai_base_url}")