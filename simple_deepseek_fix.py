#!/usr/bin/env python
# -*- coding: utf-8 -*-

def test_deepseek_config():
    """测试DeepSeek配置"""
    
    print("Testing DeepSeek Configuration")
    print("=" * 40)
    
    try:
        from config import get_config
        
        config = get_config()
        print(f"Current model: {config.ai_model}")
        
        # 检查DeepSeek模型配置
        available_models = config.available_models
        deepseek_models = {k: v for k, v in available_models.items() if 'deepseek' in k.lower()}
        
        print(f"DeepSeek models in config: {list(deepseek_models.keys())}")
        
        for model_name, model_info in deepseek_models.items():
            print(f"\nModel: {model_name}")
            print(f"  Name: {model_info.get('name')}")
            print(f"  Base URL: {model_info.get('base_url')}")
            print(f"  Provider: {model_info.get('api_provider')}")
        
        # 测试切换到DeepSeek模型
        if deepseek_models:
            deepseek_model_name = list(deepseek_models.keys())[0]
            print(f"\nTesting model switch to: {deepseek_model_name}")
            
            try:
                config.set_model(deepseek_model_name)
                print(f"Successfully switched to: {config.ai_model}")
                
                # 获取模型信息
                model_info = config.get_model_info()
                print(f"Model config after switch:")
                print(f"  Base URL: {model_info.get('base_url')}")
                print(f"  Supports vision: {model_info.get('supports_vision')}")
                
                return True
                
            except Exception as e:
                print(f"Failed to switch model: {str(e)}")
                return False
        else:
            print("No DeepSeek models found in config")
            return False
            
    except Exception as e:
        print(f"Config test failed: {str(e)}")
        return False

def show_deepseek_instructions():
    """显示DeepSeek使用说明"""
    
    print("\n" + "=" * 40)
    print("DeepSeek API Usage Instructions")
    print("=" * 40)
    print()
    print("PROBLEM: 'Model Not Exist' error with DeepSeek API")
    print()
    print("SOLUTION:")
    print("1. In the UI, make sure to select 'DeepSeek Chat' model")
    print("2. Enter your DeepSeek API key (format: sk-xxxxxxxx)")
    print("3. The system will use 'deepseek-chat' as model name")
    print()
    print("If still getting errors, try these model names:")
    print("- deepseek-chat (most common)")
    print("- deepseek-coder (for coding tasks)")
    print("- deepseek-reasoner (for reasoning)")
    print()
    print("Check DeepSeek documentation for the latest model names:")
    print("https://platform.deepseek.com/api-docs/")
    print()
    print("IMPORTANT:")
    print("- Make sure your API key is active")
    print("- Verify you have sufficient balance")
    print("- Check if the model name matches DeepSeek's current offerings")

def test_fallback_pagination():
    """测试备用分页功能"""
    
    print("\n" + "=" * 40)
    print("Testing Fallback Pagination")
    print("=" * 40)
    
    try:
        from ai_page_splitter import AIPageSplitter
        
        # 使用测试API密钥
        splitter = AIPageSplitter("test-key")
        
        test_text = """DeepSeek API测试文档

DeepSeek是一个强大的AI模型，支持对话和推理功能。它提供了高质量的文本生成能力。

主要特点包括：强大的推理能力、高效的处理速度、优秀的中文支持。

应用场景广泛，包括对话系统、文本生成、代码辅助等多个领域。"""
        
        result = splitter._create_fallback_split(test_text)
        
        if result.get('success'):
            pages = result.get('pages', [])
            print(f"Pagination successful: {len(pages)} pages")
            
            for i, page in enumerate(pages):
                page_type = page.get('page_type', 'unknown')
                title = page.get('title', 'No title')
                segment = page.get('original_text_segment', '')
                
                print(f"\nPage {i+1} ({page_type}):")
                print(f"  Title: {title}")
                print(f"  Content length: {len(segment)} chars")
                
                if page_type == 'title':
                    print(f"  Content: '{segment}'")
                    if len(segment) < 50:
                        print("  ✅ Title page correctly extracts only title")
                    else:
                        print("  ⚠️ Title page contains too much content")
            
            return True
        else:
            print("Pagination failed")
            return False
            
    except Exception as e:
        print(f"Test failed: {str(e)}")
        return False

if __name__ == "__main__":
    print("DeepSeek API Configuration Test")
    print("=" * 40)
    
    # 测试配置
    config_ok = test_deepseek_config()
    
    # 测试分页功能
    pagination_ok = test_fallback_pagination()
    
    # 显示使用说明
    show_deepseek_instructions()
    
    print("\n" + "=" * 40)
    print("Test Results:")
    print(f"  Configuration: {'OK' if config_ok else 'FAILED'}")
    print(f"  Pagination: {'OK' if pagination_ok else 'FAILED'}")
    
    if config_ok and pagination_ok:
        print("\nStatus: Ready to use DeepSeek API")
        print("Next: Run 'streamlit run user_app.py' and select DeepSeek model")
    else:
        print("\nStatus: Issues detected, check messages above")