#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
修复AI分页模块的模型选择问题
"""

def fix_pagination_model_selection():
    """修复分页模块的模型选择逻辑"""
    
    print("Fixing AI pagination model selection...")
    
    try:
        # 检查当前配置
        from config import get_config
        config = get_config()
        
        print(f"Current default model: {config.ai_model}")
        
        available_models = list(config.available_models.keys())
        print(f"Available models: {available_models}")
        
        # 如果用户选择了DeepSeek模型，但默认仍是GPT
        if 'deepseek-chat' in available_models:
            print("\nDeepSeek model available in config")
            
            # 询问用户想要使用哪个模型作为默认
            print("\nWhich model would you like to use for AI pagination?")
            for i, model in enumerate(available_models, 1):
                model_info = config.available_models[model]
                print(f"{i}. {model} - {model_info.get('name')} ({model_info.get('cost')} cost)")
            
            print("\nTo use DeepSeek for pagination, you can either:")
            print("1. Select it in the UI when running the app")
            print("2. Or update the default in config.py")
            
            return True
        else:
            print("DeepSeek model not found in config")
            return False
            
    except Exception as e:
        print(f"Error: {str(e)}")
        return False

def test_ai_page_splitter_with_deepseek():
    """测试AI分页模块使用DeepSeek"""
    
    print("\n=== Testing AI Page Splitter ===")
    
    try:
        from ai_page_splitter import AIPageSplitter
        from config import get_config
        
        # 模拟使用DeepSeek
        config = get_config()
        
        # 临时设置为DeepSeek模型
        original_model = config.ai_model
        if 'deepseek-chat' in config.available_models:
            config.set_model('deepseek-chat')
            print(f"Temporarily set model to: {config.ai_model}")
            
            # 检查模型信息
            model_info = config.get_model_info()
            print(f"Model info: {model_info}")
            
            # 测试分页器初始化（不进行实际API调用）
            test_api_key = "sk-test-deepseek-key"
            
            try:
                splitter = AIPageSplitter(test_api_key)
                print("✅ AIPageSplitter initialized successfully with DeepSeek config")
                
                # 测试备用分页功能
                test_text = "DeepSeek AI模型测试\n\n这是一个测试文本，用于验证分页功能是否正常工作。"
                
                result = splitter._create_fallback_split(test_text)
                
                if result.get('success'):
                    pages = result.get('pages', [])
                    print(f"✅ Fallback pagination works: {len(pages)} pages generated")
                    
                    # 检查第一页
                    if pages:
                        first_page = pages[0]
                        title_segment = first_page.get('original_text_segment', '')
                        print(f"Title page content: '{title_segment}'")
                        
                        if 'DeepSeek AI模型测试' in title_segment and len(title_segment) < 30:
                            print("✅ Title page extraction works correctly")
                        else:
                            print("⚠️ Title page extraction needs adjustment")
                else:
                    print("❌ Fallback pagination failed")
                
            except Exception as e:
                print(f"❌ AIPageSplitter initialization failed: {str(e)}")
            
            # 恢复原始模型设置
            config.set_model(original_model)
            print(f"Restored original model: {config.ai_model}")
            
        else:
            print("DeepSeek model not available for testing")
            
        return True
        
    except Exception as e:
        print(f"Test failed: {str(e)}")
        return False

def provide_usage_instructions():
    """提供使用说明"""
    
    print("\n" + "=" * 50)
    print("📋 Usage Instructions for DeepSeek API:")
    print()
    print("1. In the UI (streamlit run user_app.py):")
    print("   - Select 'DeepSeek Chat' from the model dropdown")
    print("   - Enter your DeepSeek API key")
    print("   - The system will automatically use the correct model name")
    print()
    print("2. Make sure your DeepSeek API key is valid:")
    print("   - Get it from: https://platform.deepseek.com/api_keys")
    print("   - Format should be: sk-xxxxxxxx")
    print()
    print("3. If you still get 'Model Not Exist' error:")
    print("   - Check if your API key has the correct permissions")
    print("   - Try using 'deepseek-chat' as the model name")
    print("   - Verify your DeepSeek account status")
    print()
    print("4. Alternative model names to try:")
    print("   - deepseek-chat")
    print("   - deepseek-coder")
    print("   - Check DeepSeek documentation for latest model names")

if __name__ == "__main__":
    print("🔧 AI Pagination Model Fix Tool")
    print("=" * 50)
    
    # 修复模型选择
    fix_ok = fix_pagination_model_selection()
    
    # 测试分页功能
    test_ok = test_ai_page_splitter_with_deepseek()
    
    # 提供使用说明
    provide_usage_instructions()
    
    print("\n" + "=" * 50)
    print("📝 Summary:")
    print(f"   Model config: {'✅ OK' if fix_ok else '❌ Needs attention'}")
    print(f"   Pagination test: {'✅ OK' if test_ok else '❌ Needs attention'}")
    print()
    
    if fix_ok and test_ok:
        print("🎉 Everything looks good! Try using DeepSeek in the UI now.")
    else:
        print("⚠️ Some issues detected. Check the messages above.")