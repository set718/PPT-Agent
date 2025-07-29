#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
验证Dify API集成设置
"""

def verify_imports():
    """验证所有必要的模块导入"""
    try:
        print("🔍 检查模块导入...")
        
        # 核心模块
        import dify_api_client
        print("✅ dify_api_client 导入成功")
        
        # 具体类导入
        from dify_api_client import DifyAPIConfig, DifyAPIClient, DifyIntegrationService, process_pages_with_dify
        print("✅ Dify API相关类导入成功")
        
        # AI分页模块
        from ai_page_splitter import AIPageSplitter, PageContentFormatter
        print("✅ AI分页模块导入成功")
        
        # 异步HTTP客户端
        import aiohttp
        print("✅ aiohttp 模块导入成功")
        
        return True
    except ImportError as e:
        print(f"❌ 模块导入失败: {e}")
        return False
    except Exception as e:
        print(f"❌ 验证过程出错: {e}")
        return False

def verify_config():
    """验证Dify API配置"""
    try:
        print("\n🔧 检查Dify API配置...")
        
        from dify_api_client import DifyAPIConfig
        config = DifyAPIConfig()
        
        print(f"✅ API服务器: {config.base_url}")
        print(f"✅ API密钥: {config.api_key[:20]}...")
        print(f"✅ 端点: {config.endpoint}")
        print(f"✅ 超时设置: {config.timeout}秒")
        print(f"✅ 最大重试: {config.max_retries}次")
        
        return True
    except Exception as e:
        print(f"❌ 配置检查失败: {e}")
        return False

def verify_integration():
    """验证集成功能"""
    try:
        print("\n🔗 检查集成功能...")
        
        from dify_api_client import DifyIntegrationService
        service = DifyIntegrationService()
        print("✅ Dify集成服务初始化成功")
        
        # 创建模拟数据
        mock_result = {
            "success": True,
            "pages": [
                {
                    "page_number": 1,
                    "page_type": "content",
                    "title": "测试页面",
                    "key_points": ["测试要点1", "测试要点2"],
                    "original_text_segment": "这是测试文本"
                }
            ]
        }
        
        # 测试结果摘要格式化
        mock_processing_result = {
            "success": True,
            "processing_summary": {
                "total_pages": 1,
                "successful_api_calls": 1,
                "failed_api_calls": 0,
                "processing_time": 1.5,
                "success_rate": 100.0
            }
        }
        
        summary = service.format_results_summary(mock_processing_result)
        print("✅ 结果摘要格式化正常")
        
        return True
    except Exception as e:
        print(f"❌ 集成功能检查失败: {e}")
        return False

def main():
    """主验证函数"""
    print("🧪 Dify API集成设置验证")
    print("=" * 50)
    
    # 验证导入
    import_ok = verify_imports()
    
    # 验证配置
    config_ok = verify_config()
    
    # 验证集成
    integration_ok = verify_integration()
    
    print("\n" + "=" * 50)
    print("📋 验证结果:")
    print(f"   模块导入: {'✅ 通过' if import_ok else '❌ 失败'}")
    print(f"   API配置: {'✅ 通过' if config_ok else '❌ 失败'}")
    print(f"   集成功能: {'✅ 通过' if integration_ok else '❌ 失败'}")
    
    if import_ok and config_ok and integration_ok:
        print("\n🎉 所有验证通过！Dify API集成已准备就绪")
        print("\n💡 使用方法:")
        print("   1. 运行: streamlit run user_app.py")
        print("   2. 在'AI智能分页'选项卡中")
        print("   3. 勾选'启用Dify API调用'")
        print("   4. 输入文本并开始处理")
        return True
    else:
        print("\n⚠️ 部分验证未通过，请检查安装和配置")
        return False

if __name__ == "__main__":
    main()