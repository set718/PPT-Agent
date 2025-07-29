#!/usr/bin/env python
# -*- coding: utf-8 -*-

def test_multi_key_config():
    """测试多密钥配置"""
    
    print("Testing Multi-Key Configuration")
    print("=" * 40)
    
    try:
        from dify_api_client import DifyAPIConfig, APIKeyBalancer
        
        # 创建多密钥配置
        config = DifyAPIConfig()
        
        print(f"Total API keys: {len(config.api_keys)}")
        for i, key in enumerate(config.api_keys, 1):
            print(f"  Key {i}: {key[:20]}...")
        
        print(f"Load balance strategy: {config.load_balance_strategy}")
        print(f"Max concurrent: {config.max_concurrent}")
        
        # 测试负载均衡器
        balancer = APIKeyBalancer(config.api_keys, config.load_balance_strategy)
        
        print(f"\nTesting key selection (10 requests):")
        for i in range(10):
            key = balancer.get_next_key()
            print(f"  Request {i+1}: {key[:20]}...")
        
        # 显示统计信息
        stats = balancer.get_usage_stats()
        print(f"\nUsage statistics:")
        print(f"  Total keys: {stats['total_keys']}")
        print(f"  Available keys: {stats['available_keys']}")
        print(f"  Strategy: {stats['strategy']}")
        
        usage_count = stats['usage_count']
        for key, count in usage_count.items():
            print(f"  {key[:20]}...: {count} times")
        
        return True
        
    except Exception as e:
        print(f"Test failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_multi_key_config()
    
    if success:
        print("\n✅ Multi-key configuration working!")
        print("\nTo use in UI:")
        print("1. Run: streamlit run user_app.py")
        print("2. Enable 'Dify API调用' in AI pagination")
        print("3. The system will automatically use all 3 API keys")
        print("4. Requests will be distributed using round-robin")
    else:
        print("\n❌ Configuration test failed")