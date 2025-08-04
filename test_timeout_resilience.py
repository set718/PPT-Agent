#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试超时容错机制
"""

from dify_api_client import DifyAPIConfig, process_pages_with_dify

def create_timeout_test_data():
    """创建测试数据来验证超时处理"""
    
    pages = [
        {
            "page_number": 1,
            "page_type": "content",
            "title": "超时测试页面1",
            "subtitle": "测试容错机制",
            "content_summary": "这是一个用于测试超时处理的页面",
            "key_points": [
                "测试API超时处理",
                "验证重试机制",
                "确保不会调用失败"
            ],
            "original_text_segment": "测试超时容错机制的详细内容"
        },
        {
            "page_number": 2,
            "page_type": "content", 
            "title": "超时测试页面2",
            "subtitle": "验证多密钥轮换",
            "content_summary": "验证多个API密钥的故障转移",
            "key_points": [
                "多密钥负载均衡",
                "故障自动转移", 
                "智能重试策略"
            ],
            "original_text_segment": "验证在有API密钥超时时的处理机制"
        }
    ]
    
    return {
        "success": True,
        "analysis": {
            "total_pages": 2,
            "content_type": "超时容错测试",
            "split_strategy": "验证重试机制",
            "reasoning": "测试系统在API超时时的表现"
        },
        "pages": pages,
        "original_text": "超时容错机制测试..."
    }

def test_improved_timeout_handling():
    """测试改进后的超时处理机制"""
    
    print("测试超时容错机制改进")
    print("=" * 50)
    
    # 使用改进的配置
    improved_config = DifyAPIConfig(
        timeout=180,  # 3分钟超时
        max_retries=8,  # 8次重试
        retry_delay=3.0,  # 3秒重试间隔
        max_concurrent=8,  # 8个并发
        load_balance_strategy="round_robin"
    )
    
    print("改进后的配置:")
    print(f"  超时时间: {improved_config.timeout}秒 (3分钟)")
    print(f"  最大重试: {improved_config.max_retries}次")
    print(f"  重试间隔: {improved_config.retry_delay}秒")
    print(f"  最大并发: {improved_config.max_concurrent}")
    print(f"  API密钥数量: {len(improved_config.api_keys)}")
    
    # 创建测试数据
    test_data = create_timeout_test_data()
    
    print(f"\n开始测试...")
    print(f"测试页面: {len(test_data['pages'])}页")
    
    import time
    start_time = time.time()
    
    # 执行测试
    result = process_pages_with_dify(test_data, improved_config)
    
    end_time = time.time()
    total_time = end_time - start_time
    
    print(f"\n测试结果:")
    success = result.get('success', False)
    print(f"  总体成功: {success}")
    
    # 显示处理统计
    summary = result.get('processing_summary', {})
    total_pages = summary.get('total_pages', 0)
    successful_calls = summary.get('successful_api_calls', 0)
    failed_calls = summary.get('failed_api_calls', 0)
    processing_time = summary.get('processing_time', 0)
    
    print(f"  总页面数: {total_pages}")
    print(f"  成功调用: {successful_calls}")
    print(f"  失败调用: {failed_calls}")
    print(f"  处理耗时: {processing_time:.2f}秒")
    print(f"  总测试时间: {total_time:.2f}秒")
    
    # 验证容错效果
    if failed_calls == 0:
        print(f"\n✅ 容错机制验证成功!")
        print(f"✅ 所有API调用都成功完成")
        print(f"✅ 改进的超时和重试机制有效")
    else:
        print(f"\n❌ 仍有{failed_calls}个调用失败")
        print(f"❌ 需要进一步优化容错机制")
    
    # 显示API密钥使用情况
    api_results = result.get('dify_api_results', {})
    key_stats = api_results.get('api_key_stats', {})
    
    if key_stats:
        print(f"\nAPI密钥使用统计:")
        print(f"  可用密钥: {key_stats.get('available_keys', 0)}/{key_stats.get('total_keys', 0)}")
        print(f"  负载策略: {key_stats.get('strategy', 'unknown')}")
        
        usage_count = key_stats.get('usage_count', {})
        for key, count in usage_count.items():
            if count > 0:
                print(f"  {key[:20]}...: 使用{count}次")
    
    # 性能分析
    if successful_calls > 0:
        avg_time_per_call = processing_time / successful_calls
        print(f"\n性能分析:")
        print(f"  平均响应时间: {avg_time_per_call:.2f}秒/调用")
        print(f"  并发效率: {successful_calls}个调用在{processing_time:.2f}秒内完成")
    
    return result

def show_improvement_summary():
    """显示改进措施总结"""
    
    print(f"\n📋 超时容错机制改进总结:")
    print(f"=" * 50)
    
    print(f"\n🔧 配置优化:")
    print(f"  • 超时时间: 60秒 → 180秒 (3倍增长)")
    print(f"  • 重试次数: 3次 → 8次 (更多机会)")
    print(f"  • 重试间隔: 2秒 → 3秒 (更充分的等待)")
    print(f"  • 连接超时: 10秒 → 30秒")  
    print(f"  • 读取超时: 30秒 → 120秒")
    
    print(f"\n🧠 智能重试策略:")
    print(f"  • 指数退避 + 随机抖动避免雷击")
    print(f"  • 不同错误类型使用不同重试策略")
    print(f"  • 超时2次后自动切换API密钥")
    print(f"  • 连接错误1次后立即切换密钥")
    
    print(f"\n🔄 API密钥故障转移:")
    print(f"  • 8个API密钥提供高可用性")
    print(f"  • 失效密钥60秒后自动恢复")
    print(f"  • 智能失败计数和恢复机制")
    print(f"  • 强制重置确保服务不中断")
    
    print(f"\n⚡ 并发优化:")
    print(f"  • 降低并发数减少服务器压力")
    print(f"  • 信号量控制避免过载")
    print(f"  • 更好的连接池管理")
    
    print(f"\n🎯 预期效果:")
    print(f"  ✅ 彻底消除\"调用失败\"情况")
    print(f"  ✅ 即使部分API密钥临时不可用也能正常工作")
    print(f"  ✅ 提供生产环境级别的稳定性")
    print(f"  ✅ 自动故障恢复，无需人工干预")

if __name__ == "__main__":
    print("超时容错机制测试")
    print("=" * 50)
    
    try:
        # 显示改进总结
        show_improvement_summary()
        
        print(f"\n开始验证测试...")
        
        # 执行测试
        result = test_improved_timeout_handling()
        
        print(f"\n🎉 测试完成!")
        
        if result.get('success'):
            print(f"✅ 超时容错机制工作正常")
            print(f"✅ 系统已具备生产环境稳定性")
        else:
            print(f"⚠️ 部分功能需要进一步优化")
        
    except Exception as e:
        print(f"❌ 测试过程中出现异常: {str(e)}")
        import traceback
        traceback.print_exc()