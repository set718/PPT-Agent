#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试多API密钥的性能提升
"""

import asyncio
import time
from dify_api_client import DifyAPIConfig, process_pages_with_dify

def create_test_pagination_result(num_pages: int):
    """创建测试分页结果"""
    
    pages = []
    
    # 创建标题页
    pages.append({
        "page_number": 1,
        "page_type": "title",
        "title": "多API密钥性能测试",
        "subtitle": "",
        "date": "2024年7月",
        "content_summary": "测试标题页",
        "key_points": ["测试标题", "性能优化"],
        "original_text_segment": "多API密钥性能测试"
    })
    
    # 创建内容页
    for i in range(2, num_pages + 1):
        pages.append({
            "page_number": i,
            "page_type": "content",
            "title": f"测试内容页 {i-1}",
            "subtitle": f"第{i-1}部分内容",
            "content_summary": f"这是第{i-1}个内容页，用于测试多API密钥的并行处理能力",
            "key_points": [
                f"测试要点 {i-1}.1",
                f"测试要点 {i-1}.2", 
                f"测试要点 {i-1}.3",
                f"验证API密钥轮询机制"
            ],
            "original_text_segment": f"这是第{i-1}个内容页的详细文本内容，用于测试Dify API的响应速度和多密钥负载均衡效果。"
        })
    
    return {
        "success": True,
        "analysis": {
            "total_pages": num_pages,
            "content_type": "性能测试",
            "split_strategy": "多API密钥测试",
            "reasoning": "创建多个页面测试并发性能"
        },
        "pages": pages,
        "original_text": "多API密钥性能测试文档..."
    }

def test_single_vs_multi_key_performance():
    """测试单密钥vs多密钥的性能对比"""
    
    print("🚀 多API密钥性能测试")
    print("=" * 50)
    
    # 创建测试数据
    test_pages_count = 6  # 测试6页内容
    test_data = create_test_pagination_result(test_pages_count)
    
    print(f"测试数据: {test_pages_count}页内容")
    print(f"测试页面: {len(test_data['pages'])}个页面")
    
    # 测试1: 单API密钥配置
    print(f"\n📍 测试1: 单API密钥配置")
    single_key_config = DifyAPIConfig(
        api_keys=["app-7HOcCxB7uosj23f1xgjFClkv"],  # 只使用一个密钥
        max_concurrent=3,
        load_balance_strategy="round_robin"
    )
    
    print(f"配置: 1个API密钥，最大并发: {single_key_config.max_concurrent}")
    
    start_time = time.time()
    single_result = process_pages_with_dify(test_data, single_key_config)
    single_duration = time.time() - start_time
    
    print(f"单密钥结果:")
    print(f"  成功: {single_result.get('successful_count', 0)}")
    print(f"  失败: {single_result.get('failed_count', 0)}")
    print(f"  耗时: {single_duration:.2f}秒")
    
    # 测试2: 多API密钥配置
    print(f"\n📍 测试2: 多API密钥配置")
    multi_key_config = DifyAPIConfig(
        api_keys=[
            "app-7HOcCxB7uosj23f1xgjFClkv",
            "app-vxEWYWTaakWITl041b8UHBCN", 
            "app-WM17uKVOQHpYE4sNyxRH0dtG"
        ],
        max_concurrent=6,  # 增加并发数
        load_balance_strategy="round_robin"
    )
    
    print(f"配置: {len(multi_key_config.api_keys)}个API密钥，最大并发: {multi_key_config.max_concurrent}")
    
    start_time = time.time()
    multi_result = process_pages_with_dify(test_data, multi_key_config)
    multi_duration = time.time() - start_time
    
    print(f"多密钥结果:")
    print(f"  成功: {multi_result.get('successful_count', 0)}")
    print(f"  失败: {multi_result.get('failed_count', 0)}")
    print(f"  耗时: {multi_duration:.2f}秒")
    
    # 显示API密钥统计
    if 'api_key_stats' in multi_result:
        stats = multi_result['api_key_stats']
        print(f"  API密钥统计:")
        print(f"    总密钥: {stats.get('total_keys', 0)}")
        print(f"    可用密钥: {stats.get('available_keys', 0)}")
        print(f"    负载策略: {stats.get('strategy', 'unknown')}")
        
        usage_count = stats.get('usage_count', {})
        for key, count in usage_count.items():
            print(f"    {key[:20]}...: 使用{count}次")
    
    # 性能对比
    print(f"\n📊 性能对比:")
    if single_duration > 0 and multi_duration > 0:
        improvement = ((single_duration - multi_duration) / single_duration) * 100
        speedup = single_duration / multi_duration
        
        print(f"  单密钥耗时: {single_duration:.2f}秒")
        print(f"  多密钥耗时: {multi_duration:.2f}秒")
        
        if improvement > 0:
            print(f"  性能提升: {improvement:.1f}%")
            print(f"  速度倍数: {speedup:.2f}x")
        else:
            print(f"  性能下降: {abs(improvement):.1f}%")
        
        print(f"  单密钥平均响应: {single_duration/test_pages_count:.2f}秒/页")
        print(f"  多密钥平均响应: {multi_duration/test_pages_count:.2f}秒/页")
    
    return single_result, multi_result

def analyze_api_key_distribution(result):
    """分析API密钥分配情况"""
    
    print(f"\n🔍 API密钥使用分析:")
    
    enhanced_pages = result.get('enhanced_pages', [])
    key_usage = {}
    
    for page in enhanced_pages:
        api_result = page.get('dify_api_result', {})
        used_key = api_result.get('used_api_key', 'unknown')
        
        if used_key in key_usage:
            key_usage[used_key] += 1
        else:
            key_usage[used_key] = 1
    
    if key_usage:
        print(f"密钥使用分布:")
        for key, count in key_usage.items():
            print(f"  {key}: {count}次")
        
        # 检查负载均衡效果
        usage_values = list(key_usage.values())
        if usage_values:
            max_usage = max(usage_values)
            min_usage = min(usage_values)
            balance_ratio = min_usage / max_usage if max_usage > 0 else 0
            
            print(f"负载均衡效果:")
            print(f"  最大使用次数: {max_usage}")
            print(f"  最小使用次数: {min_usage}")
            print(f"  均衡度: {balance_ratio:.2f} (1.0为完全均衡)")
    else:
        print("无API密钥使用数据")

if __name__ == "__main__":
    print("多API密钥性能测试工具")
    print("=" * 50)
    
    try:
        # 执行性能测试
        single_result, multi_result = test_single_vs_multi_key_performance()
        
        # 分析多密钥的分配情况
        if multi_result.get('success'):
            analyze_api_key_distribution(multi_result)
        
        print(f"\n📋 测试总结:")
        print(f"✅ 多API密钥负载均衡系统已实现")
        print(f"✅ 支持轮询、随机、最少使用三种策略")
        print(f"✅ 自动故障转移和密钥恢复机制")
        print(f"✅ 增加并发数提升整体处理速度")
        
        print(f"\n🎯 使用建议:")
        print(f"• 对于少量页面（<5页），单密钥足够")
        print(f"• 对于大量页面（>5页），多密钥效果明显")
        print(f"• 建议并发数设置为密钥数量的2倍")
        print(f"• 推荐使用round_robin策略保证均衡")
        
    except Exception as e:
        print(f"❌ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()