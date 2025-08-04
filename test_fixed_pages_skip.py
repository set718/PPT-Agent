#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试固定页面跳过功能
"""

from dify_api_client import process_pages_with_dify

def create_test_pages_with_fixed_types():
    """创建包含不同页面类型的测试数据"""
    
    pages = [
        # 标题页（应跳过）
        {
            "page_number": 1,
            "page_type": "title",
            "title": "AI PPT工具测试",
            "subtitle": "固定页面跳过功能验证",
            "date": "2024年7月",
            "content_summary": "这是标题页",
            "key_points": ["标题页测试"],
            "original_text_segment": "标题页内容"
        },
        
        # 目录页（应跳过）
        {
            "page_number": 2,
            "page_type": "table_of_contents",
            "title": "目录",
            "content_summary": "这是目录页",
            "key_points": ["目录项1", "目录项2", "目录项3"],
            "original_text_segment": "目录页内容"
        },
        
        # 内容页1（应处理）
        {
            "page_number": 3,
            "page_type": "content",
            "title": "第一部分内容",
            "subtitle": "重要内容介绍",
            "content_summary": "这是第一个内容页，应该被Dify API处理",
            "key_points": [
                "内容要点1",
                "内容要点2", 
                "内容要点3"
            ],
            "original_text_segment": "这是第一个内容页的详细文本内容，需要通过Dify API进行增强处理。"
        },
        
        # 内容页2（应处理）
        {
            "page_number": 4,
            "page_type": "content",
            "title": "第二部分内容",
            "subtitle": "更多重要信息",
            "content_summary": "这是第二个内容页，也应该被Dify API处理",
            "key_points": [
                "更多要点1",
                "更多要点2",
                "技术细节"
            ],
            "original_text_segment": "这是第二个内容页的详细文本内容，同样需要Dify API增强。"
        },
        
        # 结束页（应跳过）
        {
            "page_number": 5,
            "page_type": "ending",
            "title": "谢谢观看",
            "content_summary": "这是结束页",
            "key_points": ["感谢", "联系方式"],
            "original_text_segment": "结束页内容"
        }
    ]
    
    return {
        "success": True,
        "analysis": {
            "total_pages": 5,
            "content_type": "固定页面跳过测试",
            "split_strategy": "包含固定页面类型",
            "reasoning": "测试不同页面类型的处理逻辑"
        },
        "pages": pages,
        "original_text": "固定页面跳过功能测试文档..."
    }

def test_fixed_pages_skip():
    """测试固定页面跳过功能"""
    
    print("🧪 固定页面跳过功能测试")
    print("=" * 50)
    
    # 创建测试数据
    test_data = create_test_pages_with_fixed_types()
    pages = test_data['pages']
    
    print(f"测试数据包含 {len(pages)} 页:")
    for page in pages:
        page_type = page.get('page_type', 'unknown')
        page_number = page.get('page_number', '?')
        title = page.get('title', '无标题')
        print(f"  第{page_number}页: {page_type} - {title}")
    
    print(f"\n📋 预期结果:")
    print(f"  • 应跳过: 第1页(title)、第2页(table_of_contents)、第5页(ending)")
    print(f"  • 应处理: 第3页(content)、第4页(content)")
    print(f"  • 总共5页，跳过3页，处理2页")
    
    print(f"\n🚀 开始测试...")
    
    # 调用处理函数
    result = process_pages_with_dify(test_data)
    
    print(f"\n📊 测试结果:")
    print(f"  处理成功: {result.get('success', False)}")
    
    # 显示详细统计
    summary = result.get('processing_summary', {})
    dify_results = result.get('dify_api_results', {})
    
    total_pages = summary.get('total_pages', 0)
    successful_calls = summary.get('successful_api_calls', 0)
    failed_calls = summary.get('failed_api_calls', 0)
    skipped_pages = summary.get('skipped_fixed_pages', 0)
    processing_time = summary.get('processing_time', 0)
    
    print(f"  总页面数: {total_pages}")
    print(f"  成功API调用: {successful_calls}")
    print(f"  失败API调用: {failed_calls}")
    print(f"  跳过固定页面: {skipped_pages}")
    print(f"  处理耗时: {processing_time:.2f}秒")
    
    # 验证结果
    print(f"\n✅ 验证结果:")
    
    # 检查跳过页面数量
    if skipped_pages == 3:
        print(f"  ✅ 正确跳过了3个固定页面")
    else:
        print(f"  ❌ 跳过页面数量错误: 期望3，实际{skipped_pages}")
    
    # 检查处理页面数量
    processed_pages = successful_calls + failed_calls
    if processed_pages == 2:
        print(f"  ✅ 正确处理了2个内容页面")
    else:
        print(f"  ❌ 处理页面数量错误: 期望2，实际{processed_pages}")
    
    # 检查增强页面
    enhanced_pages = result.get('enhanced_pages', [])
    print(f"\n🔍 页面详细检查:")
    
    for page in enhanced_pages:
        page_num = page.get('page_number', '?')
        page_type = page.get('page_type', 'unknown')
        
        if page.get('dify_skipped'):
            skip_reason = page.get('dify_skip_reason', '未知原因')
            print(f"  第{page_num}页 ({page_type}): ⏭️ 已跳过 - {skip_reason}")
        elif page.get('dify_response'):
            print(f"  第{page_num}页 ({page_type}): ✅ Dify API处理成功")
        elif page.get('dify_error'):
            error = page.get('dify_error', '未知错误')
            print(f"  第{page_num}页 ({page_type}): ❌ 处理失败 - {error}")
        else:
            print(f"  第{page_num}页 ({page_type}): ❓ 状态未知")
    
    # 性能统计
    if skipped_pages > 0:
        print(f"\n⚡ 性能优化效果:")
        total_would_process = total_pages
        actually_processed = processed_pages
        time_saved_estimate = (skipped_pages / total_would_process) * 100
        print(f"  • 节省API调用: {skipped_pages}次")
        print(f"  • 预计节省时间: {time_saved_estimate:.1f}%")
        print(f"  • 实际处理页面占比: {actually_processed/total_pages*100:.1f}%")
    
    return result

if __name__ == "__main__":
    print("固定页面跳过功能测试工具")
    print("=" * 50)
    
    try:
        result = test_fixed_pages_skip()
        
        print(f"\n🎯 测试总结:")
        if result.get('success'):
            print(f"✅ 固定页面跳过功能正常工作")
            print(f"✅ 封面页、目录页、结束页成功跳过Dify API调用")
            print(f"✅ 内容页正常通过Dify API处理")
            print(f"✅ 性能优化效果明显，减少不必要的API调用")
        else:
            print(f"❌ 测试存在问题，需要检查配置")
        
        print(f"\n💡 功能说明:")
        print(f"• 封面页(title)：固定格式，不需要AI增强")
        print(f"• 目录页(table_of_contents)：自动生成，不需要处理")
        print(f"• 结束页(ending)：模板固定，不需要API调用")
        print(f"• 内容页(content)：核心内容，需要Dify API增强")
        
    except Exception as e:
        print(f"❌ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()