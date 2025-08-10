#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPT格式保留测试脚本
测试不同合并方法对源格式的保留效果
"""

import os
import sys
from typing import List, Dict, Any

def create_test_page_results() -> List[Dict[str, Any]]:
    """创建测试用的页面结果"""
    template_dir = os.path.join("templates", "ppt_template")
    
    # 选择几个不同的模板进行测试
    test_templates = [
        "split_presentations_1.pptx",   # 通常是标题页
        "split_presentations_5.pptx",   # 内容页
        "split_presentations_10.pptx",  # 可能包含图表
        "split_presentations_20.pptx"   # 其他样式
    ]
    
    page_results = []
    for i, template_name in enumerate(test_templates):
        template_path = os.path.join(template_dir, template_name)
        if os.path.exists(template_path):
            page_results.append({
                'page_number': i + 1,
                'template_number': template_name.replace('split_presentations_', '').replace('.pptx', ''),
                'template_path': template_path,
                'template_filename': template_name
            })
        else:
            print(f"⚠️ 模板文件不存在: {template_path}")
    
    return page_results

def test_merge_method(method_name: str, merge_function, page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """测试特定的合并方法"""
    print(f"\n测试合并方法: {method_name}")
    print("=" * 50)
    
    try:
        result = merge_function(page_results)
        
        if result["success"]:
            print(f"{method_name} 合并成功！")
            print(f"   总页数: {result['total_pages']}")
            print(f"   处理页数: {result['processed_pages']}")
            print(f"   跳过页数: {result['skipped_pages']}")
            
            if result.get("output_path"):
                print(f"   输出文件: {result['output_path']}")
            
            if result.get("presentation_bytes"):
                # 保存测试文件
                test_filename = f"test_{method_name.lower().replace(' ', '_')}.pptx"
                with open(test_filename, "wb") as f:
                    f.write(result["presentation_bytes"])
                print(f"   测试文件已保存: {test_filename}")
                
                # 简单的格式检查
                file_size = len(result["presentation_bytes"])
                print(f"   文件大小: {file_size:,} 字节")
                
                # 文件大小可以作为格式复杂度的粗略指标
                if file_size > 50000:  # 50KB以上通常包含较多格式信息
                    print("   文件大小表明可能保留了较多格式信息")
                else:
                    print("   文件较小，可能丢失了部分格式信息")
        else:
            print(f"{method_name} 合并失败: {result.get('error')}")
            if result.get("errors"):
                for error in result["errors"]:
                    print(f"     - {error}")
        
        return result
        
    except Exception as e:
        print(f"❌ {method_name} 测试异常: {str(e)}")
        return {"success": False, "error": str(e)}

def main():
    """主测试函数"""
    print("PPT格式保留测试")
    print("=" * 60)
    
    # 创建测试数据
    page_results = create_test_page_results()
    
    if not page_results:
        print("❌ 没有找到可用的测试模板文件")
        return
    
    print(f"找到 {len(page_results)} 个测试模板:")
    for result in page_results:
        print(f"   - {result['template_filename']}")
    
    results = {}
    
    # 测试增强版合并器
    try:
        from ppt_merger import merge_dify_templates_to_ppt_enhanced
        results["增强版合并器"] = test_merge_method(
            "增强版合并器", merge_dify_templates_to_ppt_enhanced, page_results
        )
    except ImportError as e:
        print(f"警告: 无法导入增强版合并器: {e}")
    
    # 基础合并器已移除（不再测试）
    
    # 测试Win32COM合并器（仅Windows）
    if sys.platform.startswith('win'):
        try:
            from ppt_merger_win32 import merge_dify_templates_to_ppt_win32
            results["Win32COM合并器"] = test_merge_method(
                "Win32COM合并器", merge_dify_templates_to_ppt_win32, page_results
            )
        except ImportError as e:
            print(f"警告: Win32COM合并器不可用: {e}")
    
    # 测试Spire合并器
    try:
        from ppt_merger_spire import merge_dify_templates_to_ppt_spire
        results["Spire合并器"] = test_merge_method(
            "Spire合并器", merge_dify_templates_to_ppt_spire, page_results
        )
    except ImportError as e:
        print(f"警告: Spire合并器不可用: {e}")
    
    # 输出测试总结
    print("\n测试总结")
    print("=" * 50)
    
    successful_methods = [name for name, result in results.items() if result.get("success")]
    failed_methods = [name for name, result in results.items() if not result.get("success")]
    
    if successful_methods:
        print("成功的合并方法:")
        for method in successful_methods:
            result = results[method]
            file_size = len(result.get("presentation_bytes", b"")) if result.get("presentation_bytes") else 0
            print(f"   - {method}: {result['processed_pages']}/{len(page_results)} 页, {file_size:,} 字节")
    
    if failed_methods:
        print("\n失败的合并方法:")
        for method in failed_methods:
            print(f"   - {method}: {results[method].get('error', '未知错误')}")
    
    print(f"\n建议:")
    if "Win32COM合并器" in successful_methods:
        print("   - Win32COM合并器通常提供最佳的格式保留效果")
    elif "Spire合并器" in successful_methods:
        print("   - Spire合并器提供良好的格式保留效果")
    elif "增强版合并器" in successful_methods:
        print("   - 增强版合并器会自动选择最佳可用方法")
    else:
        print("   - 请检查依赖包安装情况，考虑安装 pywin32 或 Spire.Presentation")

if __name__ == "__main__":
    main()
