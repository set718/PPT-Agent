#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试所有PPT合并器的格式保留效果
"""

import os
import sys
from typing import List, Dict, Any

def create_test_data() -> List[Dict[str, Any]]:
    """创建测试数据"""
    return [
        {
            'page_number': 1,
            'template_number': 1,
            'template_path': os.path.join("templates", "ppt_template", "split_presentations_1.pptx"),
            'template_filename': "split_presentations_1.pptx"
        },
        {
            'page_number': 2,
            'template_number': 5,
            'template_path': os.path.join("templates", "ppt_template", "split_presentations_5.pptx"),
            'template_filename': "split_presentations_5.pptx"
        }
    ]

def test_merger(name: str, func, page_results: List[Dict[str, Any]], output_file: str):
    """测试单个合并器"""
    print(f"\n测试 {name}")
    print("-" * 50)
    
    try:
        result = func(page_results)
        
        if result.get("success"):
            print(f"{name} 成功")
            print(f"   页数: {result.get('total_pages', 0)}")
            print(f"   处理: {result.get('processed_pages', 0)}")
            print(f"   跳过: {result.get('skipped_pages', 0)}")
            
            # 保存测试文件
            if result.get("presentation_bytes"):
                with open(output_file, "wb") as f:
                    f.write(result["presentation_bytes"])
                
                file_size = len(result["presentation_bytes"])
                print(f"   文件大小: {file_size:,} 字节")
                print(f"   保存为: {output_file}")
                
                return {"success": True, "file_size": file_size}
            
        else:
            print(f"{name} 失败: {result.get('error', '未知错误')}")
            return {"success": False, "error": result.get('error')}
            
    except Exception as e:
        print(f"❌ {name} 异常: {str(e)}")
        return {"success": False, "error": str(e)}

def main():
    """主测试函数"""
    print("PPT合并器格式保留效果对比测试")
    print("=" * 60)
    
    page_results = create_test_data()
    
    # 检查测试文件是否存在
    missing_files = []
    for result in page_results:
        if not os.path.exists(result['template_path']):
            missing_files.append(result['template_path'])
    
    if missing_files:
        print("❌ 缺少测试文件:")
        for file in missing_files:
            print(f"   - {file}")
        return
    
    print(f"测试模板: {len(page_results)} 个")
    for result in page_results:
        print(f"   - {result['template_filename']}")
    
    results = {}
    
    # 1. 测试格式基准合并器 (新增：以split_presentations_1.pptx为设计基准)
    try:
        from format_base_merger import merge_with_split_presentations_1_format
        results["格式基准合并"] = test_merger(
            "格式基准合并器",
            merge_with_split_presentations_1_format,
            page_results,
            "test_format_base_merger.pptx"
        )
    except ImportError as e:
        print(f"警告: 格式基准合并器不可用: {e}")

    # 2. 测试增强版合并器
    try:
        from ppt_merger import merge_dify_templates_to_ppt_enhanced
        results["增强版合并"] = test_merger(
            "增强版合并器",
            merge_dify_templates_to_ppt_enhanced,
            page_results,
            "test_enhanced_merger.pptx"
        )
    except ImportError as e:
        print(f"警告: 增强版合并器不可用: {e}")
    
    # 3. 基础合并器已移除（不再测试）
    
    # 4. 测试Win32COM合并器（仅Windows）
    if sys.platform.startswith('win'):
        try:
            from ppt_merger_win32 import merge_dify_templates_to_ppt_win32
            results["Win32COM合并"] = test_merger(
                "Win32COM合并器",
                merge_dify_templates_to_ppt_win32,
                page_results,
                "test_win32_merger.pptx"
            )
        except ImportError as e:
            print(f"警告: Win32COM合并器不可用: {e}")
    
    # 5. 测试Spire合并器
    try:
        from ppt_merger_spire import merge_dify_templates_to_ppt_spire
        results["Spire合并"] = test_merger(
            "Spire合并器",
            merge_dify_templates_to_ppt_spire,
            page_results,
            "test_spire_merger.pptx"
        )
    except ImportError as e:
        print(f"警告: Spire合并器不可用: {e}")
    
    # 输出对比结果
    print("\n合并器效果对比")
    print("=" * 60)
    
    successful_mergers = [(name, result) for name, result in results.items() if result.get("success")]
    failed_mergers = [(name, result) for name, result in results.items() if not result.get("success")]
    
    if successful_mergers:
        print("成功的合并器:")
        # 按文件大小排序（文件大小通常反映格式保留程度）
        successful_mergers.sort(key=lambda x: x[1].get("file_size", 0), reverse=True)
        
        for i, (name, result) in enumerate(successful_mergers, 1):
            file_size = result.get("file_size", 0)
            print(f"   {i}. {name}: {file_size:,} 字节")
            
            # 根据文件大小给出评价
            if file_size > 80000:
                print("      优秀 - 可能保留了丰富的格式信息")
            elif file_size > 50000:
                print("      良好 - 保留了基本格式信息")
            else:
                print("      一般 - 可能丢失了部分格式信息")
    
    if failed_mergers:
        print("\n失败的合并器:")
        for name, result in failed_mergers:
            print(f"   - {name}: {result.get('error', '未知错误')}")
    
    # 推荐使用方案
    print(f"\n推荐方案:")
    if successful_mergers:
        best_merger = successful_mergers[0][0]
        print(f"   最佳选择: {best_merger}")
        
        if "Win32COM" in best_merger:
            print("   安装要求: Windows系统 + pip install pywin32")
        elif "Spire" in best_merger:
            print("   安装要求: pip install Spire.Presentation")
        elif "改进格式保留" in best_merger:
            print("   安装要求: 无额外依赖，基于split_presentations_1.pptx的格式")
        else:
            print("   安装要求: 仅需python-pptx（已安装）")
            
        print(f"\n说明:")
        print(f"   - 改进格式保留合并器会以第一个模板(split_presentations_1.pptx)为基准")
        print(f"   - 保留其颜色、布局、字体等设计元素")
        print(f"   - 后续页面会尽可能匹配第一个模板的风格")
    else:
        print("   请检查模板文件是否存在，或尝试安装额外的依赖包")

if __name__ == "__main__":
    main()
