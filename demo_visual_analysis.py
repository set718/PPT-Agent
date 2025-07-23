#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPT视觉分析功能演示脚本
展示如何使用GPT-4V分析PPT美观度并提供优化建议
"""

import os
import sys
from datetime import datetime
from config import get_config
from ppt_visual_analyzer import PPTVisualAnalyzer, VisualLayoutOptimizer
from logger import get_logger

def main():
    """主函数"""
    print("=" * 60)
    print("🎨 PPT视觉分析功能演示")
    print("=" * 60)
    print()
    
    # 获取配置和日志
    config = get_config()
    logger = get_logger()
    
    # 检查API密钥
    api_key = input("请输入您的OpenRouter API密钥: ").strip()
    if not api_key:
        print("❌ 需要提供OpenRouter API密钥")
        return
    
    # 检查PPT文件
    ppt_path = config.default_ppt_template
    if not os.path.exists(ppt_path):
        print(f"❌ PPT模板文件不存在: {ppt_path}")
        print("请检查config.py中的default_ppt_template设置")
        return
    
    print(f"📄 使用PPT文件: {ppt_path}")
    print()
    
    try:
        # 初始化视觉分析器
        print("🔧 初始化视觉分析器...")
        visual_analyzer = PPTVisualAnalyzer(api_key)
        
        # 执行视觉分析
        print("🔍 开始分析PPT视觉质量...")
        print("⏳ 这可能需要几分钟时间，请稍候...")
        print()
        
        analysis_result = visual_analyzer.analyze_presentation_visual_quality(ppt_path)
        
        if "error" in analysis_result:
            print(f"❌ 分析失败: {analysis_result['error']}")
            return
        
        # 显示分析结果
        display_analysis_results(analysis_result)
        
        # 询问是否应用优化
        print("\n" + "=" * 60)
        apply_optimization = input("是否要应用布局优化建议? (y/n): ").strip().lower()
        
        if apply_optimization == 'y':
            print("\n🎨 开始应用布局优化...")
            apply_layout_optimizations(visual_analyzer, ppt_path, analysis_result)
        
    except Exception as e:
        logger.error(f"演示过程中出错: {e}")
        print(f"❌ 演示失败: {e}")

def display_analysis_results(analysis_result):
    """显示分析结果"""
    overall_analysis = analysis_result.get("overall_analysis", {})
    slide_analyses = analysis_result.get("slide_analyses", [])
    total_slides = analysis_result.get("total_slides", 0)
    
    print("📊 整体分析结果")
    print("-" * 40)
    
    # 显示总体评分
    scores = overall_analysis.get("scores", {})
    weighted_score = overall_analysis.get("weighted_score", 0)
    grade = overall_analysis.get("grade", "未知")
    
    print(f"📋 总体评分: {weighted_score}/10 ({grade})")
    print()
    
    print("📈 各项评分详情:")
    score_descriptions = {
        "layout_balance": "布局平衡度",
        "color_harmony": "色彩协调性", 
        "typography": "字体排版",
        "visual_hierarchy": "视觉层次",
        "white_space": "留白使用",
        "overall_aesthetics": "整体美观度"
    }
    
    for criterion, score in scores.items():
        desc = score_descriptions.get(criterion, criterion)
        bar = "█" * int(score) + "░" * (10 - int(score))
        print(f"  {desc:12} | {bar} {score:.1f}/10")
    print()
    
    # 显示优点
    strengths = overall_analysis.get("strengths", [])
    if strengths:
        print("✅ 设计优点:")
        for strength in strengths[:3]:  # 显示前3个
            print(f"  • {strength}")
        print()
    
    # 显示待改进点
    weaknesses = overall_analysis.get("weaknesses", [])
    if weaknesses:
        print("⚠️  待改进点:")
        for weakness in weaknesses[:3]:  # 显示前3个
            print(f"  • {weakness}")
        print()
    
    # 显示改进建议
    suggestions = overall_analysis.get("improvement_suggestions", [])
    if suggestions:
        print("💡 改进建议:")
        high_priority = [s for s in suggestions if s.get("priority") == "high"][:3]
        for suggestion in high_priority:
            category = suggestion.get("category", "")
            description = suggestion.get("description", "")
            implementation = suggestion.get("implementation", "")
            print(f"  🔸 {description}")
            if implementation:
                print(f"     实施方法: {implementation}")
        print()
    
    # 显示各页面分析摘要
    if slide_analyses:
        print(f"📑 各页面分析摘要 (共{total_slides}页):")
        print("-" * 40)
        for i, slide_analysis in enumerate(slide_analyses):
            slide_score = slide_analysis.get("weighted_score", 0)
            slide_strengths = slide_analysis.get("strengths", [])
            print(f"  第{i+1}页: {slide_score:.1f}/10")
            if slide_strengths:
                print(f"    优点: {slide_strengths[0] if slide_strengths else '无'}")
        print()

def apply_layout_optimizations(visual_analyzer, ppt_path, analysis_result):
    """应用布局优化"""
    try:
        from pptx import Presentation
        
        # 加载PPT
        presentation = Presentation(ppt_path)
        optimizer = VisualLayoutOptimizer(visual_analyzer)
        
        slide_analyses = analysis_result.get("slide_analyses", [])
        optimizations_applied = []
        
        for slide_analysis in slide_analyses:
            slide_index = slide_analysis.get("slide_index", 0)
            print(f"🔧 优化第{slide_index + 1}页...")
            
            # 应用优化
            optimization_result = optimizer.optimize_slide_layout(
                presentation, slide_index, slide_analysis
            )
            
            if optimization_result.get("success"):
                optimizations = optimization_result.get("optimizations_applied", [])
                optimizations_applied.extend(optimizations)
                for opt in optimizations:
                    print(f"  ✅ {opt}")
            else:
                error = optimization_result.get("error", "未知错误")
                print(f"  ❌ 优化失败: {error}")
        
        # 保存优化后的PPT
        if optimizations_applied:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"output/optimized_ppt_{timestamp}.pptx"
            os.makedirs("output", exist_ok=True)
            presentation.save(output_path)
            
            print(f"\n🎉 优化完成!")
            print(f"📁 已保存优化后的PPT: {output_path}")
            print(f"📊 共应用了{len(optimizations_applied)}项优化")
        else:
            print("\n📋 没有可应用的优化建议")
            
    except Exception as e:
        print(f"❌ 应用优化时出错: {e}")

def test_image_conversion():
    """测试图片转换功能"""
    print("\n🖼️  测试PPT图片转换功能...")
    
    config = get_config()
    ppt_path = config.default_ppt_template
    
    if not os.path.exists(ppt_path):
        print("❌ PPT文件不存在，跳过图片转换测试")
        return False
    
    try:
        # 测试基本图片转换
        api_key = "test"  # 测试用，不需要真实API密钥
        visual_analyzer = PPTVisualAnalyzer(api_key)
        
        print("📸 正在转换PPT页面为图片...")
        image_paths = visual_analyzer.convert_ppt_to_images(ppt_path)
        
        if image_paths:
            print(f"✅ 成功转换{len(image_paths)}张图片")
            for i, path in enumerate(image_paths):
                print(f"  第{i+1}页: {os.path.basename(path)}")
            
            # 清理测试文件
            for path in image_paths:
                try:
                    os.remove(path)
                except:
                    pass
            
            return True
        else:
            print("❌ 图片转换失败")
            return False
            
    except Exception as e:
        print(f"❌ 图片转换测试失败: {e}")
        return False

if __name__ == "__main__":
    # 选择运行模式
    print("选择运行模式:")
    print("1. 完整视觉分析演示 (需要OpenAI API密钥)")
    print("2. 图片转换功能测试 (无需API密钥)")
    
    choice = input("请输入选择 (1/2): ").strip()
    
    if choice == "1":
        main()
    elif choice == "2":
        test_image_conversion()
    else:
        print("无效选择")