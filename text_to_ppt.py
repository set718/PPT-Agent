#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文本转PPT工具
使用DeepSeek API处理文本并生成PowerPoint演示文稿
"""

import os
import sys
from datetime import datetime
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
import json
import re

class TextToPPTGenerator:
    def __init__(self, api_key=None, ppt_path=None):
        """
        初始化文本转PPT生成器
        
        Args:
            api_key (str): DeepSeek API密钥
            ppt_path (str): 现有PPT文件路径
        """
        self.api_key = api_key or os.getenv('DEEPSEEK_API_KEY')
        if not self.api_key:
            raise ValueError("请设置DEEPSEEK_API_KEY环境变量或提供API密钥")
        
        # 初始化DeepSeek客户端
        self.client = OpenAI(
            api_key=self.api_key,
            base_url="https://api.deepseek.com"
        )
        
        # 设置PPT文件路径
        self.ppt_path = ppt_path
        if not ppt_path or not os.path.exists(ppt_path):
            raise ValueError(f"PPT文件不存在: {ppt_path}")
        
        # 创建输出目录
        self.output_dir = "output"
        os.makedirs(self.output_dir, exist_ok=True)
        
        # 加载现有PPT
        self.presentation = Presentation(self.ppt_path)
        self.ppt_structure = self.analyze_existing_ppt()
    
    def analyze_existing_ppt(self):
        """
        分析现有PPT的结构
        
        Returns:
            dict: PPT结构信息
        """
        slides_info = []
        for i, slide in enumerate(self.presentation.slides):
            slide_info = {
                "slide_index": i,
                "title": "",
                "text_shapes": [],
                "has_content": False
            }
            
            # 分析幻灯片中的文本框
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    if shape.text.strip():
                        if hasattr(shape, "text_frame") and shape.text_frame:
                            # 这是标题
                            if slide_info["title"] == "" and len(shape.text.strip()) < 100:
                                slide_info["title"] = shape.text.strip()
                            slide_info["has_content"] = True
                    
                    # 记录可编辑的文本形状
                    if hasattr(shape, "text_frame"):
                        slide_info["text_shapes"].append({
                            "shape_id": shape.shape_id if hasattr(shape, "shape_id") else len(slide_info["text_shapes"]),
                            "current_text": shape.text,
                            "shape": shape
                        })
            
            slides_info.append(slide_info)
        
        return {
            "total_slides": len(self.presentation.slides),
            "slides": slides_info
        }
    
    def process_text_with_deepseek(self, user_text):
        """
        使用DeepSeek API分析如何将用户文本填入现有PPT的合适位置
        
        Args:
            user_text (str): 用户输入的文本
            
        Returns:
            dict: 文本分配方案
        """
        # 创建现有PPT结构的描述
        ppt_description = f"现有PPT共有{self.ppt_structure['total_slides']}张幻灯片:\n"
        for slide in self.ppt_structure['slides']:
            ppt_description += f"第{slide['slide_index']+1}页: 标题「{slide['title']}」, "
            ppt_description += f"有{len(slide['text_shapes'])}个文本区域\n"
        
        system_prompt = f"""你是一个专业的PPT内容填充专家。我有一个现有的PPT文件和用户提供的文本，请分析如何将用户文本合理地填入现有PPT的合适位置。

现有PPT结构：
{ppt_description}

**重要原则：**
1. 保持用户原始文本内容完全不变
2. 根据现有幻灯片的标题和结构，合理分配文本内容
3. 如果文本内容超出现有幻灯片容量，可以建议添加新幻灯片

请按照以下JSON格式返回：
{{
  "assignments": [
    {{
      "slide_index": 0,
      "action": "update",
      "content": "要填入该幻灯片的原始文本片段",
      "reason": "选择该幻灯片的原因"
    }},
    {{
      "slide_index": -1,
      "action": "add_new",
      "title": "新幻灯片标题",
      "content": "原始文本片段",
      "reason": "需要新增幻灯片的原因"
    }}
  ]
}}

分析要求：
1. 仔细分析现有PPT的主题和结构
2. 将用户文本按逻辑分段，保持原文不变
3. 为每段文本选择最合适的现有幻灯片，或建议新增幻灯片
4. action可以是"update"（更新现有幻灯片）或"add_new"（新增幻灯片）
5. 提供清晰的分配理由
6. 只返回JSON格式"""
        
        try:
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_text}
                ],
                temperature=0.7,
                max_tokens=2000
            )
            
            content = response.choices[0].message.content
            if content:
                content = content.strip()
            else:
                content = ""
            
            # 提取JSON内容（如果有代码块包围）
            json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
            if json_match:
                content = json_match.group(1)
            
            try:
                return json.loads(content)
            except json.JSONDecodeError:
                # 如果JSON解析失败，创建一个默认分配方案
                return {
                    "assignments": [
                        {
                            "slide_index": 0,
                            "action": "update",
                            "content": user_text,
                            "reason": "默认填入第一张幻灯片"
                        }
                    ]
                }
        
        except Exception as e:
            print(f"调用DeepSeek API时出错: {e}")
            # 返回基础分配方案
            return {
                "assignments": [
                    {
                        "slide_index": 0,
                        "action": "update",
                        "content": user_text,
                        "reason": f"API调用失败，默认填入第一张幻灯片。错误: {e}"
                    }
                ]
            }
    
    def apply_text_assignments(self, assignments):
        """
        根据分配方案修改现有PPT
        
        Args:
            assignments (dict): 文本分配方案
            
        Returns:
            str: 修改后的PPT文件路径
        """
        assignments_list = assignments.get('assignments', [])
        
        for assignment in assignments_list:
            action = assignment.get('action')
            content = assignment.get('content', '')
            
            if action == 'update':
                slide_index = assignment.get('slide_index', 0)
                if 0 <= slide_index < len(self.presentation.slides):
                    slide = self.presentation.slides[slide_index]
                    self.update_slide_content(slide, content)
                    print(f"✓ 已更新第{slide_index+1}页: {assignment.get('reason', '')}")
                
            elif action == 'add_new':
                title = assignment.get('title', '新增内容')
                self.add_new_slide(title, content)
                print(f"✓ 已新增幻灯片「{title}」: {assignment.get('reason', '')}")
        
        # 保存修改后的PPT
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"updated_ppt_{timestamp}.pptx"
        filepath = os.path.join(self.output_dir, filename)
        
        self.presentation.save(filepath)
        return filepath
    
    def update_slide_content(self, slide, content):
        """
        更新幻灯片内容
        """
        # 查找可用的文本框
        text_shapes = []
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                text_shapes.append(shape)
        
        if text_shapes:
            # 使用第一个可用的文本框（通常是主要内容区域）
            target_shape = text_shapes[-1] if len(text_shapes) > 1 else text_shapes[0]
            
            # 清空现有内容并添加新内容
            tf = target_shape.text_frame
            tf.clear()
            
            # 添加内容
            p = tf.paragraphs[0]
            p.text = content
            p.font.size = Pt(16)
    
    def add_new_slide(self, title, content):
        """
        添加新幻灯片
        """
        # 使用标题和内容布局
        slide_layout = self.presentation.slide_layouts[1]
        slide = self.presentation.slides.add_slide(slide_layout)
        
        # 设置标题
        if slide.shapes.title:
            slide.shapes.title.text = title
        
        # 设置内容
        if len(slide.placeholders) > 1:
            content_placeholder = slide.placeholders[1]
            tf = content_placeholder.text_frame
            tf.clear()
            
            p = tf.paragraphs[0]
            p.text = content
            p.font.size = Pt(16)
    
    def generate_ppt_from_text(self, user_text):
        """
        将用户文本填入现有PPT的完整流程
        
        Args:
            user_text (str): 用户输入的文本
            
        Returns:
            str: 修改后的PPT文件路径
        """
        print("正在使用DeepSeek API分析文本结构...")
        assignments = self.process_text_with_deepseek(user_text)
        
        print("正在将您的原始文本填入现有PPT...")
        filepath = self.apply_text_assignments(assignments)
        
        return filepath

def main():
    """主函数"""
    print("=" * 50)
    print("        文本转PPT填充器")
    print("    使用DeepSeek AI将文本填入现有PPT")
    print("=" * 50)
    
    # 设置PPT文件路径
    ppt_path = r"D:\jiayihan\Desktop\ppt format V1_2.pptx"
    
    # 检查PPT文件是否存在
    if not os.path.exists(ppt_path):
        print(f"\n⚠️  指定的PPT文件不存在:")
        print(f"路径: {ppt_path}")
        
        # 询问是否使用测试文件
        print("\n是否创建并使用测试PPT文件进行演示？(y/n)")
        choice = input().strip().lower()
        
        if choice in ['y', 'yes', '是', '是的']:
            # 创建测试PPT
            from create_test_ppt import create_test_ppt
            ppt_path = create_test_ppt()
        else:
            print("\n程序退出。请确认PPT文件路径或创建测试文件。")
            sys.exit(1)
    else:
        print(f"\n✅ 已找到PPT文件: {os.path.basename(ppt_path)}")
    
    # 检查API密钥
    api_key = os.getenv('DEEPSEEK_API_KEY')
    if not api_key:
        print("\n❌ 错误：未找到DEEPSEEK_API_KEY环境变量")
        print("\n请设置您的DeepSeek API密钥：")
        print("方法1: 设置环境变量")
        print("   export DEEPSEEK_API_KEY=your_api_key_here")
        print("\n方法2: 创建.env文件")
        print("   DEEPSEEK_API_KEY=your_api_key_here")
        print("\n获取API密钥：https://platform.deepseek.com/api_keys")
        sys.exit(1)
    
    try:
        # 初始化生成器
        generator = TextToPPTGenerator(api_key, ppt_path)
        
        # 显示现有PPT信息
        ppt_info = generator.ppt_structure
        print(f"\n📊 PPT信息:")
        print(f"   总共 {ppt_info['total_slides']} 张幻灯片")
        for slide in ppt_info['slides'][:3]:  # 只显示前3张
            title = slide['title'] if slide['title'] else "（无标题）"
            print(f"   第{slide['slide_index']+1}页: {title}")
        if ppt_info['total_slides'] > 3:
            print(f"   ... 还有 {ppt_info['total_slides']-3} 张幻灯片")
        
        print("\n请输入您想要填入PPT的文本内容：")
        print("(输入'quit'或'exit'退出)")
        print("-" * 50)
        
        while True:
            try:
                user_input = input("\n请输入文本: ").strip()
                
                if user_input.lower() in ['quit', 'exit', '退出']:
                    print("\n感谢使用！再见！")
                    break
                
                if not user_input:
                    print("请输入有效的文本内容。")
                    continue
                
                # 填入PPT
                filepath = generator.generate_ppt_from_text(user_input)
                
                print(f"\n✅ PPT更新成功！")
                print(f"📁 文件路径: {filepath}")
                print(f"📊 您可以在 {os.path.abspath(filepath)} 找到更新的PPT文件")
                
                print("\n是否继续添加更多文本？(y/n)")
                continue_choice = input().strip().lower()
                if continue_choice in ['n', 'no', '否', '不']:
                    print("\n感谢使用！再见！")
                    break
                    
            except KeyboardInterrupt:
                print("\n\n程序被用户中断。再见！")
                break
                
            except Exception as e:
                print(f"\n❌ 生成过程中出现错误: {e}")
                print("请重试或检查您的输入。")
    
    except Exception as e:
        print(f"\n❌ 初始化失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 