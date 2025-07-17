#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
工具函数模块
包含项目中的共用工具函数
"""

import os
import re
import json
import tempfile
import time
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from config import get_config
from ppt_beautifier import PPTBeautifier

class PPTAnalyzer:
    """PPT分析器"""
    
    @staticmethod
    def analyze_ppt_structure(presentation: Presentation) -> Dict[str, Any]:
        """
        分析PPT结构，提取占位符和文本信息
        
        Args:
            presentation: PPT演示文稿对象
            
        Returns:
            Dict: PPT结构信息
        """
        slides_info = []
        
        for i, slide in enumerate(presentation.slides):
            slide_info = {
                "slide_index": i,
                "title": "",
                "placeholders": {},
                "text_shapes": [],
                "has_content": False
            }
            
            # 分析幻灯片中的文本框和占位符
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    current_text = shape.text.strip()
                    if current_text:
                        # 检查是否包含占位符
                        placeholder_pattern = r'\{([^}]+)\}'
                        placeholders = re.findall(placeholder_pattern, current_text)
                        
                        if placeholders:
                            # 这个文本框包含占位符
                            for placeholder in placeholders:
                                slide_info["placeholders"][placeholder] = {
                                    "shape": shape,
                                    "original_text": current_text,
                                    "placeholder": placeholder
                                }
                        
                        # 如果是简短文本且没有占位符，可能是标题
                        if not placeholders and len(current_text) < 100:
                            if slide_info["title"] == "":
                                slide_info["title"] = current_text
                        
                        slide_info["has_content"] = True
                    
                    # 记录所有可编辑的文本形状
                    if hasattr(shape, "text_frame"):
                        slide_info["text_shapes"].append({
                            "shape_id": shape.shape_id if hasattr(shape, "shape_id") else len(slide_info["text_shapes"]),
                            "current_text": shape.text,
                            "shape": shape,
                            "has_placeholder": bool(re.findall(r'\{([^}]+)\}', shape.text)) if shape.text else False
                        })
            
            slides_info.append(slide_info)
        
        return {
            "total_slides": len(presentation.slides),
            "slides": slides_info
        }

class AIProcessor:
    """AI处理器"""
    
    def __init__(self, api_key: str = None):
        """初始化AI处理器"""
        config = get_config()
        self.api_key = api_key or config.deepseek_api_key
        if not self.api_key:
            raise ValueError("请设置DeepSeek API密钥")
        
        self.client = OpenAI(
            api_key=self.api_key,
            base_url=config.deepseek_base_url
        )
        self.config = config
    
    def analyze_text_for_ppt(self, user_text: str, ppt_structure: Dict[str, Any]) -> Dict[str, Any]:
        """
        使用AI分析文本并生成PPT填充方案
        
        Args:
            user_text: 用户输入的文本
            ppt_structure: PPT结构信息
            
        Returns:
            Dict: 文本分配方案
        """
        # 创建PPT结构描述
        ppt_description = self._create_ppt_description(ppt_structure)
        
        # 构建系统提示
        system_prompt = self._build_system_prompt(ppt_description)
        
        try:
            response = self.client.chat.completions.create(
                model=self.config.ai_model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_text}
                ],
                temperature=self.config.ai_temperature,
                max_tokens=self.config.ai_max_tokens
            )
            
            content = response.choices[0].message.content
            if content:
                content = content.strip()
            else:
                content = ""
            
            # 提取JSON内容
            return self._extract_json_from_response(content, user_text)
            
        except Exception as e:
            print(f"调用AI API时出错: {e}")
            return self._create_fallback_assignment(user_text, str(e))
    
    def _create_ppt_description(self, ppt_structure: Dict[str, Any]) -> str:
        """创建PPT结构描述"""
        description = f"现有PPT共有{ppt_structure['total_slides']}张幻灯片:\n"
        
        for slide in ppt_structure['slides']:
            description += f"\n第{slide['slide_index']+1}页:"
            if slide['title']:
                description += f" 标题「{slide['title']}」"
            
            # 列出所有占位符
            if slide['placeholders']:
                description += f"\n  包含占位符: "
                for placeholder_name in slide['placeholders'].keys():
                    description += f"{{{placeholder_name}}} "
                description += "\n"
            else:
                description += f" (无占位符)\n"
        
        return description
    
    def _build_system_prompt(self, ppt_description: str) -> str:
        """构建系统提示"""
        return f"""你是一个专业的PPT内容优化专家。请分析用户文本，并根据PPT模板结构进行智能适配和优化。

现有PPT结构：
{ppt_description}

**核心任务：**
1. **结构化适配**：根据PPT模板的占位符结构，将用户文本进行合理的结构化调整
2. **内容优化**：可以适当精简、重组或格式化文本，使其更适合PPT呈现
3. **语言润色**：可以优化语言表达，使其更加简洁明了，适合幻灯片展示

**操作原则：**
- ✅ **可以做的**：重新组织文本结构、精简冗余内容、优化表达方式、调整语言风格
- ✅ **可以做的**：根据占位符特点调整内容长度和格式（如将长段落拆分为要点）
- ❌ **不能做的**：添加用户未提供的信息、编造数据、从外部知识添加内容
- ❌ **不能做的**：改变用户文本的核心意思和关键信息

**占位符语义规则：**
- `title` = 主标题或文档标题（简洁有力）
- `subtitle` = 副标题（补充说明）
- `content_X` = 分类标题、章节标题、时间点等结构性内容（清晰明确）
- `content_X_bullet_Y` = 属于特定content的具体要点（简洁扼要）
- `bullet_X` = 独立的要点列表（重点突出）
- `description` = 描述性文字（详细但不冗长）
- `conclusion` = 结论性内容（总结性强）

请按照以下JSON格式返回：
{{
  "assignments": [
    {{
      "slide_index": 0,
      "action": "replace_placeholder",
      "placeholder": "title",
      "content": "优化后的标题内容",
      "reason": "提炼核心概念，适配标题占位符"
    }}
  ]
}}

分析要求：
1. 基于用户文本进行结构化分析和适配优化
2. 根据占位符语义特点调整内容呈现方式
3. 保持核心信息完整，但可优化表达形式
4. action必须是"replace_placeholder"
5. placeholder必须是模板中实际存在的占位符名称
6. 只返回JSON格式，不要其他文字"""
    
    def _extract_json_from_response(self, content: str, user_text: str) -> Dict[str, Any]:
        """从AI响应中提取JSON"""
        # 提取JSON内容（如果有代码块包围）
        json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
        if json_match:
            content = json_match.group(1)
        
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            print(f"AI返回的JSON格式有误，内容：{content}")
            return self._create_fallback_assignment(user_text, "JSON解析失败")
    
    def _create_fallback_assignment(self, user_text: str, error_msg: str) -> Dict[str, Any]:
        """创建备用分配方案"""
        return {
            "assignments": [
                {
                    "slide_index": 0,
                    "action": "replace_placeholder",
                    "placeholder": "content",
                    "content": user_text,
                    "reason": f"API调用失败或解析错误，默认填入content占位符。错误: {error_msg}"
                }
            ]
        }

class PPTProcessor:
    """PPT处理器"""
    
    def __init__(self, presentation: Presentation):
        """初始化PPT处理器"""
        self.presentation = presentation
        self.ppt_structure = PPTAnalyzer.analyze_ppt_structure(presentation)
        self.beautifier = PPTBeautifier(presentation)
        self.filled_placeholders = {}  # 记录已填充的占位符
    
    def apply_assignments(self, assignments: Dict[str, Any]) -> List[str]:
        """
        应用文本分配方案
        
        Args:
            assignments: 分配方案
            
        Returns:
            List[str]: 处理结果列表
        """
        assignments_list = assignments.get('assignments', [])
        results = []
        
        for assignment in assignments_list:
            action = assignment.get('action')
            content = assignment.get('content', '')
            slide_index = assignment.get('slide_index', 0)
            
            if action == 'replace_placeholder':
                placeholder = assignment.get('placeholder', '')
                if 0 <= slide_index < len(self.presentation.slides):
                    slide = self.presentation.slides[slide_index]
                    slide_info = self.ppt_structure['slides'][slide_index]
                    
                    # 检查该占位符是否存在
                    if placeholder in slide_info['placeholders']:
                        success = self._replace_placeholder_in_slide(
                            slide_info['placeholders'][placeholder], 
                            content
                        )
                        if success:
                            # 记录已填充的占位符
                            if slide_index not in self.filled_placeholders:
                                self.filled_placeholders[slide_index] = set()
                            self.filled_placeholders[slide_index].add(placeholder)
                            
                            results.append(f"✓ 已替换第{slide_index+1}页的 {{{placeholder}}} 占位符: {assignment.get('reason', '')}")
                        else:
                            results.append(f"✗ 替换第{slide_index+1}页的 {{{placeholder}}} 占位符失败")
                    else:
                        results.append(f"✗ 第{slide_index+1}页不存在 {{{placeholder}}} 占位符")
                else:
                    results.append(f"✗ 幻灯片索引 {slide_index+1} 超出范围")
            
            elif action == 'update':
                if 0 <= slide_index < len(self.presentation.slides):
                    slide = self.presentation.slides[slide_index]
                    self._update_slide_content(slide, content)
                    results.append(f"✓ 已更新第{slide_index+1}页: {assignment.get('reason', '')}")
                
            elif action == 'add_new':
                title = assignment.get('title', '新增内容')
                self._add_new_slide(title, content)
                results.append(f"✓ 已新增幻灯片「{title}」: {assignment.get('reason', '')}")
        
        return results
    
    def beautify_presentation(self) -> Dict[str, Any]:
        """
        美化演示文稿，清理未填充的占位符并重新排版
        
        Returns:
            Dict: 美化结果
        """
        beautify_results = self.beautifier.cleanup_and_beautify(self.filled_placeholders)
        optimization_results = self.beautifier.optimize_slide_sequence()
        
        return {
            'beautify_results': beautify_results,
            'optimization_results': optimization_results,
            'summary': {
                'removed_placeholders_count': sum(
                    item['removed_count'] for item in beautify_results['removed_placeholders']
                ),
                'reorganized_slides_count': len(beautify_results['reorganized_slides']),
                'removed_empty_slides_count': len(optimization_results['removed_empty_slides']),
                'final_slide_count': optimization_results['final_slide_count']
            }
        }
    
    def _replace_placeholder_in_slide(self, placeholder_info: Dict[str, Any], new_content: str) -> bool:
        """在特定的文本框中替换占位符"""
        try:
            shape = placeholder_info['shape']
            placeholder_name = placeholder_info['placeholder']
            
            # 检查当前文本框的实际内容
            current_text = shape.text if hasattr(shape, 'text') else ""
            
            # 构建要替换的占位符模式
            placeholder_pattern = f"{{{placeholder_name}}}"
            
            # 使用当前文本框内容进行替换
            if placeholder_pattern in current_text:
                updated_text = current_text.replace(placeholder_pattern, new_content)
                
                # 更新文本框内容
                if hasattr(shape, "text_frame") and shape.text_frame:
                    tf = shape.text_frame
                    tf.clear()
                    
                    # 添加新内容
                    p = tf.paragraphs[0]
                    p.text = updated_text
                    
                    # 保持字体大小
                    if hasattr(p, 'font') and hasattr(p.font, 'size'):
                        if not p.font.size:
                            p.font.size = Pt(16)
                else:
                    # 直接设置text属性
                    shape.text = updated_text
                
                return True
            else:
                return False
                
        except Exception as e:
            print(f"替换占位符时出错: {e}")
            return False
    
    def _update_slide_content(self, slide, content: str):
        """更新幻灯片内容"""
        # 查找可用的文本框
        text_shapes = []
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                text_shapes.append(shape)
        
        if text_shapes:
            # 使用最后一个可用的文本框（通常是主要内容区域）
            target_shape = text_shapes[-1] if len(text_shapes) > 1 else text_shapes[0]
            
            # 清空现有内容并添加新内容
            tf = target_shape.text_frame
            tf.clear()
            
            # 添加内容
            p = tf.paragraphs[0]
            p.text = content
            p.font.size = Pt(16)
    
    def _add_new_slide(self, title: str, content: str):
        """添加新幻灯片"""
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

class FileManager:
    """文件管理器"""
    
    @staticmethod
    def save_ppt_to_bytes(presentation: Presentation) -> bytes:
        """
        将PPT保存为字节数据
        
        Args:
            presentation: PPT演示文稿对象
            
        Returns:
            bytes: PPT文件的字节数据
        """
        config = get_config()
        
        # 创建临时文件
        timestamp = str(int(time.time() * 1000))
        temp_filename = f"temp_ppt_{timestamp}.pptx"
        temp_filepath = os.path.join(config.temp_output_dir, temp_filename)
        
        try:
            # 保存文件
            presentation.save(temp_filepath)
            
            # 读取字节数据
            with open(temp_filepath, 'rb') as f:
                ppt_bytes = f.read()
            
            return ppt_bytes
        finally:
            # 清理临时文件
            try:
                if os.path.exists(temp_filepath):
                    os.remove(temp_filepath)
            except Exception:
                pass
    
    @staticmethod
    def save_ppt_to_file(presentation: Presentation, filename: str = None) -> str:
        """
        将PPT保存到文件
        
        Args:
            presentation: PPT演示文稿对象
            filename: 文件名（可选）
            
        Returns:
            str: 保存的文件路径
        """
        config = get_config()
        
        if not filename:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"updated_ppt_{timestamp}.pptx"
        
        filepath = os.path.join(config.output_dir, filename)
        presentation.save(filepath)
        return filepath
    
    @staticmethod
    def validate_ppt_file(file_path: str) -> Tuple[bool, str]:
        """
        验证PPT文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            Tuple[bool, str]: (是否有效, 错误信息)
        """
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        if not file_path.lower().endswith('.pptx'):
            return False, "文件格式不支持，请使用.pptx格式"
        
        try:
            # 尝试打开文件
            presentation = Presentation(file_path)
            if len(presentation.slides) == 0:
                return False, "PPT文件为空"
            return True, ""
        except Exception as e:
            return False, f"文件损坏或格式错误: {e}"

def format_timestamp(timestamp: float = None) -> str:
    """
    格式化时间戳
    
    Args:
        timestamp: 时间戳（可选，默认当前时间）
        
    Returns:
        str: 格式化的时间字符串
    """
    if timestamp is None:
        timestamp = time.time()
    return datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')

def sanitize_filename(filename: str) -> str:
    """
    清理文件名，移除非法字符
    
    Args:
        filename: 原始文件名
        
    Returns:
        str: 清理后的文件名
    """
    # 移除或替换非法字符
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    # 移除开头和结尾的空白字符和点
    filename = filename.strip('. ')
    # 如果文件名为空，使用默认名称
    if not filename:
        filename = 'untitled'
    return filename

def is_valid_api_key(api_key: str) -> bool:
    """
    验证API密钥格式
    
    Args:
        api_key: API密钥
        
    Returns:
        bool: 是否有效
    """
    if not api_key:
        return False
    
    # 简单验证：以sk-开头，长度合理
    return api_key.startswith('sk-') and len(api_key) > 20