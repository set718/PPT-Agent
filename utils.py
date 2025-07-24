#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
工具函数模块
包含项目中的共用工具函数
"""

import os
import re
import json
import time
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from config import get_config
from ppt_beautifier import PPTBeautifier
from ppt_advanced_analyzer import PPTStructureAnalyzer, PositionExtractor, SmartLayoutAdjuster, create_advanced_ppt_analysis
from ppt_visual_analyzer import PPTVisualAnalyzer, VisualLayoutOptimizer

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
                            # 为了避免多个占位符在同一个文本框中的冲突，
                            # 我们记录这个文本框的所有占位符
                            for placeholder in placeholders:
                                slide_info["placeholders"][placeholder] = {
                                    "shape": shape,
                                    "original_text": current_text,
                                    "placeholder": placeholder,
                                    "all_placeholders": placeholders  # 记录同一文本框中的所有占位符
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
        self.api_key = api_key or config.openai_api_key
        if not self.api_key:
            raise ValueError("请设置API密钥")
        
        # 根据当前选择的模型获取对应的base_url
        model_info = config.get_model_info()
        self.base_url = model_info.get('base_url', config.openai_base_url)
        
        # 延迟初始化client，避免在创建时就验证API密钥
        self.client = None
        self.config = config
    
    def _ensure_client(self):
        """确保client已初始化"""
        if self.client is None:
            try:
                self.client = OpenAI(
                    api_key=self.api_key,
                    base_url=self.base_url
                )
            except Exception as e:
                raise ValueError(f"API密钥验证失败: {str(e)}")
    
    def analyze_text_for_ppt(self, user_text: str, ppt_structure: Dict[str, Any], enhanced_info: Dict[str, Any] = None) -> Dict[str, Any]:
        """
        使用AI分析文本并生成PPT填充方案
        
        Args:
            user_text: 用户输入的文本
            ppt_structure: PPT结构信息
            enhanced_info: 增强的结构信息（可选）
            
        Returns:
            Dict: 文本分配方案
        """
        # 确保client已初始化
        self._ensure_client()
        
        # 创建PPT结构描述
        if enhanced_info:
            ppt_description = self._create_enhanced_ppt_description(enhanced_info)
        else:
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
            
        except ConnectionError as e:
            print("网络连接错误: %s", str(e))
            return self._create_fallback_assignment(user_text, f"网络连接错误: {str(e)}")
        except TimeoutError as e:
            print("请求超时: %s", str(e))
            return self._create_fallback_assignment(user_text, f"请求超时: {str(e)}")
        except Exception as e:
            print("调用AI API时出错: %s", str(e))
            return self._create_fallback_assignment(user_text, f"API调用失败: {str(e)}")
    
    def _create_ppt_description(self, ppt_structure: Dict[str, Any]) -> str:
        """创建PPT结构描述"""
        description = f"现有PPT共有{ppt_structure['total_slides']}张幻灯片，模板设计意图分析:\n"
        
        # 分析整体结构
        total_placeholders = sum(len(slide.get('placeholders', {})) for slide in ppt_structure['slides'])
        description += f"总占位符数量: {total_placeholders}个，需要智能分配用户文本\n"
        
        # 分析各类占位符分布
        placeholder_types = {'title': 0, 'subtitle': 0, 'content': 0, 'bullet': 0, 'description': 0, 'conclusion': 0}
        for slide in ppt_structure['slides']:
            for placeholder_name in slide.get('placeholders', {}).keys():
                if 'title' in placeholder_name.lower():
                    placeholder_types['title'] += 1
                elif 'subtitle' in placeholder_name.lower():
                    placeholder_types['subtitle'] += 1
                elif 'content' in placeholder_name.lower():
                    placeholder_types['content'] += 1
                elif 'bullet' in placeholder_name.lower():
                    placeholder_types['bullet'] += 1
                elif 'description' in placeholder_name.lower():
                    placeholder_types['description'] += 1
                elif 'conclusion' in placeholder_name.lower():
                    placeholder_types['conclusion'] += 1
        
        description += f"占位符类型分布: {dict(placeholder_types)}\n"
        
        # 详细描述每张幻灯片
        for slide in ppt_structure['slides']:
            description += f"\n第{slide['slide_index']+1}页:"
            
            # 幻灯片标题分析
            if slide['title']:
                description += f" 现有标题「{slide['title']}」"
            else:
                description += f" (无现有标题)"
            
            # 占位符详细分析
            if slide['placeholders']:
                description += f"\n  占位符详情:"
                
                # 按重要性排序显示占位符
                sorted_placeholders = sorted(
                    slide['placeholders'].items(),
                    key=lambda x: self._get_placeholder_priority(x[0])
                )
                
                for placeholder_name, placeholder_info in sorted_placeholders:
                    placeholder_type = self._analyze_placeholder_type(placeholder_name)
                    description += f"\n    - {{{placeholder_name}}} [{placeholder_type}]"
                    
                description += f"\n  设计意图: {self._analyze_slide_design_intent(slide)}\n"
            else:
                description += f" (无占位符)\n"
        
        return description
    
    def _get_placeholder_priority(self, placeholder_name: str) -> int:
        """获取占位符优先级（数字越小优先级越高）"""
        name_lower = placeholder_name.lower()
        if 'title' in name_lower:
            return 1
        elif 'subtitle' in name_lower:
            return 2
        elif 'content' in name_lower and 'bullet' not in name_lower:
            return 3
        elif 'bullet' in name_lower:
            return 4
        elif 'description' in name_lower:
            return 5
        elif 'conclusion' in name_lower:
            return 6
        else:
            return 7
    
    def _analyze_placeholder_type(self, placeholder_name: str) -> str:
        """分析占位符类型"""
        name_lower = placeholder_name.lower()
        if 'title' in name_lower:
            return "标题类-高视觉权重"
        elif 'subtitle' in name_lower:
            return "副标题类-中高视觉权重"
        elif 'content' in name_lower and 'bullet' not in name_lower:
            return "内容类-框架构建"
        elif 'bullet' in name_lower:
            return "要点类-核心信息"
        elif 'description' in name_lower:
            return "描述类-详细说明"
        elif 'conclusion' in name_lower:
            return "结论类-总结升华"
        else:
            return "通用类-灵活使用"
    
    def _analyze_slide_design_intent(self, slide: Dict[str, Any]) -> str:
        """分析幻灯片设计意图"""
        placeholders = slide.get('placeholders', {})
        if not placeholders:
            return "纯展示页面，无需填充"
        
        placeholder_names = list(placeholders.keys())
        
        # 分析设计意图
        has_title = any('title' in name.lower() for name in placeholder_names)
        has_bullets = any('bullet' in name.lower() for name in placeholder_names)
        has_content = any('content' in name.lower() for name in placeholder_names)
        has_description = any('description' in name.lower() for name in placeholder_names)
        
        if has_title and has_bullets:
            return "标题要点型页面，适合概要展示"
        elif has_content and has_bullets:
            return "内容详解型页面，适合分点阐述"
        elif has_description:
            return "描述详解型页面，适合详细说明"
        elif has_title and has_content:
            return "标题内容型页面，适合主题阐述"
        else:
            return "复合型页面，需要灵活安排内容"
    
    def _create_enhanced_ppt_description(self, enhanced_info: Dict[str, Any]) -> str:
        """创建增强的PPT结构描述"""
        basic_structure = enhanced_info.get('basic_structure', {})
        advanced_analysis = enhanced_info.get('advanced_analysis', {})
        position_analysis = enhanced_info.get('position_analysis', {})
        layout_suggestions = enhanced_info.get('layout_suggestions', [])
        
        # 基础信息
        total_slides = basic_structure.get('total_slides', 0)
        description = f"现有PPT共有{total_slides}张幻灯片，高级结构分析如下:\n"
        
        # 添加整体结构分析
        if advanced_analysis:
            overall_structure = advanced_analysis.get('overall_structure', {})
            if overall_structure:
                description += f"\n【整体设计分析】\n"
                description += f"• 整体风格：{overall_structure.get('overall_style', '未知')}\n"
                description += f"• 设计一致性：{overall_structure.get('design_consistency', 0):.2f}/1.0\n"
                
                avg_metrics = overall_structure.get('average_metrics', {})
                if avg_metrics:
                    description += f"• 平均内容密度：{avg_metrics.get('content_density', 0):.2f}/1.0\n"
                    description += f"• 平均视觉平衡度：{avg_metrics.get('visual_balance', 0):.2f}/1.0\n"
                    description += f"• 平均层次清晰度：{avg_metrics.get('hierarchy_clarity', 0):.2f}/1.0\n"
                
                layout_dist = overall_structure.get('layout_distribution', {})
                if layout_dist:
                    description += f"• 布局类型分布：{layout_dist}\n"
        
        # 添加详细的幻灯片分析
        slide_layouts = advanced_analysis.get('slide_layouts', [])
        for i, slide_layout in enumerate(slide_layouts):
            description += f"\n第{i+1}页详细分析：\n"
            description += f"• 布局类型：{slide_layout.layout_type}\n"
            description += f"• 设计意图：{slide_layout.design_intent}\n"
            description += f"• 内容密度：{slide_layout.content_density:.2f}/1.0\n"
            description += f"• 视觉平衡度：{slide_layout.visual_balance:.2f}/1.0\n"
            description += f"• 层次清晰度：{slide_layout.hierarchy_clarity:.2f}/1.0\n"
            
            # 添加元素信息
            elements = slide_layout.elements
            if elements:
                description += f"• 包含{len(elements)}个元素：\n"
                for element in elements:
                    if element.placeholder_name:
                        description += f"  - {{{element.placeholder_name}}} [{element.element_type}] 视觉权重:{element.visual_weight}/5\n"
                        description += f"    位置:(x:{element.position.left:.0f}, y:{element.position.top:.0f}, w:{element.position.width:.0f}, h:{element.position.height:.0f})\n"
            
            # 添加视觉区域分析
            visual_regions = slide_layout.visual_regions
            if visual_regions:
                description += f"• 视觉区域分布：\n"
                for region_name, region_elements in visual_regions.items():
                    if region_elements:
                        description += f"  - {region_name}区域：{len(region_elements)}个元素\n"
        
        # 添加布局建议
        if layout_suggestions:
            description += f"\n【布局优化建议】\n"
            for suggestion in layout_suggestions:
                slide_idx = suggestion.get('slide_index', 0)
                suggestions = suggestion.get('suggestions', {})
                
                layout_sugg = suggestions.get('layout_suggestions', [])
                if layout_sugg:
                    description += f"第{slide_idx+1}页建议：\n"
                    for sugg in layout_sugg:
                        description += f"• {sugg.get('description', '')}\n"
        
        # 添加位置分析摘要
        if position_analysis:
            description += f"\n【空间布局分析】\n"
            spatial_relationships = position_analysis.get('spatial_relationships', {})
            if spatial_relationships:
                description += f"• 幻灯片间布局一致性分析已完成\n"
                # 可以添加更多空间关系的描述
        
        return description
    
    def _build_system_prompt(self, ppt_description: str) -> str:
        """构建系统提示"""
        return """你是一个专业的PPT内容分析专家。你的任务是将用户提供的文本内容智能分配到PPT模板的合适占位符中。

**重要原则：**
1. 只使用用户提供的文本内容，不生成新内容
2. 可以对文本进行适当的优化、精简或重新组织
3. 根据占位符的语义含义选择最合适的内容片段
4. 不是所有占位符都必须填充，只填充有合适内容的占位符

现有PPT高级结构分析：
%s""" % ppt_description + """

**核心任务：**
1. **内容分析**：理解用户提供的文本结构和主要信息点
2. **智能匹配**：将文本内容分配到最合适的占位符中
3. **适度优化**：对文本进行必要的精简和重组，但保持原意
4. **结构清晰**：确保分配后的内容逻辑清晰，层次分明

**操作原则：**
- ✅ **可以做的**：从用户文本中提取合适的片段填入占位符
- ✅ **可以做的**：适当精简、重组文本使其更适合PPT展示
- ✅ **可以做的**：调整语言表达，使其更简洁明了
- ❌ **不能做的**：生成用户未提供的新信息
- ❌ **不能做的**：强行填满所有占位符
- ❌ **不能做的**：改变用户文本的核心含义

**高级分析信息使用指南：**
1. **整体设计分析**：参考整体风格、设计一致性和平均指标，确保内容风格匹配
2. **布局类型识别**：根据每页的布局类型（如title_with_bullets、content_grid等）调整内容结构
3. **视觉权重优化**：将最重要的内容分配给视觉权重高的占位符
4. **内容密度控制**：根据当前内容密度调整文本长度，避免过于拥挤或空旷
5. **视觉平衡考量**：在分配内容时考虑视觉平衡度，避免内容过于集中
6. **层次清晰度优化**：确保内容层次清晰，与现有的层次结构保持一致
7. **布局建议应用**：参考提供的布局优化建议，调整内容分配策略

**占位符语义规则与视觉层次：**
- `title` = 主标题或文档标题（简洁有力，建议8-15字）
  * 视觉权重：★★★★★ 最高优先级，是视觉焦点
  * 设计要求：突出核心主题，用词精炼有力，避免冗长表述
- `subtitle` = 副标题（补充说明，建议15-25字）
  * 视觉权重：★★★★ 高优先级，支撑主标题
  * 设计要求：与主标题形成呼应，提供必要补充信息
- `content_X` = 分类标题、章节标题、时间点等结构性内容（清晰明确，建议10-20字）
  * 视觉权重：★★★★ 高优先级，构建内容框架
  * 设计要求：逻辑清晰，层次分明，便于读者理解结构
- `content_X_bullet_Y` = 属于特定content的具体要点（简洁扼要，建议20-40字）
  * 视觉权重：★★★ 中高优先级，支撑章节内容
  * 设计要求：要点明确，表述简洁，与对应content形成逻辑层次
- `bullet_X` = 独立的要点列表（重点突出，建议15-35字）
  * 视觉权重：★★★ 中高优先级，关键信息载体
  * 设计要求：并列关系清晰，每个要点独立且完整
- `description` = 描述性文字（详细但不冗长，建议30-80字）
  * 视觉权重：★★ 中等优先级，提供详细说明
  * 设计要求：信息丰富但不冗长，支撑主要内容
- `conclusion` = 结论性内容（总结性强，建议20-50字）
  * 视觉权重：★★★★ 高优先级，总结升华
  * 设计要求：总结有力，呼应主题，给人深刻印象

**美观性设计原则：**
1. **视觉层次清晰**：
   - 标题类（title, subtitle）：用词精炼，突出核心概念
   - 内容类（content_X）：条理清晰，逻辑分明
   - 要点类（bullet_X）：简洁有力，易于快速理解

2. **文本长度控制与格式约束**：
   - 标题类占位符：
     * title: 8-15字为佳，最多不超过20字
     * subtitle: 15-25字为佳，避免超过30字
     * 要求：简洁有力，避免冗长描述
   - 内容类占位符：
     * content_X: 10-20字为佳，构建清晰框架
     * content_X_bullet_Y: 20-40字为佳，保持单行显示
     * 要求：逻辑清晰，层次分明
   - 要点类占位符：
     * bullet_X: 15-35字为佳，确保单行完整显示
     * 要求：并列关系明确，避免换行影响美观
   - 描述类占位符：
     * description: 30-80字为佳，提供适度详细说明
     * conclusion: 20-50字为佳，总结有力
     * 要求：信息丰富但不冗长，保持可读性

3. **语言风格统一与表达优化**：
   - 保持同一张PPT内语言风格的一致性
   - 使用简洁明了的表达方式
   - 避免冗长的句子和复杂的语法结构
   - 专业术语适度使用，确保可读性
   - 使用主动语态，增强表达力
   - 避免重复用词，保持语言丰富性

4. **内容平衡分布与版式协调**：
   - 合理分配内容到各个占位符，避免内容集中在少数占位符
   - 确保同一张幻灯片内容量相对均衡，避免头重脚轻
   - 标题与内容比例协调，标题简洁，内容充实但不冗长
   - 并列要点长度相近，保持视觉整齐美观
   - 考虑占位符的空间位置，重要内容优先填充显眼位置

5. **可读性优化与信息层次**：
   - 使用易于理解的词汇和表达
   - 避免过于专业的术语堆砌
   - 确保关键信息突出显示
   - 重要概念优先分配到高权重占位符
   - 支撑信息合理分配到中低权重占位符
   - 避免信息重复，每个占位符承担独特功能

6. **版式设计原则**：
   - **对比原则**：标题与内容、主要与次要信息形成明显对比
   - **对齐原则**：保持内容逻辑对齐，增强整体感
   - **重复原则**：在多张幻灯片中保持风格一致性
   - **接近原则**：相关内容放置在相近位置，形成视觉关联
   - **留白原则**：避免信息过密，适当留白增强可读性

请按照以下JSON格式返回：
{{
  "assignments": [
    {{
      "slide_index": 0,
      "action": "replace_placeholder",
      "placeholder": "title",
      "content": "优化后的标题内容",
      "reason": "提炼核心概念，适配标题占位符，符合美观性要求"
    }}
  ]
}}

**具体格式要求：**
1. **标点符号规范**：
   - 标题类占位符：避免使用句号，可使用感叹号或问号增强表达力
   - 要点类占位符：使用句号结尾，保持格式一致
   - 描述类占位符：使用标准标点，增强可读性

2. **数字和符号处理**：
   - 优先使用阿拉伯数字，简洁明了
   - 适当使用符号（如：→、●、★）增强视觉效果
   - 避免过多特殊符号，保持整洁

3. **换行和分段**：
   - 单个占位符内容避免内部换行
   - 长内容优先通过精简语言控制长度
   - 必要时可使用分号分隔多个要点

**分析要求：**
1. 仔细阅读用户提供的文本内容
2. 分析PPT模板中各个占位符的语义含义
3. 从用户文本中提取最合适的内容片段分配给相应占位符
4. 优先填充重要的占位符（如title、主要content）
5. 对于细节性的占位符（如bullet点），只在有合适内容时才填充
6. 保持原文的核心意思，只做必要的格式调整

**输出格式：**
只返回JSON格式，包含assignments数组，每个元素包含：
- slide_index: 幻灯片索引（从0开始）
- action: "replace_placeholder"
- placeholder: 占位符名称（必须存在于模板中）
- content: 从用户文本提取的内容（经过适当优化）
- reason: 选择此内容的理由

**示例：**
如果用户文本是"人工智能发展历程包括三个阶段"，模板有title和content_1占位符，则：
```json
{
  "assignments": [
    {
      "slide_index": 0,
      "action": "replace_placeholder",
      "placeholder": "title",
      "content": "人工智能发展历程",
      "reason": "提取主题作为标题"
    },
    {
      "slide_index": 0,
      "action": "replace_placeholder", 
      "placeholder": "content_1",
      "content": "包括三个重要发展阶段",
      "reason": "提取核心内容并简化表达"
    }
  ]
}
```

只返回JSON，不要其他文字。"""
    
    def _extract_json_from_response(self, content: str, user_text: str) -> Dict[str, Any]:
        """从AI响应中提取JSON"""
        # 提取JSON内容（如果有代码块包围）
        json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
        if json_match:
            content = json_match.group(1)
        
        try:
            return json.loads(content)
        except json.JSONDecodeError as e:
            print("AI返回的JSON格式有误，错误：%s", str(e))
            print("返回内容：%s", content)
            return self._create_fallback_assignment(user_text, f"JSON解析失败: {str(e)}")
    
    def _create_fallback_assignment(self, user_text: str, error_msg: str) -> Dict[str, Any]:
        """创建备用分配方案"""
        return {
            "assignments": [
                {
                    "slide_index": 0,
                    "action": "replace_placeholder",
                    "placeholder": "content",
                    "content": user_text,
                    "reason": "API调用失败或解析错误，默认填入content占位符。错误: " + str(error_msg)
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
        
        # 初始化高级分析器
        self.advanced_analysis = create_advanced_ppt_analysis(presentation)
        self.structure_analyzer = self.advanced_analysis['analyzers']['structure_analyzer'] if 'analyzers' in self.advanced_analysis else None
        self.position_extractor = self.advanced_analysis['analyzers']['position_extractor'] if 'analyzers' in self.advanced_analysis else None
        self.layout_adjuster = self.advanced_analysis['analyzers']['layout_adjuster'] if 'analyzers' in self.advanced_analysis else None
        
        # 视觉分析器（需要API密钥时才初始化）
        self.visual_analyzer = None
        self.visual_optimizer = None
    
    def get_enhanced_structure_info(self) -> Dict[str, Any]:
        """获取增强的PPT结构信息"""
        if not self.structure_analyzer:
            return self.ppt_structure
        
        # 合并基础分析和高级分析结果
        enhanced_info = {
            'basic_structure': self.ppt_structure,
            'advanced_analysis': self.advanced_analysis.get('structure_analysis', {}),
            'position_analysis': self.advanced_analysis.get('position_analysis', {}),
            'layout_suggestions': []
        }
        
        # 为每张幻灯片生成布局建议
        if self.layout_adjuster and 'structure_analysis' in self.advanced_analysis:
            slide_layouts = self.advanced_analysis['structure_analysis'].get('slide_layouts', [])
            for i, layout in enumerate(slide_layouts):
                # 模拟一些内容来生成建议
                mock_content = {}
                if i < len(self.ppt_structure['slides']):
                    slide_info = self.ppt_structure['slides'][i]
                    for placeholder in slide_info.get('placeholders', {}).keys():
                        mock_content[placeholder] = f"示例内容_{placeholder}"
                
                if mock_content:
                    suggestions = self.layout_adjuster.suggest_optimal_layout(i, mock_content)
                    enhanced_info['layout_suggestions'].append({
                        'slide_index': i,
                        'suggestions': suggestions
                    })
        
        return enhanced_info
    
    def initialize_visual_analyzer(self, api_key: str) -> bool:
        """
        初始化视觉分析器（仅在启用视觉分析时）
        
        Args:
            api_key: OpenAI API密钥
            
        Returns:
            bool: 初始化是否成功
        """
        # 检查配置是否启用视觉分析
        config = get_config()
        if not config.enable_visual_analysis:
            print(f"[INFO] 当前模型 {config.ai_model} 不支持视觉分析，跳过视觉分析器初始化")
            self.visual_analyzer = None
            self.visual_optimizer = None
            return True  # 返回True表示按配置正确初始化
        
        try:
            self.visual_analyzer = PPTVisualAnalyzer(api_key)
            self.visual_optimizer = VisualLayoutOptimizer(self.visual_analyzer)
            print(f"[INFO] 视觉分析器初始化成功，使用模型: {config.ai_model}")
            return True
        except Exception as e:
            print("视觉分析器初始化失败: %s", str(e))
            return False
    
    def analyze_visual_quality(self, ppt_path: str) -> Dict[str, Any]:
        """
        分析PPT视觉质量（如果启用了视觉分析功能）
        
        Args:
            ppt_path: PPT文件路径
            
        Returns:
            Dict: 视觉分析结果
        """
        config = get_config()
        
        if not config.enable_visual_analysis:
            # 视觉分析被禁用，返回简单的默认分析结果
            return {
                "analysis_skipped": True,
                "reason": f"当前使用的模型 {config.ai_model} 不支持视觉分析功能",
                "slides_analysis": [],
                "overall_quality": {
                    "visual_appeal": 0.5,
                    "content_balance": 0.5,
                    "consistency": 0.5
                }
            }
        
        if not self.visual_analyzer:
            return {"error": "视觉分析器未初始化，请先提供API密钥"}
        
        try:
            return self.visual_analyzer.analyze_presentation_visual_quality(ppt_path)
        except Exception as e:
            return {"error": f"视觉分析失败: {e}"}
    
    def apply_visual_optimizations(self, visual_analysis: Dict[str, Any]) -> Dict[str, Any]:
        """
        应用视觉优化建议
        
        Args:
            visual_analysis: 视觉分析结果
            
        Returns:
            Dict: 优化结果
        """
        if not self.visual_optimizer:
            return {"error": "视觉优化器未初始化"}
        
        try:
            slide_analyses = visual_analysis.get("slide_analyses", [])
            optimization_results = []
            
            for slide_analysis in slide_analyses:
                slide_index = slide_analysis.get("slide_index", 0)
                result = self.visual_optimizer.optimize_slide_layout(
                    self.presentation, slide_index, slide_analysis
                )
                optimization_results.append(result)
            
            return {
                "success": True,
                "optimization_results": optimization_results,
                "total_optimizations": sum(
                    len(r.get("optimizations_applied", [])) 
                    for r in optimization_results 
                    if r.get("success")
                )
            }
            
        except Exception as e:
            return {"error": f"视觉优化失败: {e}"}
    
    def apply_assignments(self, assignments: Dict[str, Any], user_text: str = "") -> List[str]:
        """
        应用文本分配方案
        
        Args:
            assignments: 分配方案
            user_text: 用户原始文本（可选，用于添加到幻灯片备注）
            
        Returns:
            List[str]: 处理结果列表
        """
        assignments_list = assignments.get('assignments', [])
        results = []
        
        # 如果提供了用户原始文本，则为幻灯片添加备注
        if user_text.strip():
            notes_results = self._add_notes_to_slides(assignments_list, user_text)
            results.extend(notes_results)
        
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
                            
                            results.append(f"SUCCESS: 已替换第{slide_index+1}页的 {{{placeholder}}} 占位符: {assignment.get('reason', '')}")
                        else:
                            results.append(f"ERROR: 替换第{slide_index+1}页的 {{{placeholder}}} 占位符失败")
                    else:
                        results.append(f"ERROR: 第{slide_index+1}页不存在 {{{placeholder}}} 占位符")
                else:
                    results.append(f"ERROR: 幻灯片索引 {slide_index+1} 超出范围")
            
            elif action == 'update':
                if 0 <= slide_index < len(self.presentation.slides):
                    slide = self.presentation.slides[slide_index]
                    self._update_slide_content(slide, content)
                    results.append(f"SUCCESS: 已更新第{slide_index+1}页: {assignment.get('reason', '')}")
                
            elif action == 'add_new':
                title = assignment.get('title', '新增内容')
                self._add_new_slide(title, content)
                results.append(f"SUCCESS: 已新增幻灯片「{title}」: {assignment.get('reason', '')}")
        
        return results
    
    def _add_notes_to_slides(self, assignments_list: List[Dict], user_text: str) -> List[str]:
        """
        为幻灯片添加用户原始文本备注
        
        Args:
            assignments_list: 分配方案列表
            user_text: 用户原始文本
            
        Returns:
            List[str]: 备注添加结果
        """
        results = []
        
        # 获取涉及的幻灯片索引
        involved_slides = set()
        for assignment in assignments_list:
            slide_index = assignment.get('slide_index', 0)
            if 0 <= slide_index < len(self.presentation.slides):
                involved_slides.add(slide_index)
        
        # 如果只有一张幻灯片被涉及，将完整的用户文本添加到该幻灯片
        if len(involved_slides) == 1:
            slide_index = list(involved_slides)[0]
            success = self._add_note_to_slide(slide_index, user_text)
            if success:
                results.append(f"NOTES: 已将原始文本添加到第{slide_index+1}页备注")
            else:
                results.append(f"NOTES ERROR: 添加备注到第{slide_index+1}页失败")
        
        # 如果涉及多张幻灯片，智能分割用户文本
        elif len(involved_slides) > 1:
            text_segments = self._split_text_for_slides(user_text, involved_slides, assignments_list)
            for slide_index, text_segment in text_segments.items():
                if text_segment.strip():
                    success = self._add_note_to_slide(slide_index, text_segment)
                    if success:
                        results.append(f"NOTES: 已将相关文本添加到第{slide_index+1}页备注")
                    else:
                        results.append(f"NOTES ERROR: 添加备注到第{slide_index+1}页失败")
        
        return results
    
    def _add_note_to_slide(self, slide_index: int, note_text: str) -> bool:
        """
        为指定幻灯片添加备注
        
        Args:
            slide_index: 幻灯片索引
            note_text: 备注文本
            
        Returns:
            bool: 是否成功添加备注
        """
        try:
            slide = self.presentation.slides[slide_index]
            
            # 获取或创建备注页
            if slide.has_notes_slide:
                notes_slide = slide.notes_slide
            else:
                notes_slide = slide.notes_slide  # 这会自动创建notes_slide
            
            # 获取备注文本框
            notes_text_frame = notes_slide.notes_text_frame
            
            # 设置备注内容
            if notes_text_frame.text.strip():
                # 如果已有备注，添加分隔符和新内容
                notes_text_frame.text += f"\n\n【原始文本】\n{note_text}"
            else:
                # 如果没有备注，直接添加
                notes_text_frame.text = f"【原始文本】\n{note_text}"
            
            return True
            
        except Exception as e:
            print(f"添加备注失败: {e}")
            return False
    
    def _split_text_for_slides(self, user_text: str, involved_slides: set, assignments_list: List[Dict]) -> Dict[int, str]:
        """
        智能分割用户文本，为不同幻灯片分配相关的文本段落
        
        Args:
            user_text: 用户原始文本
            involved_slides: 涉及的幻灯片索引集合
            assignments_list: 分配方案列表
            
        Returns:
            Dict[int, str]: 每张幻灯片对应的文本段落
        """
        # 按段落分割用户文本
        paragraphs = [p.strip() for p in user_text.split('\n\n') if p.strip()]
        if not paragraphs:
            paragraphs = [user_text]
        
        # 为每张幻灯片分配文本段落
        slide_texts = {}
        sorted_slides = sorted(involved_slides)
        
        # 如果段落数量 >= 幻灯片数量，平均分配
        if len(paragraphs) >= len(sorted_slides):
            paragraphs_per_slide = len(paragraphs) // len(sorted_slides)
            remainder = len(paragraphs) % len(sorted_slides)
            
            start_idx = 0
            for i, slide_index in enumerate(sorted_slides):
                end_idx = start_idx + paragraphs_per_slide
                if i < remainder:
                    end_idx += 1
                
                slide_paragraphs = paragraphs[start_idx:end_idx]
                slide_texts[slide_index] = '\n\n'.join(slide_paragraphs)
                start_idx = end_idx
        else:
            # 如果段落少于幻灯片，优先为前几张幻灯片分配
            for i, slide_index in enumerate(sorted_slides):
                if i < len(paragraphs):
                    slide_texts[slide_index] = paragraphs[i]
                else:
                    # 剩余幻灯片分享最后一个段落或完整文本
                    slide_texts[slide_index] = user_text if len(paragraphs) == 1 else paragraphs[-1]
        
        return slide_texts
    
    def beautify_presentation(self, enable_visual_optimization: bool = False, ppt_path: str = None) -> Dict[str, Any]:
        """
        美化演示文稿，清理未填充的占位符并重新排版
        
        Args:
            enable_visual_optimization: 是否启用视觉优化
            ppt_path: PPT文件路径（视觉分析需要）
            
        Returns:
            Dict: 美化结果
        """
        beautify_results = self.beautifier.cleanup_and_beautify(self.filled_placeholders)
        optimization_results = self.beautifier.optimize_slide_sequence()
        
        # 基础美化结果
        result = {
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
        
        # 如果启用视觉优化且视觉分析器可用
        if enable_visual_optimization and self.visual_analyzer and ppt_path:
            try:
                print("🎨 执行视觉质量分析...")
                visual_analysis = self.analyze_visual_quality(ppt_path)
                
                if "error" not in visual_analysis:
                    print("🔧 应用视觉优化建议...")
                    visual_optimization = self.apply_visual_optimizations(visual_analysis)
                    
                    result['visual_analysis'] = visual_analysis
                    result['visual_optimization'] = visual_optimization
                    result['summary']['visual_optimizations_applied'] = visual_optimization.get('total_optimizations', 0)
                    
                    overall_score = visual_analysis.get('overall_analysis', {}).get('weighted_score', 0)
                    result['summary']['visual_quality_score'] = overall_score
                else:
                    result['visual_analysis'] = {"error": visual_analysis.get("error")}
                    
            except Exception as e:
                result['visual_analysis'] = {"error": f"视觉分析过程中出错: {e}"}
        
        return result
    
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
                # 对于多个占位符的情况，只替换第一次出现的
                updated_text = current_text.replace(placeholder_pattern, new_content, 1)
                
                print(f"替换占位符: {placeholder_pattern}")
                print(f"原文本: '{current_text}'")
                print(f"新内容: '{new_content}'")
                print(f"更新后: '{updated_text}'")
                
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
                print(f"占位符 {placeholder_pattern} 在文本 '{current_text}' 中未找到")
                return False
                
        except Exception as e:
            print("替换占位符时出错: %s", str(e))
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
    
    # 简单验证：支持OpenRouter (sk-or-) 和标准 (sk-) 格式
    return (api_key.startswith('sk-or-') or api_key.startswith('sk-')) and len(api_key) > 20