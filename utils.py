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
        self.api_key = api_key or config.openai_api_key
        if not self.api_key:
            raise ValueError("请设置OpenRouter API密钥")
        
        self.client = OpenAI(
            api_key=self.api_key,
            base_url=config.openai_base_url
        )
        self.config = config
    
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
            print(f"网络连接错误: {e}")
            return self._create_fallback_assignment(user_text, f"网络连接错误: {e}")
        except TimeoutError as e:
            print(f"请求超时: {e}")
            return self._create_fallback_assignment(user_text, f"请求超时: {e}")
        except Exception as e:
            print(f"调用AI API时出错: {e}")
            return self._create_fallback_assignment(user_text, f"API调用失败: {e}")
    
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
        return f"""你是一个专业的PPT内容优化专家，具备丰富的视觉设计经验和高级布局分析能力。请分析用户文本，并根据PPT模板的深度结构分析进行智能适配和优化，重点关注内容的美观性和视觉效果。

现有PPT高级结构分析：
{ppt_description}

**核心任务：**
1. **结构化适配**：根据PPT模板的占位符结构，将用户文本进行合理的结构化调整
2. **内容优化**：可以适当精简、重组或格式化文本，使其更适合PPT呈现
3. **语言润色**：可以优化语言表达，使其更加简洁明了，适合幻灯片展示
4. **美观性设计**：确保文本内容符合视觉美观要求，提升整体呈现效果
5. **高级布局优化**：利用提供的高级分析信息，优化内容分配和视觉层次

**操作原则：**
- ✅ **可以做的**：重新组织文本结构、精简冗余内容、优化表达方式、调整语言风格
- ✅ **可以做的**：根据占位符特点调整内容长度和格式（如将长段落拆分为要点）
- ✅ **可以做的**：根据高级分析结果调整内容优先级和分配策略
- ✅ **可以做的**：利用视觉权重信息优化重要内容的位置分配
- ❌ **不能做的**：添加用户未提供的信息、编造数据、从外部知识添加内容
- ❌ **不能做的**：改变用户文本的核心意思和关键信息

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

分析要求：
1. 基于用户文本进行结构化分析和适配优化
2. 根据占位符语义特点调整内容呈现方式
3. 保持核心信息完整，但可优化表达形式
4. 严格遵循美观性设计原则，确保视觉效果
5. 严格控制文本长度，遵循字数限制建议
6. 确保格式规范，符合PPT展示要求
7. **充分利用高级分析信息**：
   - 优先填充视觉权重高的占位符
   - 根据布局类型调整内容结构
   - 考虑内容密度避免过度拥挤
   - 保持视觉平衡和层次清晰度
   - 参考布局建议优化分配策略
8. **智能内容分配**：
   - 将最重要的核心信息分配给高权重占位符
   - 根据元素位置信息合理安排内容层次
   - 考虑视觉区域分布优化阅读体验
9. action必须是"replace_placeholder"
10. placeholder必须是模板中实际存在的占位符名称
11. reason字段应该体现高级分析的考量
12. 只返回JSON格式，不要其他文字"""
    
    def _extract_json_from_response(self, content: str, user_text: str) -> Dict[str, Any]:
        """从AI响应中提取JSON"""
        # 提取JSON内容（如果有代码块包围）
        json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
        if json_match:
            content = json_match.group(1)
        
        try:
            return json.loads(content)
        except json.JSONDecodeError as e:
            print(f"AI返回的JSON格式有误，错误：{e}")
            print(f"返回内容：{content}")
            return self._create_fallback_assignment(user_text, f"JSON解析失败: {e}")
    
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
        初始化视觉分析器
        
        Args:
            api_key: OpenAI API密钥
            
        Returns:
            bool: 初始化是否成功
        """
        try:
            self.visual_analyzer = PPTVisualAnalyzer(api_key)
            self.visual_optimizer = VisualLayoutOptimizer(self.visual_analyzer)
            return True
        except Exception as e:
            print(f"视觉分析器初始化失败: {e}")
            return False
    
    def analyze_visual_quality(self, ppt_path: str) -> Dict[str, Any]:
        """
        分析PPT视觉质量
        
        Args:
            ppt_path: PPT文件路径
            
        Returns:
            Dict: 视觉分析结果
        """
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
    
    # 简单验证：支持OpenRouter (sk-or-) 和标准 (sk-) 格式
    return (api_key.startswith('sk-or-') or api_key.startswith('sk-')) and len(api_key) > 20