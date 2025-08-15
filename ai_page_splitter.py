#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AI智能分页模块
将用户输入的长文本智能分割为适合PPT展示的多个页面
"""

import re
import json
import requests
from typing import Dict, List, Any, Optional, Tuple
from openai import OpenAI
from config import get_config
from logger import log_user_action

class AIPageSplitter:
    """AI智能分页处理器"""
    
    def __init__(self, api_key: Optional[str] = None):
        """初始化AI分页处理器"""
        config = get_config()
        
        # 根据当前选择的模型获取对应的配置
        model_info = config.get_model_info()
        
        # 自动获取对应的API密钥
        if api_key is None:
            api_key_env = model_info.get('api_key_env')
            if api_key_env:
                # 从环境变量获取对应的API密钥
                import os
                self.api_key = os.getenv(api_key_env) or config.openai_api_key or ""
            else:
                self.api_key = config.openai_api_key or ""
        else:
            self.api_key = api_key
            
        if not self.api_key:
            raise ValueError("请设置API密钥")
        
        self.base_url = model_info.get('base_url', config.openai_base_url)
        
        # 针对不同API提供商优化超时设置
        if 'groq.com' in self.base_url.lower():
            timeout = 45.0
        else:
            timeout = 30.0
        
        self.client = OpenAI(
            api_key=self.api_key,
            base_url=self.base_url,
            timeout=timeout
        )
        self.config = config
        
        # 创建持久化session用于HTTP连接复用
        self.session = requests.Session()
        # 针对API提供商优化适配器配置
        if 'groq.com' in self.base_url.lower():
            from requests.adapters import HTTPAdapter
            from urllib3.util.retry import Retry
            
            retry_strategy = Retry(
                total=3,
                backoff_factor=1,
                status_forcelist=[429, 500, 502, 503, 504],
                allowed_methods=["POST"]
            )
            adapter = HTTPAdapter(
                max_retries=retry_strategy,
                pool_connections=5,
                pool_maxsize=10
            )
            self.session.mount("http://", adapter)
            self.session.mount("https://", adapter)
            # 设置keep-alive
            self.session.headers.update({'Connection': 'keep-alive'})
        
        # 简单的内存缓存
        self._cache = {}
    
    def split_text_to_pages(self, user_text: str, target_pages: Optional[int] = None) -> Dict[str, Any]:
        """
        将用户文本智能分割为多个PPT页面
        
        Args:
            user_text: 用户输入的原始文本
            target_pages: 目标页面数量（可选，由AI自动判断）
            
        Returns:
            Dict: 分页结果，包含每页的内容和分析
        """
        log_user_action("AI智能分页", f"文本长度: {len(user_text)}")
        
        try:
            # 构建AI提示
            system_prompt = self._build_system_prompt(target_pages)
            
            # 检查是否为Liai API
            model_info = self.config.get_model_info()
            if model_info.get('request_format') == 'dify_compatible':
                # 使用Liai API格式
                content = self._call_liai_api(system_prompt, user_text)
            else:
                # 针对不同API提供商优化请求参数
                request_timeout = 60
                
                # 获取实际使用的模型名称
                actual_model = model_info.get('actual_model', self.config.ai_model)
                
                # 检查是否支持流式输出（Groq Kimi K2）
                use_streaming = 'groq.com' in self.base_url.lower()
                
                if use_streaming:
                    # 使用流式输出
                    stream_options = {}
                    
                    response = self.client.chat.completions.create(
                        model=actual_model,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": user_text}
                        ],
                        temperature=0.3,
                        max_tokens=4000,
                        stream=True,
                        stream_options=stream_options,
                        timeout=request_timeout
                    )
                    
                    # 收集流式响应内容
                    content = ""
                    for chunk in response:
                        if chunk.choices and chunk.choices[0].delta.content:
                            content += chunk.choices[0].delta.content
                    
                    content = content.strip() if content else ""
                else:
                    # 使用传统非流式调用
                    response = self.client.chat.completions.create(
                        model=actual_model,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": user_text}
                        ],
                        temperature=0.3,
                        max_tokens=4000,
                        stream=False,
                        timeout=request_timeout
                    )
                    
                    content = response.choices[0].message.content
                    content = content.strip() if content else ""
            
            # 解析AI返回的结果
            return self._parse_ai_response(content, user_text)
            
        except Exception as e:
            print(f"AI分页分析失败: {e}")
            return self._create_fallback_split(user_text)
    
    def _call_liai_api(self, system_prompt: str, user_text: str) -> str:
        """调用Liai API"""
        model_info = self.config.get_model_info()
        base_url = model_info.get('base_url', '')
        endpoint = model_info.get('chat_endpoint', '/chat-messages')
        
        url = base_url + endpoint
        
        # 构建Liai API请求格式
        combined_query = f"{system_prompt}\n\n用户输入：{user_text}"
        
        payload = {
            "inputs": {},
            "query": combined_query,
            "response_mode": "streaming",  # 改为streaming模式提升响应速度
            "conversation_id": "",
            "user": "ai-ppt-user",
            "files": []
        }
        
        headers = {
            'Authorization': f'Bearer {self.api_key}',
            'Content-Type': 'application/json',
            'Connection': 'keep-alive'  # 保持连接
        }
        
        try:
            # 使用持久化会话复用连接，增加超时处理
            response = self.session.post(url, headers=headers, json=payload, timeout=120, stream=True)
            response.raise_for_status()
            
            # 处理streaming响应，特别处理阿里云API的keep-alive
            content = ""
            for line in response.iter_lines():
                if line:
                    try:
                        line_text = line.decode('utf-8').strip()
                        # 忽略阿里云的keep-alive注释
                        if line_text == ': keep-alive' or line_text == '':
                            continue
                        if line_text.startswith('data: '):
                            json_str = line_text[6:]  # 去掉'data: '前缀
                            if json_str.strip() == '[DONE]':
                                break
                            data = json.loads(json_str)
                            if 'answer' in data:
                                content += data['answer']
                            elif 'data' in data and 'answer' in data['data']:
                                content += data['data']['answer']
                    except (json.JSONDecodeError, UnicodeDecodeError):
                        continue
            
            # 如果streaming失败，尝试作为普通JSON处理
            if not content:
                try:
                    result = response.json()
                    content = result.get('answer', '') or result.get('data', {}).get('answer', '')
                except:
                    pass
            
            return content.strip() if content else ""
            
        except Exception as e:
            print(f"Liai API调用失败: {e}")
            raise e
    
    def _build_system_prompt(self, target_pages: Optional[int] = None) -> str:
        """构建AI系统提示"""
        target_instruction = ""
        if target_pages:
            target_instruction = f"目标分为{target_pages}页，"
        
        return f"""你是一个专业的PPT内容分析专家。你的任务是将用户提供的文本内容智能分割为适合PPT展示的多个页面。

**核心原则：**
1. **逻辑清晰**：确保每页内容有明确的主题和逻辑关系
2. **信息完整**：不遗漏原文的重要信息
3. **层次分明**：按照重要性和逻辑顺序安排页面
4. **适量分配**：每页内容量适中，避免过于拥挤或空洞

**分页策略：**
- **标题页（第1页）**：仅提取文档标题和日期信息，其他所有文本内容都延后到第二页开始处理
- **内容页（第2页开始）**：从第二页开始处理所有实际内容，按主要观点、时间顺序或逻辑结构分页
- **概述页**：如果内容复杂，可在第二页作为概述页
- **结尾页**：不生成结尾页（使用预设的固定结尾页模板）

**标题页处理规则：**
- 只从文本开头提取标题信息（通常是第一行或最醒目的文字）
- 自动生成或提取日期信息
- 其余所有文本内容（包括副标题、简介、正文等）都保留给后续内容页处理
- 标题页的original_text_segment只包含提取的标题部分

**页面内容要求：**
- 每页应该有清晰的**主题**
- 每页包含2-4个**要点**
- 每个要点20-50字为宜
- 保持内容的**连贯性**和**完整性**

**分页建议：**
- 短文本（<200字）：建议2-3页（1个标题页 + 1-2个内容页）
- 中等文本（200-800字）：建议3-6页（1个标题页 + 2-5个内容页）  
- 长文本（800-2000字）：建议6-12页（1个标题页 + 5-11个内容页）
- 超长文本（2000-5000字）：建议12-20页（1个标题页 + 11-19个内容页）
- 特长文本（>5000字）：建议20-25页（1个标题页 + 19-24个内容页）

**演示时间参考：**
- 5分钟演示：3-5页（标题页 + 2-4个内容页，每页1-2分钟）
- 10分钟演示：5-8页（标题页 + 4-7个内容页，每页1-2分钟）
- 15分钟演示：8-12页（标题页 + 7-11个内容页，每页1-2分钟）
- 30分钟演示：15-20页（标题页 + 14-19个内容页，每页1-2分钟）
- 学术报告：20-25页（标题页 + 19-24个内容页，根据内容深度调整）

**页面结构说明：**
- 结尾页使用固定模板，不需要AI生成，因此总页数 = 生成页数 + 1个固定结尾页
- **重要限制：总页数不得超过25页**（包括标题页、内容页和结尾页）

{target_instruction}请分析用户文本的结构和内容，智能分割为合适的页面数量。

**严格要求：生成的页面数量必须控制在24页以内，为固定结尾页预留位置，确保总数不超过25页。**

**输出格式要求：**
请严格按照以下JSON格式返回：

```json
{{
  "analysis": {{
    "total_pages": 4,
    "content_type": "技术介绍",
    "split_strategy": "按发展阶段分页",
    "reasoning": "文本描述了技术发展的多个阶段，适合按时间线分页展示"
  }},
  "pages": [
         {{
       "page_number": 1,
       "page_type": "title",
       "title": "人工智能发展历程",
       "subtitle": "",
       "date": "2024年7月",
       "content_summary": "文档标题页，仅从文本开头提取标题和日期，其他内容延后处理",
       "key_points": [
         "文档标题：人工智能发展历程",
         "日期：2024年7月"
       ],
       "original_text_segment": "人工智能发展历程"
     }},
    {{
      "page_number": 2,
      "page_type": "content",
      "title": "AI发展概述",
      "subtitle": "技术演进的主要阶段",
      "content_summary": "从第二页开始处理所有实际内容，介绍AI发展的各个阶段",
      "key_points": [
        "1950年代符号主义起始",
        "1980年代专家系统兴起", 
        "2010年代深度学习突破",
        "当前大语言模型时代"
      ],
      "original_text_segment": "人工智能技术发展经历了多个重要阶段。从1950年代的符号主义开始..."
    }}
  ]
}}
```

**页面类型说明：**
- `title`: 标题页，仅包含文档标题和日期（其他内容固定）
- `overview`: 概述页，总体介绍内容框架（可选）
- `content`: 内容页，具体的要点和详细内容（分页重点）

**重要说明：**
- **标题页处理**：第1页只提取文档标题和日期，所有其他文本内容（包括简介、背景、正文等）都保留给第2页及后续页面处理
- **内容分配**：确保除了标题信息外，原文的所有实际内容都从第2页开始分配，不要在标题页中包含任何正文内容
- **日期处理**：如果原文没有明确日期，自动生成当前年月
- 不要生成结尾页，系统将使用预设的固定结尾页模板
- 专注于从第2页开始的内容页智能分割，确保逻辑清晰、内容均衡

只返回JSON格式，不要其他文字。"""
    
    def _parse_ai_response(self, content: str, user_text: str) -> Dict[str, Any]:
        """解析AI响应结果"""
        try:
            # 提取JSON内容
            json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                # 如果没有代码块，尝试直接解析
                json_str = content
            
            # 解析JSON
            result = json.loads(json_str)
            
            # 验证结果格式
            if self._validate_split_result(result):
                result['success'] = True
                result['original_text'] = user_text
                
                # 添加固定的结尾页
                self._add_ending_page(result)
                
                return result
            else:
                print("AI返回结果格式验证失败")
                return self._create_fallback_split(user_text)
                
        except json.JSONDecodeError as e:
            print(f"JSON解析失败: {e}")
            return self._create_fallback_split(user_text)
        except Exception as e:
            print(f"结果解析失败: {e}")
            return self._create_fallback_split(user_text)
    
    def _validate_split_result(self, result: Dict[str, Any]) -> bool:
        """验证分页结果的格式"""
        try:
            # 检查必需的字段
            if 'analysis' not in result or 'pages' not in result:
                return False
            
            analysis = result['analysis']
            pages = result['pages']
            
            # 检查analysis字段
            required_analysis_fields = ['total_pages', 'content_type', 'split_strategy']
            for field in required_analysis_fields:
                if field not in analysis:
                    return False
            
            # 检查pages数组
            if not isinstance(pages, list) or len(pages) == 0:
                return False
            
            # 检查每个页面的字段
            required_page_fields = ['page_number', 'page_type', 'title', 'key_points']
            for page in pages:
                for field in required_page_fields:
                    if field not in page:
                        return False
                
                # 检查key_points是数组
                if not isinstance(page['key_points'], list):
                    return False
            
            return True
            
        except Exception:
            return False
    
    def _create_fallback_split(self, user_text: str) -> Dict[str, Any]:
        """创建备用分页方案"""
        # 按行分割，找到标题
        lines = [line.strip() for line in user_text.split('\n') if line.strip()]
        if not lines:
            lines = [user_text.strip()]
        
        # 提取标题（通常是第一行，且相对较短）
        title = lines[0] if lines else "内容展示"
        if len(title) > 50:  # 如果第一行太长，可能不是标题，截取前面部分
            title = title[:30] + "..."
        
        pages = []
        
        # 创建标题页（仅包含从文本开头提取的标题和日期）
        import datetime
        current_date = datetime.datetime.now().strftime("%Y年%m月")
        
        pages.append({
            "page_number": 1,
            "page_type": "title", 
            "title": title,
            "subtitle": "",
            "date": current_date,
            "content_summary": "文档标题页，仅从文本开头提取标题和日期，其他内容延后处理",
            "key_points": [f"文档标题：{title}", f"日期：{current_date}"],
            "original_text_segment": title  # 只包含标题部分
        })
        
        # 将除标题外的所有内容分配到第2页开始的内容页
        # 重新组织内容：去掉标题行后的所有文本
        remaining_text = user_text
        if lines and len(lines) > 1:
            # 去掉第一行（标题），保留其余内容
            title_end_pos = user_text.find(lines[0]) + len(lines[0])
            remaining_text = user_text[title_end_pos:].strip()
        
        # 按段落分割剩余内容
        remaining_paragraphs = [p.strip() for p in remaining_text.split('\n\n') if p.strip()]
        if not remaining_paragraphs and remaining_text:
            remaining_paragraphs = [remaining_text]
        
        page_num = 2
        if remaining_paragraphs:
            for i, paragraph in enumerate(remaining_paragraphs):
                # 限制总页数不超过24页（为结尾页预留空间）
                if page_num > 24:
                    print(f"警告：内容过多，已达到24页上限，剩余{len(remaining_paragraphs) - i}段内容将被省略")
                    break
                    
                pages.append({
                    "page_number": page_num,
                    "page_type": "content",
                    "title": f"内容 {page_num - 1}",
                    "subtitle": "",
                    "content_summary": f"第{page_num - 1}部分内容（从第2页开始处理实际内容）",
                    "key_points": [paragraph[:50] + "..." if len(paragraph) > 50 else paragraph],
                    "original_text_segment": paragraph
                })
                page_num += 1
        else:
            # 如果没有剩余内容，至少创建一个空的内容页
            pages.append({
                "page_number": 2,
                "page_type": "content",
                "title": "内容页",
                "subtitle": "",
                "content_summary": "从第2页开始处理实际内容，但当前文本只包含标题",
                "key_points": ["内容待补充"],
                "original_text_segment": "无额外内容"
            })
        
        result = {
            "success": True,
            "analysis": {
                "total_pages": len(pages),
                "content_type": "通用内容",
                "split_strategy": "按段落分页",
                "reasoning": "采用备用分页策略，按段落自动分割"
            },
            "pages": pages,
            "original_text": user_text,
            "is_fallback": True
        }
        
        # 添加固定的结尾页
        self._add_ending_page(result)
        
        return result
    
    def _add_ending_page(self, result: Dict[str, Any]) -> None:
        """添加固定的结尾页"""
        import os
        
        pages = result.get('pages', [])
        if not pages:
            return
        
        # 计算结尾页的页码
        ending_page_number = len(pages) + 1
        
        # 添加结尾页信息
        ending_page = {
            "page_number": ending_page_number,
            "page_type": "ending",
            "title": "谢谢观看",
            "subtitle": "感谢您的聆听",
            "content_summary": "固定结尾页模板，表达感谢",
            "key_points": [
                "感谢观看",
                "欢迎交流讨论"
            ],
            "original_text_segment": "",
            "template_path": os.path.join("templates", "ending_slides.pptx"),
            "is_fixed_template": True,
            "skip_dify_api": True  # 标记为跳过Dify API调用
        }
        
        pages.append(ending_page)
        
        # 更新总页数
        if 'analysis' in result:
            result['analysis']['total_pages'] = len(pages)

class PageContentFormatter:
    """页面内容格式化工具"""
    
    @staticmethod
    def format_page_preview(page: Dict[str, Any]) -> str:
        """格式化页面预览文本"""
        page_type_map = {
            "title": "🏷️ 标题页",
            "overview": "📋 概述页",
            "table_of_contents": "📑 目录页", 
            "content": "📄 内容页",
            "ending": "🔚 结束页"
        }
        
        page_type_display = page_type_map.get(page.get('page_type', 'content'), "📄 内容页")
        
        preview = f"**{page_type_display} - 第{page.get('page_number', 1)}页**\n\n"
        preview += f"**标题：** {page.get('title', '未设置标题')}\n"
        
        # 标题页特殊处理
        if page.get('page_type') == 'title':
            if page.get('date'):
                preview += f"**日期：** {page.get('date')}\n"
            preview += f"**说明：** 标题页使用固定模板，其他内容（作者、机构等）将自动填充\n\n"
        else:
            if page.get('subtitle'):
                preview += f"**副标题：** {page.get('subtitle')}\n"
            preview += f"**内容摘要：** {page.get('content_summary', '无摘要')}\n\n"
        
        key_points = page.get('key_points', [])
        if key_points:
            preview += "**主要要点：**\n"
            for i, point in enumerate(key_points, 1):
                preview += f"{i}. {point}\n"
        
        return preview
    
    @staticmethod
    def format_analysis_summary(analysis: Dict[str, Any]) -> str:
        """格式化分析摘要"""
        summary = f"**📊 分页分析结果**\n\n"
        summary += f"• **总页数：** {analysis.get('total_pages', 0)} 页\n"
        summary += f"• **内容类型：** {analysis.get('content_type', '未知')}\n"
        summary += f"• **分页策略：** {analysis.get('split_strategy', '未知')}\n"
        
        if analysis.get('reasoning'):
            summary += f"• **分析说明：** {analysis.get('reasoning')}\n"
        
        return summary 