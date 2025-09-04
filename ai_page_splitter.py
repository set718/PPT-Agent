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
        
        # 初始化多密钥管理
        self._initialize_api_keys(model_info, config, api_key)
        
        self.base_url = model_info.get('base_url', config.openai_base_url)
        self.config = config
        
        # 创建持久化session用于HTTP连接复用
        self.session = requests.Session()
        
        # 简单的内存缓存
        self._cache = {}
        
        # 密钥轮询索引
        self._current_key_index = 0
        
    
    def _initialize_api_keys(self, model_info, config, api_key):
        """初始化API密钥列表"""
        import os
        
        if api_key:
            self.api_keys = [api_key]
            return
        
        if model_info.get('api_provider') == 'OpenRouter':
            # 从环境变量获取OpenRouter密钥（用户自定义）
            self.api_keys = []
            for i in range(1, 6):  # 支持1-5个密钥
                key_name = f'OPENROUTER_API_KEY_{i}'
                key_value = os.getenv(key_name)
                if key_value:
                    self.api_keys.append(key_value)
            
            # 如果没有找到编号密钥，尝试单个密钥
            if not self.api_keys:
                single_key = os.getenv('OPENROUTER_API_KEY')
                if single_key:
                    self.api_keys = [single_key]
        elif model_info.get('api_provider') == 'Volces' and model_info.get('use_multiple_keys'):
            # 从环境变量获取火山引擎密钥（多密钥负载均衡）
            self.api_keys = []
            for i in range(1, 6):  # 支持1-5个密钥
                key_name = f'ARK_API_KEY_{i}'
                key_value = os.getenv(key_name)
                if key_value:
                    self.api_keys.append(key_value)
            
            # 如果没有找到编号密钥，尝试单个密钥
            if not self.api_keys:
                single_key = os.getenv('ARK_API_KEY')
                if single_key:
                    self.api_keys = [single_key]
        else:
            # 其他API使用单密钥
            api_key_env = model_info.get('api_key_env')
            if api_key_env:
                key_value = os.getenv(api_key_env) or config.openai_api_key or ""
                if key_value:
                    self.api_keys = [key_value]
                else:
                    self.api_keys = []
            else:
                self.api_keys = [config.openai_api_key] if config.openai_api_key else []
        
        if not self.api_keys:
            raise ValueError("请设置API密钥")
        
        print(f"初始化完成，可用API密钥数量: {len(self.api_keys)}")
    
    def _get_next_api_key(self):
        """获取下一个API密钥（轮询）"""
        if not self.api_keys:
            raise ValueError("没有可用的API密钥")
        
        key = self.api_keys[self._current_key_index]
        self._current_key_index = (self._current_key_index + 1) % len(self.api_keys)
        return key
    
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
            
            # 检查API类型，决定调用方式
            model_info = self.config.get_model_info()
            if model_info.get('request_format') == 'dify_compatible':
                # 使用Liai API格式
                content = self._call_liai_api(system_prompt, user_text)
            elif model_info.get('request_format') == 'streaming_compatible':
                # 使用OpenRouter API格式（类似Liai的分批处理）
                content = self._call_openrouter_api(system_prompt, user_text)
            elif model_info.get('request_format') == 'openai_responses_api':
                # 使用GPT-5 Responses API格式
                content = self._call_gpt5_responses_api(system_prompt, user_text)
            else:
                # 标准OpenAI API格式
                request_timeout = 60
                actual_model = model_info.get('actual_model', self.config.ai_model)
                
                response = self.client.chat.completions.create(
                    model=actual_model,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_text}
                    ],
                    temperature=0.3,
                    max_tokens=4000,
                    stream=True,
                    timeout=request_timeout
                )
                
                # 收集流式响应内容
                content = ""
                for chunk in response:
                    if chunk.choices and chunk.choices[0].delta.content:
                        content += chunk.choices[0].delta.content
                
                content = content.strip() if content else ""
            
            # 解析AI返回的结果
            return self._parse_ai_response(content, user_text)
            
        except Exception as e:
            print(f"AI分页分析失败: {e}")
            raise e
    
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
    
    def _call_openrouter_api(self, system_prompt: str, user_text: str) -> str:
        """调用OpenRouter API（带故障转移的多密钥负载均衡）"""
        model_info = self.config.get_model_info()
        
        # 获取实际模型名称和额外头部
        actual_model = model_info.get('actual_model', 'openai/gpt-5')
        extra_headers = model_info.get('extra_headers', {})
        
        # 尝试所有可用密钥
        last_exception = None
        for attempt in range(len(self.api_keys)):
            current_api_key = self._get_next_api_key()
            
            try:
                # 为当前密钥创建临时客户端
                temp_client = OpenAI(
                    api_key=current_api_key,
                    base_url=self.base_url,
                    timeout=120
                )
                
                print(f"尝试使用API密钥 {attempt + 1}/{len(self.api_keys)} (末尾: ...{current_api_key[-8:]})")
                
                # 使用持久化会话复用连接，类似Liai的处理方式
                response = temp_client.chat.completions.create(
                    model=actual_model,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_text}
                    ],
                    temperature=0.3,
                    max_tokens=4000,
                    stream=True,  # 使用流式响应，类似Liai
                    extra_headers=extra_headers,
                    extra_body={},  # OpenRouter兼容
                    timeout=120  # 与Liai相同的超时时间
                )
                
                # 处理streaming响应，类似Liai的逐行处理
                content = ""
                for chunk in response:
                    if chunk.choices and chunk.choices[0].delta and chunk.choices[0].delta.content:
                        chunk_content = chunk.choices[0].delta.content
                        if chunk_content:
                            content += chunk_content
                
                result_content = content.strip() if content else ""
                print(f"✅ API调用成功，使用密钥: ...{current_api_key[-8:]}")
                return result_content
                
            except Exception as e:
                last_exception = e
                print(f"❌ API密钥 ...{current_api_key[-8:]} 调用失败: {e}")
                
                # 如果还有其他密钥可以尝试，继续下一个
                if attempt < len(self.api_keys) - 1:
                    print(f"⏳ 尝试下一个API密钥...")
                    continue
        
        # 所有密钥都失败了
        print(f"❌ 所有{len(self.api_keys)}个OpenRouter API密钥都失败了")
        raise last_exception or Exception("所有OpenRouter API密钥调用失败")
    
    def _call_gpt5_responses_api(self, system_prompt: str, user_text: str) -> str:
        """调用GPT-5 Responses API进行文本分析"""
        from openai import OpenAI
        
        client = OpenAI(
            api_key="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"
        )
        
        # 组合系统提示和用户文本
        full_input = f"{system_prompt}\n\n用户文本：\n{user_text}"
        
        response = client.responses.create(
            model="gpt-5",
            input=full_input,
            store=True
        )
        
        return response.output_text
    
    def _build_system_prompt(self, target_pages: Optional[int] = None) -> str:
        """构建AI系统提示"""
        target_instruction = ""
        if target_pages:
            target_instruction = f"目标分为{target_pages}页，"
        
        return f"""你是一个专业的PPT内容分析专家。你的任务是将用户提供的文本内容智能分割为适合PPT展示的多个页面。

**核心原则：**
1. **逻辑结构优先**：按内容的逻辑主题分页，同一主题和相关主题必须合并
2. **内容充实性**：每页必须有足够内容量，严禁薄页面，AI倾向过度分页需主动抵制
3. **强制合并策略**：相似、相关、关联主题必须合并，只有完全不同主题才分页
4. **信息完整性**：不遗漏重要信息，保持逻辑完整

**分页策略：**
- **标题页（第1页）**：仅提取文档标题和日期信息，其他所有文本内容都延后到第三页开始处理
- **目录页（第2页）**：AI必须生成完整的目录内容，包括各章节标题，格式如"第一章节\n第二章节\n第三章节"
- **内容页（第3页开始）**：从第三页开始处理所有实际内容，按主要观点、时间顺序或逻辑结构分页
- **结尾页**：不生成结尾页（使用预设的固定结尾页模板）

**标题页处理规则：**
- 只从文本开头提取标题信息（通常是第一行或最醒目的文字）
- 自动生成或提取日期信息
- 其余所有文本内容（包括副标题、简介、正文等）都保留给后续内容页处理
- 标题页的original_text_segment只包含提取的标题部分

**页面内容要求：**
- 每页应该有清晰的**主题**（通过title字段体现）
- **优先按逻辑分配**：属于同一个主题、概念或章节的内容应该放在同一页
- **重点保留原文**：original_text_segment字段必须包含该页对应的完整原文片段
- **内容量优先级**：适中 >> 过多 >> 过少（宁可内容多一些，也不要让页面显得空洞）
- 保持内容的**连贯性**和**完整性**

**分页建议（极简策略 - 最大化内容合并）：**
- 极短文本（<300字）：仅1页内容（全部内容放在一页）
- 短文本（300-1000字）：1页内容（强制合并为1页）
- 中等文本（1000-2000字）：1-2页内容（优先合并为1页，仅在逻辑完全不相关时分为2页）
- 长文本（2000-4000字）：2-3页内容（按主要章节分页，大量合并小节）
- 超长文本（>4000字）：3-6页内容（仅按主要章节分页，严格合并子主题）
- **核心原则：能合并必须合并，宁可单页内容丰富也不要页面分散**
- **最小阈值：每页至少300字，低于此阈值必须与相邻页面合并**

{target_instruction}请分析用户文本的结构和内容，按逻辑主题智能分割为合适的页面数量。

**输出格式要求：**
请严格按照以下JSON格式返回：

```json
{{{{
  "analysis": {{{{
    "total_pages": 4,
    "content_type": "技术介绍",
    "split_strategy": "按发展阶段分页",
    "reasoning": "文本描述了技术发展的多个阶段，适合按时间线分页展示"
  }}}},
  "pages": [
    {{{{
      "page_number": 1,
      "page_type": "title",
      "title": "人工智能发展历程",
      "date": "2024年7月",
      "original_text_segment": "人工智能发展历程"
    }}}},
    {{{{
      "page_number": 2,
      "page_type": "table_of_contents",
      "title": "目录",
      "original_text_segment": "AI发展概述\n技术突破阶段\n当前发展趋势\n未来展望"
    }}}},
    {{{{
      "page_number": 3,
      "page_type": "content", 
      "title": "AI发展概述",
      "original_text_segment": "人工智能技术发展经历了多个重要阶段。从1950年代的符号主义开始，到1980年代专家系统的兴起，再到2010年代深度学习的突破，以及当前大语言模型时代的到来..."
    }}}}
  ]
}}}}
```

**页面类型说明：**
- `title`: 标题页，仅包含文档标题和日期
- `table_of_contents`: 目录页，必须包含各章节标题（不含页码）
- `content`: 内容页，具体的要点和详细内容（分页重点）

**关键注意事项：**
- **title字段**：必须准确概括该页内容（用于生成目录）
- **original_text_segment字段最重要**：必须包含该页对应的完整原文片段，不能遗漏或截断
- **标题页original_text_segment**：只包含提取的标题部分
- **目录页original_text_segment**：包含各章节标题，每行一个标题
- **内容页original_text_segment**：包含该页面对应的所有原文内容，确保完整性
- 不要生成结尾页，系统将使用预设的固定结尾页模板

只返回JSON格式，不要其他文字。"""
    
    def _parse_ai_response(self, content: str, user_text: str) -> Dict[str, Any]:
        """解析AI响应结果"""
        try:
            # 检查返回内容是否为空
            if not content or not content.strip():
                error_detail = f"AI返回内容为空。原始内容: '{content}'"
                print(f"❌ {error_detail}")
                raise ValueError(error_detail)
            
            
            # 提取JSON内容
            json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                # 如果没有代码块，尝试直接解析
                json_str = content.strip()
            
            if not json_str or not json_str.strip():
                error_detail = "提取的JSON字符串为空"
                print(f"❌ {error_detail}")
                raise ValueError(error_detail)
            
            # 解析JSON
            result = json.loads(json_str)
            
            # 验证结果格式
            if not self._validate_split_result(result):
                error_detail = "AI返回的JSON格式不符合要求"
                print(f"❌ {error_detail}")
                raise ValueError(error_detail)
            
            result['success'] = True
            result['original_text'] = user_text
            
            # 添加固定的结尾页
            self._add_ending_page(result)
            
            return result
                
        except json.JSONDecodeError as e:
            json_str_safe = json_str[:500] if 'json_str' in locals() else '未获取到'
            error_msg = f"JSON解析失败: {e}\n尝试解析的内容: {json_str_safe}"
            print(f"❌ {error_msg}")
            
            
            raise ValueError(error_msg)
        except Exception as e:
            content_safe = content[:500] if content else 'N/A'
            error_msg = f"AI分页解析失败: {e}\n原始AI返回内容: {content_safe}..."
            print(f"❌ {error_msg}")
            
            
            raise e
    
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
            required_page_fields = ['page_number', 'page_type', 'title', 'original_text_segment']
            for page in pages:
                for field in required_page_fields:
                    if field not in page:
                        return False
                
                # 检查original_text_segment是字符串
                if not isinstance(page['original_text_segment'], str):
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
            "date": current_date,
            "original_text_segment": title  # 只包含标题部分
        })
        
        # 将除标题外的所有内容分配到第3页开始的内容页（第2页是固定目录页）
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
        
        page_num = 3  # 从第3页开始（第2页是固定目录页）
        if remaining_paragraphs:
            for i, paragraph in enumerate(remaining_paragraphs):
                # 限制总页数不超过23页（为目录页和结尾页预留空间）
                if page_num > 23:
                    print(f"警告：内容过多，已达到23页上限，剩余{len(remaining_paragraphs) - i}段内容将被省略")
                    break
                    
                pages.append({
                    "page_number": page_num,
                    "page_type": "content",
                    "title": f"内容 {page_num - 2}",
                    "original_text_segment": paragraph
                })
                page_num += 1
        else:
            # 如果没有剩余内容，至少创建一个空的内容页
            pages.append({
                "page_number": 3,
                "page_type": "content",
                "title": "内容页",
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
        
        # 添加固定的目录页和结尾页
        self._add_table_of_contents_page(result)
        self._add_ending_page(result)
        
        return result
    
    def _add_table_of_contents_page(self, result: Dict[str, Any]) -> None:
        """添加动态目录页（第2页）"""
        import os
        
        pages = result.get('pages', [])
        if not pages:
            return
        
        # 调整现有页面的页码（为目录页腾出第2页位置）
        for page in pages:
            if page.get('page_number', 1) > 1:
                page['page_number'] = page['page_number'] + 1
        
        # 提取所有内容页的标题信息，生成动态目录
        content_titles = []
        for page in pages:
            if page.get('page_type') == 'content' and page.get('title'):
                page_number = page.get('page_number', 0)
                title = page.get('title', '').strip()
                subtitle = page.get('subtitle', '').strip()
                
                # 构建目录项
                if subtitle:
                    toc_item = f"{page_number}. {title} - {subtitle}"
                else:
                    toc_item = f"{page_number}. {title}"
                content_titles.append(toc_item)
        
        # 如果没有提取到标题，使用默认目录
        if not content_titles:
            content_titles = [
                "演示内容导航",
                "章节结构预览"
            ]
        
        # 创建动态目录页信息
        table_of_contents_page = {
            "page_number": 2,
            "page_type": "table_of_contents",
            "title": "目录",
            "original_text_segment": "",
            "template_path": os.path.join("templates", "table_of_contents_slides.pptx"),
            "is_toc_page": True,  # 标记为目录页
            "skip_dify_api": True,  # 不需要调用Dify API，但内容已动态提取
            "toc_items": content_titles  # 将目录项单独存储
        }
        
        # 将目录页插入到第2位
        pages.insert(1, table_of_contents_page)
        
        # 更新分析信息中的总页数
        if 'analysis' in result:
            result['analysis']['total_pages'] = len(pages)
    
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
        
        # 显示原文片段
        original_text = page.get('original_text_segment', '')
        if original_text and original_text.strip():
            preview += "**原文内容：**\n"
            # 如果原文太长，显示前200字符
            if len(original_text) > 200:
                preview += f"{original_text[:200]}...\n"
            else:
                preview += f"{original_text}\n"
        
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