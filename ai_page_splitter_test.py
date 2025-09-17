#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AI智能分页模块（测试版本）
专门用于测试两次调用策略的独立版本
"""

import re
import json
import requests
from typing import Dict, List, Any, Optional, Tuple
from openai import OpenAI
from config import get_config
from logger import log_user_action

class AIPageSplitterTest:
    """AI智能分页处理器（测试版本）"""
    
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
        
        if model_info.get('api_provider') == 'Volces' and model_info.get('use_multiple_keys'):
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
        elif model_info.get('api_provider') == 'Liai':
            # Liai API多密钥负载均衡
            self.api_keys = []
            for i in range(1, 6):  # 支持1-5个密钥
                key_name = f'LIAI_API_KEY_{i}'
                key_value = os.getenv(key_name)
                if key_value:
                    self.api_keys.append(key_value)
            
            # 如果没有找到编号密钥，尝试单个密钥
            if not self.api_keys:
                single_key = os.getenv('LIAI_API_KEY')
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
        将用户文本智能分割为多个PPT页面（测试版本 - 固定使用两次调用策略）
        
        Args:
            user_text: 用户输入的原始文本
            target_pages: 目标页面数量（可选，由AI自动判断）
            
        Returns:
            Dict: 分页结果，包含每页的内容和分析
        """
        log_user_action("AI智能分页测试", f"文本长度: {len(user_text)}, 两次调用策略")
        
        try:
            # 固定使用两次调用策略
            return self._split_with_two_pass(user_text, target_pages)
            
        except Exception as e:
            print(f"AI分页分析失败: {e}")
            raise e
    
    def _split_with_two_pass(self, user_text: str, target_pages: Optional[int]) -> Dict[str, Any]:
        """两次调用分页策略：第一次注重逻辑性，第二次注重分页数"""
        print(f"🔄 开始两次调用AI分页策略，目标页数: {target_pages}")
        
        # 第一次调用：注重逻辑结构，不强制页数
        print("📝 第一次调用：分析内容逻辑结构...")
        first_system_prompt = self._build_logical_structure_prompt()
        first_content = self._call_api_with_prompt(first_system_prompt, user_text)
        first_result = self._parse_ai_response_without_ending(first_content, user_text)  # 不添加结尾页
        
        print(f"✅ 第一次调用完成，生成 {first_result['analysis']['total_pages']} 页")
        
        # 第二次调用：基于第一次结果，调整页数
        if target_pages:
            print(f"🎯 第二次调用：调整页数至目标 {target_pages} 页...")
        else:
            print(f"🎯 第二次调用：优化页数（当前 {first_result['analysis']['total_pages']} 页，减少过度分页）...")
        second_system_prompt = self._build_page_adjustment_prompt(target_pages)
        
        # 将第一次的结果作为上下文传给第二次调用
        first_result_text = self._format_first_result_for_second_call(first_result)
        second_content = self._call_api_with_prompt(second_system_prompt, first_result_text)
        second_result = self._parse_ai_response(second_content, user_text)
        
        print(f"✅ 第二次调用完成，最终生成 {second_result['analysis']['total_pages']} 页")
        
        # 标记为两次调用结果
        second_result['is_two_pass_result'] = True
        second_result['first_pass_pages'] = first_result['analysis']['total_pages'] + 1  # 第一次页数 + 结尾页
        second_result['final_pass_pages'] = second_result['analysis']['total_pages']  # 第二次页数已包含结尾页
        
        return second_result
    
    def _call_api_with_prompt(self, system_prompt: str, user_text: str) -> str:
        """根据配置调用相应的API"""
        model_info = self.config.get_model_info()
        if model_info.get('request_format') == 'dify_compatible':
            # 使用Liai API格式
            return self._call_liai_api(system_prompt, user_text)
        elif model_info.get('request_format') == 'streaming_compatible':
            # 使用火山引擎DeepSeek API格式
            return self._call_deepseek_api(system_prompt, user_text)
        else:
            # 标准OpenAI API格式
            request_timeout = 60
            actual_model = model_info.get('actual_model', self.config.ai_model)
            
            # 创建临时客户端（如果还没有）
            if not hasattr(self, 'client'):
                from openai import OpenAI
                self.client = OpenAI(
                    api_key=self._get_next_api_key(),
                    base_url=self.base_url,
                    timeout=request_timeout
                )
            
            response = self.client.chat.completions.create(
                model=actual_model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_text}
                ],
                temperature=self.config.ai_temperature,
                stream=True,
                timeout=request_timeout
            )
            
            # 收集流式响应内容
            content = ""
            for chunk in response:
                if chunk.choices and chunk.choices[0].delta.content:
                    content += chunk.choices[0].delta.content
            
            return content.strip() if content else ""
    
    def _build_logical_structure_prompt(self) -> str:
        """构建第一次调用的逻辑结构分析提示（不强制页数）"""
        return f"""你是一个资深PPT架构师。请按照以下**严格流程**将文本转化为PPT分页大纲：

**第一步：全局分析**
- 首先，通读全文，识别出文本的**核心逻辑结构**（如：引言->问题分析->数据论证->解决方案->总结）
- 将整个文本划分为几个主要部分

**第二步：逐部分分页**
- 对于**每一个主要部分**，执行以下操作：
  1. **提取核心论点**：找出这部分要证明的1个最终观点
  2. **收集论据**：将所有支持该论点的段落、数据和论据集合起来
  3. **合并成一页**：**将上述所有内容（核心论点+所有论据）共同作为一页PPT的文本内容**。即使内容很长，也先放在一起
  4. **保留完整文本**：无论怎么分页，每一页都必须包含该页对应的完整用户原始文本，不能遗漏或截断

**第三步：拆分例外规则**
- **仅在以下情况下**，才允许将一页内容拆分成多页：
  a. 包含了**两个完全独立的核心论点**

**分页策略：**
- **标题页（第1页）**：PPT封面页，不对应任何原文内容，自动生成标题和日期
- **目录页（第2页）**：AI根据内容结构生成完整目录
- **内容页（第3页开始）**：处理所有原文内容，按逻辑结构分页
- **结尾页**：不生成结尾页（使用预设模板）

**标题页处理规则：**
- 标题页是PPT的封面，生成合适的PPT标题
- 自动生成标题（基于内容主题）
- original_text_segment与title相同，包含PPT标题
- 所有原文内容都从第2页（目录）和第3页开始处理

**页面类型说明：**
- `title`: 标题页，仅包含文档标题和日期
- `table_of_contents`: 目录页，必须包含各章节标题（不含页码）
- `content`: 内容页，具体的要点和详细内容（分页重点）

**字段要求：**
pages字段里只需要包含：page_number/page_type/title/original_text_segment字段
- **title字段**：必须准确概括该页内容（用于生成目录）
- **original_text_segment字段最重要**：必须包含该页对应的完整原文片段，不能遗漏或截断

**关键注意事项：**
- **标题页original_text_segment**：与title相同，包含PPT标题
- **目录页original_text_segment**：包含各章节标题，每行一个标题
- **内容页original_text_segment**：包含该页面对应的所有原文内容，确保完整性
- 不要生成结尾页，系统将使用预设的固定结尾页模板

**输出格式要求：**
严格按照以下JSON格式返回：

```json
[
  {{{{
    "page_number": 1,
    "page_type": "title",
    "title": "PPT标题（基于内容主题生成）",
    "original_text_segment": "PPT标题（基于内容主题生成）"
  }}}},
  {{{{
    "page_number": 2,
    "page_type": "table_of_contents",
    "title": "目录",
    "original_text_segment": "主题一\n主题二\n主题三"
  }}}},
  {{{{
    "page_number": 3,
    "page_type": "content",
    "title": "主题一标题",
    "original_text_segment": "完整的主题一内容..."
  }}}}
]
```

只返回JSON格式，不要其他文字。"""

    def _build_page_adjustment_prompt(self, target_pages: Optional[int]) -> str:
        """构建第二次调用的页数调整提示"""
        if target_pages:
            # 有指定目标页数：精确调整
            ai_pages = target_pages - 1  # AI生成页数 = 总页数 - 结尾页
            return f"""你是PPT页数精确调整专家。用户明确要求PPT总共{target_pages}页，你必须严格满足这个需求。

【严格要求】你只需生成{ai_pages}页内容，系统会自动添加第{target_pages}页结尾页！

**PPT页数调整任务：**
基于第一次AI分析结果，重新组织PPT内容以精确满足用户的{target_pages}页要求：

**页面分配：**
- 你负责生成：{ai_pages}页内容（第1页到第{ai_pages}页）
- 系统自动添加：第{target_pages}页结尾页
- 最终PPT总页数：{target_pages}页（完全符合用户要求）

**调整策略：**
- 保持标题页(第1页)和目录页(第2页)不变
- 内容页范围：第3页到第{ai_pages}页
- 通过合并或拆分内容页来精确达到{ai_pages}页
- 确保每页内容充实，符合PPT展示标准
- **【严格300字限制】除了标题页、目录页和结尾页，所有内容页的original_text_segment必须包含至少300字原始文本，不足300字的页面必须与相邻页面合并**

**字段要求：**
pages字段里只需要包含：page_number/page_type/title/original_text_segment字段
- **title字段**：必须准确概括该页内容
- **original_text_segment字段**：包含该页对应的完整原文片段，不能遗漏

严格按JSON格式返回，必须生成{ai_pages}页：

```json
[
  {{{{
    "page_number": 1,
    "page_type": "title",
    "title": "PPT标题",
    "original_text_segment": "PPT标题"
  }}}},
  {{{{
    "page_number": 2,
    "page_type": "table_of_contents",
    "title": "目录",
    "original_text_segment": "目录内容"
  }}}},
  {{{{
    "page_number": 3,
    "page_type": "content",
    "title": "内容页标题",
    "original_text_segment": "页面内容"
  }}}}
]
```

只返回JSON，不要其他文字。"""
        else:
            # 无指定目标页数：优化减少页数
            return f"""你是PPT内容优化专家。基于第一次AI分析结果，优化PPT页数分配，解决过度分页问题。

【PPT分页优化任务】
PPT制作中，AI容易过度分页导致页面内容稀薄。你需要通过合并相关主题的内容页来优化页数：

**分页原则：**
- 保持标题页(第1页)和目录页(第2页)不变
- 合并逻辑相关的内容页（如"产品介绍"+"产品特点"合并为一页）
- 确保每页内容充实，避免内容过少或过多
- 优化后的AI生成页数应比第一次结果更少（系统会自动添加结尾页）
- **【严格300字限制】除了标题页、目录页和结尾页，所有内容页的original_text_segment必须包含至少300字原始文本，不足300字的页面必须与相邻页面合并**

**字段要求：**
pages字段里只需要包含：page_number/page_type/title/original_text_segment字段
- **title字段**：必须准确概括该页内容
- **original_text_segment字段**：包含该页对应的完整原文片段，不能遗漏

严格按JSON格式返回：

```json
[
  {{{{
    "page_number": 1,
    "page_type": "title",
    "title": "PPT标题",
    "original_text_segment": "PPT标题"
  }}}},
  {{{{
    "page_number": 2,
    "page_type": "table_of_contents",
    "title": "目录",
    "original_text_segment": "目录内容"
  }}}},
  {{{{
    "page_number": 3,
    "page_type": "content",
    "title": "内容页标题",
    "original_text_segment": "页面内容"
  }}}}
]
```

只返回JSON，不要其他文字。"""

    def _format_first_result_for_second_call(self, first_result: Dict[str, Any]) -> str:
        """将第一次调用结果格式化为第二次调用的输入"""
        formatted_text = "【第一次AI分析结果】\n\n"
        
        # 添加分析信息
        analysis = first_result.get('analysis', {})
        formatted_text += f"原始分析：总页数{analysis.get('total_pages', 0)}页，{analysis.get('split_strategy', '未知策略')}\n\n"
        
        # 添加每页的详细内容
        pages = first_result.get('pages', [])
        formatted_text += "【页面详情】\n"
        for page in pages:
            page_num = page.get('page_number', 0)
            page_type = page.get('page_type', 'content')
            title = page.get('title', '无标题')
            original_text = page.get('original_text_segment', '')
            
            formatted_text += f"\n第{page_num}页 ({page_type}): {title}\n"
            formatted_text += f"内容: {original_text}\n"
            formatted_text += "---\n"
        
        # 添加原始文本
        formatted_text += f"\n【原始文本】\n{first_result.get('original_text', '')}"
        
        return formatted_text

    def _call_liai_api(self, system_prompt: str, user_text: str) -> str:
        """调用Liai API（支持多密钥负载均衡）"""
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
        
        # 尝试所有可用密钥
        last_exception = None
        for attempt in range(len(self.api_keys)):
            current_api_key = self._get_next_api_key()
            
            headers = {
                'Authorization': f'Bearer {current_api_key}',
                'Content-Type': 'application/json',
                'Connection': 'keep-alive'  # 保持连接
            }
            
            try:
                print(f"尝试使用Liai API密钥 {attempt + 1}/{len(self.api_keys)} (末尾: ...{current_api_key[-8:]})")
                
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
                
                # 成功获取内容，返回结果
                if content.strip():
                    print(f"✅ Liai API密钥 ...{current_api_key[-8:]} 调用成功")
                    return content.strip()
                else:
                    raise Exception("API返回空内容")
                    
            except Exception as e:
                last_exception = e
                print(f"❌ Liai API密钥 ...{current_api_key[-8:]} 调用失败: {e}")
                
                # 如果还有其他密钥可以尝试，继续下一个
                if attempt < len(self.api_keys) - 1:
                    print(f"⏳ 尝试下一个Liai API密钥...")
                    continue
        
        # 所有密钥都失败了
        print(f"❌ 所有{len(self.api_keys)}个Liai API密钥都失败了")
        raise last_exception or Exception("所有Liai API密钥调用失败")
    
    def _call_deepseek_api(self, system_prompt: str, user_text: str) -> str:
        """调用DeepSeek API（带故障转移的多密钥负载均衡）"""
        model_info = self.config.get_model_info()
        
        # 获取实际模型名称和额外头部
        actual_model = model_info.get('actual_model', 'deepseek-v3-250324')
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
                    temperature=self.config.ai_temperature,
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
    
    def _parse_ai_response_without_ending(self, content: str, user_text: str) -> Dict[str, Any]:
        """解析AI响应结果（不添加结尾页）"""
        result = self._parse_ai_response_base(content, user_text)
        return result
    
    def _parse_ai_response(self, content: str, user_text: str) -> Dict[str, Any]:
        """解析AI响应结果（添加结尾页）"""
        result = self._parse_ai_response_base(content, user_text)
        
        # 添加固定的结尾页
        self._add_ending_page(result)
        
        return result
    
    def _parse_ai_response_base(self, content: str, user_text: str) -> Dict[str, Any]:
        """解析AI响应结果的基础方法"""
        try:
            # 检查返回内容是否为空
            if not content or not content.strip():
                error_detail = f"AI返回内容为空。原始内容: '{content}'"
                print(f"❌ {error_detail}")
                raise ValueError(error_detail)
            
            # 提取JSON内容（支持对象{}和数组[]）
            json_match = re.search(r'```(?:json)?\s*([{\[].*?[}\]])\s*```', content, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                # 如果没有代码块，尝试直接解析
                json_str = content.strip()
            
            if not json_str or not json_str.strip():
                error_detail = "提取的JSON字符串为空"
                print(f"❌ {error_detail}")
                raise ValueError(error_detail)
            
            # 检查JSON是否被截断
            if json_str.strip().endswith(',"page_nu') or not json_str.strip().endswith((']', '}')):
                error_detail = f"JSON响应被截断，可能是token限制导致。JSON末尾: ...{json_str[-50:]}"
                print(f"❌ {error_detail}")
                raise ValueError(error_detail)
            
            # 解析JSON
            parsed_data = json.loads(json_str)
            
            # 如果返回的是数组，转换为标准格式
            if isinstance(parsed_data, list):
                result = {
                    'pages': parsed_data,
                    'analysis': {
                        'total_pages': len(parsed_data),
                        'content_type': '自动生成',
                        'split_strategy': '智能分页'
                    }
                }
            else:
                result = parsed_data
            
            # 验证结果格式
            validation_result = self._validate_split_result(result)
            if not validation_result['is_valid']:
                error_detail = f"AI返回的JSON格式不符合要求: {validation_result['error']}"
                print(f"❌ {error_detail}")
                print(f"🔍 JSON内容: {json.dumps(result, ensure_ascii=False, indent=2)[:1000]}...")
                raise ValueError(error_detail)
            
            result['success'] = True
            result['original_text'] = user_text
            
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
    
    def _validate_split_result(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """验证分页结果的格式"""
        try:
            # 检查必需的字段
            if 'analysis' not in result:
                return {'is_valid': False, 'error': '缺少analysis字段'}
            if 'pages' not in result:
                return {'is_valid': False, 'error': '缺少pages字段'}
            
            analysis = result['analysis']
            pages = result['pages']
            
            # 检查analysis字段
            required_analysis_fields = ['total_pages', 'content_type', 'split_strategy']
            for field in required_analysis_fields:
                if field not in analysis:
                    return {'is_valid': False, 'error': f'analysis缺少字段: {field}'}
            
            # 检查pages数组
            if not isinstance(pages, list):
                return {'is_valid': False, 'error': 'pages不是数组类型'}
            if len(pages) == 0:
                return {'is_valid': False, 'error': 'pages数组为空'}
            
            # 检查每个页面的字段
            required_page_fields = ['page_number', 'page_type', 'title', 'original_text_segment']
            for i, page in enumerate(pages):
                for field in required_page_fields:
                    if field not in page:
                        return {'is_valid': False, 'error': f'第{i+1}个页面缺少字段: {field}'}
                
                # 检查original_text_segment是字符串
                if not isinstance(page['original_text_segment'], str):
                    return {'is_valid': False, 'error': f'第{i+1}个页面的original_text_segment不是字符串类型'}
            
            return {'is_valid': True, 'error': None}
            
        except Exception as e:
            return {'is_valid': False, 'error': f'验证过程中出现异常: {str(e)}'}
    
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