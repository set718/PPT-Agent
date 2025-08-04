#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Dify API客户端模块
用于在文本分页后调用Dify API，对每页内容进行处理
"""

import asyncio
import aiohttp
import time
from typing import Dict, List, Any, Optional, Tuple
import json
from dataclasses import dataclass, field
from logger import get_logger, log_user_action

logger = get_logger()

class APIKeyBalancer:
    """API密钥负载均衡器"""
    
    def __init__(self, api_keys: List[str], strategy: str = "round_robin"):
        """
        初始化负载均衡器
        
        Args:
            api_keys: API密钥列表
            strategy: 负载均衡策略 (round_robin, random, least_used)
        """
        self.api_keys = api_keys
        self.strategy = strategy
        self.current_index = 0
        self.usage_count = {key: 0 for key in api_keys}
        self.failed_keys = set()
        self.failure_count = {key: 0 for key in api_keys}  # 记录每个密钥的失败次数
        self.last_failure_time = {key: 0 for key in api_keys}  # 记录最后失败时间
        
        logger.info(f"初始化API密钥负载均衡器，策略: {strategy}, 密钥数量: {len(api_keys)}")
    
    def get_next_key(self) -> str:
        """获取下一个API密钥"""
        import time
        current_time = time.time()
        
        # 智能恢复机制：检查失效的密钥是否可以恢复
        keys_to_recover = []
        for key in list(self.failed_keys):
            last_fail_time = self.last_failure_time.get(key, 0)
            # 如果密钥失效超过60秒，尝试恢复
            if current_time - last_fail_time > 60:
                keys_to_recover.append(key)
        
        for key in keys_to_recover:
            self.failed_keys.discard(key)
            logger.info(f"API密钥自动恢复: {key[:20]}... (失效时间超过60秒)")
        
        available_keys = [key for key in self.api_keys if key not in self.failed_keys]
        
        if not available_keys:
            # 如果所有密钥都失败了，强制重置失败列表
            logger.warning(f"所有{len(self.api_keys)}个API密钥都失效，强制重置失败列表以确保服务可用")
            self.failed_keys.clear()
            # 重置失败计数
            self.failure_count = {key: 0 for key in self.api_keys}
            available_keys = self.api_keys
        
        if self.strategy == "round_robin":
            key = self._round_robin_select(available_keys)
        elif self.strategy == "random":
            key = self._random_select(available_keys)
        elif self.strategy == "least_used":
            key = self._least_used_select(available_keys)
        else:
            key = available_keys[0]  # 默认选择第一个
        
        self.usage_count[key] += 1
        logger.debug(f"选择API密钥: {key[:20]}..., 使用次数: {self.usage_count[key]}")
        return key
    
    def _round_robin_select(self, available_keys: List[str]) -> str:
        """轮询选择"""
        if not available_keys:
            return self.api_keys[0]
        
        key = available_keys[self.current_index % len(available_keys)]
        self.current_index += 1
        return key
    
    def _random_select(self, available_keys: List[str]) -> str:
        """随机选择"""
        import random
        return random.choice(available_keys) if available_keys else self.api_keys[0]
    
    def _least_used_select(self, available_keys: List[str]) -> str:
        """选择使用次数最少的密钥"""
        if not available_keys:
            return self.api_keys[0]
        
        return min(available_keys, key=lambda k: self.usage_count.get(k, 0))
    
    def mark_key_failed(self, api_key: str):
        """标记密钥失败"""
        import time
        self.failed_keys.add(api_key)
        self.failure_count[api_key] = self.failure_count.get(api_key, 0) + 1
        self.last_failure_time[api_key] = time.time()
        
        failure_count = self.failure_count[api_key]
        logger.warning(f"标记API密钥失败: {api_key[:20]}... (第{failure_count}次失败)")
    
    def mark_key_success(self, api_key: str):
        """标记密钥成功（从失败列表中移除）"""
        if api_key in self.failed_keys:
            self.failed_keys.remove(api_key)
            # 重置失败计数和时间
            self.failure_count[api_key] = 0
            self.last_failure_time[api_key] = 0
            logger.info(f"API密钥恢复正常: {api_key[:20]}... (失败计数已重置)")
    
    def get_usage_stats(self) -> Dict[str, Any]:
        """获取使用统计"""
        return {
            "total_keys": len(self.api_keys),
            "available_keys": len(self.api_keys) - len(self.failed_keys),
            "failed_keys": len(self.failed_keys),
            "usage_count": dict(self.usage_count),
            "strategy": self.strategy
        }

@dataclass
class DifyAPIConfig:
    """Dify API配置类 - 支持多API密钥负载均衡"""
    base_url: str = "https://api.dify.ai/v1"
    api_keys: List[str] = field(default_factory=lambda: [
        "app-7HOcCxB7uosj23f1xgjFClkv",
        "app-vxEWYWTaakWITl041b8UHBCN", 
        "app-WM17uKVOQHpYE4sNyxRH0dtG",
        "app-dmKCw2gATM5mbC2VXS0Htiyu",
        "app-wC0jXMfM1qwliIaPtrq9evuE",
        "app-0QHbweS3gsr9q3o5R08Q8bYE",
        "app-mvjmGCfgEV86qjsMihwaGB2T",
        "app-pzeGjsmkvH9oQ9vNK7DX9Z0a"
    ])
    endpoint: str = "/chat-messages"
    timeout: int = 180  # 增加到3分钟超时
    max_retries: int = 8  # 增加到8次重试
    retry_delay: float = 3.0  # 增加重试间隔
    max_concurrent: int = 8   # 降低并发数，减少服务器压力，提高成功率
    load_balance_strategy: str = "round_robin"  # round_robin, random, least_used
    
    @property
    def api_key(self) -> str:
        """向后兼容：返回第一个API密钥"""
        return self.api_keys[0] if self.api_keys else ""

class DifyAPIClient:
    """Dify API客户端 - 支持多API密钥负载均衡"""
    
    def __init__(self, config: Optional[DifyAPIConfig] = None):
        """初始化Dify API客户端"""
        self.config = config or DifyAPIConfig()
        self.session = None
        
        # 初始化负载均衡器
        self.key_balancer = APIKeyBalancer(
            self.config.api_keys, 
            self.config.load_balance_strategy
        )
        
        logger.info(f"初始化Dify API客户端，支持{len(self.config.api_keys)}个API密钥")
    
    async def __aenter__(self):
        """异步上下文管理器入口"""
        # 创建连接器，优化连接参数
        connector = aiohttp.TCPConnector(
            limit=10,  # 总连接数限制
            limit_per_host=5,  # 每个主机的连接数限制
            ttl_dns_cache=300,  # DNS缓存时间
            use_dns_cache=True,
            keepalive_timeout=60,  # 保持连接时间
            enable_cleanup_closed=True
        )
        
        self.session = aiohttp.ClientSession(
            connector=connector,
            timeout=aiohttp.ClientTimeout(
                total=self.config.timeout,
                connect=30,  # 增加连接超时到30秒
                sock_read=120  # 增加读取超时到2分钟
            ),
            headers={
                'Content-Type': 'application/json',
                'User-Agent': 'Dify-API-Client/2.0-MultiKey'
            }  # Authorization header will be set per request
        )
        return self
    
    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """异步上下文管理器出口"""
        if self.session:
            await self.session.close()
    
    async def _make_single_request(self, page_data: Dict[str, Any], page_index: int) -> Dict[str, Any]:
        """
        对单个页面内容发起API请求
        
        Args:
            page_data: 页面数据
            page_index: 页面索引
            
        Returns:
            Dict: API响应结果
        """
        # 构建请求输入内容
        input_text = self._format_page_content(page_data)
        
        request_data = {
            "inputs": {},
            "query": input_text,
            "response_mode": "blocking",
            "conversation_id": "",
            "user": f"pagination_user_{int(time.time())}"
        }
        
        url = f"{self.config.base_url}{self.config.endpoint}"
        
        # 重试逻辑（现在支持多API密钥）
        current_api_key = None
        for attempt in range(self.config.max_retries):
            # 获取下一个API密钥
            current_api_key = self.key_balancer.get_next_key()
            
            # 为当前请求设置Authorization头
            headers = {
                'Authorization': f'Bearer {current_api_key}',
                'Content-Type': 'application/json'
            }
            
            try:
                logger.info(f"开始请求第{page_index + 1}页内容 (尝试 {attempt + 1}/{self.config.max_retries}, API密钥: {current_api_key[:20]}...)")
                
                async with self.session.post(url, json=request_data, headers=headers) as response:
                    if response.status == 200:
                        # 尝试正确解码响应
                        try:
                            result = await response.json(encoding='utf-8')
                        except:
                            result = await response.json()
                        
                        logger.info(f"第{page_index + 1}页API请求成功 (使用密钥: {current_api_key[:20]}...)")
                        
                        # 标记该API密钥成功
                        self.key_balancer.mark_key_success(current_api_key)
                        
                        # 根据不同的响应格式提取文本内容
                        response_text = ""
                        if 'answer' in result:
                            response_text = result.get('answer', '')
                        elif 'message' in result and 'content' in result['message']:
                            response_text = result['message']['content']
                        elif 'data' in result and isinstance(result['data'], dict):
                            response_text = result['data'].get('answer', result['data'].get('content', ''))
                        else:
                            # 如果找不到标准字段，尝试将整个结果转为字符串
                            response_text = str(result)
                        
                        # 如果响应文本为空或看起来有问题，使用备用方案
                        if not response_text or len(response_text.strip()) == 0:
                            response_text = f"API响应成功，但内容为空。原始响应包含以下字段: {list(result.keys())}"
                        
                        return {
                            "success": True,
                            "page_index": page_index,
                            "page_number": page_data.get('page_number', page_index + 1),
                            "input_content": input_text,
                            "api_response": result,
                            "response_text": response_text,
                            "api_status": response.status,
                            "attempt": attempt + 1,
                            "used_api_key": current_api_key[:20] + "..."
                        }
                    else:
                        error_text = await response.text()
                        logger.warning(f"第{page_index + 1}页API请求失败，状态码: {response.status} (使用密钥: {current_api_key[:20]}...)")
                        
                        # 如果是认证错误，标记该API密钥失败
                        if response.status in [401, 403]:
                            self.key_balancer.mark_key_failed(current_api_key)
                            logger.warning(f"API密钥认证失败，已标记为失败: {current_api_key[:20]}...")
                        
                        if attempt < self.config.max_retries - 1:
                            delay = self.config.retry_delay * (2 ** attempt)  # 指数退避
                            logger.info(f"等待 {delay:.1f} 秒后使用下一个API密钥重试...")
                            await asyncio.sleep(delay)
                            continue
                        else:
                            return {
                                "success": False,
                                "page_index": page_index,
                                "page_number": page_data.get('page_number', page_index + 1),
                                "input_content": input_text,
                                "error": f"HTTP {response.status}: {error_text}",
                                "api_status": response.status,
                                "attempts": self.config.max_retries,
                                "last_used_api_key": current_api_key[:20] + "..."
                            }
            
            except asyncio.TimeoutError as e:
                logger.warning(f"第{page_index + 1}页API请求超时 (尝试 {attempt + 1}/{self.config.max_retries}，使用密钥: {current_api_key[:20]}...)")
                
                # 超时时也标记当前API密钥可能有问题
                if attempt >= 2:  # 超时2次后标记密钥为临时失效
                    self.key_balancer.mark_key_failed(current_api_key)
                    logger.info(f"API密钥因多次超时被临时标记为失效: {current_api_key[:20]}...")
                
                if attempt < self.config.max_retries - 1:
                    # 增强的指数退避策略
                    base_delay = self.config.retry_delay
                    exponential_delay = base_delay * (2 ** attempt)
                    jitter = exponential_delay * 0.1  # 添加10%的随机抖动
                    import random
                    delay = exponential_delay + random.uniform(-jitter, jitter)
                    delay = min(delay, 30)  # 最大延迟30秒
                    
                    logger.info(f"超时重试：等待 {delay:.1f} 秒后使用下一个API密钥重试...")
                    await asyncio.sleep(delay)
                    continue
                else:
                    # 所有重试都失败了，但不应该发生这种情况
                    logger.error(f"第{page_index + 1}页经过{self.config.max_retries}次重试仍然超时，这不应该发生！")
                    return {
                        "success": False,
                        "page_index": page_index,
                        "page_number": page_data.get('page_number', page_index + 1),
                        "input_content": input_text,
                        "error": f"经过{self.config.max_retries}次重试仍然超时 (超时限制: {self.config.timeout}秒)",
                        "attempts": self.config.max_retries,
                        "all_used_keys": [k[:20] + "..." for k in self.config.api_keys]
                    }
            
            except aiohttp.ClientConnectorError as e:
                logger.warning(f"第{page_index + 1}页连接错误 (尝试 {attempt + 1}/{self.config.max_retries}，使用密钥: {current_api_key[:20]}...): {str(e)}")
                
                # 连接错误时也可能是API密钥或服务器问题
                if attempt >= 1:  # 连接错误1次后就切换密钥
                    self.key_balancer.mark_key_failed(current_api_key)
                    logger.info(f"API密钥因连接错误被标记为失效: {current_api_key[:20]}...")
                
                if attempt < self.config.max_retries - 1:
                    delay = self.config.retry_delay * (2 ** attempt)
                    delay = min(delay, 20)  # 连接错误最大等待20秒
                    logger.info(f"连接错误重试：等待 {delay:.1f} 秒后使用下一个API密钥重试...")
                    await asyncio.sleep(delay)
                    continue
                else:
                    logger.error(f"第{page_index + 1}页经过{self.config.max_retries}次重试仍然连接失败！")
                    return {
                        "success": False,
                        "page_index": page_index,
                        "page_number": page_data.get('page_number', page_index + 1),
                        "input_content": input_text,
                        "error": f"经过{self.config.max_retries}次重试仍然连接失败: {str(e)}",
                        "attempts": self.config.max_retries
                    }
            
            except Exception as e:
                logger.error(f"第{page_index + 1}页API请求异常 (尝试 {attempt + 1}/{self.config.max_retries}，使用密钥: {current_api_key[:20]}...): {str(e)}")
                
                # 未知异常时也切换API密钥
                if attempt >= 0:  # 任何异常都立即切换密钥
                    self.key_balancer.mark_key_failed(current_api_key)
                    logger.info(f"API密钥因异常被标记为失效: {current_api_key[:20]}...")
                
                if attempt < self.config.max_retries - 1:
                    delay = self.config.retry_delay * (attempt + 1)
                    delay = min(delay, 15)  # 异常最大等待15秒
                    logger.info(f"异常重试：等待 {delay:.1f} 秒后使用下一个API密钥重试...")
                    await asyncio.sleep(delay)
                    continue
                else:
                    logger.error(f"第{page_index + 1}页经过{self.config.max_retries}次重试仍然异常！")
                    return {
                        "success": False,
                        "page_index": page_index,
                        "page_number": page_data.get('page_number', page_index + 1),
                        "input_content": input_text,
                        "error": f"经过{self.config.max_retries}次重试仍然异常: {str(e)}",
                        "attempts": self.config.max_retries,
                        "exception_type": type(e).__name__
                    }
        
        # 不应该到达这里
        return {
            "success": False,
            "page_index": page_index,
            "page_number": page_data.get('page_number', page_index + 1),
            "input_content": input_text,
            "error": "未知错误",
            "attempts": self.config.max_retries
        }
    
    def _format_page_content(self, page_data: Dict[str, Any]) -> str:
        """
        格式化页面内容为API输入
        
        Args:
            page_data: 页面数据
            
        Returns:
            str: 格式化后的输入文本
        """
        # 构建结构化的输入内容
        input_parts = []
        
        # 页面基本信息
        page_number = page_data.get('page_number', 1)
        page_type = page_data.get('page_type', 'content')
        title = page_data.get('title', '')
        
        input_parts.append(f"页面信息：第{page_number}页 ({page_type})")
        
        if title:
            input_parts.append(f"标题：{title}")
        
        # 副标题（如果有）
        subtitle = page_data.get('subtitle', '')
        if subtitle:
            input_parts.append(f"副标题：{subtitle}")
        
        # 内容摘要
        content_summary = page_data.get('content_summary', '')
        if content_summary:
            input_parts.append(f"内容摘要：{content_summary}")
        
        # 主要要点
        key_points = page_data.get('key_points', [])
        if key_points:
            input_parts.append("主要要点：")
            for i, point in enumerate(key_points, 1):
                input_parts.append(f"{i}. {point}")
        
        # 原始文本片段
        original_text = page_data.get('original_text_segment', '')
        if original_text:
            input_parts.append(f"原始文本：{original_text}")
        
        return "\n\n".join(input_parts)
    
    async def process_pages_concurrent(self, pages: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        并发处理多个页面的API请求（控制并发数量）
        
        Args:
            pages: 页面数据列表
            
        Returns:
            Dict: 处理结果
        """
        if not pages:
            return {
                "success": False,
                "error": "没有页面数据需要处理",
                "results": []
            }
        
        # 过滤出需要Dify API处理的页面（跳过固定页面）
        fixed_page_types = {'title', 'table_of_contents', 'ending'}
        
        # 分离需要处理和跳过的页面
        pages_to_process = []
        skipped_pages = []
        
        for page in pages:
            page_type = page.get('page_type', 'content')
            if page_type in fixed_page_types:
                skipped_pages.append(page)
                logger.info(f"跳过固定页面: 第{page.get('page_number', '?')}页 ({page_type})")
            else:
                pages_to_process.append(page)
        
        logger.info(f"共{len(pages)}页，需要Dify API处理: {len(pages_to_process)}页，跳过固定页面: {len(skipped_pages)}页")
        
        # 如果没有需要处理的页面，直接返回成功结果
        if not pages_to_process:
            return {
                "success": True,
                "total_pages": len(pages),
                "successful_count": 0,
                "failed_count": 0,
                "exception_count": 0,
                "skipped_count": len(skipped_pages),
                "processing_time": 0,
                "results": [],
                "successful_results": [],
                "failed_results": [],
                "skipped_results": skipped_pages,
                "exceptions": [],
                "api_key_stats": self.key_balancer.get_usage_stats()
            }
        
        start_time = time.time()
        log_user_action("Dify API并发处理", f"开始处理{len(pages_to_process)}个页面（最大并发: {self.config.max_concurrent}，跳过{len(skipped_pages)}个固定页面）")
        
        try:
            # 创建信号量来控制并发数量
            semaphore = asyncio.Semaphore(self.config.max_concurrent)
            
            async def limited_request(page_data, original_index):
                async with semaphore:
                    logger.info(f"开始处理第{page_data.get('page_number', original_index + 1)}页（并发控制）")
                    return await self._make_single_request(page_data, original_index)
            
            # 创建并发任务（只处理需要API调用的页面，但保持原始索引）
            tasks = []
            for page_data in pages_to_process:
                # 找到该页面在原始列表中的索引
                original_index = next(i for i, p in enumerate(pages) if p is page_data)
                tasks.append(limited_request(page_data, original_index))
            
            # 并发执行所有API请求
            results = await asyncio.gather(*tasks, return_exceptions=True)
            
            # 处理结果
            successful_results = []
            failed_results = []
            exceptions = []
            
            for result in results:
                if isinstance(result, Exception):
                    exceptions.append(str(result))
                elif result.get('success', False):
                    successful_results.append(result)
                else:
                    failed_results.append(result)
            
            end_time = time.time()
            processing_time = end_time - start_time
            
            # 获取API密钥使用统计
            key_stats = self.key_balancer.get_usage_stats()
            
            # 构建返回结果
            return_result = {
                "success": len(successful_results) > 0 or len(skipped_pages) > 0,  # 如果有跳过的页面也算成功
                "total_pages": len(pages),
                "successful_count": len(successful_results),
                "failed_count": len(failed_results),
                "exception_count": len(exceptions),
                "skipped_count": len(skipped_pages),  # 添加跳过页面统计
                "processing_time": processing_time,
                "results": successful_results + failed_results,
                "successful_results": successful_results,
                "failed_results": failed_results,
                "skipped_results": skipped_pages,  # 添加跳过的页面
                "exceptions": exceptions,
                "api_key_stats": key_stats  # 添加API密钥统计
            }
            
            log_user_action(
                "Dify API处理完成", 
                f"成功: {len(successful_results)}, 失败: {len(failed_results)}, 跳过: {len(skipped_pages)}, 异常: {len(exceptions)}, 耗时: {processing_time:.2f}秒"
            )
            
            return return_result
            
        except Exception as e:
            logger.error(f"并发处理异常: {str(e)}")
            return {
                "success": len(skipped_pages) > 0,  # 如果有跳过的页面，至少部分成功
                "error": f"并发处理异常: {str(e)}",
                "total_pages": len(pages),
                "successful_count": 0,
                "failed_count": 0,
                "exception_count": 1,
                "skipped_count": len(skipped_pages),
                "processing_time": time.time() - start_time,
                "results": [],
                "skipped_results": skipped_pages,
                "exceptions": [str(e)]
            }

class DifyIntegrationService:
    """Dify集成服务类"""
    
    def __init__(self, config: Optional[DifyAPIConfig] = None):
        """初始化服务"""
        self.config = config or DifyAPIConfig()
        logger.info("初始化Dify集成服务")
    
    async def process_pagination_result(self, pagination_result: Dict[str, Any]) -> Dict[str, Any]:
        """
        处理分页结果，对每页调用Dify API
        
        Args:
            pagination_result: AI分页的结果
            
        Returns:
            Dict: 包含Dify API处理结果的完整数据
        """
        if not pagination_result.get('success', False):
            return {
                "success": False,
                "error": "输入的分页结果无效",
                "original_pagination": pagination_result
            }
        
        pages = pagination_result.get('pages', [])
        if not pages:
            return {
                "success": False,
                "error": "没有页面数据需要处理",
                "original_pagination": pagination_result
            }
        
        log_user_action("Dify集成处理", f"开始处理{len(pages)}个页面的API调用")
        
        try:
            # 使用异步客户端处理页面
            async with DifyAPIClient(self.config) as client:
                api_results = await client.process_pages_concurrent(pages)
            
            # 合并原始分页结果和API处理结果
            combined_result = {
                "success": api_results.get('success', False),
                "original_pagination": pagination_result,
                "dify_api_results": api_results,
                "processing_summary": {
                    "total_pages": api_results.get('total_pages', 0),
                    "successful_api_calls": api_results.get('successful_count', 0),
                    "failed_api_calls": api_results.get('failed_count', 0),
                    "skipped_fixed_pages": api_results.get('skipped_count', 0),
                    "processing_time": api_results.get('processing_time', 0),
                    "success_rate": api_results.get('successful_count', 0) / max(api_results.get('successful_count', 0) + api_results.get('failed_count', 0), 1) * 100  # 跳过的页面不计入成功率
                }
            }
            
            # 为每个页面添加API结果
            enhanced_pages = []
            api_results_by_page = {
                result.get('page_index', -1): result 
                for result in api_results.get('results', [])
            }
            skipped_results_by_page = {
                i: page for i, page in enumerate(pages) 
                if page in api_results.get('skipped_results', [])
            }
            
            for i, page in enumerate(pages):
                enhanced_page = page.copy()
                page_type = page.get('page_type', 'content')
                
                # 检查是否是跳过的固定页面
                if i in skipped_results_by_page or page_type in ['title', 'table_of_contents', 'ending']:
                    enhanced_page['dify_skipped'] = True
                    enhanced_page['dify_skip_reason'] = f'固定页面类型({page_type})，无需Dify API处理'
                    logger.debug(f"第{page.get('page_number', i+1)}页({page_type})已跳过Dify API处理")
                else:
                    # 查找API处理结果
                    api_result = api_results_by_page.get(i)
                    
                    if api_result:
                        enhanced_page['dify_api_result'] = api_result
                        if api_result.get('success'):
                            enhanced_page['dify_response'] = api_result.get('response_text', '')
                            enhanced_page['dify_full_response'] = api_result.get('api_response', {})
                        else:
                            enhanced_page['dify_error'] = api_result.get('error', 'API调用失败')
                    else:
                        enhanced_page['dify_error'] = '未找到对应的API结果'
                
                enhanced_pages.append(enhanced_page)
            
            combined_result['enhanced_pages'] = enhanced_pages
            
            return combined_result
            
        except Exception as e:
            logger.error(f"Dify集成处理异常: {str(e)}")
            return {
                "success": False,
                "error": f"Dify集成处理异常: {str(e)}",
                "original_pagination": pagination_result
            }
    
    def format_results_summary(self, result: Dict[str, Any]) -> str:
        """
        格式化结果摘要
        
        Args:
            result: 处理结果
            
        Returns:
            str: 格式化的摘要文本
        """
        if not result.get('success', False):
            return f"❌ 处理失败: {result.get('error', '未知错误')}"
        
        summary = result.get('processing_summary', {})
        total_pages = summary.get('total_pages', 0)
        successful = summary.get('successful_api_calls', 0)
        failed = summary.get('failed_api_calls', 0)
        skipped = summary.get('skipped_fixed_pages', 0)
        processing_time = summary.get('processing_time', 0)
        success_rate = summary.get('success_rate', 0)
        
        # 添加多API密钥统计信息
        api_key_stats = result.get('api_key_stats', {})
        key_info = ""
        if api_key_stats:
            total_keys = api_key_stats.get('total_keys', 0)
            available_keys = api_key_stats.get('available_keys', 0)
            strategy = api_key_stats.get('strategy', 'unknown')
            
            key_info = f"""
🔑 API密钥统计:
• 总密钥数: {total_keys}
• 可用密钥: {available_keys}
• 负载策略: {strategy}"""
        
        summary_text = f"""✅ Dify API处理完成 (多密钥并行)

📊 处理统计:
• 总页面数: {total_pages}
• 成功调用: {successful}
• 失败调用: {failed}
• 跳过固定页面: {skipped} (封面/目录/结束页)
• 成功率: {success_rate:.1f}%
• 处理耗时: {processing_time:.2f}秒{key_info}

🚀 平均响应时间: {processing_time / max(successful + failed, 1):.2f}秒/页（不含固定页面）"""
        
        return summary_text

# 同步接口函数
def process_pages_with_dify(pagination_result: Dict[str, Any], config: Optional[DifyAPIConfig] = None) -> Dict[str, Any]:
    """
    同步接口：处理分页结果并调用Dify API
    
    Args:
        pagination_result: AI分页结果
        config: Dify API配置
        
    Returns:
        Dict: 处理结果
    """
    service = DifyIntegrationService(config)
    
    # 运行异步处理
    try:
        loop = asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
    
    try:
        return loop.run_until_complete(service.process_pagination_result(pagination_result))
    finally:
        # 清理事件循环（如果是新创建的）
        if loop != asyncio.get_event_loop():
            loop.close()