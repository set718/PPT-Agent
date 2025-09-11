#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Dify API与模板文件桥接模块
测试Dify API返回数字与模板文件的对应关系
"""

import os
import json
import asyncio
import aiohttp
import threading
import time
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path
from logger import get_logger, log_user_action
from dify_api_client import DifyAPIConfig, SmartAPIKeyPoller, APIKeyHealth
from utils import FileManager
from liai_auth import liai_m2m_auth, is_liai_m2m_configured

logger = get_logger()

class LiaiAPIKeyPoller:
    """Liai API智能密钥轮询器"""
    
    def __init__(self, api_keys: List[str]):
        self.api_keys = api_keys.copy()
        self.key_health: Dict[str, APIKeyHealth] = {}
        self.current_index = 0
        self._lock = threading.Lock()
        self.last_health_check = 0
        
        # 配置参数（与Dify保持一致）
        self.failure_threshold = 3
        self.recovery_time = 600
        self.response_time_weight = 0.3
        self.success_rate_weight = 0.7
        self.polling_strategy = "health_based"
        
        # 初始化密钥健康状态
        for api_key in self.api_keys:
            self.key_health[api_key] = APIKeyHealth(api_key)
        
        logger.info(f"初始化Liai API智能密钥轮询器，共{len(self.api_keys)}个密钥")
    
    def get_next_key(self) -> Optional[Tuple[str, int]]:
        """获取下一个API密钥"""
        if not self.api_keys:
            return None
            
        with self._lock:
            if self.polling_strategy == "round_robin":
                return self._round_robin_selection()
            elif self.polling_strategy == "health_based":
                return self._health_based_selection()
            else:
                return self._round_robin_selection()
    
    def _round_robin_selection(self) -> Tuple[str, int]:
        """轮询选择"""
        selected_key = self.api_keys[self.current_index]
        selected_index = self.current_index
        self.current_index = (self.current_index + 1) % len(self.api_keys)
        return selected_key, selected_index
    
    def _health_based_selection(self) -> Tuple[str, int]:
        """基于健康状态的选择"""
        healthy_keys = []
        
        for i, api_key in enumerate(self.api_keys):
            health = self.key_health[api_key]
            if health.is_considered_healthy(self.failure_threshold, self.recovery_time):
                healthy_keys.append((api_key, i))
        
        if not healthy_keys:
            # 如果没有健康的密钥，选择恢复时间最长的
            logger.warning("没有健康的Liai API密钥，选择恢复时间最长的密钥")
            oldest_key = min(
                self.api_keys,
                key=lambda k: self.key_health[k].last_failure_time
            )
            return oldest_key, self.api_keys.index(oldest_key)
        
        # 从健康密钥中轮询选择
        selected_key, selected_index = healthy_keys[self.current_index % len(healthy_keys)]
        self.current_index += 1
        return selected_key, selected_index
    
    def record_request_result(self, api_key: str, success: bool, response_time: float, error_type: str = None):
        """记录请求结果"""
        if api_key in self.key_health:
            self.key_health[api_key].record_request(success, response_time, error_type)
            
            # 记录日志
            health = self.key_health[api_key]
            if success:
                logger.debug(f"Liai API密钥请求成功: {api_key[:20]}... (响应时间: {response_time:.2f}s, 成功率: {health.get_success_rate():.2%})")
            else:
                logger.warning(f"Liai API密钥请求失败: {api_key[:20]}... (连续失败: {health.consecutive_failures}, 错误类型: {error_type})")
    
    def get_health_report(self) -> Dict[str, Dict]:
        """获取健康状态报告"""
        report = {}
        for api_key, health in self.key_health.items():
            masked_key = api_key[:20] + "..." if len(api_key) > 20 else api_key
            report[masked_key] = {
                "total_requests": health.total_requests,
                "successful_requests": health.successful_requests,
                "failed_requests": health.failed_requests,
                "success_rate": health.get_success_rate(),
                "avg_response_time": health.avg_response_time,
                "consecutive_failures": health.consecutive_failures,
                "health_score": health.get_health_score(self.response_time_weight, self.success_rate_weight),
                "is_healthy": health.is_considered_healthy(self.failure_threshold, self.recovery_time),
                "failure_reasons": dict(health.failure_reasons)
            }
        return report
    
    def perform_health_check(self):
        """执行健康检查"""
        current_time = time.time()
        if current_time - self.last_health_check > 300:  # 5分钟检查一次
            healthy_count = sum(1 for health in self.key_health.values() 
                              if health.is_considered_healthy(self.failure_threshold, self.recovery_time))
            logger.info(f"Liai API密钥健康检查: {healthy_count}/{len(self.api_keys)}个密钥健康")
            self.last_health_check = current_time

class DifyTemplateBridge:
    """Dify API与模板文件桥接器 - 单例模式"""
    
    _instance = None
    _lock = threading.Lock()
    _initialized = False
    
    def __new__(cls, *args, **kwargs):
        """单例模式实现"""
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self, config: Optional[DifyAPIConfig] = None, model_config=None):
        """初始化桥接器（仅初始化一次）"""
        if self._initialized:
            return
            
        self.config = config or DifyAPIConfig()
        self.model_config = model_config  # 添加模型配置支持
        self.templates_dir = os.path.join(os.path.dirname(__file__), "templates", "ppt_template")
        
        # 智能API密钥轮询器
        self.api_key_poller = SmartAPIKeyPoller(self.config) if self.config.api_keys else None
        
        # Liai API轮询器
        self.liai_api_poller = None
        if self.model_config and self.model_config.get('request_format') == 'dify_compatible':
            liai_api_keys = []
            for i in range(1, 6):
                key_name = f"LIAI_TEMPLATE_API_KEY_{i}"
                key = os.getenv(key_name)
                if key:
                    liai_api_keys.append(key)
            
            if liai_api_keys:
                self.liai_api_poller = LiaiAPIKeyPoller(liai_api_keys)
        
        # 添加缓存
        self._templates_cache = None
        self._cache_timestamp = 0
        self._cache_lock = threading.Lock()
        
        DifyTemplateBridge._initialized = True
        logger.info(f"初始化Dify模板桥接器（单例），模板目录: {self.templates_dir}")
    
    def scan_available_templates(self) -> Dict[str, Any]:
        """扫描可用的模板文件（带缓存）"""
        with self._cache_lock:
            current_time = time.time()
            # 缓存5分钟
            if self._templates_cache and (current_time - self._cache_timestamp) < 300:
                logger.debug("使用缓存的模板扫描结果")
                return self._templates_cache.copy()
        
        templates_info = {
            "template_directory": self.templates_dir,
            "templates": [],
            "total_count": 0,
            "number_range": {"min": None, "max": None}
        }
        
        if not os.path.exists(self.templates_dir):
            logger.error(f"模板目录不存在: {self.templates_dir}")
            return templates_info
        
        # 扫描所有split_presentations_*.pptx文件
        template_files = []
        template_numbers = []
        
        for filename in os.listdir(self.templates_dir):
            if filename.startswith("split_presentations_") and filename.endswith(".pptx"):
                try:
                    # 提取文件编号
                    number_str = filename.replace("split_presentations_", "").replace(".pptx", "")
                    template_number = int(number_str)
                    
                    file_path = os.path.join(self.templates_dir, filename)
                    file_size = os.path.getsize(file_path)
                    
                    template_info = {
                        "filename": filename,
                        "number": template_number,
                        "file_path": file_path,
                        "file_size": file_size,
                        "file_size_kb": round(file_size / 1024, 1)
                    }
                    
                    template_files.append(template_info)
                    template_numbers.append(template_number)
                    
                except ValueError:
                    # 跳过无法解析编号的文件
                    continue
        
        # 按编号排序
        template_files.sort(key=lambda x: x["number"])
        
        templates_info["templates"] = template_files
        templates_info["total_count"] = len(template_files)
        
        if template_numbers:
            templates_info["number_range"]["min"] = min(template_numbers)
            templates_info["number_range"]["max"] = max(template_numbers)
        
        logger.info(f"扫描到{len(template_files)}个模板文件，编号范围: {templates_info['number_range']}")
        
        # 更新缓存
        with self._cache_lock:
            self._templates_cache = templates_info.copy()
            self._cache_timestamp = time.time()
            logger.debug("更新模板扫描缓存")
        
        return templates_info
    
    def get_template_by_number(self, template_number: int) -> Dict[str, Any]:
        """根据编号获取模板文件信息"""
        result = {
            "success": False,
            "template_number": template_number,
            "filename": None,
            "file_path": None,
            "file_exists": False,
            "file_valid": False,
            "error": None
        }
        
        # 构建文件名和路径
        filename = f"split_presentations_{template_number}.pptx"
        file_path = os.path.join(self.templates_dir, filename)
        
        result["filename"] = filename
        result["file_path"] = file_path
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            result["error"] = f"模板文件不存在: {filename}"
            logger.warning(f"请求的模板文件不存在: {file_path}")
            return result
        
        result["file_exists"] = True
        
        # 验证PPT文件格式
        is_valid, error_msg = FileManager.validate_ppt_file(file_path)
        if not is_valid:
            result["error"] = f"模板文件格式无效: {error_msg}"
            logger.error(f"模板文件格式验证失败: {file_path}, 错误: {error_msg}")
            return result
        
        result["file_valid"] = True
        result["success"] = True
        
        # 添加文件详细信息
        file_size = os.path.getsize(file_path)
        result["file_size"] = file_size
        result["file_size_kb"] = round(file_size / 1024, 1)
        
        logger.info(f"成功获取模板文件: {filename} ({result['file_size_kb']}KB)")
        return result
    
    async def call_api_for_template_number(self, user_input: str) -> Dict[str, Any]:
        """
        调用API（Dify或Liai）获取模板编号
        
        Args:
            user_input: 用户输入的文本
            
        Returns:
            Dict: 包含API响应和解析的模板编号
        """
        result = {
            "success": False,
            "user_input": user_input,
            "api_response": None,
            "template_number": None,
            "error": None,
            "used_api_key": None,
            "attempt_count": 0,
            "api_type": "unknown"
        }
        
        # 根据当前用户选择的模型决定使用哪个API
        from config import get_config
        config = get_config()
        current_model_info = config.get_model_info()
        
        # 如果用户选择的是liai模型，使用Liai API
        if current_model_info.get('request_format') == 'dify_compatible':
            result["api_type"] = "liai"
            return await self._call_liai_api(user_input, result)
        else:
            # 如果用户选择的是其他模型（如deepseek），使用Dify API
            result["api_type"] = "dify"
            return await self._call_dify_api(user_input, result)
    
    async def _call_dify_api(self, user_input: str, result: Dict[str, Any]) -> Dict[str, Any]:
        """调用Dify API"""
        # 构建请求数据
        request_data = {
            "inputs": {},
            "query": user_input,
            "response_mode": "streaming",
            "conversation_id": "",
            "user": f"template_selector_{int(asyncio.get_event_loop().time())}"
        }
        
        url = f"{self.config.base_url}{self.config.endpoint}"
        
        # 创建HTTP会话
        connector = aiohttp.TCPConnector(
            limit=5,
            ttl_dns_cache=300,
            use_dns_cache=True,
            keepalive_timeout=60
        )
        
        timeout = aiohttp.ClientTimeout(
            total=self.config.timeout,
            connect=30,
            sock_read=120
        )
        
        async with aiohttp.ClientSession(
            connector=connector,
            timeout=timeout,
            headers={'Content-Type': 'application/json'}
        ) as session:
            
            # 重试逻辑
            for attempt in range(self.config.max_retries):
                result["attempt_count"] = attempt + 1
                
                # 获取API密钥（使用智能轮询器）
                if not self.api_key_poller:
                    result["error"] = "无可用的API密钥"
                    return result
                
                key_info = self.api_key_poller.get_next_key()
                if not key_info:
                    result["error"] = "所有API密钥都不可用"
                    return result
                
                current_api_key, key_index = key_info
                result["used_api_key"] = current_api_key[:20] + "..."
                
                # 执行健康检查
                self.api_key_poller.perform_health_check()
                
                headers = {
                    'Authorization': f'Bearer {current_api_key}',
                    'Content-Type': 'application/json'
                }
                
                try:
                    logger.info(f"调用Dify API获取模板编号 (尝试 {attempt + 1}/{self.config.max_retries})，使用密钥索引: {key_index}")
                    
                    request_start_time = time.time()
                    async with session.post(url, json=request_data, headers=headers) as response:
                        if response.status == 200:
                            # 处理streaming响应
                            response_text = ""
                            async for line in response.content:
                                line_text = line.decode('utf-8').strip()
                                if line_text.startswith('data: '):
                                    data_text = line_text[6:]  # 去掉'data: '前缀
                                    if data_text == '[DONE]':
                                        break
                                    try:
                                        import json
                                        data_json = json.loads(data_text)
                                        if 'answer' in data_json:
                                            response_text += data_json['answer']
                                        elif 'event' in data_json and data_json['event'] == 'agent_message':
                                            if 'answer' in data_json:
                                                response_text += data_json['answer']
                                    except json.JSONDecodeError:
                                        continue
                            
                            result["api_response"] = {"answer": response_text}
                            
                            # 尝试从响应中提取数字
                            template_number = self._extract_template_number(response_text)
                            
                            request_end_time = time.time()
                            response_time = request_end_time - request_start_time
                            
                            if template_number is not None:
                                result["success"] = True
                                result["template_number"] = template_number
                                result["response_text"] = response_text
                                
                                # 记录成功的请求
                                if self.api_key_poller:
                                    self.api_key_poller.record_request_result(
                                        current_api_key, True, response_time
                                    )
                                
                                logger.info(f"成功获取模板编号: {template_number} (响应时间: {response_time:.2f}s)")
                                return result
                            else:
                                # 记录失败的请求（解析失败）
                                if self.api_key_poller:
                                    self.api_key_poller.record_request_result(
                                        current_api_key, False, response_time, "parse_error"
                                    )
                                
                                result["error"] = f"无法从API响应中提取有效的模板编号: {response_text}"
                                logger.warning(f"API响应中未找到有效数字: {response_text}")
                        else:
                            request_end_time = time.time()
                            response_time = request_end_time - request_start_time
                            error_text = await response.text()
                            result["error"] = f"HTTP {response.status}: {error_text}"
                            logger.warning(f"API请求失败，状态码: {response.status}")
                            
                            # 记录失败的请求
                            if self.api_key_poller:
                                error_type = "auth_error" if response.status in [401, 403] else "http_error"
                                self.api_key_poller.record_request_result(
                                    current_api_key, False, response_time, error_type
                                )
                            
                            # 认证错误时记录日志
                            if response.status in [401, 403]:
                                logger.warning(f"API密钥认证失败: {current_api_key[:20]}...")
                
                except asyncio.TimeoutError:
                    request_end_time = time.time()
                    response_time = request_end_time - request_start_time
                    result["error"] = "API请求超时"
                    logger.warning(f"API请求超时 (尝试 {attempt + 1})")
                    
                    # 记录超时失败
                    if self.api_key_poller:
                        self.api_key_poller.record_request_result(
                            current_api_key, False, response_time, "timeout"
                        )
                    
                    # 超时处理
                    if attempt >= 2:
                        logger.warning(f"API密钥多次超时: {current_api_key[:20]}...")
                
                except Exception as e:
                    request_end_time = time.time()
                    response_time = request_end_time - request_start_time
                    result["error"] = f"API请求异常: {str(e)}"
                    logger.error(f"API请求异常: {str(e)}")
                    
                    # 记录异常失败
                    if self.api_key_poller:
                        self.api_key_poller.record_request_result(
                            current_api_key, False, response_time, "exception"
                        )
                    
                    # 异常处理
                    logger.warning(f"API密钥请求异常: {current_api_key[:20]}...")
                
                # 如果不是最后一次尝试，等待后重试
                if attempt < self.config.max_retries - 1:
                    delay = self.config.retry_delay * (2 ** attempt)
                    await asyncio.sleep(min(delay, 30))
        
        return result
    
    async def _call_liai_api(self, user_input: str, result: Dict[str, Any]) -> Dict[str, Any]:
        """
        调用Liai API进行模板推荐
        基于simple_liai_test.py的成功实现
        """
        # 构建Liai API请求数据，包含IDaaS用户信息
        idaas_user = os.getenv('IDAAS_USER', 'template-selector-user')
        idaas_open_id = os.getenv('IDAAS_OPEN_ID', '')
        
        request_data = {
            "inputs": {},
            "query": user_input,
            "response_mode": "streaming",
            "conversation_id": "",
            "user": idaas_user,
            "open_id": idaas_open_id
        }
        
        # 构建URL
        url = "https://liai-app.chj.cloud/v1/chat-messages"
        
        # 直接从环境变量获取API密钥（使用simple_liai_test的成功方式）
        api_key = os.getenv('LIAI_TEMPLATE_API_KEY_1')
        if not api_key:
            result["error"] = "未找到LIAI_TEMPLATE_API_KEY_1环境变量"
            return result
        
        # 获取M2M token
        m2m_token = None
        try:
            if is_liai_m2m_configured():
                m2m_token = await liai_m2m_auth.get_access_token()
                if m2m_token:
                    logger.info("M2M token获取成功")
                else:
                    logger.warning("M2M token获取失败")
        except Exception as e:
            logger.warning(f"M2M认证异常: {str(e)}")
        
        # 构建请求头（按照simple_liai_test的成功格式）
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        
        # 如果有M2M token，添加到请求头
        if m2m_token:
            headers['X-IDaaS-M2M-Token'] = m2m_token
            result["auth_method"] = "API Key + M2M Token"
        else:
            result["auth_method"] = "API Key Only"
        
        result["used_api_key"] = api_key[:20] + "..."
        if m2m_token:
            result["m2m_token_used"] = m2m_token[:20] + "..."
        
        # 创建HTTP会话（简化配置，使用与simple_liai_test相同的方式）
        timeout = aiohttp.ClientTimeout(total=30)
        
        async with aiohttp.ClientSession(timeout=timeout) as session:
            
            # 重试逻辑（最多重试3次）
            max_retries = 3
            for attempt in range(max_retries):
                result["attempt_count"] = attempt + 1
                
                try:
                    logger.info(f"调用Liai API获取模板编号 (尝试 {attempt + 1}/{max_retries})，认证方式: {result.get('auth_method', 'Unknown')}")
                    
                    request_start_time = time.time()
                    async with session.post(url, json=request_data, headers=headers) as response:
                        if response.status == 200:
                            # 处理streaming响应
                            response_text = ""
                            async for line in response.content:
                                line_text = line.decode('utf-8').strip()
                                if line_text.startswith('data: '):
                                    data_text = line_text[6:]  # 去掉'data: '前缀
                                    if data_text == '[DONE]':
                                        break
                                    try:
                                        import json
                                        data_json = json.loads(data_text)
                                        if 'answer' in data_json:
                                            response_text += data_json['answer']
                                        elif 'event' in data_json and data_json['event'] == 'agent_message':
                                            if 'answer' in data_json:
                                                response_text += data_json['answer']
                                    except json.JSONDecodeError:
                                        continue
                            
                            result["api_response"] = {"answer": response_text}
                            
                            # 尝试从响应中提取数字
                            template_number = self._extract_template_number(response_text)
                            
                            request_end_time = time.time()
                            response_time = request_end_time - request_start_time
                            
                            if template_number is not None:
                                result["success"] = True
                                result["template_number"] = template_number
                                result["response_text"] = response_text
                                
                                logger.info(f"成功获取模板编号: {template_number} (响应时间: {response_time:.2f}s, 认证: {result.get('auth_method', 'Unknown')})")
                                return result
                            else:
                                result["error"] = f"无法从API响应中提取有效的模板编号: {response_text}"
                                logger.warning(f"API响应中未找到有效数字: {response_text}")
                        else:
                            request_end_time = time.time()
                            response_time = request_end_time - request_start_time
                            error_text = await response.text()
                            result["error"] = f"HTTP {response.status}: {error_text}"
                            logger.warning(f"Liai API请求失败，状态码: {response.status}")
                            
                            # 认证错误时记录日志
                            if response.status in [401, 403]:
                                auth_info = result.get('auth_method', 'Unknown')
                                logger.warning(f"Liai API认证失败 (认证方式: {auth_info})")
                
                except asyncio.TimeoutError:
                    result["error"] = "Liai API请求超时"
                    logger.warning(f"Liai API请求超时 (尝试 {attempt + 1})")
                
                except Exception as e:
                    result["error"] = f"Liai API请求异常: {str(e)}"
                    logger.error(f"Liai API请求异常: {str(e)}")
                
                # 如果不是最后一次尝试，等待后重试
                if attempt < max_retries - 1:
                    delay = 2.0 * (2 ** attempt)  # 指数退避
                    await asyncio.sleep(min(delay, 10))
        
        return result
    
    def _extract_template_number(self, text: str) -> Optional[int]:
        """从文本中提取模板编号"""
        import re
        
        # 尝试多种模式提取数字
        patterns = [
            r'模板编号[：:]\s*(\d+)',
            r'编号[：:]\s*(\d+)',
            r'选择\s*(\d+)',
            r'返回\s*(\d+)',
            r'(\d+)',  # 任何数字
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                try:
                    number = int(match)
                    # 验证数字范围
                    if 1 <= number <= 250:
                        return number
                except ValueError:
                    continue
        
        return None
    
    async def test_dify_template_bridge(self, user_input: str) -> Dict[str, Any]:
        """
        测试完整的Dify API到模板文件的桥接流程
        
        Args:
            user_input: 用户输入文本
            
        Returns:
            Dict: 完整的测试结果
        """
        log_user_action("Dify模板桥接测试", f"输入文本长度: {len(user_input)}字符")
        
        test_result = {
            "success": False,
            "user_input": user_input,
            "step_1_dify_api": None,
            "step_2_template_lookup": None,
            "final_template_path": None,
            "error": None,
            "processing_time": 0
        }
        
        import time
        start_time = time.time()
        
        try:
            # 步骤1: 调用API获取模板编号
            api_type = self.model_config.get('request_format') if self.model_config else 'dify'
            logger.info(f"步骤1: 调用{api_type}API获取模板编号")
            dify_result = await self.call_api_for_template_number(user_input)
            test_result["step_1_dify_api"] = dify_result
            
            if not dify_result["success"]:
                test_result["error"] = f"Dify API调用失败: {dify_result.get('error', '未知错误')}"
                return test_result
            
            template_number = dify_result["template_number"]
            logger.info(f"Dify API返回模板编号: {template_number}")
            
            # 步骤2: 根据编号查找模板文件
            logger.info(f"步骤2: 查找模板文件 split_presentations_{template_number}.pptx")
            template_result = self.get_template_by_number(template_number)
            test_result["step_2_template_lookup"] = template_result
            
            if not template_result["success"]:
                test_result["error"] = f"模板文件查找失败: {template_result.get('error', '未知错误')}"
                return test_result
            
            # 成功完成桥接
            test_result["success"] = True
            test_result["final_template_path"] = template_result["file_path"]
            
            end_time = time.time()
            test_result["processing_time"] = end_time - start_time
            
            logger.info(f"Dify模板桥接测试成功完成，耗时: {test_result['processing_time']:.2f}秒")
            
            # 输出API密钥健康报告
            if test_result.get("step_1_dify_api", {}).get("api_type") == "liai" and self.liai_api_poller:
                # Liai API健康报告
                health_report = self.liai_api_poller.get_health_report()
                test_result["api_key_health_report"] = health_report
                logger.info("Liai API密钥健康状态报告:")
                for key, stats in health_report.items():
                    if stats["total_requests"] > 0:
                        logger.info(
                            f"  {key}: {stats['total_requests']}次请求, "
                            f"成功率{stats['success_rate']:.1%}, "
                            f"平均响应时间{stats['avg_response_time']:.2f}s, "
                            f"健康分数{stats['health_score']:.2f}"
                        )
            elif self.api_key_poller:
                # Dify API健康报告
                health_report = self.api_key_poller.get_health_report()
                test_result["api_key_health_report"] = health_report
                logger.info("Dify API密钥健康状态报告:")
                for key, stats in health_report.items():
                    if stats["total_requests"] > 0:
                        logger.info(
                            f"  {key}: {stats['total_requests']}次请求, "
                            f"成功率{stats['success_rate']:.1%}, "
                            f"平均响应时间{stats['avg_response_time']:.2f}s, "
                            f"健康分数{stats['health_score']:.2f}"
                        )
            
        except Exception as e:
            test_result["error"] = f"桥接测试异常: {str(e)}"
            logger.error(f"Dify模板桥接测试异常: {str(e)}")
        
        finally:
            if test_result["processing_time"] == 0:
                test_result["processing_time"] = time.time() - start_time
        
        return test_result

def sync_test_dify_template_bridge(user_input: str, config: Optional[DifyAPIConfig] = None, model_config: Optional[Dict] = None) -> Dict[str, Any]:
    """
    同步接口：测试API到模板文件的桥接（支持Dify和Liai）
    
    Args:
        user_input: 用户输入文本
        config: Dify API配置
        model_config: 模型配置（包含API类型信息）
        
    Returns:
        Dict: 测试结果
    """
    bridge = DifyTemplateBridge(config, model_config)
    
    # 运行异步测试
    try:
        loop = asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
    
    try:
        return loop.run_until_complete(bridge.test_dify_template_bridge(user_input))
    finally:
        # 清理事件循环
        if loop != asyncio.get_event_loop():
            loop.close()

if __name__ == "__main__":
    # 简单的命令行测试
    import sys
    
    if len(sys.argv) > 1:
        test_input = " ".join(sys.argv[1:])
    else:
        test_input = "人工智能技术发展趋势分析报告"
    
    print("=" * 50)
    print("Dify API与模板文件桥接测试")
    print("=" * 50)
    
    # 先扫描可用模板
    bridge = DifyTemplateBridge()
    templates_info = bridge.scan_available_templates()
    
    print(f"可用模板数量: {templates_info['total_count']}")
    print(f"编号范围: {templates_info['number_range']['min']} - {templates_info['number_range']['max']}")
    
    # 测试桥接
    result = sync_test_dify_template_bridge(test_input)
    
    print(f"\n测试输入: {test_input}")
    print(f"测试结果: {'成功' if result['success'] else '失败'}")
    
    if result["success"]:
        dify_result = result["step_1_dify_api"]
        template_result = result["step_2_template_lookup"]
        
        print(f"Dify API返回编号: {dify_result['template_number']}")
        print(f"对应模板文件: {template_result['filename']}")
        print(f"文件大小: {template_result['file_size_kb']}KB")
        print(f"处理耗时: {result['processing_time']:.2f}秒")
    else:
        print(f"错误信息: {result['error']}")