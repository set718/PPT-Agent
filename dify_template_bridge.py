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
from dify_api_client import DifyAPIConfig, APIKeyBalancer
from utils import FileManager

logger = get_logger()

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
    
    def __init__(self, config: Optional[DifyAPIConfig] = None):
        """初始化桥接器（仅初始化一次）"""
        if self._initialized:
            return
            
        self.config = config or DifyAPIConfig()
        self.templates_dir = os.path.join(os.path.dirname(__file__), "templates", "ppt_template")
        
        # 使用单例负载均衡器
        self.key_balancer = APIKeyBalancer(
            self.config.api_keys, 
            self.config.load_balance_strategy
        )
        
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
    
    async def call_dify_for_template_number(self, user_input: str) -> Dict[str, Any]:
        """
        调用Dify API获取模板编号
        
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
            "attempt_count": 0
        }
        
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
                
                # 获取API密钥
                current_api_key = self.key_balancer.get_next_key()
                result["used_api_key"] = current_api_key[:20] + "..."
                
                headers = {
                    'Authorization': f'Bearer {current_api_key}',
                    'Content-Type': 'application/json'
                }
                
                try:
                    logger.info(f"调用Dify API获取模板编号 (尝试 {attempt + 1}/{self.config.max_retries})")
                    
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
                            
                            if template_number is not None:
                                result["success"] = True
                                result["template_number"] = template_number
                                result["response_text"] = response_text
                                
                                # 标记API密钥成功
                                self.key_balancer.mark_key_success(current_api_key)
                                
                                logger.info(f"成功获取模板编号: {template_number}")
                                return result
                            else:
                                result["error"] = f"无法从API响应中提取有效的模板编号: {response_text}"
                                logger.warning(f"API响应中未找到有效数字: {response_text}")
                        else:
                            error_text = await response.text()
                            result["error"] = f"HTTP {response.status}: {error_text}"
                            logger.warning(f"API请求失败，状态码: {response.status}")
                            
                            # 认证错误时标记密钥失败
                            if response.status in [401, 403]:
                                self.key_balancer.mark_key_failed(current_api_key)
                
                except asyncio.TimeoutError:
                    result["error"] = "API请求超时"
                    logger.warning(f"API请求超时 (尝试 {attempt + 1})")
                    
                    # 超时多次后标记密钥失败
                    if attempt >= 2:
                        self.key_balancer.mark_key_failed(current_api_key)
                
                except Exception as e:
                    result["error"] = f"API请求异常: {str(e)}"
                    logger.error(f"API请求异常: {str(e)}")
                    
                    # 异常时标记密钥失败
                    self.key_balancer.mark_key_failed(current_api_key)
                
                # 如果不是最后一次尝试，等待后重试
                if attempt < self.config.max_retries - 1:
                    delay = self.config.retry_delay * (2 ** attempt)
                    await asyncio.sleep(min(delay, 30))
        
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
            # 步骤1: 调用Dify API获取模板编号
            logger.info("步骤1: 调用Dify API获取模板编号")
            dify_result = await self.call_dify_for_template_number(user_input)
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
            
        except Exception as e:
            test_result["error"] = f"桥接测试异常: {str(e)}"
            logger.error(f"Dify模板桥接测试异常: {str(e)}")
        
        finally:
            if test_result["processing_time"] == 0:
                test_result["processing_time"] = time.time() - start_time
        
        return test_result

def sync_test_dify_template_bridge(user_input: str, config: Optional[DifyAPIConfig] = None) -> Dict[str, Any]:
    """
    同步接口：测试Dify API到模板文件的桥接
    
    Args:
        user_input: 用户输入文本
        config: Dify API配置
        
    Returns:
        Dict: 测试结果
    """
    bridge = DifyTemplateBridge(config)
    
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