#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Dify API基础配置模块（简化版）
仅保留基础配置，移除内容处理功能，专注于模板推荐
"""

import os
import time
import threading
import asyncio
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Any
from collections import defaultdict
from logger import get_logger

logger = get_logger()

@dataclass
class DifyAPIConfig:
    """Dify API配置类 - 专用于模板推荐，支持智能轮询策略"""
    base_url: str = "https://api.dify.ai/v1"
    api_keys: List[str] = field(default_factory=lambda: [
        key for key in [
            os.getenv("DIFY_API_KEY_1"),
            os.getenv("DIFY_API_KEY_2"),
            os.getenv("DIFY_API_KEY_3"),
            os.getenv("DIFY_API_KEY_4"),
            os.getenv("DIFY_API_KEY_5")
        ] if key
    ])
    endpoint: str = "/chat-messages"
    timeout: int = 30
    max_retries: int = 5
    retry_delay: float = 2.0
    max_concurrent: int = 5
    load_balance_strategy: str = "health_based_polling"
    
    # 轮询策略配置
    polling_strategy: str = "round_robin"  # round_robin, health_based, weighted
    health_check_interval: int = 300  # 健康检查间隔（秒）
    key_failure_threshold: int = 3  # 密钥失败阈值
    key_recovery_time: int = 600  # 密钥恢复时间（秒）
    response_time_weight: float = 0.3  # 响应时间权重
    success_rate_weight: float = 0.7  # 成功率权重
    
    # 分批处理配置
    batch_size: int = 5  # 每批处理的请求数量
    batch_delay: float = 2.0  # 批次间延迟（秒）
    enable_batch_processing: bool = True  # 启用分批处理
    batch_timeout: int = 300  # 单批次超时时间（秒）
    
    @property
    def api_key(self) -> str:
        """向后兼容：返回第一个API密钥"""
        return self.api_keys[0] if self.api_keys else ""

    def __post_init__(self):
        """初始化后验证"""
        if not self.api_keys:
            logger.warning("未配置任何Dify API密钥，模板推荐功能将无法使用")
        else:
            logger.info(f"Dify API配置初始化完成，共{len(self.api_keys)}个密钥可用")

class APIKeyHealth:
    """API密钥健康状态跟踪"""
    
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.total_requests = 0
        self.successful_requests = 0
        self.failed_requests = 0
        self.last_request_time = 0
        self.last_success_time = 0
        self.last_failure_time = 0
        self.consecutive_failures = 0
        self.avg_response_time = 0
        self.is_healthy = True
        self.failure_reasons = defaultdict(int)
        self._lock = threading.Lock()
    
    def record_request(self, success: bool, response_time: float, error_type: str = None):
        """记录请求结果"""
        with self._lock:
            self.total_requests += 1
            self.last_request_time = time.time()
            
            if success:
                self.successful_requests += 1
                self.last_success_time = self.last_request_time
                self.consecutive_failures = 0
                
                # 更新平均响应时间（指数移动平均）
                if self.avg_response_time == 0:
                    self.avg_response_time = response_time
                else:
                    self.avg_response_time = 0.7 * self.avg_response_time + 0.3 * response_time
            else:
                self.failed_requests += 1
                self.last_failure_time = self.last_request_time
                self.consecutive_failures += 1
                
                if error_type:
                    self.failure_reasons[error_type] += 1
    
    def get_success_rate(self) -> float:
        """获取成功率"""
        if self.total_requests == 0:
            return 1.0
        return self.successful_requests / self.total_requests
    
    def get_health_score(self, response_time_weight: float = 0.3, success_rate_weight: float = 0.7) -> float:
        """计算健康分数 (0-1)"""
        success_rate = self.get_success_rate()
        
        # 响应时间分数（越快越好，1秒为基准）
        if self.avg_response_time == 0:
            response_time_score = 1.0
        else:
            response_time_score = max(0.1, min(1.0, 1.0 / self.avg_response_time))
        
        return success_rate_weight * success_rate + response_time_weight * response_time_score
    
    def is_considered_healthy(self, failure_threshold: int, recovery_time: int) -> bool:
        """判断密钥是否健康"""
        current_time = time.time()
        
        # 连续失败次数检查
        if self.consecutive_failures >= failure_threshold:
            # 检查是否超过恢复时间
            if current_time - self.last_failure_time < recovery_time:
                return False
        
        return True


class SmartAPIKeyPoller:
    """智能API密钥轮询器"""
    
    def __init__(self, config: DifyAPIConfig):
        self.config = config
        self.api_keys = config.api_keys.copy()
        self.key_health: Dict[str, APIKeyHealth] = {}
        self.current_index = 0
        self._lock = threading.Lock()
        self.last_health_check = 0
        
        # 初始化密钥健康状态
        for api_key in self.api_keys:
            self.key_health[api_key] = APIKeyHealth(api_key)
        
        logger.info(f"初始化智能API密钥轮询器，共{len(self.api_keys)}个密钥")
    
    def get_next_key(self) -> Optional[Tuple[str, int]]:
        """获取下一个API密钥"""
        if not self.api_keys:
            return None
            
        with self._lock:
            if self.config.polling_strategy == "round_robin":
                return self._round_robin_selection()
            elif self.config.polling_strategy == "health_based":
                return self._health_based_selection()
            elif self.config.polling_strategy == "weighted":
                return self._weighted_selection()
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
            if health.is_considered_healthy(
                self.config.key_failure_threshold,
                self.config.key_recovery_time
            ):
                healthy_keys.append((api_key, i))
        
        if not healthy_keys:
            # 如果没有健康的密钥，选择恢复时间最长的
            logger.warning("没有健康的API密钥，选择恢复时间最长的密钥")
            oldest_key = min(
                self.api_keys,
                key=lambda k: self.key_health[k].last_failure_time
            )
            return oldest_key, self.api_keys.index(oldest_key)
        
        # 从健康密钥中轮询选择
        selected_key, selected_index = healthy_keys[self.current_index % len(healthy_keys)]
        self.current_index += 1
        return selected_key, selected_index
    
    def _weighted_selection(self) -> Tuple[str, int]:
        """基于权重的选择"""
        if not self.api_keys:
            return None
            
        # 计算每个密钥的权重
        weights = []
        for api_key in self.api_keys:
            health = self.key_health[api_key]
            if health.is_considered_healthy(
                self.config.key_failure_threshold,
                self.config.key_recovery_time
            ):
                score = health.get_health_score(
                    self.config.response_time_weight,
                    self.config.success_rate_weight
                )
            else:
                score = 0.1  # 不健康的密钥给予很低的权重
            weights.append(score)
        
        if sum(weights) == 0:
            # 所有权重都为0，回退到轮询
            return self._round_robin_selection()
        
        # 根据权重选择
        import random
        selected_index = random.choices(range(len(self.api_keys)), weights=weights)[0]
        return self.api_keys[selected_index], selected_index
    
    def record_request_result(self, api_key: str, success: bool, response_time: float, error_type: str = None):
        """记录请求结果"""
        if api_key in self.key_health:
            self.key_health[api_key].record_request(success, response_time, error_type)
            
            # 记录日志
            health = self.key_health[api_key]
            if success:
                logger.debug(f"API密钥请求成功: {api_key[:20]}... (响应时间: {response_time:.2f}s, 成功率: {health.get_success_rate():.2%})")
            else:
                logger.warning(f"API密钥请求失败: {api_key[:20]}... (连续失败: {health.consecutive_failures}, 错误类型: {error_type})")
    
    def get_health_report(self) -> Dict[str, Dict]:
        """获取健康状态报告"""
        report = {}
        for api_key, health in self.key_health.items():
            masked_key = api_key[:20] + "..." if len(api_key) > 20 else api_key
            report[masked_key] = {
                "total_requests": health.total_requests,
                "success_rate": health.get_success_rate(),
                "avg_response_time": health.avg_response_time,
                "consecutive_failures": health.consecutive_failures,
                "health_score": health.get_health_score(
                    self.config.response_time_weight,
                    self.config.success_rate_weight
                ),
                "is_healthy": health.is_considered_healthy(
                    self.config.key_failure_threshold,
                    self.config.key_recovery_time
                ),
                "failure_reasons": dict(health.failure_reasons)
            }
        return report
    
    def perform_health_check(self):
        """执行健康检查"""
        current_time = time.time()
        if current_time - self.last_health_check < self.config.health_check_interval:
            return
            
        self.last_health_check = current_time
        report = self.get_health_report()
        
        healthy_count = sum(1 for stats in report.values() if stats["is_healthy"])
        total_count = len(report)
        
        logger.info(f"API密钥健康检查: {healthy_count}/{total_count} 个密钥健康")
        
        # 详细健康报告
        for key, stats in report.items():
            if stats["total_requests"] > 0:
                logger.debug(
                    f"密钥 {key}: 成功率 {stats['success_rate']:.2%}, "
                    f"平均响应时间 {stats['avg_response_time']:.2f}s, "
                    f"健康分数 {stats['health_score']:.2f}, "
                    f"状态: {'健康' if stats['is_healthy'] else '不健康'}"
                )


class BatchProcessor:
    """分批API调用处理器"""
    
    def __init__(self, config: DifyAPIConfig, api_key_poller: SmartAPIKeyPoller = None):
        self.config = config
        self.api_key_poller = api_key_poller
        self.batch_results = []
        self.total_requests = 0
        self.successful_requests = 0
        self.failed_requests = 0
        self.processing_start_time = 0
        self._lock = threading.Lock()
        
        logger.info(f"初始化分批处理器，批次大小: {config.batch_size}")
    
    async def process_pages_in_batches(self, pages_data: List[Dict], 
                                     api_call_func, 
                                     progress_callback=None) -> Dict[str, Any]:
        """
        分批处理页面数据
        
        Args:
            pages_data: 页面数据列表
            api_call_func: API调用函数
            progress_callback: 进度回调函数
            
        Returns:
            Dict: 批处理结果
        """
        self.processing_start_time = time.time()
        self.total_requests = len(pages_data)
        self.batch_results = []
        
        logger.info(f"开始分批处理 {self.total_requests} 个页面，每批 {self.config.batch_size} 个")
        
        if not self.config.enable_batch_processing or len(pages_data) <= self.config.batch_size:
            # 不启用分批处理或数据量小，直接处理
            return await self._process_single_batch(pages_data, api_call_func, progress_callback)
        
        # 分批处理
        batches = self._split_into_batches(pages_data)
        
        for batch_index, batch_data in enumerate(batches):
            batch_start_time = time.time()
            
            logger.info(f"处理第 {batch_index + 1}/{len(batches)} 批，包含 {len(batch_data)} 个页面")
            
            try:
                # 处理当前批次
                batch_result = await self._process_single_batch(
                    batch_data, 
                    api_call_func, 
                    progress_callback,
                    batch_index + 1
                )
                
                self.batch_results.append({
                    "batch_index": batch_index + 1,
                    "batch_size": len(batch_data),
                    "batch_result": batch_result,
                    "processing_time": time.time() - batch_start_time,
                    "success": batch_result.get("success", False)
                })
                
                # 更新统计
                with self._lock:
                    batch_successful = batch_result.get("successful_count", 0)
                    batch_failed = len(batch_data) - batch_successful
                    self.successful_requests += batch_successful
                    self.failed_requests += batch_failed
                
                # 批次间延迟（除了最后一批）
                if batch_index < len(batches) - 1:
                    logger.debug(f"批次间延迟 {self.config.batch_delay} 秒")
                    await asyncio.sleep(self.config.batch_delay)
                
            except Exception as e:
                logger.error(f"第 {batch_index + 1} 批处理异常: {str(e)}")
                self.batch_results.append({
                    "batch_index": batch_index + 1,
                    "batch_size": len(batch_data),
                    "error": str(e),
                    "success": False
                })
                
                with self._lock:
                    self.failed_requests += len(batch_data)
        
        # 汇总结果
        return self._consolidate_batch_results()
    
    def _split_into_batches(self, data: List) -> List[List]:
        """将数据分割成批次"""
        batches = []
        for i in range(0, len(data), self.config.batch_size):
            batch = data[i:i + self.config.batch_size]
            batches.append(batch)
        return batches
    
    async def _process_single_batch(self, batch_data: List[Dict], 
                                  api_call_func, 
                                  progress_callback=None,
                                  batch_number: int = 1) -> Dict[str, Any]:
        """处理单个批次"""
        batch_results = []
        successful_count = 0
        
        for i, page_data in enumerate(batch_data):
            try:
                # 调用API处理函数
                result = await api_call_func(page_data)
                
                if result.get("success", False):
                    successful_count += 1
                
                batch_results.append({
                    "page_data": page_data,
                    "result": result,
                    "success": result.get("success", False)
                })
                
                # 更新进度
                if progress_callback:
                    total_processed = (batch_number - 1) * self.config.batch_size + i + 1
                    progress_callback(total_processed, self.total_requests)
                
                # API密钥健康检查
                if self.api_key_poller:
                    self.api_key_poller.perform_health_check()
                
            except Exception as e:
                logger.error(f"处理页面异常: {str(e)}")
                batch_results.append({
                    "page_data": page_data,
                    "error": str(e),
                    "success": False
                })
        
        return {
            "success": True,
            "batch_results": batch_results,
            "successful_count": successful_count,
            "total_count": len(batch_data)
        }
    
    def _consolidate_batch_results(self) -> Dict[str, Any]:
        """汇总所有批次的结果"""
        all_page_results = []
        successful_batches = 0
        
        for batch_info in self.batch_results:
            if batch_info.get("success", False):
                successful_batches += 1
                batch_result = batch_info["batch_result"]
                all_page_results.extend(batch_result.get("batch_results", []))
            else:
                # 失败的批次也要记录
                logger.warning(f"批次 {batch_info.get('batch_index', '?')} 处理失败")
        
        total_processing_time = time.time() - self.processing_start_time
        
        result = {
            "success": True,
            "page_templates": all_page_results,
            "successful_count": self.successful_requests,
            "failed_count": self.failed_requests,
            "total_pages": self.total_requests,
            "total_batches": len(self.batch_results),
            "successful_batches": successful_batches,
            "total_processing_time": total_processing_time,
            "average_batch_time": total_processing_time / len(self.batch_results) if self.batch_results else 0,
            "batch_details": self.batch_results
        }
        
        logger.info(
            f"分批处理完成: {self.successful_requests}/{self.total_requests} 成功, "
            f"{successful_batches}/{len(self.batch_results)} 批次成功, "
            f"总耗时: {total_processing_time:.2f}秒"
        )
        
        return result
    
    def get_processing_stats(self) -> Dict[str, Any]:
        """获取处理统计信息"""
        if self.processing_start_time == 0:
            return {"status": "not_started"}
        
        current_time = time.time()
        elapsed_time = current_time - self.processing_start_time
        
        return {
            "status": "processing" if len(self.batch_results) * self.config.batch_size < self.total_requests else "completed",
            "total_requests": self.total_requests,
            "successful_requests": self.successful_requests,
            "failed_requests": self.failed_requests,
            "completed_batches": len(self.batch_results),
            "elapsed_time": elapsed_time,
            "estimated_remaining_time": self._estimate_remaining_time(elapsed_time)
        }
    
    def _estimate_remaining_time(self, elapsed_time: float) -> float:
        """估算剩余处理时间"""
        if len(self.batch_results) == 0:
            return 0
        
        completed_requests = len(self.batch_results) * self.config.batch_size
        if completed_requests >= self.total_requests:
            return 0
        
        avg_time_per_request = elapsed_time / completed_requests
        remaining_requests = self.total_requests - completed_requests
        
        return avg_time_per_request * remaining_requests


# 保持向后兼容的函数
def get_dify_config():
    """获取Dify配置实例"""
    return DifyAPIConfig()

# 导出主要类和函数
__all__ = ['DifyAPIConfig', 'APIKeyHealth', 'SmartAPIKeyPoller', 'BatchProcessor', 'get_dify_config']
