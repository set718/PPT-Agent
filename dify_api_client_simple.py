#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Dify API基础配置模块（简化版）
仅保留基础配置，移除内容处理功能，专注于模板推荐
"""

import os
from dataclasses import dataclass, field
from typing import List
from logger import get_logger

logger = get_logger()

@dataclass
class DifyAPIConfig:
    """Dify API配置类 - 专用于模板推荐"""
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
    load_balance_strategy: str = "concurrent_random"
    
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

# 保持向后兼容的函数
def get_dify_config():
    """获取Dify配置实例"""
    return DifyAPIConfig()

# 导出主要类和函数
__all__ = ['DifyAPIConfig', 'get_dify_config']
