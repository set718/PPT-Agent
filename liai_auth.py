#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Liai API认证模块
处理M2M token认证，专用于Liai的模板推荐功能

注意：M2M token仅用于Liai的模板推荐功能，不影响分页功能和其他API调用
"""

import os
import asyncio
import time
import threading
from typing import Dict, Any, Optional
from logger import get_logger

logger = get_logger()

try:
    from idaas.app import TokenManager
    IDAAS_AVAILABLE = True
except ImportError:
    IDAAS_AVAILABLE = False
    # 不在这里打印警告，等真正需要时再提示

class LiaiM2MAuth:
    """Liai M2M认证管理器"""
    
    def __init__(self):
        # 这些变量名应该与你代码中使用的环境变量名一致
        self.app_id = os.getenv('APP_ID', '')
        self.app_secret = os.getenv('APP_Secret', '')  # 注意大小写
        self.service_id = os.getenv('LIAI_SERVICE_ID', '1tmGCFoJSQiQmR46HK97rX')
        self.auth_url = os.getenv('LIAI_AUTH_URL', 'https://id.lixiang.com/api')
        
        self.access_token = None
        self.token_expires_at = 0
        self._lock = threading.Lock()
        self.manager = None
        
        # 初始化TokenManager（延迟初始化，只在需要时初始化）
        if IDAAS_AVAILABLE and self.app_id and self.app_secret:
            try:
                self.manager = TokenManager.singleton_m2m(
                    self.auth_url,
                    self.app_id,
                    self._get_secret,
                    cache=True
                )
                logger.info("Liai M2M认证管理器初始化成功")
            except Exception as e:
                logger.error(f"初始化Liai M2M认证管理器失败: {e}")
        else:
            # 不在初始化时打印警告，等真正需要使用时再提示
            pass
    
    def _get_secret(self, client_id: str) -> str:
        """获取客户端密钥"""
        return self.app_secret
    
    async def get_access_token(self) -> Optional[str]:
        """获取有效的access token"""
        with self._lock:
            # 检查token是否仍然有效（提前5分钟刷新）
            if self.access_token and time.time() < (self.token_expires_at - 300):
                return self.access_token
        
        # 需要获取新token
        if not self.manager:
            if not IDAAS_AVAILABLE:
                logger.error("idaas库未安装，M2M认证不可用")
            elif not (self.app_id and self.app_secret):
                logger.error("M2M认证配置不完整，缺少APP_ID或APP_Secret")
            else:
                logger.error("TokenManager初始化失败")
            return None
        
        try:
            logger.info(f"正在获取Liai M2M access token... APP_ID: {self.app_id[:10]}..., SERVICE_ID: {self.service_id}")
            # 对照你的示例修正参数
            bundle = await self.manager.get_token(
                self.app_id,
                self.service_id,
                "read:api"  # 这里可能需要调整为正确的scope
            )
            
            if bundle and bundle.access_token:
                with self._lock:
                    self.access_token = bundle.access_token
                    # 设置过期时间（通常是1小时，提前5分钟刷新）
                    self.token_expires_at = time.time() + 3300  # 55分钟
                
                logger.info(f"Liai M2M access token获取成功: {self.access_token[:20]}...")
                return self.access_token
            else:
                logger.error("M2M token bundle为空或无access_token")
                return None
            
        except Exception as e:
            logger.error(f"获取Liai M2M access token失败: {e}")
            return None
    
    def is_configured(self) -> bool:
        """检查是否正确配置"""
        return (IDAAS_AVAILABLE and 
                bool(self.app_id) and 
                bool(self.app_secret) and 
                self.manager is not None)
    
    async def get_auth_headers(self) -> Dict[str, str]:
        """获取认证头"""
        access_token = await self.get_access_token()
        if not access_token:
            raise Exception("无法获取Liai M2M access token")
        
        return {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

# 全局M2M认证管理器实例
liai_m2m_auth = LiaiM2MAuth()

async def get_liai_auth_headers() -> Dict[str, str]:
    """获取Liai认证头的便捷函数"""
    return await liai_m2m_auth.get_auth_headers()

def is_liai_m2m_configured() -> bool:
    """检查Liai M2M是否配置完整"""
    return liai_m2m_auth.is_configured()

if __name__ == "__main__":
    # 测试M2M认证
    async def test_auth():
        auth = LiaiM2MAuth()
        
        print(f"配置状态: {'已配置' if auth.is_configured() else '未配置'}")
        print(f"APP_ID: {auth.app_id[:10]}..." if auth.app_id else "APP_ID: 未设置")
        print(f"SERVICE_ID: {auth.service_id}")
        print(f"AUTH_URL: {auth.auth_url}")
        
        if auth.is_configured():
            try:
                token = await auth.get_access_token()
                if token:
                    print(f"Token获取成功: {token[:20]}...")
                else:
                    print("Token获取失败")
            except Exception as e:
                print(f"测试失败: {e}")
        else:
            print("请设置环境变量: APP_ID, APP_Secret")
    
    asyncio.run(test_auth())