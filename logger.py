#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
日志管理模块
提供统一的日志记录功能
"""

import os
import logging
import sys
from datetime import datetime
from typing import Optional
from logging.handlers import RotatingFileHandler
from config import get_config

class ColoredFormatter(logging.Formatter):
    """彩色日志格式化器"""
    
    # 颜色代码
    COLORS = {
        'DEBUG': '\033[36m',    # 青色
        'INFO': '\033[32m',     # 绿色
        'WARNING': '\033[33m',  # 黄色
        'ERROR': '\033[31m',    # 红色
        'CRITICAL': '\033[35m', # 紫色
        'RESET': '\033[0m'      # 重置
    }
    
    def format(self, record):
        # 获取原始消息
        message = super().format(record)
        
        # 只在终端输出时使用颜色
        if hasattr(sys.stderr, 'isatty') and sys.stderr.isatty():
            level_name = record.levelname
            if level_name in self.COLORS:
                color = self.COLORS[level_name]
                reset = self.COLORS['RESET']
                return f"{color}{message}{reset}"
        
        return message

class Logger:
    """日志管理器"""
    
    _instance = None
    _initialized = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not self._initialized:
            self._setup_logger()
            self._initialized = True
    
    def _setup_logger(self):
        """设置日志器"""
        config = get_config()
        
        # 创建日志器
        self.logger = logging.getLogger('ppt_generator')
        
        # 设置日志级别
        level = getattr(logging, config.log_level.upper(), logging.INFO)
        self.logger.setLevel(level)
        
        # 清除现有处理器
        self.logger.handlers.clear()
        
        # 创建格式化器
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        colored_formatter = ColoredFormatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 控制台处理器
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(colored_formatter)
        console_handler.setLevel(level)
        self.logger.addHandler(console_handler)
        
        # 文件处理器
        if config.log_file:
            try:
                file_handler = RotatingFileHandler(
                    config.log_file,
                    maxBytes=10*1024*1024,  # 10MB
                    backupCount=5,
                    encoding='utf-8'
                )
                file_handler.setFormatter(formatter)
                file_handler.setLevel(level)
                self.logger.addHandler(file_handler)
            except Exception as e:
                self.logger.error(f"无法创建文件日志处理器: {e}")
        
        # 防止日志重复
        self.logger.propagate = False
    
    def debug(self, message: str, *args, **kwargs):
        """调试日志"""
        self.logger.debug(message, *args, **kwargs)
    
    def info(self, message: str, *args, **kwargs):
        """信息日志"""
        self.logger.info(message, *args, **kwargs)
    
    def warning(self, message: str, *args, **kwargs):
        """警告日志"""
        self.logger.warning(message, *args, **kwargs)
    
    def error(self, message: str, *args, **kwargs):
        """错误日志"""
        self.logger.error(message, *args, **kwargs)
    
    def critical(self, message: str, *args, **kwargs):
        """严重错误日志"""
        self.logger.critical(message, *args, **kwargs)
    
    def exception(self, message: str, *args, **kwargs):
        """异常日志（包含堆栈跟踪）"""
        self.logger.exception(message, *args, **kwargs)

# 全局日志器实例
_logger_instance = Logger()

def get_logger() -> Logger:
    """获取日志器实例"""
    return _logger_instance

def log_function_call(func_name: str, args: tuple = None, kwargs: dict = None):
    """记录函数调用"""
    logger = get_logger()
    args_str = f"args={args}" if args else ""
    kwargs_str = f"kwargs={kwargs}" if kwargs else ""
    param_str = ", ".join(filter(None, [args_str, kwargs_str]))
    logger.debug(f"调用函数: {func_name}({param_str})")

def log_api_call(api_name: str, status: str, duration: float = None, error: str = None):
    """记录API调用"""
    logger = get_logger()
    duration_str = f"耗时: {duration:.2f}s" if duration else ""
    
    if status == "success":
        logger.info(f"API调用成功: {api_name} {duration_str}")
    elif status == "error":
        logger.error(f"API调用失败: {api_name} {duration_str} 错误: {error}")
    else:
        logger.warning(f"API调用状态未知: {api_name} {duration_str}")

def log_file_operation(operation: str, file_path: str, status: str, error: str = None):
    """记录文件操作"""
    logger = get_logger()
    
    if status == "success":
        logger.info(f"文件操作成功: {operation} - {file_path}")
    elif status == "error":
        logger.error(f"文件操作失败: {operation} - {file_path} 错误: {error}")
    else:
        logger.warning(f"文件操作状态未知: {operation} - {file_path}")

def log_user_action(action: str, details: str = None):
    """记录用户操作"""
    logger = get_logger()
    details_str = f" - {details}" if details else ""
    logger.info(f"用户操作: {action}{details_str}")

def log_system_info(info: str):
    """记录系统信息"""
    logger = get_logger()
    logger.info(f"系统信息: {info}")

def log_performance(operation: str, duration: float, additional_info: str = None):
    """记录性能信息"""
    logger = get_logger()
    info_str = f" - {additional_info}" if additional_info else ""
    logger.info(f"性能: {operation} 耗时 {duration:.2f}s{info_str}")

# 装饰器
def log_execution_time(func):
    """记录函数执行时间的装饰器"""
    import time
    import functools
    
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        logger = get_logger()
        
        try:
            logger.debug(f"开始执行: {func.__name__}")
            result = func(*args, **kwargs)
            end_time = time.time()
            duration = end_time - start_time
            logger.debug(f"执行完成: {func.__name__} 耗时 {duration:.2f}s")
            return result
        except Exception as e:
            end_time = time.time()
            duration = end_time - start_time
            logger.error(f"执行失败: {func.__name__} 耗时 {duration:.2f}s 错误: {e}")
            raise
    
    return wrapper

def log_errors(func):
    """记录函数错误的装饰器"""
    import functools
    
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        logger = get_logger()
        
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger.exception(f"函数 {func.__name__} 发生异常: {e}")
            raise
    
    return wrapper

# 上下文管理器
class LogContext:
    """日志上下文管理器"""
    
    def __init__(self, operation: str, log_level: str = "INFO"):
        self.operation = operation
        self.log_level = log_level
        self.logger = get_logger()
        self.start_time = None
    
    def __enter__(self):
        self.start_time = datetime.now()
        self.logger.info(f"开始操作: {self.operation}")
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        end_time = datetime.now()
        duration = (end_time - self.start_time).total_seconds()
        
        if exc_type is None:
            self.logger.info(f"操作完成: {self.operation} 耗时 {duration:.2f}s")
        else:
            self.logger.error(f"操作失败: {self.operation} 耗时 {duration:.2f}s 错误: {exc_val}")
        
        return False  # 不抑制异常

# 使用示例
if __name__ == "__main__":
    # 测试日志功能
    logger = get_logger()
    
    logger.debug("这是调试信息")
    logger.info("这是信息")
    logger.warning("这是警告")
    logger.error("这是错误")
    logger.critical("这是严重错误")
    
    # 测试上下文管理器
    with LogContext("测试操作"):
        import time
        time.sleep(1)
        print("操作进行中...")
    
    # 测试装饰器
    @log_execution_time
    @log_errors
    def test_function():
        import time
        time.sleep(0.5)
        return "测试完成"
    
    result = test_function()
    print(f"结果: {result}")