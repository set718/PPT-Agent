#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPT页面整合器
将多个Dify API匹配的模板页面原样整合成一个完整的PPT文件
仅保留高级合并策略（格式基准、Win32COM、Spire），基础版合并已弃用。
"""

import os
import sys
from typing import List, Dict, Any
from logger import get_logger, log_user_action

logger = get_logger()

# 导入其他合并器
try:
    from ppt_merger_win32 import merge_dify_templates_to_ppt_win32, WIN32_AVAILABLE
except ImportError:
    WIN32_AVAILABLE = False
    logger.warning("Win32COM合并器不可用")

try:
    from ppt_merger_spire import merge_dify_templates_to_ppt_spire, SPIRE_AVAILABLE
except ImportError:
    SPIRE_AVAILABLE = False
    logger.warning("Spire.Presentation合并器不可用")


def merge_dify_templates_to_ppt_enhanced(page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    增强版PPT合并函数：根据页面数量智能选择最佳合并器
    
    智能选择策略：
    - 10页及以下：使用Spire.Presentation合并器（轻量快速）
    - 10页以上：使用Win32COM合并器（大文件处理能力强）
    - 回退方案：格式基准合并器（兼容性保障）
    
    Args:
        page_results: 页面处理结果列表
        
    Returns:
        Dict: 整合结果
    """
    page_count = len(page_results)
    log_user_action("增强版PPT合并", f"根据页面数量({page_count}页)智能选择合并器")

    # 根据页面数量选择最佳合并器
    if page_count <= 10:
        # 10页及以下：优先使用Spire合并器（轻量快速）
        if SPIRE_AVAILABLE:
            logger.info(f"页面数量{page_count}页(≤10页)，使用Spire.Presentation合并器")
            try:
                result = merge_dify_templates_to_ppt_spire(page_results)
                if result.get("success"):
                    logger.info("Spire合并成功")
                    return result
                else:
                    logger.warning(f"Spire合并失败: {result.get('error')}")
            except Exception as e:
                logger.warning(f"Spire合并异常: {str(e)}")
        
        # Spire不可用时回退到Win32COM
        if WIN32_AVAILABLE and sys.platform.startswith('win'):
            logger.info("Spire不可用，回退到Win32COM合并器")
            try:
                result = merge_dify_templates_to_ppt_win32(page_results)
                if result.get("success"):
                    logger.info("Win32COM合并成功")
                    return result
                else:
                    logger.warning(f"Win32COM合并失败: {result.get('error')}")
            except Exception as e:
                logger.warning(f"Win32COM合并异常: {str(e)}")
    else:
        # 10页以上：优先使用Win32COM合并器（大文件处理能力强）
        if WIN32_AVAILABLE and sys.platform.startswith('win'):
            logger.info(f"页面数量{page_count}页(>10页)，使用Win32COM合并器")
            try:
                result = merge_dify_templates_to_ppt_win32(page_results)
                if result.get("success"):
                    logger.info("Win32COM合并成功")
                    return result
                else:
                    logger.warning(f"Win32COM合并失败: {result.get('error')}")
            except Exception as e:
                logger.warning(f"Win32COM合并异常: {str(e)}")
        
        # Win32COM不可用时回退到Spire
        if SPIRE_AVAILABLE:
            logger.info("Win32COM不可用，回退到Spire.Presentation合并器")
            try:
                result = merge_dify_templates_to_ppt_spire(page_results)
                if result.get("success"):
                    logger.info("Spire合并成功")
                    return result
                else:
                    logger.warning(f"Spire合并失败: {result.get('error')}")
            except Exception as e:
                logger.warning(f"Spire合并异常: {str(e)}")

    # 3. 最后回退到格式基准合并器（统一格式风格）
    logger.info("使用格式基准合并器（统一为split_presentations_1格式风格）")
    try:
        from format_base_merger import merge_with_split_presentations_1_format
        result = merge_with_split_presentations_1_format(page_results)
        if result.get("success"):
            logger.info("格式基准合并成功")
            return result
        else:
            logger.warning(f"格式基准合并失败: {result.get('error')}")
    except Exception as e:
        logger.warning(f"格式基准合并异常: {str(e)}")

    # 所有高级合并方法都失败
    error_msg = "所有可用的合并方法都失败了，请检查模板文件和系统环境"
    logger.error(error_msg)
    return {
        "success": False,
        "error": error_msg,
        "total_pages": 0,
        "processed_pages": 0,
        "skipped_pages": 0,
        "errors": ["格式基准合并器、Spire合并器、Win32COM合并器均不可用"]
    }


def merge_dify_templates_to_ppt(page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    基础版合并已弃用。请使用 merge_dify_templates_to_ppt_enhanced。
    """
    return {
        "success": False,
        "error": "基础版合并已弃用，请使用 merge_dify_templates_to_ppt_enhanced",
    } 