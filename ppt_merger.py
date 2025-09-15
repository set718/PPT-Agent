#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPT页面整合器
将多个Dify API匹配的模板页面原样整合成一个完整的PPT文件

合并策略：
- 优先使用Spire.Presentation合并器（无论页数多少）
- 超过10页时自动分批处理（每批10页，适配Spire免费版限制）
- Win32COM合并器仅作为备用方案（当Spire失败时）
- 支持生成多个批次文件供用户下载
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
    增强版PPT合并函数：优先使用Spire合并器，支持分批处理
    
    合并策略：
    - 优先使用Spire.Presentation合并器（无论页数多少）
    - 超过10页时，Spire分批处理（每批10页，考虑免费版限制）
    - Win32COM合并器仅作为备用方案（当Spire失败时）
    - 如果两种合并器都不可用，直接报错
    
    Args:
        page_results: 页面处理结果列表
        
    Returns:
        Dict: 整合结果
    """
    page_count = len(page_results)
    log_user_action("增强版PPT合并", f"优先使用Spire合并器，页面数量({page_count}页)")

    # 1. 优先使用Spire合并器（无论页数多少）
    if SPIRE_AVAILABLE:
        if page_count <= 10:
            # 10页及以下：直接使用Spire合并
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
        else:
            # 10页以上：使用Spire分批处理
            logger.info(f"页面数量{page_count}页(>10页)，使用Spire分批合并策略")
            try:
                result = merge_dify_templates_to_ppt_spire_batch(page_results)
                if result.get("success"):
                    logger.info("Spire分批合并成功")
                    return result
                else:
                    logger.warning(f"Spire分批合并失败: {result.get('error')}")
            except Exception as e:
                logger.warning(f"Spire分批合并异常: {str(e)}")
    
    # 2. Spire失败时，使用Win32COM作为备用方案
    if WIN32_AVAILABLE and sys.platform.startswith('win'):
        logger.info("Spire合并失败，使用Win32COM作为备用方案")
        try:
            result = merge_dify_templates_to_ppt_win32(page_results)
            if result.get("success"):
                logger.info("Win32COM备用合并成功")
                return result
            else:
                logger.warning(f"Win32COM备用合并失败: {result.get('error')}")
        except Exception as e:
            logger.warning(f"Win32COM备用合并异常: {str(e)}")

    # 所有可用的合并方法都失败
    error_msg = "PPT合并失败：Spire.Presentation和Win32COM合并器均不可用，请检查系统环境和依赖安装"
    logger.error(error_msg)
    return {
        "success": False,
        "error": error_msg,
        "total_pages": 0,
        "processed_pages": 0,
        "skipped_pages": 0,
        "errors": ["Spire.Presentation合并器和Win32COM合并器均不可用"]
    }


def merge_dify_templates_to_ppt_spire_batch(page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Spire分批处理合并函数：按AI分页顺序分批处理，每批最多10页
    
    重要：严格按照AI分页的page_number顺序分批（1-10页第一批，11-20页第二批等），
    确保PPT内容的逻辑连贯性。考虑到Spire.Presentation免费版有10页限制。
    
    Args:
        page_results: 页面处理结果列表
        
    Returns:
        Dict: 整合结果，包含多个文件的信息
    """
    if not SPIRE_AVAILABLE:
        return {
            "success": False,
            "error": "Spire.Presentation不可用，无法进行分批合并",
            "total_pages": 0,
            "processed_pages": 0,
            "skipped_pages": 0,
            "errors": ["Spire.Presentation不可用"]
        }
    
    # 首先按page_number排序，确保页面顺序正确
    sorted_page_results = sorted(page_results, key=lambda x: x.get('page_number', 0))
    
    # 分离结尾页和其他页面
    ending_pages = [page for page in sorted_page_results if page.get('page_type') == 'ending']
    non_ending_pages = [page for page in sorted_page_results if page.get('page_type') != 'ending']
    
    page_count = len(sorted_page_results)
    batch_size = 10
    
    # 重新计算批次逻辑：确保结尾页在最后一批
    if len(non_ending_pages) > 0:
        # 计算非结尾页需要多少批次
        non_ending_batch_count = (len(non_ending_pages) + batch_size - 1) // batch_size
        
        # 如果最后一批不满10页，可以放入结尾页
        last_batch_capacity = batch_size - (len(non_ending_pages) % batch_size or batch_size)
        if last_batch_capacity >= len(ending_pages):
            # 结尾页可以放入最后一批
            batch_count = non_ending_batch_count
        else:
            # 结尾页需要单独一批
            batch_count = non_ending_batch_count + 1
    else:
        # 只有结尾页
        batch_count = 1
    
    log_user_action("Spire分批合并", f"按页面顺序将{page_count}页分成{batch_count}个批次处理（结尾页在最后）")
    logger.info(f"开始Spire分批合并：{page_count}页 → {batch_count}个批次（每批最多{batch_size}页，结尾页在最后一批）")
    logger.info(f"页面分离结果：{len(non_ending_pages)}个非结尾页，{len(ending_pages)}个结尾页")
    
    result = {
        "success": False,
        "total_pages": page_count,
        "processed_pages": 0,
        "skipped_pages": 0,
        "errors": [],
        "batch_files": [],  # 存储多个批次文件信息
        "batch_count": batch_count,
        "presentation_bytes": None  # 对于分批处理，这个字段将包含主要文件
    }
    
    successful_batches = 0
    
    for batch_index in range(batch_count):
        if batch_index < batch_count - 1:
            # 非最后一批：只包含非结尾页
            start_idx = batch_index * batch_size
            end_idx = min(start_idx + batch_size, len(non_ending_pages))
            batch_pages = non_ending_pages[start_idx:end_idx]
        else:
            # 最后一批：包含剩余的非结尾页 + 所有结尾页
            remaining_start = batch_index * batch_size
            remaining_non_ending = non_ending_pages[remaining_start:] if remaining_start < len(non_ending_pages) else []
            batch_pages = remaining_non_ending + ending_pages
            
            # 确保最后一批内部也按页面顺序排序（结尾页自然会在最后，因为page_number最大）
            batch_pages = sorted(batch_pages, key=lambda x: x.get('page_number', 0))
        
        if not batch_pages:
            continue
            
        # 获取实际的页面号范围
        actual_start_page = batch_pages[0].get('page_number', 1)
        actual_end_page = batch_pages[-1].get('page_number', 1)
        
        logger.info(f"处理第{batch_index + 1}/{batch_count}批次：第{actual_start_page}-{actual_end_page}页（共{len(batch_pages)}页，包含{len([p for p in batch_pages if p.get('page_type') == 'ending'])}个结尾页）")
        
        try:
            # 使用现有的Spire合并器处理这一批
            batch_result = merge_dify_templates_to_ppt_spire(batch_pages)
            
            if batch_result.get("success"):
                successful_batches += 1
                result["processed_pages"] += batch_result.get("processed_pages", 0)
                
                # 保存批次文件信息
                batch_info = {
                    "batch_index": batch_index + 1,
                    "batch_name": f"PPT_第{batch_index + 1}批次_第{actual_start_page}-{actual_end_page}页",
                    "pages_in_batch": len(batch_pages),
                    "actual_start_page": actual_start_page,
                    "actual_end_page": actual_end_page,
                    "presentation_bytes": batch_result.get("presentation_bytes"),
                    "file_size_mb": len(batch_result.get("presentation_bytes", b"")) / (1024 * 1024)
                }
                result["batch_files"].append(batch_info)
                
                # 第一个成功的批次作为主要文件
                if result["presentation_bytes"] is None:
                    result["presentation_bytes"] = batch_result.get("presentation_bytes")
                
                logger.info(f"第{batch_index + 1}批次合并成功，包含{len(batch_pages)}页")
            else:
                result["skipped_pages"] += len(batch_pages)
                error_msg = f"第{batch_index + 1}批次合并失败: {batch_result.get('error', '未知错误')}"
                result["errors"].append(error_msg)
                logger.warning(error_msg)
                
        except Exception as e:
            result["skipped_pages"] += len(batch_pages)
            error_msg = f"第{batch_index + 1}批次处理异常: {str(e)}"
            result["errors"].append(error_msg)
            logger.error(error_msg)
    
    # 判断整体成功状态
    if successful_batches > 0:
        result["success"] = True
        result["successful_batches"] = successful_batches
        result["failed_batches"] = batch_count - successful_batches
        
        logger.info(f"Spire分批合并完成：成功{successful_batches}/{batch_count}个批次，"
                   f"处理{result['processed_pages']}/{page_count}页")
        
        if result["failed_batches"] > 0:
            logger.warning(f"有{result['failed_batches']}个批次处理失败")
    else:
        result["error"] = f"所有{batch_count}个批次都合并失败"
        logger.error(result["error"])
    
    return result


def merge_dify_templates_to_ppt(page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    基础版合并已弃用。请使用 merge_dify_templates_to_ppt_enhanced。
    """
    return {
        "success": False,
        "error": "基础版合并已弃用，请使用 merge_dify_templates_to_ppt_enhanced",
    } 