#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPT水印去除模块
专门用于去除Spire.Presentation免费版生成的水印
"""

import os
import zipfile
import tempfile
import shutil
import xml.etree.ElementTree as ET
import re
from typing import Optional
import logging

logger = logging.getLogger(__name__)

def remove_spire_watermark(pptx_path: str, output_path: Optional[str] = None) -> str:
    """
    去除Spire.Presentation生成的水印
    
    Args:
        pptx_path: 输入的PPTX文件路径
        output_path: 输出文件路径，如果为None则覆盖原文件
    
    Returns:
        处理后的文件路径
    """
    try:
        # 如果没有指定输出路径，则覆盖原文件
        if output_path is None:
            output_path = pptx_path
        
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            # 解压PPTX文件
            extract_dir = os.path.join(temp_dir, "pptx_content")
            with zipfile.ZipFile(pptx_path, 'r') as zip_file:
                zip_file.extractall(extract_dir)
            
            # 处理所有幻灯片
            slides_dir = os.path.join(extract_dir, "ppt", "slides")
            if os.path.exists(slides_dir):
                removed_count = 0
                
                # 遍历所有slide文件
                for filename in os.listdir(slides_dir):
                    if filename.endswith('.xml') and filename.startswith('slide'):
                        slide_path = os.path.join(slides_dir, filename)
                        if _remove_watermark_from_slide(slide_path):
                            removed_count += 1
                            logger.info(f"从 {filename} 中移除了水印")
                
                logger.info(f"总共从 {removed_count} 个幻灯片中移除了水印")
            
            # 重新打包PPTX文件
            _create_pptx_from_directory(extract_dir, output_path)
            
        return output_path
        
    except Exception as e:
        logger.error(f"去除水印时发生错误: {e}")
        raise

def _remove_watermark_from_slide(slide_xml_path: str) -> bool:
    """
    从单个幻灯片XML文件中移除水印
    
    Args:
        slide_xml_path: 幻灯片XML文件路径
    
    Returns:
        是否找到并移除了水印
    """
    try:
        # 定义命名空间
        namespaces = {
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
        
        # 注册命名空间
        for prefix, uri in namespaces.items():
            ET.register_namespace(prefix, uri)
        
        # 解析XML
        tree = ET.parse(slide_xml_path)
        root = tree.getroot()
        
        removed = False
        
        # 查找所有形状元素
        shape_elements = root.findall('.//p:sp', namespaces)
        
        for shape in shape_elements:
            # 检查是否为水印形状
            if _is_spire_watermark(shape, namespaces):
                # 找到父元素并移除这个形状
                parent = root.find('.//p:sp/..', namespaces)
                if parent is None:
                    # 如果没有找到直接父元素，在整个树中搜索
                    for elem in root.iter():
                        if shape in list(elem):
                            elem.remove(shape)
                            removed = True
                            break
                else:
                    parent.remove(shape)
                    removed = True
        
        # 如果移除了水印，保存文件
        if removed:
            tree.write(slide_xml_path, encoding='utf-8', xml_declaration=True)
        
        return removed
        
    except Exception as e:
        logger.error(f"处理幻灯片 {slide_xml_path} 时发生错误: {e}")
        return False

def _is_spire_watermark(shape_element, namespaces: dict) -> bool:
    """
    判断一个形状元素是否为Spire水印
    
    Args:
        shape_element: 形状XML元素
        namespaces: XML命名空间字典
    
    Returns:
        是否为Spire水印
    """
    try:
        # 检查1: 查找文本内容
        text_elements = shape_element.findall('.//a:t', namespaces)
        for text_elem in text_elements:
            if text_elem.text and 'Spire.Presentation' in text_elem.text:
                logger.debug(f"找到Spire水印文本: {text_elem.text}")
                return True
        
        # 检查2: 查找红色文本 (FF0000)
        color_elements = shape_element.findall('.//a:srgbClr[@val="FF0000"]', namespaces)
        if color_elements:
            # 进一步检查是否包含相关文本
            for text_elem in text_elements:
                if text_elem.text and ('Evaluation Warning' in text_elem.text or 'document was created' in text_elem.text):
                    logger.debug("找到红色的评估警告文本")
                    return True
        
        # 检查3: 查找具有特定锁定属性的形状
        lock_elements = shape_element.findall('.//a:spLocks', namespaces)
        for lock_elem in lock_elements:
            if (lock_elem.get('noSelect') == '1' and 
                lock_elem.get('noMove') == '1' and 
                lock_elem.get('noResize') == '1' and 
                lock_elem.get('noTextEdit') == '1'):
                # 如果有锁定属性，再检查是否包含相关文本
                for text_elem in text_elements:
                    if text_elem.text and ('Warning' in text_elem.text or 'Evaluation' in text_elem.text):
                        logger.debug("找到锁定的警告形状")
                        return True
        
        # 检查4: 查找名称为"New shape"且包含警告文本的形状
        name_elements = shape_element.findall('.//p:cNvPr', namespaces)
        for name_elem in name_elements:
            if name_elem.get('name') == 'New shape':
                for text_elem in text_elements:
                    if text_elem.text and 'Warning' in text_elem.text:
                        logger.debug("找到名为'New shape'的警告形状")
                        return True
        
        return False
        
    except Exception as e:
        logger.error(f"判断水印时发生错误: {e}")
        return False

def _create_pptx_from_directory(source_dir: str, output_path: str):
    """
    从目录创建PPTX文件
    
    Args:
        source_dir: 源目录路径
        output_path: 输出PPTX文件路径
    """
    try:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for root, dirs, files in os.walk(source_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    # 计算相对路径
                    arcname = os.path.relpath(file_path, source_dir)
                    # 使用正斜杠作为路径分隔符（ZIP标准）
                    arcname = arcname.replace(os.path.sep, '/')
                    zip_file.write(file_path, arcname)
        
        logger.info(f"成功创建处理后的PPTX文件: {output_path}")
        
    except Exception as e:
        logger.error(f"创建PPTX文件时发生错误: {e}")
        raise

def batch_remove_watermarks(input_dir: str, output_dir: Optional[str] = None) -> list:
    """
    批量去除目录中所有PPTX文件的水印
    
    Args:
        input_dir: 输入目录
        output_dir: 输出目录，如果为None则覆盖原文件
    
    Returns:
        处理结果列表
    """
    results = []
    
    try:
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        for filename in os.listdir(input_dir):
            if filename.lower().endswith('.pptx'):
                input_path = os.path.join(input_dir, filename)
                
                if output_dir:
                    output_path = os.path.join(output_dir, filename)
                else:
                    output_path = None
                
                try:
                    processed_path = remove_spire_watermark(input_path, output_path)
                    results.append({
                        'file': filename,
                        'status': 'success',
                        'output_path': processed_path
                    })
                    logger.info(f"成功处理文件: {filename}")
                    
                except Exception as e:
                    results.append({
                        'file': filename,
                        'status': 'error',
                        'error': str(e)
                    })
                    logger.error(f"处理文件 {filename} 时出错: {e}")
        
        return results
        
    except Exception as e:
        logger.error(f"批量处理时发生错误: {e}")
        return []

# 命令行接口
if __name__ == "__main__":
    import sys
    import argparse
    
    parser = argparse.ArgumentParser(description="PPT水印去除工具")
    parser.add_argument("input", help="输入PPTX文件或目录路径")
    parser.add_argument("-o", "--output", help="输出文件或目录路径")
    parser.add_argument("-v", "--verbose", action="store_true", help="详细输出")
    
    args = parser.parse_args()
    
    # 设置日志级别
    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=level, format='%(asctime)s - %(levelname)s - %(message)s')
    
    try:
        if os.path.isfile(args.input):
            # 处理单个文件
            result = remove_spire_watermark(args.input, args.output)
            print(f"处理完成: {result}")
        elif os.path.isdir(args.input):
            # 批量处理
            results = batch_remove_watermarks(args.input, args.output)
            print(f"批量处理完成，共处理 {len(results)} 个文件")
            for result in results:
                print(f"  {result['file']}: {result['status']}")
        else:
            print(f"错误: 输入路径不存在: {args.input}")
            sys.exit(1)
            
    except Exception as e:
        print(f"处理时发生错误: {e}")
        sys.exit(1)