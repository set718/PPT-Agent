#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
基于Spire.Presentation的PPT页面整合器
使用AppendBySlide方法完整保留格式
"""

import os
import tempfile
from typing import List, Dict, Any, Optional
from logger import get_logger, log_user_action

logger = get_logger()

try:
    from spire.presentation.common import *
    from spire.presentation import *
    SPIRE_AVAILABLE = True
except ImportError:
    SPIRE_AVAILABLE = False
    logger.warning("Spire.Presentation不可用，请安装: pip install Spire.Presentation")

class SpirePPTMerger:
    """基于Spire.Presentation的PPT页面整合器"""
    
    def __init__(self):
        """初始化整合器"""
        if not SPIRE_AVAILABLE:
            raise ImportError("需要安装Spire.Presentation: pip install Spire.Presentation")
    
    def merge_template_pages_to_ppt(self, page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        使用Spire.Presentation的AppendBySlide方法合并PPT
        
        Args:
            page_results: 页面处理结果列表，每个元素包含模板路径信息
            
        Returns:
            Dict: 整合结果
        """
        log_user_action("PPT Spire合并", f"开始合并{len(page_results)}个模板页面")
        
        result = {
            "success": False,
            "total_pages": 0,
            "processed_pages": 0,
            "skipped_pages": 0,
            "errors": [],
            "presentation_bytes": None,
            "output_path": None,
            "error": None
        }
        
        if len(page_results) == 0:
            result["error"] = "没有页面需要合并"
            return result
        
        try:
            # 创建主演示文稿对象
            main_presentation = None
            
            # 逐个处理每个模板文件
            for i, page_result in enumerate(page_results):
                template_path = page_result.get('template_path')
                
                if not template_path or not os.path.exists(template_path):
                    result["errors"].append(f"第{i+1}页模板文件不存在: {template_path}")
                    result["skipped_pages"] += 1
                    continue
                
                try:
                    # 加载当前模板演示文稿
                    template_presentation = Presentation()
                    template_presentation.LoadFromFile(os.path.abspath(template_path))
                    
                    if len(template_presentation.Slides) == 0:
                        result["errors"].append(f"第{i+1}页模板文件为空: {os.path.basename(template_path)}")
                        result["skipped_pages"] += 1
                        continue
                    
                    if i == 0:
                        # 第一个模板作为主演示文稿
                        main_presentation = template_presentation
                        result["processed_pages"] += 1
                        logger.info(f"使用第一个模板作为基础: {os.path.basename(template_path)}")
                    else:
                        # 从当前模板获取第一张幻灯片
                        source_slide = template_presentation.Slides[0]
                        
                        # 使用AppendBySlide方法将幻灯片添加到主演示文稿
                        # 这个方法会保留原始格式和样式
                        main_presentation.Slides.AppendBySlide(source_slide)
                        
                        result["processed_pages"] += 1
                        logger.info(f"使用AppendBySlide合并: {os.path.basename(template_path)}")
                        
                        # 释放当前模板的资源
                        template_presentation.Dispose()
                    
                except Exception as e:
                    error_msg = f"第{i+1}页Spire处理失败: {str(e)}"
                    result["errors"].append(error_msg)
                    result["skipped_pages"] += 1
                    logger.error(error_msg)
            
            # 保存合并后的演示文稿
            if main_presentation is not None and result["processed_pages"] > 0:
                output_path = self._save_presentation(main_presentation)
                
                if output_path:
                    result["output_path"] = output_path
                    result["success"] = True
                    result["total_pages"] = len(main_presentation.Slides) if main_presentation is not None else 0
                    
                    # 读取文件字节数据
                    try:
                        with open(output_path, 'rb') as f:
                            result["presentation_bytes"] = f.read()
                    except Exception as read_e:
                        logger.warning(f"读取PPT字节数据失败: {str(read_e)}")
                    
                    log_user_action("PPT Spire合并完成", 
                                   f"成功: {result['processed_pages']}页, 跳过: {result['skipped_pages']}页")
                else:
                    result["error"] = "Spire PPT文件保存失败"
                
                # 释放主演示文稿资源
                if main_presentation is not None:
                    main_presentation.Dispose()
            else:
                result["error"] = "没有成功处理的页面"
                
        except Exception as e:
            result["error"] = f"Spire PPT整合异常: {str(e)}"
            logger.error(f"Spire PPT整合异常: {str(e)}")
            
            # 清理可能的资源
            try:
                if 'main_presentation' in locals() and main_presentation is not None:
                    main_presentation.Dispose()
            except:
                pass
            
        return result
    
    def _save_presentation(self, presentation) -> Optional[str]:
        """
        保存演示文稿到文件
        
        Args:
            presentation: Spire.Presentation对象
            
        Returns:
            Optional[str]: 保存的文件路径
        """
        try:
            # 生成临时文件路径
            import datetime
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = tempfile.gettempdir()
            output_filename = f"merged_presentation_spire_{timestamp}.pptx"
            output_path = os.path.join(output_dir, output_filename)
            
            # 保存为PowerPoint格式
            presentation.SaveToFile(output_path, FileFormat.Pptx2016)
            
            logger.info(f"Spire PPT文件已保存: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Spire保存PPT文件失败: {str(e)}")
            return None

def merge_dify_templates_to_ppt_spire(page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    便捷函数：使用Spire.Presentation将Dify API匹配的模板页面整合为PPT
    
    Args:
        page_results: 页面处理结果列表
        
    Returns:
        Dict: 整合结果
    """
    if not SPIRE_AVAILABLE:
        return {
            "success": False,
            "error": "Spire.Presentation不可用，请安装: pip install Spire.Presentation"
        }
    
    try:
        merger = SpirePPTMerger()
        return merger.merge_template_pages_to_ppt(page_results)
    except Exception as e:
        return {
            "success": False,
            "error": f"Spire PPT整合失败: {str(e)}"
        }

if __name__ == "__main__":
    # 测试用例
    test_results = [
        {
            'page_number': 1,
            'template_number': 'title',
            'template_path': os.path.join("templates", "title_slides.pptx"),
            'template_filename': "title_slides.pptx"
        },
        {
            'page_number': 2,
            'template_number': 5,
            'template_path': os.path.join("templates", "ppt_template", "split_presentations_5.pptx"),
            'template_filename': "split_presentations_5.pptx"
        }
    ]
    
    result = merge_dify_templates_to_ppt_spire(test_results)
    
    if result["success"]:
        print("PPT Spire合并成功！")
        print(f"总页数: {result['total_pages']}")
        print(f"处理页数: {result['processed_pages']}")
        print(f"跳过页数: {result['skipped_pages']}")
        print(f"输出文件: {result['output_path']}")
    else:
        print(f"PPT合并失败: {result['error']}")
        if result.get("errors"):
            for error in result["errors"]:
                print(f"  - {error}")