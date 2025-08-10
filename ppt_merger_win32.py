#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
基于win32com的PPT页面整合器
使用Microsoft PowerPoint COM接口来完整保留格式
"""

import os
import sys
from typing import List, Dict, Any, Optional
from logger import get_logger, log_user_action

logger = get_logger()

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    logger.warning("win32com不可用，无法使用COM接口")

class Win32PPTMerger:
    """基于win32com的PPT页面整合器"""
    
    def __init__(self):
        """初始化整合器"""
        if not WIN32_AVAILABLE:
            raise ImportError("需要安装pywin32: pip install pywin32")
        
        self.ppt_app = None
        self.merged_presentation = None
        self.temp_presentations = []  # 存储临时打开的演示文稿
        
    def __enter__(self):
        """上下文管理器入口"""
        self._start_powerpoint()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器退出，清理资源"""
        self._cleanup()
    
    def _start_powerpoint(self):
        """启动PowerPoint应用程序"""
        try:
            # 初始化COM
            import pythoncom
            pythoncom.CoInitialize()
            
            self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            # PowerPoint 2016+ 不允许设置Visible=False，保持默认可见状态
            
            # 创建新的演示文稿
            self.merged_presentation = self.ppt_app.Presentations.Add()
            
            logger.info("PowerPoint COM接口初始化成功")
            
        except Exception as e:
            logger.error(f"PowerPoint启动失败: {str(e)}")
            raise
    
    def merge_template_pages_to_ppt_file_level(self, page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        文件级别的PPT合并 - 完全保留每个模板的原始格式
        """
        log_user_action("PPT文件级别合并(Win32COM)", f"开始合并{len(page_results)}个模板文件")
        
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
            # 关闭默认演示文稿
            if self.merged_presentation:
                self.merged_presentation.Close()
                self.merged_presentation = None
            
            # 创建一个临时的合并演示文稿
            self.merged_presentation = self.ppt_app.Presentations.Add()
            
            # 删除默认空白页
            if self.merged_presentation.Slides.Count > 0:
                self.merged_presentation.Slides(1).Delete()
            
            # 逐个插入每个模板文件的第一页，完全保留格式
            for i, page_result in enumerate(page_results):
                template_path = page_result.get('template_path')
                
                if not template_path or not os.path.exists(template_path):
                    result["errors"].append(f"第{i+1}页模板文件不存在")
                    result["skipped_pages"] += 1
                    continue
                
                try:
                    abs_path = os.path.abspath(template_path)
                    
                    # 使用InsertFromFile，但指定保持源格式的参数
                    slide_index = self.merged_presentation.Slides.Count
                    self.merged_presentation.Slides.InsertFromFile(abs_path, slide_index, 1, 1)
                    
                    result["processed_pages"] += 1
                    logger.info(f"文件级别合并页面: {os.path.basename(template_path)}")
                    
                except Exception as e:
                    error_msg = f"第{i+1}页文件合并失败: {str(e)}"
                    result["errors"].append(error_msg)
                    result["skipped_pages"] += 1
                    logger.error(error_msg)
            
            # 保存合并结果
            output_path = self._save_presentation()
            if output_path:
                result["output_path"] = output_path
                result["success"] = True
                result["total_pages"] = self.merged_presentation.Slides.Count
                
                # 读取文件字节数据
                try:
                    with open(output_path, 'rb') as f:
                        result["presentation_bytes"] = f.read()
                except Exception as read_e:
                    logger.warning(f"读取PPT字节数据失败: {str(read_e)}")
                
                log_user_action("PPT文件级别合并完成", 
                               f"成功: {result['processed_pages']}页, 跳过: {result['skipped_pages']}页")
            else:
                result["error"] = "PPT文件保存失败"
        
        except Exception as e:
            result["error"] = f"文件级别合并异常: {str(e)}"
            logger.error(f"文件级别合并异常: {str(e)}")
        
        return result

    def merge_template_pages_to_ppt_append_style(self, page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        使用AppendSlide风格的合并，类似Spire.Presentation的AppendBySlide
        """
        log_user_action("PPT AppendSlide风格合并", f"开始合并{len(page_results)}个模板文件")
        
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
            # 关闭默认演示文稿
            if self.merged_presentation:
                self.merged_presentation.Close()
                self.merged_presentation = None
            
            # 使用第一个模板作为基础
            first_template_path = os.path.abspath(page_results[0].get('template_path'))
            self.merged_presentation = self.ppt_app.Presentations.Open(first_template_path)
            result["processed_pages"] = 1
            
            logger.info(f"使用第一个模板作为基础: {os.path.basename(first_template_path)}")
            
            # 为后续模板使用Range.Copy和Paste方法
            for i in range(1, len(page_results)):
                template_path = page_results[i].get('template_path')
                
                if not template_path or not os.path.exists(template_path):
                    result["errors"].append(f"第{i+1}页模板文件不存在")
                    result["skipped_pages"] += 1
                    continue
                
                try:
                    abs_path = os.path.abspath(template_path)
                    
                    # 创建临时演示文稿
                    temp_ppt = self.ppt_app.Presentations.Open(abs_path, ReadOnly=True, WithWindow=False)
                    
                    if temp_ppt.Slides.Count > 0:
                        # 使用Range方法选择第一张slide
                        slide_range = temp_ppt.Slides.Range([1])
                        
                        # 复制选中的slide
                        slide_range.Copy()
                        
                        # 在目标位置粘贴，保持源格式
                        slide_index = self.merged_presentation.Slides.Count + 1
                        self.merged_presentation.Slides.Paste(slide_index)
                        
                        result["processed_pages"] += 1
                        logger.info(f"Range.Copy合并: {os.path.basename(template_path)}")
                    else:
                        result["skipped_pages"] += 1
                        result["errors"].append(f"第{i+1}页模板文件为空")
                    
                    # 关闭临时演示文稿
                    temp_ppt.Close()
                    
                except Exception as e:
                    error_msg = f"第{i+1}页Range.Copy失败: {str(e)}"
                    result["errors"].append(error_msg)
                    result["skipped_pages"] += 1
                    logger.error(error_msg)
            
            # 保存合并结果
            output_path = self._save_presentation()
            if output_path:
                result["output_path"] = output_path
                result["success"] = True
                result["total_pages"] = self.merged_presentation.Slides.Count
                
                # 读取文件字节数据
                try:
                    with open(output_path, 'rb') as f:
                        result["presentation_bytes"] = f.read()
                except Exception as read_e:
                    logger.warning(f"读取PPT字节数据失败: {str(read_e)}")
                
                log_user_action("PPT AppendSlide风格合并完成", 
                               f"成功: {result['processed_pages']}页, 跳过: {result['skipped_pages']}页")
            else:
                result["error"] = "PPT文件保存失败"
        
        except Exception as e:
            result["error"] = f"AppendSlide风格合并异常: {str(e)}"
            logger.error(f"AppendSlide风格合并异常: {str(e)}")
        
        return result

    def merge_template_pages_to_ppt(self, page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        将模板页面原样整合成完整的PPT（使用AppendSlide风格策略）
        
        Args:
            page_results: 页面处理结果列表，每个元素包含模板路径信息
            
        Returns:
            Dict: 整合结果
        """
        # 使用AppendSlide风格合并，更好地保留格式
        return self.merge_template_pages_to_ppt_append_style(page_results)
    
    def _copy_template_page_win32(self, page_result: Dict[str, Any], page_number: int) -> bool:
        """
        使用win32com复制模板页面（完整保留格式和颜色）
        
        Args:
            page_result: 页面结果数据
            page_number: 页面编号
            
        Returns:
            bool: 是否成功复制
        """
        try:
            template_path = page_result.get('template_path')
            
            if not template_path or not os.path.exists(template_path):
                logger.warning(f"模板文件不存在: {template_path}")
                return False
            
            # 获取绝对路径
            abs_template_path = os.path.abspath(template_path)
            
            # 打开模板文件
            template_ppt = self.ppt_app.Presentations.Open(abs_template_path, ReadOnly=True, WithWindow=False)
            self.temp_presentations.append(template_ppt)
            
            if template_ppt.Slides.Count == 0:
                logger.warning(f"模板文件为空: {template_path}")
                return False
            
            # 获取模板的第一页
            template_slide = template_ppt.Slides(1)
            
            # 方法1: 使用Copy和Paste保留完整格式
            template_slide.Copy()
            
            # 粘贴到合并的演示文稿中，保持源格式
            if self.merged_presentation.Slides.Count == 0:
                # 如果是第一页，直接粘贴
                self.merged_presentation.Slides.Paste()
            else:
                # 在指定位置粘贴
                paste_index = self.merged_presentation.Slides.Count + 1
                self.merged_presentation.Slides.Paste(paste_index)
            
            # 获取刚粘贴的幻灯片
            pasted_slide = self.merged_presentation.Slides(self.merged_presentation.Slides.Count)
            
            # 确保保留源幻灯片的背景和主题
            try:
                # 复制背景格式
                if hasattr(template_slide, 'Background') and hasattr(pasted_slide, 'Background'):
                    # 复制背景填充
                    pasted_slide.Background.Fill.ForeColor.RGB = template_slide.Background.Fill.ForeColor.RGB
                    pasted_slide.Background.Fill.BackColor.RGB = template_slide.Background.Fill.BackColor.RGB
                    pasted_slide.Background.Fill.Type = template_slide.Background.Fill.Type
                    
                # 复制颜色方案
                if hasattr(template_slide, 'ColorScheme') and hasattr(pasted_slide, 'ColorScheme'):
                    for i in range(1, 9):  # PowerPoint有8种颜色方案
                        try:
                            pasted_slide.ColorScheme.Colors(i).RGB = template_slide.ColorScheme.Colors(i).RGB
                        except:
                            pass
                            
            except Exception as format_e:
                logger.warning(f"格式复制异常: {str(format_e)}")
            
            logger.info(f"成功复制模板页面(Win32COM): {os.path.basename(template_path)}")
            return True
            
        except Exception as e:
            logger.error(f"Win32COM页面复制失败: {str(e)}")
            return False
    
    def _save_presentation(self) -> Optional[str]:
        """
        保存演示文稿到文件
        
        Returns:
            Optional[str]: 保存的文件路径
        """
        try:
            # 生成临时文件路径
            import tempfile
            import datetime
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = tempfile.gettempdir()
            output_filename = f"merged_presentation_{timestamp}.pptx"
            output_path = os.path.join(output_dir, output_filename)
            
            # 保存为PowerPoint格式
            # ppSaveAsOpenXMLPresentation = 24 (PowerPoint 2007-2019 format)
            self.merged_presentation.SaveAs(output_path, 24)
            
            logger.info(f"PPT文件已保存: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"保存PPT文件失败: {str(e)}")
            return None
    
    def _cleanup(self):
        """清理资源"""
        try:
            # 关闭所有临时打开的演示文稿
            for temp_ppt in self.temp_presentations:
                try:
                    temp_ppt.Close()
                except:
                    pass
            
            # 关闭合并的演示文稿
            if self.merged_presentation:
                try:
                    self.merged_presentation.Close()
                except:
                    pass
            
            # 退出PowerPoint应用程序
            if self.ppt_app:
                try:
                    self.ppt_app.Quit()
                except:
                    pass
            
            # 清理COM
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except:
                pass
                    
            logger.info("PowerPoint COM资源已清理")
            
        except Exception as e:
            logger.warning(f"资源清理异常: {str(e)}")

def merge_dify_templates_to_ppt_win32(page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    便捷函数：使用win32com将Dify API匹配的模板页面整合为PPT
    
    Args:
        page_results: 页面处理结果列表
        
    Returns:
        Dict: 整合结果
    """
    if not WIN32_AVAILABLE:
        return {
            "success": False,
            "error": "win32com不可用，请安装: pip install pywin32"
        }
    
    try:
        with Win32PPTMerger() as merger:
            return merger.merge_template_pages_to_ppt(page_results)
    except Exception as e:
        return {
            "success": False,
            "error": f"Win32COM PPT整合失败: {str(e)}"
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
    
    result = merge_dify_templates_to_ppt_win32(test_results)
    
    if result["success"]:
        print("PPT模板整合成功(Win32COM)！")
        print(f"总页数: {result['total_pages']}")
        print(f"处理页数: {result['processed_pages']}")
        print(f"跳过页数: {result['skipped_pages']}")
        print(f"输出文件: {result['output_path']}")
    else:
        print(f"PPT整合失败: {result['error']}")
        if result.get("errors"):
            for error in result["errors"]:
                print(f"  - {error}")