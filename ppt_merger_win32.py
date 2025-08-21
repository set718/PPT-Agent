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
    


    def merge_template_pages_to_ppt_perfect_format(self, page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        完美格式保留合并 - 使用ImportSlides保持原始设计
        """
        log_user_action("PPT完美格式保留合并", f"开始合并{len(page_results)}个模板文件")
        
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
            
            # 按页面编号排序，确保结尾页在最后
            sorted_page_results = sorted(page_results, key=lambda x: x.get("page_number", 0))
            
            # 关键改进：以第一个模板为基础创建演示文稿，而不是空白演示文稿
            if sorted_page_results:
                first_template_path = os.path.abspath(sorted_page_results[0].get('template_path'))
                self.merged_presentation = self.ppt_app.Presentations.Open(first_template_path, ReadOnly=False, WithWindow=False)
                result["processed_pages"] = 1
                logger.info(f"以第一个模板为基础创建演示文稿: {os.path.basename(first_template_path)}")
                
                # 从第二个模板开始处理剩余页面
                remaining_results = sorted_page_results[1:]
            else:
                # 如果没有模板，创建空白演示文稿（备用方案）
                self.merged_presentation = self.ppt_app.Presentations.Add()
                if self.merged_presentation.Slides.Count > 0:
                    self.merged_presentation.Slides(1).Delete()
                remaining_results = sorted_page_results
            
            # 逐个导入剩余模板页面，保持完整原始格式
            for i, page_result in enumerate(remaining_results):
                template_path = page_result.get('template_path')
                
                if not template_path or not os.path.exists(template_path):
                    result["errors"].append(f"第{i+1}页模板文件不存在")
                    result["skipped_pages"] += 1
                    continue
                
                try:
                    abs_path = os.path.abspath(template_path)
                    
                    # 打开源模板文件
                    temp_ppt = self.ppt_app.Presentations.Open(abs_path, ReadOnly=True, WithWindow=False)
                    
                    if temp_ppt.Slides.Count > 0:
                        # 获取源幻灯片
                        source_slide = temp_ppt.Slides(1)
                        
                        # 复制整个幻灯片
                        source_slide.Copy()
                        
                        # 粘贴到目标演示文稿
                        paste_index = self.merged_presentation.Slides.Count + 1
                        self.merged_presentation.Slides.Paste(paste_index)
                        
                        # 获取刚粘贴的幻灯片，进行额外的颜色保留验证
                        pasted_slide = self.merged_presentation.Slides(paste_index)
                        try:
                            self._preserve_slide_colors_and_format(source_slide, pasted_slide)
                            logger.debug(f"完美格式合并额外颜色保留检查完成: {os.path.basename(template_path)}")
                        except Exception as color_e:
                            logger.debug(f"完美格式合并额外颜色保留异常: {color_e}")
                        
                        result["processed_pages"] += 1
                        logger.info(f"完美格式保留合并: {os.path.basename(template_path)}")
                    
                    # 关闭临时演示文稿
                    temp_ppt.Close()
                    
                except Exception as e:
                    error_msg = f"第{i+1}页完美格式合并失败: {str(e)}"
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
                    logger.info(f"成功读取PPT字节数据，大小: {len(result['presentation_bytes'])} 字节")
                except Exception as read_e:
                    error_msg = f"读取PPT字节数据失败: {str(read_e)}"
                    logger.error(error_msg)
                    result["error"] = error_msg
                    result["success"] = False
                
                log_user_action("PPT完美格式保留合并完成", 
                               f"成功: {result['processed_pages']}页, 跳过: {result['skipped_pages']}页")
            else:
                result["error"] = "PPT文件保存失败"
        
        except Exception as e:
            result["error"] = f"完美格式保留合并异常: {str(e)}"
            logger.error(f"完美格式保留合并异常: {str(e)}")
        
        return result

    def merge_template_pages_to_ppt(self, page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        将模板页面原样整合成完整的PPT（使用完美格式保留策略）
        
        Args:
            page_results: 页面处理结果列表，每个元素包含模板路径信息
            
        Returns:
            Dict: 整合结果
        """
        # 优先使用完美格式保留合并
        return self.merge_template_pages_to_ppt_perfect_format(page_results)
    
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
            
            # 增强的颜色和格式保留机制
            try:
                self._preserve_slide_colors_and_format(template_slide, pasted_slide)
            except Exception as format_e:
                logger.warning(f"增强格式保留异常: {str(format_e)}")
            
            logger.info(f"成功复制模板页面(Win32COM): {os.path.basename(template_path)}")
            return True
            
        except Exception as e:
            logger.error(f"Win32COM页面复制失败: {str(e)}")
            return False
    
    def _preserve_slide_colors_and_format(self, source_slide, target_slide):
        """
        增强的颜色和格式保留方法
        
        Args:
            source_slide: 源幻灯片
            target_slide: 目标幻灯片
        """
        try:
            # 1. 保留设计模板和颜色方案
            if hasattr(source_slide, 'Design') and hasattr(target_slide, 'Design'):
                try:
                    target_slide.Design = source_slide.Design
                    logger.debug("设计模板保留成功")
                except Exception as e:
                    logger.debug(f"设计模板保留失败: {e}")
            
            if hasattr(source_slide, 'ColorScheme') and hasattr(target_slide, 'ColorScheme'):
                try:
                    target_slide.ColorScheme = source_slide.ColorScheme
                    logger.debug("颜色方案保留成功")
                except Exception as e:
                    logger.debug(f"颜色方案保留失败: {e}")
                    # 尝试逐个复制颜色
                    try:
                        for i in range(1, 9):  # PowerPoint标准8色方案
                            target_slide.ColorScheme.Colors(i).RGB = source_slide.ColorScheme.Colors(i).RGB
                    except:
                        pass
            
            # 2. 保留背景格式
            if hasattr(source_slide, 'Background') and hasattr(target_slide, 'Background'):
                try:
                    # 复制背景类型
                    if hasattr(source_slide.Background, 'Type'):
                        target_slide.Background.Type = source_slide.Background.Type
                    
                    # 复制背景填充
                    if hasattr(source_slide.Background, 'Fill') and hasattr(target_slide.Background, 'Fill'):
                        target_slide.Background.Fill.Type = source_slide.Background.Fill.Type
                        
                        # 复制前景色和背景色
                        if hasattr(source_slide.Background.Fill, 'ForeColor'):
                            target_slide.Background.Fill.ForeColor.RGB = source_slide.Background.Fill.ForeColor.RGB
                        if hasattr(source_slide.Background.Fill, 'BackColor'):
                            target_slide.Background.Fill.BackColor.RGB = source_slide.Background.Fill.BackColor.RGB
                        
                        # 复制渐变属性（如果有）
                        if hasattr(source_slide.Background.Fill, 'GradientAngle'):
                            target_slide.Background.Fill.GradientAngle = source_slide.Background.Fill.GradientAngle
                    
                    logger.debug("背景格式保留成功")
                except Exception as e:
                    logger.debug(f"背景格式保留失败: {e}")
            
            # 3. 保留主题信息
            if hasattr(source_slide, 'ThemeColorScheme') and hasattr(target_slide, 'ThemeColorScheme'):
                try:
                    target_slide.ThemeColorScheme = source_slide.ThemeColorScheme
                    logger.debug("主题颜色方案保留成功")
                except Exception as e:
                    logger.debug(f"主题颜色方案保留失败: {e}")
            
            # 4. 尝试保留形状颜色（仅作为备用措施）
            try:
                self._preserve_shapes_colors(source_slide, target_slide)
            except Exception as e:
                logger.debug(f"形状颜色保留过程异常: {e}")
                
        except Exception as e:
            logger.warning(f"颜色格式保留总体异常: {e}")
    
    def _preserve_shapes_colors(self, source_slide, target_slide):
        """
        保留形状的颜色属性（备用方案）
        
        Args:
            source_slide: 源幻灯片
            target_slide: 目标幻灯片
        """
        try:
            if (hasattr(source_slide, 'Shapes') and hasattr(target_slide, 'Shapes') and 
                source_slide.Shapes.Count == target_slide.Shapes.Count):
                
                for i in range(1, source_slide.Shapes.Count + 1):
                    try:
                        source_shape = source_slide.Shapes(i)
                        target_shape = target_slide.Shapes(i)
                        
                        # 保留填充颜色
                        if (hasattr(source_shape, 'Fill') and hasattr(target_shape, 'Fill') and
                            hasattr(source_shape.Fill, 'ForeColor') and hasattr(target_shape.Fill, 'ForeColor')):
                            target_shape.Fill.ForeColor.RGB = source_shape.Fill.ForeColor.RGB
                        
                        # 保留线条颜色
                        if (hasattr(source_shape, 'Line') and hasattr(target_shape, 'Line') and
                            hasattr(source_shape.Line, 'ForeColor') and hasattr(target_shape.Line, 'ForeColor')):
                            target_shape.Line.ForeColor.RGB = source_shape.Line.ForeColor.RGB
                        
                        # 保留文本颜色
                        if (hasattr(source_shape, 'TextFrame') and hasattr(target_shape, 'TextFrame') and
                            hasattr(source_shape.TextFrame, 'TextRange') and hasattr(target_shape.TextFrame, 'TextRange')):
                            if (hasattr(source_shape.TextFrame.TextRange, 'Font') and 
                                hasattr(target_shape.TextFrame.TextRange, 'Font')):
                                target_shape.TextFrame.TextRange.Font.Color.RGB = source_shape.TextFrame.TextRange.Font.Color.RGB
                                
                    except Exception as shape_e:
                        # 单个形状颜色保留失败不影响整体
                        continue
                        
        except Exception as e:
            logger.debug(f"形状颜色保留异常: {e}")
    
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