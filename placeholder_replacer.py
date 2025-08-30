    def _replace_single_placeholder_in_table_cell(self, placeholder_info: Dict[str, Any], new_content: str) -> bool:
        """在表格单元格中替换单个占位符，应用缓存的格式"""
        try:
            cell = placeholder_info['cell']
            placeholder_pattern = f"{{{placeholder_info['placeholder']}}}"
            cached_format = placeholder_info.get('cached_format', {})
            
            # 直接在单元格文本中替换
            current_text = cell.text
            if placeholder_pattern not in current_text:
                return False
            
            updated_text = current_text.replace(placeholder_pattern, new_content, 1)
            cell.text = updated_text
            
            # 应用格式到单元格的文本框
            if cached_format and hasattr(cell, 'text_frame') and cell.text_frame:
                self._apply_format_to_cell(cell, cached_format)
            
            return True
            
        except Exception as e:
            print(f"表格占位符替换失败: {e}")
            return False
    
    def _replace_single_placeholder_in_run(self, placeholder_info: Dict[str, Any], new_content: str) -> bool:
        """在run中替换单个占位符，保持run的格式"""
        try:
            run = placeholder_info['run']
            placeholder_pattern = f"{{{placeholder_info['placeholder']}}}"
            cached_format = placeholder_info.get('cached_format', {})
            
            # 在run文本中替换
            if placeholder_pattern not in run.text:
                return False
            
            run.text = run.text.replace(placeholder_pattern, new_content, 1)
            
            # 应用格式到run
            if cached_format:
                self._apply_format_to_run(run, cached_format)
            
            return True
            
        except Exception as e:
            print(f"Run级占位符替换失败: {e}")
            return False
    
    def _replace_single_placeholder_in_shape(self, placeholder_info: Dict[str, Any], new_content: str) -> bool:
        """在文本框形状中替换单个占位符，保持其他占位符不变"""
        try:
            shape = placeholder_info['shape']
            placeholder_pattern = f"{{{placeholder_info['placeholder']}}}"
            cached_format = placeholder_info.get('cached_format', {})
            
            # 直接在shape文本中替换
            current_text = shape.text if hasattr(shape, 'text') else ""
            if placeholder_pattern not in current_text:
                return False
            
            updated_text = current_text.replace(placeholder_pattern, new_content, 1)
            shape.text = updated_text
            
            # 应用格式到整个文本框（但只影响新替换的内容）
            if cached_format and hasattr(shape, 'text_frame') and shape.text_frame:
                self._apply_format_to_shape_text(shape, cached_format, new_content)
            
            return True
            
        except Exception as e:
            print(f"Shape级占位符替换失败: {e}")
            return False
    
    def _apply_format_to_cell(self, cell, format_info: Dict[str, Any]):
        """应用格式到表格单元格"""
        try:
            if hasattr(cell, 'text_frame') and cell.text_frame:
                for paragraph in cell.text_frame.paragraphs:
                    font = paragraph.font
                    if format_info.get('font_name'):
                        font.name = format_info['font_name']
                    if format_info.get('font_size'):
                        font.size = Pt(format_info['font_size'])
                    if format_info.get('font_bold') is not None:
                        font.bold = format_info['font_bold']
                    if format_info.get('font_italic') is not None:
                        font.italic = format_info['font_italic']
        except Exception as e:
            print(f"应用单元格格式失败: {e}")
    
    def _apply_format_to_run(self, run, format_info: Dict[str, Any]):
        """应用格式到run"""
        try:
            font = run.font
            if format_info.get('font_name'):
                font.name = format_info['font_name']
            if format_info.get('font_size'):
                font.size = Pt(format_info['font_size'])
            if format_info.get('font_bold') is not None:
                font.bold = format_info['font_bold']
            if format_info.get('font_italic') is not None:
                font.italic = format_info['font_italic']
        except Exception as e:
            print(f"应用run格式失败: {e}")
    
    def _apply_format_to_shape_text(self, shape, format_info: Dict[str, Any], new_content: str):
        """应用格式到文本框中的特定内容"""
        try:
            if hasattr(shape, 'text_frame') and shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    # 寻找包含新内容的runs并应用格式
                    for run in paragraph.runs:
                        if new_content in run.text:
                            font = run.font
                            if format_info.get('font_name'):
                                font.name = format_info['font_name']
                            if format_info.get('font_size'):
                                font.size = Pt(format_info['font_size'])
                            if format_info.get('font_bold') is not None:
                                font.bold = format_info['font_bold']
                            if format_info.get('font_italic') is not None:
                                font.italic = format_info['font_italic']
        except Exception as e:
            print(f"应用文本框格式失败: {e}")