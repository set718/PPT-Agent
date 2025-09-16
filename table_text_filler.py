#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
表格文本填充器模块
专门用于处理PPT模板中表格占位符的智能填充
"""

import re
import json
from typing import Dict, List, Any, Optional, Tuple
from pptx import Presentation
from pptx.util import Pt
from utils import AIProcessor, PPTProcessor, FileManager
from logger import log_user_action, log_file_operation, LogContext


class TableTextFiller:
    """表格文本智能填充器"""
    
    def __init__(self, api_key: Optional[str] = None):
        """初始化表格文本填充器"""
        self.api_key = api_key
        self.ai_processor = AIProcessor(api_key)
        self.presentation = None
        self.ppt_processor = None
        self.ppt_structure = None
        self.table_info = {}  # 存储表格信息
        
    def load_ppt_template(self, ppt_path: str) -> Tuple[bool, str]:
        """加载PPT模板"""
        with LogContext(f"表格填充器加载PPT模板"):
            try:
                # 验证文件
                is_valid, error_msg = FileManager.validate_ppt_file(ppt_path)
                if not is_valid:
                    return False, error_msg
                
                self.presentation = Presentation(ppt_path)
                self.ppt_processor = PPTProcessor(self.presentation)
                self.ppt_structure = self.ppt_processor.ppt_structure
                
                # 分析表格结构
                self._analyze_table_structure()
                
                log_file_operation("load_ppt_table_filler", ppt_path, "success")
                return True, "模板加载成功"
                
            except Exception as e:
                log_file_operation("load_ppt_table_filler", ppt_path, "error", str(e))
                return False, f"加载失败: {str(e)}"
    
    def _analyze_table_structure(self):
        """分析PPT中所有表格的结构"""
        self.table_info = {}
        
        if not self.presentation:
            return
            
        for slide_idx, slide in enumerate(self.presentation.slides):
            slide_tables = []
            
            for shape_idx, shape in enumerate(slide.shapes):
                # 检查是否为表格
                if hasattr(shape, 'shape_type') and shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE = 19
                    table = shape.table
                    table_data = {
                        'slide_idx': slide_idx,
                        'shape_idx': shape_idx,
                        'rows': len(table.rows),
                        'cols': len(table.columns),
                        'placeholders': []
                    }
                    
                    # 分析表格中的占位符
                    for row_idx, row in enumerate(table.rows):
                        for col_idx, cell in enumerate(row.cells):
                            cell_text = cell.text.strip()
                            if cell_text:
                                placeholders = re.findall(r'\{([^}]+)\}', cell_text)
                                for placeholder in placeholders:
                                    table_data['placeholders'].append({
                                        'placeholder': placeholder,
                                        'row': row_idx,
                                        'col': col_idx,
                                        'cell_text': cell_text,
                                        'position': f"行{row_idx+1}列{col_idx+1}"
                                    })
                    
                    if table_data['placeholders']:  # 只记录包含占位符的表格
                        slide_tables.append(table_data)
            
            if slide_tables:
                self.table_info[slide_idx] = slide_tables
    
    def get_table_analysis(self) -> Dict[str, Any]:
        """获取表格分析结果"""
        analysis = {
            'total_tables': 0,
            'total_table_placeholders': 0,
            'slides_with_tables': [],
            'table_details': []
        }
        
        for slide_idx, tables in self.table_info.items():
            slide_info = {
                'slide_number': slide_idx + 1,
                'table_count': len(tables),
                'tables': []
            }
            
            for table_idx, table_data in enumerate(tables):
                table_detail = {
                    'table_index': table_idx + 1,
                    'size': f"{table_data['rows']}行{table_data['cols']}列",
                    'placeholder_count': len(table_data['placeholders']),
                    'placeholders': [
                        f"{{{p['placeholder']}}}"
                        for p in table_data['placeholders']
                    ]
                }
                slide_info['tables'].append(table_detail)
                analysis['total_table_placeholders'] += len(table_data['placeholders'])
            
            analysis['slides_with_tables'].append(slide_info)
            analysis['total_tables'] += len(tables)
        
        return analysis
    
    def process_table_text_with_ai(self, user_text: str) -> Dict[str, Any]:
        """使用AI分析文本并生成表格填充方案"""
        if not self.table_info:
            return {"assignments": []}
        
        # 构建专门针对表格填充的系统提示
        system_prompt = self._build_table_filling_prompt()
        
        # 获取表格结构信息
        table_structure_info = self._get_table_structure_for_ai()
        
        # 组合完整的分析提示
        full_prompt = f"{system_prompt}\n\n{table_structure_info}\n\n用户文本内容：\n{user_text}"
        
        log_user_action("表格文本AI分析", f"文本长度: {len(user_text)}字符, 表格数量: {sum(len(tables) for tables in self.table_info.values())}")
        
        try:
            # 调用AI进行分析，使用类似analyze_text_for_ppt的方式
            self.ai_processor._ensure_client()
            
            # 检查API类型
            model_info = self.ai_processor.config.get_model_info()
            
            if model_info.get('request_format') == 'dify_compatible':
                # 使用Liai API格式
                content = self.ai_processor._call_liai_api(system_prompt, user_text)
            else:
                # 使用OpenAI格式
                actual_model = model_info.get('actual_model', self.ai_processor.config.ai_model)
                
                response = self.ai_processor.client.chat.completions.create(
                    model=actual_model,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": f"表格结构信息：\n{table_structure_info}\n\n用户文本：\n{user_text}"}
                    ],
                    temperature=0.3,
                    stream=True
                )
                
                # 收集流式响应内容
                content = ""
                for chunk in response:
                    if chunk.choices and chunk.choices[0].delta and chunk.choices[0].delta.content:
                        content += chunk.choices[0].delta.content
                
                content = content.strip() if content else ""
            
            # 解析AI返回结果
            return self._parse_table_assignments(content)
            
        except Exception as e:
            log_user_action("表格文本AI分析失败", str(e))
            return {"error": f"AI分析失败: {str(e)}"}
    
    def _build_table_filling_prompt(self) -> str:
        """构建表格填充的AI提示"""
        return """你是一个专业的表格内容填充专家。你的任务是分析用户提供的文本内容，并智能地将这些内容分配到PPT模板的表格占位符中。

**核心任务：**
1. 理解用户文本的结构和内容
2. 识别文本中可以提取的数据、概念、分类等信息
3. 将这些信息合理地分配到表格的各个占位符位置

**表格填充原则：**
1. **逻辑匹配**：根据占位符名称和位置，推断应该填入什么类型的内容
2. **数据提取**：从用户文本中提取结构化信息（如：名称、数值、分类、描述等）
3. **合理分配**：确保表格内容的逻辑性和一致性
4. **完整利用**：尽可能利用用户提供的文本信息
5. **格式适配**：生成适合表格单元格的简洁内容

**特殊处理规则：**
- 表头占位符（如第一行）：填入分类、标题、字段名等
- 数据行占位符：填入具体的数据、案例、内容等
- 统计占位符：填入数量、百分比、总计等统计信息
- 描述占位符：填入详细说明、备注等

**输出格式：**
请返回JSON格式的分配方案：

```json
{
  "assignments": [
    {
      "placeholder": "占位符名称",
      "content": "填充内容",
      "slide_number": 1,
      "table_position": "行1列2",
      "reasoning": "分配理由"
    }
  ]
}
```"""
    
    def _get_table_structure_for_ai(self) -> str:
        """获取表格结构信息，供AI分析使用"""
        structure_info = "**当前模板中的表格结构：**\n\n"
        
        for slide_idx, tables in self.table_info.items():
            structure_info += f"**第{slide_idx + 1}页：**\n"
            
            for table_idx, table_data in enumerate(tables):
                structure_info += f"  表格{table_idx + 1}（{table_data['rows']}行{table_data['cols']}列）：\n"
                
                for placeholder_info in table_data['placeholders']:
                    placeholder = placeholder_info['placeholder']
                    position = placeholder_info['position']
                    structure_info += f"    - {{{placeholder}}} (位置: {position})\n"
                
                structure_info += "\n"
        
        return structure_info
    
    def _parse_table_assignments(self, ai_response: str) -> Dict[str, Any]:
        """解析AI返回的表格分配方案"""
        try:
            # 尝试提取JSON内容
            json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', ai_response, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                # 如果没有代码块，尝试直接解析
                json_str = ai_response.strip()
            
            result = json.loads(json_str)
            
            # 验证格式
            if 'assignments' in result and isinstance(result['assignments'], list):
                return result
            else:
                raise ValueError("AI返回格式不正确")
                
        except Exception as e:
            log_user_action("表格分配方案解析失败", str(e))
            return {
                "error": f"解析AI响应失败: {str(e)}",
                "assignments": []
            }
    
    def apply_table_assignments(self, assignments: Dict[str, Any]) -> Tuple[bool, List[str]]:
        """应用表格填充方案"""
        if not self.presentation or not assignments.get('assignments'):
            return False, ["没有可应用的分配方案"]
        
        results = []
        filled_count = 0
        
        try:
            for assignment in assignments['assignments']:
                placeholder = assignment.get('placeholder', '')
                content = assignment.get('content', '')
                slide_number = assignment.get('slide_number', 1) - 1  # 转换为0索引
                
                if not placeholder or not content:
                    continue
                
                # 在指定幻灯片的表格中查找并替换占位符
                success = self._fill_table_placeholder(slide_number, placeholder, content)
                if success:
                    filled_count += 1
                    results.append(f"✅ {{{placeholder}}} -> {content[:30]}...")
                else:
                    results.append(f"❌ {{{placeholder}}} 填充失败")
            
            # 填充完成后自动清理未使用的占位符
            if filled_count > 0:
                cleanup_result = self.cleanup_unfilled_table_placeholders()
                if cleanup_result.get('success'):
                    cleaned_count = cleanup_result.get('cleaned_count', 0)
                    if cleaned_count > 0:
                        results.append(f"🧹 自动清理了 {cleaned_count} 个未使用的占位符")
                
                results.append(f"\n总共成功填充 {filled_count} 个表格占位符")
                return True, results
            else:
                return False, ["没有成功填充任何占位符"]
                
        except Exception as e:
            log_user_action("应用表格分配方案失败", str(e))
            return False, [f"应用失败: {str(e)}"]
    
    def _fill_table_placeholder(self, slide_idx: int, placeholder: str, content: str) -> bool:
        """在指定幻灯片的表格中填充占位符"""
        try:
            if slide_idx >= len(self.presentation.slides):
                return False
            
            slide = self.presentation.slides[slide_idx]
            placeholder_pattern = f"{{{placeholder}}}"
            
            # 遍历该幻灯片中的所有表格
            for shape in slide.shapes:
                if hasattr(shape, 'shape_type') and shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE
                    table = shape.table
                    
                    for row in table.rows:
                        for cell in row.cells:
                            if placeholder_pattern in cell.text:
                                # 替换占位符，保持其他内容不变
                                cell.text = cell.text.replace(placeholder_pattern, content)
                                return True
            
            return False
            
        except Exception as e:
            log_user_action(f"填充表格占位符失败", f"{placeholder}: {str(e)}")
            return False
    
    def cleanup_unfilled_table_placeholders(self) -> Dict[str, Any]:
        """清理表格中未填充的占位符"""
        if not self.presentation:
            return {"error": "PPT未加载"}
        
        cleaned_count = 0
        cleaned_placeholders = []
        
        try:
            for slide_idx, slide in enumerate(self.presentation.slides):
                for shape in slide.shapes:
                    if hasattr(shape, 'shape_type') and shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE
                        table = shape.table
                        
                        for row_idx, row in enumerate(table.rows):
                            for col_idx, cell in enumerate(row.cells):
                                original_text = cell.text
                                
                                # 查找未填充的占位符
                                placeholders = re.findall(r'\{([^}]+)\}', original_text)
                                if placeholders:
                                    cleaned_text = original_text
                                    for placeholder in placeholders:
                                        pattern = f"{{{placeholder}}}"
                                        cleaned_text = cleaned_text.replace(pattern, "")
                                        cleaned_placeholders.append(
                                            f"第{slide_idx+1}页表格({row_idx+1},{col_idx+1}): {{{placeholder}}}"
                                        )
                                    
                                    # 清理多余空白
                                    cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()
                                    cell.text = cleaned_text
                                    cleaned_count += 1
            
            return {
                "success": True,
                "cleaned_count": len(cleaned_placeholders),
                "cleaned_placeholders": cleaned_placeholders,
                "message": f"清理了{len(cleaned_placeholders)}个表格占位符"
            }
            
        except Exception as e:
            return {"error": f"清理失败: {str(e)}"}
    
    def get_ppt_bytes(self) -> bytes:
        """获取处理后的PPT文件字节数据"""
        if not self.presentation:
            raise ValueError("PPT未加载")
        
        import io
        ppt_bytes = io.BytesIO()
        self.presentation.save(ppt_bytes)
        ppt_bytes.seek(0)
        return ppt_bytes.getvalue()


class TableTextProcessor:
    """表格文本处理器 - 负责文本解析和结构化"""
    
    @staticmethod
    def parse_table_data_from_text(text: str) -> Dict[str, Any]:
        """从文本中解析可能的表格数据"""
        # 尝试识别文本中的表格结构数据
        # 例如：列表、分类、键值对等
        
        # 按行分割
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        # 尝试识别不同的数据模式
        patterns = {
            'list_items': [],      # 列表项
            'key_values': [],      # 键值对
            'categories': [],      # 分类信息
            'numbers': [],         # 数值信息
            'names': [],           # 名称信息
            'descriptions': []     # 描述信息
        }
        
        for line in lines:
            # 检测列表项（如：- xxx, 1. xxx）
            if re.match(r'^[-•\d+\.]\s*', line):
                patterns['list_items'].append(re.sub(r'^[-•\d+\.]\s*', '', line))
            
            # 检测键值对（如：名称：值）
            elif ':' in line or '：' in line:
                parts = re.split(r'[:：]', line, 1)
                if len(parts) == 2:
                    patterns['key_values'].append({
                        'key': parts[0].strip(),
                        'value': parts[1].strip()
                    })
            
            # 检测数值信息
            elif re.search(r'\d+[%％]?', line):
                patterns['numbers'].append(line)
            
            # 其他内容作为描述
            else:
                patterns['descriptions'].append(line)
        
        return patterns
    
    @staticmethod
    def suggest_table_structure(data_patterns: Dict[str, Any], table_size: Tuple[int, int]) -> List[Dict[str, Any]]:
        """根据数据模式建议表格填充结构"""
        rows, cols = table_size
        suggestions = []
        
        # 如果有键值对数据，建议用于两列表格
        if data_patterns['key_values'] and cols >= 2:
            for i, kv in enumerate(data_patterns['key_values'][:rows]):
                suggestions.extend([
                    {'row': i, 'col': 0, 'content': kv['key'], 'type': 'key'},
                    {'row': i, 'col': 1, 'content': kv['value'], 'type': 'value'}
                ])
        
        # 如果有列表项，建议填充到单列或多列
        elif data_patterns['list_items']:
            items_per_col = (len(data_patterns['list_items']) + cols - 1) // cols
            for i, item in enumerate(data_patterns['list_items'][:rows * cols]):
                row = i // cols
                col = i % cols
                if row < rows:
                    suggestions.append({
                        'row': row, 'col': col, 'content': item, 'type': 'list_item'
                    })
        
        return suggestions