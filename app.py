#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文本转PPT填充器 - Streamlit Web界面
使用DeepSeek AI将文本填入现有PPT文件
"""

import streamlit as st
import os
import tempfile
import io
from datetime import datetime
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
import json
import re

# 预设的PPT模板路径
PRESET_PPT_PATH = r"D:\jiayihan\Desktop\ppt format V1_2.pptx"

# 页面配置
st.set_page_config(
    page_title="文本转PPT填充器",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1ecf1;
        border: 1px solid #b6d4ea;
        color: #0c5460;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        color: #856404;
    }
    .error-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
    }
</style>
""", unsafe_allow_html=True)

class StreamlitPPTGenerator:
    def __init__(self, api_key):
        """初始化生成器"""
        self.api_key = api_key
        self.client = OpenAI(
            api_key=self.api_key,
            base_url="https://api.deepseek.com"
        )
        self.presentation = None
        self.ppt_structure = None
    
    def load_ppt_from_path(self, ppt_path):
        """从文件路径加载PPT"""
        try:
            if not os.path.exists(ppt_path):
                st.error(f"PPT模板文件不存在: {ppt_path}")
                return False
            
            self.presentation = Presentation(ppt_path)
            self.ppt_structure = self.analyze_existing_ppt()
            return True
        except Exception as e:
            st.error(f"加载PPT文件失败: {e}")
            return False
    
    def analyze_existing_ppt(self):
        """分析现有PPT的结构，特别关注占位符"""
        slides_info = []
        for i, slide in enumerate(self.presentation.slides):
            slide_info = {
                "slide_index": i,
                "title": "",
                "placeholders": {},  # 存储占位符信息
                "text_shapes": [],
                "has_content": False
            }
            
            # 分析幻灯片中的文本框和占位符
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    current_text = shape.text.strip()
                    if current_text:
                        # 检查是否包含占位符
                        import re
                        placeholder_pattern = r'\{([^}]+)\}'
                        placeholders = re.findall(placeholder_pattern, current_text)
                        
                        if placeholders:
                            # 这个文本框包含占位符
                            for placeholder in placeholders:
                                slide_info["placeholders"][placeholder] = {
                                    "shape": shape,
                                    "original_text": current_text,
                                    "placeholder": placeholder
                                }
                        
                        # 如果是简短文本且没有占位符，可能是标题
                        if not placeholders and len(current_text) < 100:
                            if slide_info["title"] == "":
                                slide_info["title"] = current_text
                        
                        slide_info["has_content"] = True
                    
                    # 记录所有可编辑的文本形状
                    if hasattr(shape, "text_frame"):
                        slide_info["text_shapes"].append({
                            "shape_id": shape.shape_id if hasattr(shape, "shape_id") else len(slide_info["text_shapes"]),
                            "current_text": shape.text,
                            "shape": shape,
                            "has_placeholder": bool(re.findall(r'\{([^}]+)\}', shape.text)) if shape.text else False
                        })
            
            slides_info.append(slide_info)
        
        return {
            "total_slides": len(self.presentation.slides),
            "slides": slides_info
        }
    
    def process_text_with_deepseek(self, user_text):
        """使用DeepSeek API分析如何将用户文本填入PPT模板的占位符"""
        # 创建现有PPT结构的描述，重点关注占位符
        ppt_description = f"现有PPT共有{self.ppt_structure['total_slides']}张幻灯片:\n"
        
        for slide in self.ppt_structure['slides']:
            ppt_description += f"\n第{slide['slide_index']+1}页:"
            if slide['title']:
                ppt_description += f" 标题「{slide['title']}」"
            
            # 列出所有占位符
            if slide['placeholders']:
                ppt_description += f"\n  包含占位符: "
                for placeholder_name, placeholder_info in slide['placeholders'].items():
                    ppt_description += f"{{{placeholder_name}}} "
                ppt_description += "\n"
            else:
                ppt_description += f" (无占位符)\n"
        
        system_prompt = f"""你是一个专业的PPT模板填充专家。我有一个包含占位符的PPT模板和用户提供的文本，请分析如何将用户文本精确填入对应的占位符位置。

现有PPT结构：
{ppt_description}

**占位符说明：**
- {{title}} = 主标题内容
- {{content}} = 主要内容/正文
- {{bullet_1}}, {{bullet_2}}, {{bullet_3}} = 要点列表
- {{subtitle}} = 副标题
- {{description}} = 描述性文字
- {{conclusion}} = 结论
- 其他 {{占位符}} = 根据名称推断用途

**重要原则：**
1. 仔细分析用户文本的结构和内容
2. 将文本内容精确匹配到合适的占位符
3. 保持用户原始文本内容完全不变
4. 优先填充已有的占位符，而不是创建新幻灯片

请按照以下JSON格式返回：
{{
  "assignments": [
    {{
      "slide_index": 0,
      "action": "replace_placeholder",
      "placeholder": "title",
      "content": "要填入该占位符的原始文本片段",
      "reason": "选择该占位符的原因"
    }},
    {{
      "slide_index": 1,
      "action": "replace_placeholder", 
      "placeholder": "content",
      "content": "要填入该占位符的原始文本片段",
      "reason": "选择该占位符的原因"
    }}
  ]
}}

分析要求：
1. 识别用户文本中的标题、内容、要点等部分
2. 将每部分内容匹配到最合适的占位符
3. action必须是"replace_placeholder"
4. placeholder必须是模板中实际存在的占位符名称
5. 提供清晰的匹配理由
6. 只返回JSON格式，不要其他文字"""
        
        try:
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_text}
                ],
                temperature=0.3,  # 降低温度以获得更精确的结果
                max_tokens=2000
            )
            
            content = response.choices[0].message.content
            if content:
                content = content.strip()
            else:
                content = ""
            
            # 提取JSON内容（如果有代码块包围）
            json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
            if json_match:
                content = json_match.group(1)
            
            try:
                return json.loads(content)
            except json.JSONDecodeError:
                st.error(f"AI返回的JSON格式有误，内容：{content}")
                # 返回基础分配方案
                return {
                    "assignments": [
                        {
                            "slide_index": 0,
                            "action": "replace_placeholder",
                            "placeholder": "content",
                            "content": user_text,
                            "reason": "JSON解析失败，默认填入content占位符"
                        }
                    ]
                }
        
        except Exception as e:
            st.error(f"调用DeepSeek API时出错: {e}")
            # 返回基础分配方案
            return {
                "assignments": [
                    {
                        "slide_index": 0,
                        "action": "replace_placeholder", 
                        "placeholder": "content",
                        "content": user_text,
                        "reason": f"API调用失败，默认填入content占位符。错误: {e}"
                    }
                ]
            }
    
    def apply_text_assignments(self, assignments):
        """根据分配方案替换PPT模板中的占位符"""
        assignments_list = assignments.get('assignments', [])
        results = []
        
        for assignment in assignments_list:
            action = assignment.get('action')
            content = assignment.get('content', '')
            slide_index = assignment.get('slide_index', 0)
            
            if action == 'replace_placeholder':
                placeholder = assignment.get('placeholder', '')
                if 0 <= slide_index < len(self.presentation.slides):
                    slide = self.presentation.slides[slide_index]
                    slide_info = self.ppt_structure['slides'][slide_index]
                    
                    # 检查该占位符是否存在
                    if placeholder in slide_info['placeholders']:
                        success = self.replace_placeholder_in_slide(
                            slide_info['placeholders'][placeholder], 
                            content
                        )
                        if success:
                            results.append(f"✓ 已替换第{slide_index+1}页的 {{{placeholder}}} 占位符: {assignment.get('reason', '')}")
                        else:
                            results.append(f"✗ 替换第{slide_index+1}页的 {{{placeholder}}} 占位符失败")
                    else:
                        results.append(f"✗ 第{slide_index+1}页不存在 {{{placeholder}}} 占位符")
                else:
                    results.append(f"✗ 幻灯片索引 {slide_index+1} 超出范围")
            
            elif action == 'update':  # 兼容旧的格式
                if 0 <= slide_index < len(self.presentation.slides):
                    slide = self.presentation.slides[slide_index]
                    self.update_slide_content(slide, content)
                    results.append(f"✓ 已更新第{slide_index+1}页: {assignment.get('reason', '')}")
                
            elif action == 'add_new':  # 兼容旧的格式
                title = assignment.get('title', '新增内容')
                self.add_new_slide(title, content)
                results.append(f"✓ 已新增幻灯片「{title}」: {assignment.get('reason', '')}")
        
        return results
    
    def replace_placeholder_in_slide(self, placeholder_info, new_content):
        """在特定的文本框中替换占位符"""
        try:
            shape = placeholder_info['shape']
            original_text = placeholder_info['original_text']
            placeholder_name = placeholder_info['placeholder']
            
            # 替换占位符
            updated_text = original_text.replace(f"{{{placeholder_name}}}", new_content)
            
            # 更新文本框内容
            if hasattr(shape, "text_frame") and shape.text_frame:
                tf = shape.text_frame
                tf.clear()
                
                # 添加新内容
                p = tf.paragraphs[0]
                p.text = updated_text
                
                # 保持字体大小
                if hasattr(p, 'font') and hasattr(p.font, 'size'):
                    if not p.font.size:
                        p.font.size = Pt(16)
            else:
                # 直接设置text属性
                shape.text = updated_text
            
            return True
        except Exception as e:
            st.error(f"替换占位符时出错: {e}")
            return False
    
    def update_slide_content(self, slide, content):
        """更新幻灯片内容"""
        # 查找可用的文本框
        text_shapes = []
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                text_shapes.append(shape)
        
        if text_shapes:
            # 使用最后一个可用的文本框（通常是主要内容区域）
            target_shape = text_shapes[-1] if len(text_shapes) > 1 else text_shapes[0]
            
            # 清空现有内容并添加新内容
            tf = target_shape.text_frame
            tf.clear()
            
            # 添加内容
            p = tf.paragraphs[0]
            p.text = content
            p.font.size = Pt(16)
    
    def add_new_slide(self, title, content):
        """添加新幻灯片"""
        # 使用标题和内容布局
        slide_layout = self.presentation.slide_layouts[1]
        slide = self.presentation.slides.add_slide(slide_layout)
        
        # 设置标题
        if slide.shapes.title:
            slide.shapes.title.text = title
        
        # 设置内容
        if len(slide.placeholders) > 1:
            content_placeholder = slide.placeholders[1]
            tf = content_placeholder.text_frame
            tf.clear()
            
            p = tf.paragraphs[0]
            p.text = content
            p.font.size = Pt(16)
    
    def get_ppt_bytes(self):
        """获取修改后的PPT字节数据"""
        # 创建output目录
        output_dir = "temp_output"
        os.makedirs(output_dir, exist_ok=True)
        
        # 保存到项目目录下的临时文件
        import time
        timestamp = str(int(time.time() * 1000))
        temp_filename = f"temp_ppt_{timestamp}.pptx"
        temp_filepath = os.path.join(output_dir, temp_filename)
        
        try:
            # 保存文件
            self.presentation.save(temp_filepath)
            
            # 读取字节数据
            with open(temp_filepath, 'rb') as f:
                ppt_bytes = f.read()
            
            return ppt_bytes
        finally:
            # 清理临时文件
            try:
                if os.path.exists(temp_filepath):
                    os.remove(temp_filepath)
            except Exception:
                pass  # 如果删除失败也没关系，只是临时文件

def main():
    # 页面标题
    st.markdown('<div class="main-header">📊 文本转PPT填充器</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">使用DeepSeek AI智能将您的文本填入预设PPT模板</div>', unsafe_allow_html=True)
    
    # 侧边栏配置
    with st.sidebar:
        st.header("⚙️ 配置")
        
        # API密钥输入
        api_key = st.text_input(
            "DeepSeek API密钥",
            type="password",
            help="请输入您的DeepSeek API密钥",
            placeholder="sk-..."
        )
        
        if not api_key:
            st.markdown('<div class="warning-box">⚠️ 请先输入API密钥才能使用功能</div>', unsafe_allow_html=True)
            st.markdown("获取API密钥：[DeepSeek平台](https://platform.deepseek.com/api_keys)")
        
        st.markdown("---")
        
        # 模板信息
        st.subheader("📄 PPT模板")
        st.markdown(f"**当前模板：** `{os.path.basename(PRESET_PPT_PATH)}`")
        st.markdown(f"**模板路径：** `{PRESET_PPT_PATH}`")
        
        # 检查模板文件状态
        if os.path.exists(PRESET_PPT_PATH):
            st.markdown('<div class="success-box">✅ 模板文件存在</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="error-box">❌ 模板文件不存在</div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        # 使用说明
        st.subheader("📖 使用说明")
        st.markdown("""
        1. 输入DeepSeek API密钥
        2. 确保PPT模板文件存在
        3. 输入要填入的文本内容
        4. 点击"开始处理"按钮
        5. 下载更新后的PPT文件
        """)
    
    # 主界面
    if api_key:
        # 检查模板文件
        if not os.path.exists(PRESET_PPT_PATH):
            st.markdown('<div class="error-box">❌ PPT模板文件不存在</div>', unsafe_allow_html=True)
            st.error(f"找不到模板文件: {PRESET_PPT_PATH}")
            st.info("请确保模板文件存在于指定路径")
            return
        
        # 初始化生成器
        generator = StreamlitPPTGenerator(api_key)
        
        # 加载PPT模板
        with st.spinner("正在加载PPT模板..."):
            if generator.load_ppt_from_path(PRESET_PPT_PATH):
                st.success("✅ PPT模板加载成功！")
                
                # 显示PPT信息
                ppt_info = generator.ppt_structure
                st.markdown('<div class="info-box">', unsafe_allow_html=True)
                st.markdown(f"**📊 PPT信息：** 共有 {ppt_info['total_slides']} 张幻灯片")
                
                # 显示幻灯片和占位符信息
                total_placeholders = 0
                for i, slide in enumerate(ppt_info['slides']):
                    title = slide['title'] if slide['title'] else "（无标题）"
                    placeholders = slide.get('placeholders', {})
                    total_placeholders += len(placeholders)
                    
                    if placeholders:
                        placeholder_list = ', '.join([f"{{{name}}}" for name in placeholders.keys()])
                        st.markdown(f"• 第{slide['slide_index']+1}页: {title} - 占位符: {placeholder_list}")
                    else:
                        st.markdown(f"• 第{slide['slide_index']+1}页: {title} - 无占位符")
                
                st.markdown(f"**🎯 总共找到 {total_placeholders} 个占位符**")
                st.markdown('</div>', unsafe_allow_html=True)
                
                st.markdown("---")
                
                # 文本输入
                st.subheader("✏️ 输入文本内容")
                user_text = st.text_area(
                    "请输入您想要填入PPT的文本内容",
                    height=200,
                    placeholder="请在这里输入您的文本内容...\n\n例如：\n人工智能技术的发展经历了多个重要阶段。从1950年代的符号主义开始，强调逻辑推理和知识表示...",
                    help="保持您的原文不变，AI会智能分析并填入合适的位置"
                )
                
                # 处理按钮
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    process_button = st.button(
                        "🚀 开始处理",
                        type="primary",
                        use_container_width=True,
                        disabled=not user_text.strip()
                    )
                
                # 处理文本
                if process_button and user_text.strip():
                    with st.spinner("正在使用DeepSeek AI分析文本结构..."):
                        assignments = generator.process_text_with_deepseek(user_text)
                    
                    with st.spinner("正在将文本填入PPT..."):
                        results = generator.apply_text_assignments(assignments)
                    
                    # 显示处理结果
                    st.markdown('<div class="success-box">', unsafe_allow_html=True)
                    st.markdown("**✅ 处理完成！**")
                    for result in results:
                        st.markdown(f"• {result}")
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # 提供下载
                    st.markdown("---")
                    st.subheader("💾 下载更新后的PPT")
                    
                    try:
                        updated_ppt_bytes = generator.get_ppt_bytes()
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        filename = f"updated_ppt_{timestamp}.pptx"
                        
                        st.download_button(
                            label="📥 下载更新后的PPT",
                            data=updated_ppt_bytes,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary",
                            use_container_width=True
                        )
                        
                        st.success(f"📁 文件准备完成：{filename}")
                        
                    except Exception as e:
                        st.error(f"生成下载文件时出错: {e}")
            
            else:
                st.error("❌ PPT模板加载失败")
    
    else:
        # 未输入API密钥时的提示
        st.info("👈 请在左侧输入您的DeepSeek API密钥开始使用")
        
        # 功能介绍
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 🎯 核心功能")
            st.markdown("""
            - **预设模板**：使用指定的PPT模板文件
            - **保持原文**：完全保留您的文本内容
            - **智能分析**：AI分析PPT结构和文本逻辑
            - **合理分配**：将文本填入最适合的位置
            - **灵活处理**：更新现有或新增幻灯片
            """)
        
        with col2:
            st.markdown("### 📝 适用场景")
            st.markdown("""
            - **学术报告**：研究内容填入模板
            - **商业计划**：项目信息填入格式
            - **教学课件**：课程内容填入框架
            - **工作汇报**：数据结果填入模板
            """)

if __name__ == "__main__":
    main() 