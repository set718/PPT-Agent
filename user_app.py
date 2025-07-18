#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文本转PPT填充器 - 用户版Web界面
使用DeepSeek AI将文本填入现有PPT文件
"""

import streamlit as st
import os
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
import json
import re
from config import get_config
from utils import AIProcessor, PPTProcessor, FileManager, PPTAnalyzer
from logger import get_logger, log_user_action, log_file_operation, LogContext

# 获取配置
config = get_config()
logger = get_logger()

# 页面配置
st.set_page_config(
    page_title="AI PPT助手",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #2E86AB;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.3rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1.5rem;
        border-radius: 0.8rem;
        background-color: #d4edda;
        border: 2px solid #c3e6cb;
        color: #155724;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1.5rem;
        border-radius: 0.8rem;
        background-color: #e8f4fd;
        border: 2px solid #bee5eb;
        color: #0c5460;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1.5rem;
        border-radius: 0.8rem;
        background-color: #fff3cd;
        border: 2px solid #ffeaa7;
        color: #856404;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1.5rem;
        border-radius: 0.8rem;
        background-color: #f8d7da;
        border: 2px solid #f5c6cb;
        color: #721c24;
        margin: 1rem 0;
    }
    .feature-box {
        padding: 1.5rem;
        border-radius: 0.8rem;
        background-color: #f8f9fa;
        border: 2px solid #e9ecef;
        margin: 1rem 0;
    }
    .steps-container {
        background-color: #f8f9fa;
        padding: 2rem;
        border-radius: 1rem;
        margin: 2rem 0;
    }
</style>
""", unsafe_allow_html=True)

class UserPPTGenerator:
    def __init__(self, api_key):
        """初始化生成器"""
        self.api_key = api_key
        self.ai_processor = AIProcessor(api_key)
        self.presentation = None
        self.ppt_processor = None
        self.ppt_structure = None
        logger.info(f"用户界面初始化PPT生成器")
    
    def load_ppt_from_path(self, ppt_path):
        """从文件路径加载PPT"""
        with LogContext(f"用户界面加载PPT文件"):
            try:
                # 验证文件
                is_valid, error_msg = FileManager.validate_ppt_file(ppt_path)
                if not is_valid:
                    return False, error_msg
                
                self.presentation = Presentation(ppt_path)
                self.ppt_processor = PPTProcessor(self.presentation)
                self.ppt_structure = self.ppt_processor.ppt_structure
                
                log_file_operation("load_ppt_user", ppt_path, "success")
                return True, "成功"
            except Exception as e:
                log_file_operation("load_ppt_user", ppt_path, "error", str(e))
                return False, str(e)
    
    def process_text_with_deepseek(self, user_text):
        """使用DeepSeek API分析如何将用户文本填入PPT模板的占位符"""
        if not self.ppt_structure:
            return {"assignments": []}
        
        log_user_action("用户界面AI文本分析", f"文本长度: {len(user_text)}字符")
        return self.ai_processor.analyze_text_for_ppt(user_text, self.ppt_structure)
    
    def apply_text_assignments(self, assignments):
        """根据分配方案替换PPT模板中的占位符"""
        if not self.presentation or not self.ppt_processor:
            return False, ["PPT文件未正确加载"]
        
        log_user_action("用户界面应用文本分配", f"分配数量: {len(assignments.get('assignments', []))}")
        results = self.ppt_processor.apply_assignments(assignments)
        
        # 美化演示文稿（静默执行）
        beautify_results = self.ppt_processor.beautify_presentation()
        
        return True, results
    
    def get_ppt_bytes(self):
        """获取修改后的PPT字节数据"""
        if not self.presentation:
            raise ValueError("PPT文件未正确加载")
        
        log_user_action("用户界面获取PPT字节数据")
        return FileManager.save_ppt_to_bytes(self.presentation)

def main():
    # 页面标题
    st.markdown('<div class="main-header">🎨 AI PPT助手</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">智能将您的文本内容转换为精美的PPT演示文稿</div>', unsafe_allow_html=True)
    
    # API密钥输入区域
    st.markdown("### 🔑 开始使用")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        api_key = st.text_input(
            "请输入您的DeepSeek API密钥",
            type="password",
            placeholder="sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
            help="API密钥用于AI文本分析，不会被保存"
        )
    with col2:
        st.markdown("**获取API密钥**")
        st.markdown("[🔗 DeepSeek平台](https://platform.deepseek.com/api_keys)")
    
    # 检查API密钥
    if not api_key or not api_key.strip():
        # 显示功能介绍
        st.markdown("---")
        
        # 使用步骤
        st.markdown('<div class="steps-container">', unsafe_allow_html=True)
        st.markdown("### 📝 三步轻松制作PPT")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("""
            **第一步：准备API密钥** 🔑
            - 注册DeepSeek账号
            - 获取API密钥
            - 在上方输入密钥
            """)
        
        with col2:
            st.markdown("""
            **第二步：输入内容** ✏️
            - 粘贴您的文本内容
            - 可以是任何主题
            - 无需特殊格式
            """)
        
        with col3:
            st.markdown("""
            **第三步：生成下载** 🚀
            - 点击开始处理
            - 等待AI智能分析
            - 下载精美PPT
            """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 功能特色
        st.markdown("### ✨ 产品特色")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown('<div class="feature-box">', unsafe_allow_html=True)
            st.markdown("""
            **🤖 AI智能分析**
            - 自动理解文本结构
            - 智能匹配PPT模板
            - 保持内容完整性
            """)
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown('<div class="feature-box">', unsafe_allow_html=True)
            st.markdown("""
            **🎨 专业美化**
            - 自动优化布局
            - 清理多余元素
            - 统一设计风格
            """)
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="feature-box">', unsafe_allow_html=True)
            st.markdown("""
            **⚡ 快速高效**
            - 一键生成PPT
            - 无需手动排版
            - 节省大量时间
            """)
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown('<div class="feature-box">', unsafe_allow_html=True)
            st.markdown("""
            **📱 简单易用**
            - 界面简洁明了
            - 操作步骤清晰
            - 适合所有用户
            """)
            st.markdown('</div>', unsafe_allow_html=True)
        
        # 适用场景
        st.markdown("### 🎯 适用场景")
        
        scenario_col1, scenario_col2, scenario_col3, scenario_col4 = st.columns(4)
        with scenario_col1:
            st.markdown("**📚 学术报告**\n研究成果展示")
        with scenario_col2:
            st.markdown("**💼 商业提案**\n项目方案介绍")
        with scenario_col3:
            st.markdown("**🎓 教学课件**\n课程内容整理")
        with scenario_col4:
            st.markdown("**📊 工作汇报**\n数据结果展示")
        
        return
    
    # 验证API密钥格式
    if not api_key.startswith('sk-'):
        st.markdown('<div class="warning-box">⚠️ API密钥格式可能不正确，请确认是否以"sk-"开头</div>', unsafe_allow_html=True)
        return
    
    # 检查模板文件
    is_valid, error_msg = FileManager.validate_ppt_file(config.default_ppt_template)
    if not is_valid:
        st.markdown('<div class="error-box">❌ 系统模板文件暂时不可用，请稍后再试</div>', unsafe_allow_html=True)
        return
    
    # 初始化生成器并加载模板
    generator = UserPPTGenerator(api_key)
    
    with st.spinner("正在准备AI助手..."):
        success, message = generator.load_ppt_from_path(config.default_ppt_template)
        
    if not success:
        st.markdown('<div class="error-box">❌ 系统初始化失败，请稍后再试</div>', unsafe_allow_html=True)
        return
    
    st.markdown('<div class="success-box">✅ AI助手已准备就绪！</div>', unsafe_allow_html=True)
    
    # 主要功能区域
    st.markdown("---")
    st.markdown("### 📝 输入您的内容")
    
    # 文本输入
    user_text = st.text_area(
        "请输入您想要制作成PPT的文本内容：",
        height=250,
        placeholder="""例如：

人工智能的发展历程

人工智能技术的发展经历了多个重要阶段。从1950年代的符号主义开始，强调逻辑推理和知识表示，到1980年代的专家系统兴起，再到近年来深度学习的突破性进展。

当前，大语言模型如GPT、Claude等展现出了前所未有的能力，能够进行复杂的文本理解、生成和推理。这些技术正在革新各个行业，从教育、医疗到金融、娱乐，都能看到AI的身影。

未来，人工智能将继续向更加智能化、人性化的方向发展，为人类社会带来更多便利和创新可能性。""",
        help="请输入您的完整内容，AI会自动分析并合理分配到PPT的各个部分"
    )
    
    # 字数统计
    if user_text:
        char_count = len(user_text)
        word_count = len(user_text.split())
        st.caption(f"📊 字符数：{char_count} | 词数：{word_count}")
    
    # 处理按钮
    st.markdown("### 🚀 生成PPT")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        process_button = st.button(
            "🎨 开始制作PPT",
            type="primary",
            use_container_width=True,
            disabled=not user_text.strip()
        )
    
    # 处理逻辑
    if process_button and user_text.strip():
        # 显示处理进度
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # 步骤1：AI分析
            status_text.text("🤖 AI正在分析您的内容...")
            progress_bar.progress(25)
            
            assignments = generator.process_text_with_deepseek(user_text)
            
            # 步骤2：填充PPT
            status_text.text("📝 正在将内容填入PPT模板...")
            progress_bar.progress(50)
            
            success, results = generator.apply_text_assignments(assignments)
            
            if not success:
                st.error("处理过程中出现错误，请重试")
                return
            
            # 步骤3：美化优化
            status_text.text("🎨 正在美化PPT布局...")
            progress_bar.progress(75)
            
            # 步骤4：准备下载
            status_text.text("📦 正在准备下载文件...")
            progress_bar.progress(100)
            
            # 清除进度显示
            progress_bar.empty()
            status_text.empty()
            
            # 显示成功信息
            st.markdown('<div class="success-box">', unsafe_allow_html=True)
            st.markdown("**🎉 PPT制作完成！**")
            st.markdown("您的内容已成功转换为精美的PPT演示文稿")
            st.markdown('</div>', unsafe_allow_html=True)
            
            # 提供下载
            st.markdown("### 💾 下载您的PPT")
            
            try:
                updated_ppt_bytes = generator.get_ppt_bytes()
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"AI生成PPT_{timestamp}.pptx"
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.download_button(
                        label="📥 立即下载PPT",
                        data=updated_ppt_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True
                    )
                
                st.markdown('<div class="info-box">', unsafe_allow_html=True)
                st.markdown(f"📁 **文件名：** {filename}")
                st.markdown("📋 **温馨提示：** 下载后您可以继续在PowerPoint中编辑和完善")
                st.markdown('</div>', unsafe_allow_html=True)
                
            except Exception as e:
                st.error("文件准备失败，请重试")
                logger.error(f"用户界面文件生成错误: {e}")
                
        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error("处理过程中出现错误，请检查您的API密钥或稍后重试")
            logger.error(f"用户界面处理错误: {e}")
    
    # 页脚信息
    st.markdown("---")
    st.markdown(
        '<div style="text-align: center; color: #666; padding: 2rem;">'
        '💡 由DeepSeek AI驱动 | 🎨 专业PPT自动生成'
        '</div>', 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()