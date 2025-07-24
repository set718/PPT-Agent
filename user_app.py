#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文本转PPT填充器 - 用户版Web界面
使用OpenAI GPT-4V将文本填入现有PPT文件
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
    
    def process_text_with_openai(self, user_text):
        """使用OpenAI API分析如何将用户文本填入PPT模板的占位符"""
        if not self.ppt_structure:
            return {"assignments": []}
        
        log_user_action("用户界面AI文本分析", f"文本长度: {len(user_text)}字符")
        return self.ai_processor.analyze_text_for_ppt(user_text, self.ppt_structure)
    
    def apply_text_assignments(self, assignments, user_text: str = ""):
        """根据分配方案替换PPT模板中的占位符，并将原始文本添加到备注"""
        if not self.presentation or not self.ppt_processor:
            return False, ["PPT文件未正确加载"]
        
        log_user_action("用户界面应用文本分配", f"分配数量: {len(assignments.get('assignments', []))}")
        # 传递用户原始文本，用于添加到幻灯片备注
        results = self.ppt_processor.apply_assignments(assignments, user_text)
        
        # 文本填充完成，不立即美化
        return True, results
    
    def cleanup_unfilled_placeholders(self):
        """清理未填充的占位符"""
        if not self.ppt_processor:
            return {"error": "PPT处理器未初始化"}
        
        try:
            log_user_action("用户界面清理占位符", f"已填充: {len(self.ppt_processor.filled_placeholders)}")
            
            # 手动清理占位符
            cleanup_count = 0
            for slide_idx, slide in enumerate(self.presentation.slides):
                for shape in slide.shapes:
                    if hasattr(shape, 'text') and shape.text:
                        original_text = shape.text
                        # 移除所有剩余的占位符模式 {xxx}
                        import re
                        cleaned_text = re.sub(r'\{[^}]+\}', '', original_text)
                        if cleaned_text != original_text:
                            shape.text = cleaned_text.strip()
                            cleanup_count += 1
            
            return {
                "success": True,
                "cleaned_placeholders": cleanup_count,
                "message": f"清理了{cleanup_count}个未填充的占位符"
            }
            
        except Exception as e:
            log_user_action("用户界面清理占位符失败", str(e))
            return {"error": f"清理占位符失败: {e}"}
    
    def apply_basic_beautification(self):
        """应用基础美化（不包含视觉分析）"""
        if not self.ppt_processor:
            return {"error": "PPT处理器未初始化"}
        
        try:
            log_user_action("用户界面基础美化")
            # 只进行基础的美化处理，不启用视觉优化
            beautify_results = self.ppt_processor.beautify_presentation(
                enable_visual_optimization=False
            )
            
            return beautify_results
            
        except Exception as e:
            log_user_action("用户界面基础美化失败", str(e))
            return {"error": f"基础美化失败: {e}"}
    
    def apply_visual_optimization(self, ppt_path: str, enable_visual_optimization: bool = True):
        """
        应用视觉优化
        
        Args:
            ppt_path: PPT文件路径
            enable_visual_optimization: 是否启用视觉优化
            
        Returns:
            Dict: 优化结果
        """
        if not self.ppt_processor:
            return {"error": "PPT处理器未初始化"}
        
        try:
            # 初始化视觉分析器
            if enable_visual_optimization:
                success = self.ppt_processor.initialize_visual_analyzer(self.api_key)
                if not success:
                    return {"error": "视觉分析器初始化失败"}
            
            # 执行美化，包含视觉优化
            log_user_action("用户界面视觉优化", f"启用状态: {enable_visual_optimization}")
            beautify_results = self.ppt_processor.beautify_presentation(
                enable_visual_optimization=enable_visual_optimization,
                ppt_path=ppt_path if enable_visual_optimization else None
            )
            
            return beautify_results
            
        except Exception as e:
            log_user_action("用户界面视觉优化失败", str(e))
            return {"error": f"视觉优化失败: {e}"}
    
    def get_ppt_bytes(self):
        """获取修改后的PPT字节数据"""
        if not self.presentation:
            raise ValueError("PPT文件未正确加载")
        
        log_user_action("用户界面获取PPT字节数据")
        return FileManager.save_ppt_to_bytes(self.presentation)

def display_processing_summary(optimization_results, cleanup_results, enable_visual_optimization):
    """显示处理结果摘要"""
    if not optimization_results or "error" in optimization_results:
        if "error" in optimization_results:
            st.warning(f"⚠️ 处理过程中出现问题: {optimization_results['error']}")
        return
    
    summary = optimization_results.get('summary', {})
    
    # 基础处理信息
    st.markdown("### 📊 处理结果摘要")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        final_slide_count = summary.get('final_slide_count', 1)  # 默认至少1页
        st.metric("📑 最终页数", final_slide_count)
    
    with col2:
        # 显示手动清理的占位符数量
        cleanup_count = cleanup_results.get('cleaned_placeholders', 0) if cleanup_results else 0
        st.metric("🧹 清理占位符", cleanup_count)
    
    with col3:
        removed_empty_slides = summary.get('removed_empty_slides_count', 0)
        st.metric("🗑️ 删除空页", removed_empty_slides)
    
    with col4:
        reorganized_slides = summary.get('reorganized_slides_count', 0)
        st.metric("🔄 重新排版", reorganized_slides)
    
    # 视觉优化结果（如果启用）
    if enable_visual_optimization:
        visual_analysis = optimization_results.get('visual_analysis')
        visual_optimization = optimization_results.get('visual_optimization')
        
        if visual_analysis and "error" not in visual_analysis:
            st.markdown("### 🎨 视觉质量分析")
            
            overall_analysis = visual_analysis.get('overall_analysis', {})
            visual_score = overall_analysis.get('weighted_score', 0)
            grade = overall_analysis.get('grade', '未知')
            
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.metric("🏆 视觉质量评分", f"{visual_score:.1f}/10", grade)
                
                if visual_optimization and visual_optimization.get('success'):
                    optimizations_applied = visual_optimization.get('total_optimizations', 0)
                    st.metric("🔧 应用优化", f"{optimizations_applied}项")
            
            with col2:
                # 显示评分详情
                scores = overall_analysis.get('scores', {})
                if scores:
                    st.markdown("**各项评分详情:**")
                    score_descriptions = {
                        "layout_balance": "布局平衡度",
                        "color_harmony": "色彩协调性", 
                        "typography": "字体排版",
                        "visual_hierarchy": "视觉层次",
                        "white_space": "留白使用",
                        "overall_aesthetics": "整体美观度"
                    }
                    
                    for criterion, score in scores.items():
                        if criterion in score_descriptions:
                            desc = score_descriptions[criterion]
                            progress_value = min(score / 10.0, 1.0)
                            st.progress(progress_value, f"{desc}: {score:.1f}/10")
            
            # 显示改进建议
            strengths = overall_analysis.get('strengths', [])
            weaknesses = overall_analysis.get('weaknesses', [])
            
            if strengths or weaknesses:
                with st.expander("📋 详细分析结果", expanded=False):
                    if strengths:
                        st.markdown("**✅ 设计优点:**")
                        for strength in strengths[:3]:
                            st.markdown(f"• {strength}")
                    
                    if weaknesses:
                        st.markdown("**⚠️ 待改进点:**")
                        for weakness in weaknesses[:3]:
                            st.markdown(f"• {weakness}")
        
        elif visual_analysis and "error" in visual_analysis:
            st.warning(f"🔍 视觉分析遇到问题: {visual_analysis['error']}")
    
    else:
        st.info("💡 提示：启用AI视觉优化可获得更详细的美观度分析和自动布局优化")

def main():
    # 页面标题
    st.markdown('<div class="main-header">🎨 AI PPT助手</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">智能将您的文本内容转换为精美的PPT演示文稿</div>', unsafe_allow_html=True)
    
    # 模型选择区域
    st.markdown("### 🤖 选择AI模型")
    
    available_models = config.available_models
    model_options = {}
    for model_key, model_info in available_models.items():
        display_name = f"{model_info['name']} ({model_info['cost']}成本)"
        if not model_info['supports_vision']:
            display_name += " - ⚠️ 无视觉分析"
        model_options[display_name] = model_key
    
    model_col1, model_col2 = st.columns([2, 1])
    with model_col1:
        selected_display = st.selectbox(
            "选择适合您需求的AI模型：",
            options=list(model_options.keys()),
            index=0,
            help="不同模型有不同的功能和成本特点"
        )
        
        selected_model = model_options[selected_display]
        model_info = available_models[selected_model]
        
        # 动态更新配置
        if selected_model != config.ai_model:
            config.set_model(selected_model)
    
    with model_col2:
        st.markdown("**模型对比**")
        if model_info['supports_vision']:
            st.success("✅ 支持视觉分析\n✅ 效果更佳\n💰 成本较高")
        else:
            st.info("⚡ 响应更快\n💸 成本更低\n❌ 无视觉分析")
    
    # 显示当前选择的模型信息
    if model_info['supports_vision']:
        st.markdown('<div class="success-box">🎨 已选择具备视觉分析功能的模型，将为您提供最佳的PPT美化效果</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="info-box">⚡ 已选择高效文本处理模型，将专注于内容智能分配，节省您的成本</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # API密钥输入区域
    st.markdown("### 🔑 开始使用")
    
    # 根据选择的模型动态显示API密钥输入信息
    current_model_info = config.get_model_info()
    api_provider = current_model_info.get('api_provider', 'OpenRouter')
    api_key_url = current_model_info.get('api_key_url', 'https://openrouter.ai/keys')
    
    col1, col2 = st.columns([2, 1])
    with col1:
        if api_provider == "OpenRouter":
            placeholder_text = "sk-or-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
            help_text = "通过OpenRouter访问AI模型，API密钥不会被保存"
        else:  # DeepSeek
            placeholder_text = "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
            help_text = f"通过{api_provider}平台访问AI模型，API密钥不会被保存"
            
        api_key = st.text_input(
            f"请输入您的{api_provider} API密钥",
            type="password",
            placeholder=placeholder_text,
            help=help_text
        )
    with col2:
        st.markdown("**获取API密钥**")
        st.markdown(f"[🔗 {api_provider}平台]({api_key_url})")
        
        # API密钥测试按钮
        if api_key and api_key.strip():
            if st.button("🔍 测试API密钥", help="快速验证密钥是否有效"):
                with st.spinner("正在验证API密钥..."):
                    try:
                        # 创建一个临时的AIProcessor来测试
                        test_processor = AIProcessor(api_key.strip())
                        test_processor._ensure_client()
                        st.success("✅ API密钥验证通过！")
                    except ValueError as e:
                        st.error(f"❌ API密钥验证失败: {str(e)}")
                    except Exception as e:
                        error_msg = str(e)
                        if "authentication" in error_msg.lower() or "unauthorized" in error_msg.lower():
                            st.error("❌ API密钥认证失败，请检查密钥是否正确")
                        elif "network" in error_msg.lower() or "connection" in error_msg.lower():
                            st.error("❌ 网络连接异常，请检查网络连接")
                        else:
                            st.error("❌ 验证过程出现异常")
                        st.error(f"详细错误: {error_msg}")
    
    # 检查API密钥
    if not api_key or not api_key.strip():
        # 显示功能介绍
        st.markdown("---")
        
        # 使用步骤
        st.markdown('<div class="steps-container">', unsafe_allow_html=True)
        st.markdown("### 📝 四步轻松制作PPT")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown("""
            **第一步：选择模型** 🤖
            - GPT-4o：功能完整，支持视觉分析
            - DeepSeek R1：成本更低，专注推理处理
            """)
        
        with col2:
            st.markdown("""
            **第二步：准备API密钥** 🔑
            - 根据选择的模型注册相应平台账号
            - OpenRouter/DeepSeek获取API密钥
            - 在上方输入密钥
            """)
        
        with col3:
            st.markdown("""
            **第三步：输入内容** ✏️
            - 粘贴您的文本内容
            - 可以是任何主题
            - 无需特殊格式
            """)
        
        with col4:
            st.markdown("""
            **第四步：生成下载** 🚀
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
    
    # 验证API密钥格式（根据选择的API提供商）
    if api_provider == "OpenRouter":
        if not (api_key.startswith('sk-or-') or api_key.startswith('sk-')):
            st.markdown('<div class="warning-box">⚠️ OpenRouter API密钥格式可能不正确，通常以"sk-or-"开头</div>', unsafe_allow_html=True)
            return
    elif api_provider == "DeepSeek":
        if not api_key.startswith('sk-'):
            st.markdown('<div class="warning-box">⚠️ DeepSeek API密钥格式可能不正确，请检查密钥格式</div>', unsafe_allow_html=True)
            return
    
    # 检查模板文件
    is_valid, error_msg = FileManager.validate_ppt_file(config.default_ppt_template)
    if not is_valid:
        st.markdown('<div class="error-box">❌ 系统模板文件暂时不可用，请稍后再试</div>', unsafe_allow_html=True)
        return
    
    # 初始化生成器并加载模板
    try:
        with st.spinner("正在验证API密钥..."):
            generator = UserPPTGenerator(api_key)
        
        with st.spinner("正在准备AI助手..."):
            success, message = generator.load_ppt_from_path(config.default_ppt_template)
            
        if not success:
            st.markdown('<div class="error-box">❌ 系统初始化失败，请稍后再试</div>', unsafe_allow_html=True)
            return
            
    except ValueError as e:
        if "API密钥" in str(e):
            st.markdown('<div class="error-box">❌ API密钥验证失败，请检查密钥是否正确</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="error-box">❌ 初始化失败: {str(e)}</div>', unsafe_allow_html=True)
        return
    except Exception as e:
        error_msg = str(e)
        if "authentication" in error_msg.lower() or "unauthorized" in error_msg.lower():
            st.markdown('<div class="error-box">❌ API密钥认证失败，请检查密钥是否正确或是否有足够余额</div>', unsafe_allow_html=True)
        elif "network" in error_msg.lower() or "connection" in error_msg.lower():
            st.markdown('<div class="error-box">❌ 网络连接异常，请检查网络连接后重试</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="error-box">❌ 系统初始化异常，请稍后重试</div>', unsafe_allow_html=True)
        st.error(f"详细错误: {error_msg}")
        return
    
    st.markdown('<div class="success-box">✅ AI助手已准备就绪！</div>', unsafe_allow_html=True)
    
    # 功能选择选项卡
    st.markdown("---")
    tab1, tab2 = st.tabs(["🎨 智能PPT生成", "📑 AI智能分页（预览）"])
    
    with tab1:
        # 现有的PPT生成功能
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
    
    # 高级选项（根据模型能力动态显示）
    st.markdown("### ⚙️ 高级选项")
    
    current_model_info = config.get_model_info()
    supports_vision = current_model_info.get('supports_vision', False)
    
    col1, col2 = st.columns(2)
    with col1:
        if supports_vision:
            enable_visual_optimization = st.checkbox(
                "🎨 启用AI视觉优化",
                value=False,
                help=f"使用{current_model_info['name']}分析PPT视觉效果并自动优化布局（需要额外时间）"
            )
        else:
            enable_visual_optimization = False
            st.info(f"⚠️ 当前模型 {current_model_info['name']} 不支持视觉优化功能")
    
    with col2:
        if supports_vision:
            if enable_visual_optimization:
                st.info("🔍 视觉优化将分析每页PPT的美观度并自动调整布局")
            else:
                st.info("✨ 只进行基础美化处理")
        else:
            st.info("🚀 将进行高效的文本内容分配和基础美化")
    
    # 处理按钮
    st.markdown("### 🚀 生成PPT")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # 根据模型和选项动态生成按钮文本
        if supports_vision and enable_visual_optimization:
            button_text = f"🎨 开始制作PPT (含{current_model_info['name']}视觉优化)"
        elif supports_vision:
            button_text = f"🎨 开始制作PPT (使用{current_model_info['name']})"
        else:
            button_text = f"⚡ 开始制作PPT (使用{current_model_info['name']})"
        
        process_button = st.button(
            button_text,
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
            progress_bar.progress(20)
            
            try:
                assignments = generator.process_text_with_openai(user_text)
            except ValueError as e:
                if "API密钥" in str(e):
                    st.error("❌ API密钥验证失败，请检查密钥是否正确")
                else:
                    st.error(f"❌ AI分析失败: {str(e)}")
                return
            except Exception as e:
                error_msg = str(e)
                if "rate limit" in error_msg.lower():
                    st.error("❌ API请求频率超限，请稍后重试")
                elif "insufficient" in error_msg.lower() or "quota" in error_msg.lower():
                    st.error("❌ API额度不足，请检查账户余额")
                else:
                    st.error("❌ AI分析过程出现异常，请稍后重试")
                st.error(f"详细错误: {error_msg}")
                return
            
            # 步骤2：填充PPT
            status_text.text("📝 正在将内容填入PPT模板...")
            progress_bar.progress(40)
            
            success, results = generator.apply_text_assignments(assignments, user_text)
            
            if not success:
                st.error("处理过程中出现错误，请重试")
                return
            
            # 步骤3：清理未填充的占位符
            status_text.text("🧹 正在清理未使用的占位符...")
            progress_bar.progress(60)
            
            # 手动清理未填充的占位符
            cleanup_results = generator.cleanup_unfilled_placeholders()
            
            # 步骤4：视觉优化（如果启用）
            if enable_visual_optimization:
                status_text.text("🔍 正在进行AI视觉分析...")
                progress_bar.progress(70)
                
                # 应用视觉优化
                optimization_results = generator.apply_visual_optimization(
                    config.default_ppt_template, 
                    enable_visual_optimization
                )
                
                status_text.text("🎨 正在应用视觉优化建议...")
                progress_bar.progress(80)
            else:
                status_text.text("🎨 正在进行基础美化...")
                progress_bar.progress(70)
                # 只进行基础美化，不包含视觉分析
                optimization_results = generator.apply_basic_beautification()
            
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
            
            # 显示处理结果摘要
            display_processing_summary(optimization_results, cleanup_results, enable_visual_optimization)
            
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
                logger.error("用户界面文件生成错误: %s", str(e))
                
        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error("处理过程中出现错误，请检查您的API密钥或稍后重试")
            logger.error("用户界面处理错误: %s", str(e))
    
    with tab2:
        # AI智能分页功能（新功能预览）
        st.markdown("### 🚀 AI智能分页功能")
        
        st.markdown('<div class="info-box">📢 <strong>功能预览</strong><br>这是即将推出的完整工作流程的第一步。完整流程包括：用户输入 → AI智能分页 → Dify API模板匹配 → 文本填充 → 整合输出<br><br>🎨 <strong>模板库扩展：</strong>目前支持单页模板，后续将提供200+丰富模板，支持各种风格和场景的PPT制作<br><br>📋 <strong>分页规范：</strong>标题页仅提取标题和日期（其他内容固定），结尾页使用预设模板（无需生成），重点关注中间内容页的智能分割</div>', unsafe_allow_html=True)
        
        # 文本输入区域
        st.markdown("#### 📝 输入要分页的文本内容")
        
        page_split_text = st.text_area(
            "请输入您想要进行智能分页的文本内容：",
            height=200,
            placeholder="""例如：

区块链技术发展报告

区块链技术作为一种分布式账本技术，近年来得到了广泛关注和应用。它通过去中心化的方式，为数字资产交易和数据存储提供了新的解决方案。

技术原理方面，区块链采用加密哈希、数字签名和共识机制等核心技术，确保数据的不可篡改性和系统的安全性。每个区块包含若干交易记录，通过链式结构连接形成完整的交易历史。

应用场景非常广泛，包括数字货币、供应链管理、身份认证、智能合约等领域。比特币是最早的区块链应用，展示了这项技术的巨大潜力。

未来发展趋势显示，区块链技术将向着更高的性能、更好的可扩展性和更广泛的应用场景发展。技术标准化、监管政策的完善也将推动整个行业的健康发展。""",
            help="AI将分析文本结构，智能分割为适合PPT展示的多个页面",
            key="page_split_text"
        )
        
        # 分页选项和建议
        col1, col2 = st.columns(2)
        with col1:
            target_pages = st.number_input(
                "目标页面数量（可选）",
                min_value=0,
                max_value=30,
                value=0,
                help="设置为0时，AI将自动判断最佳页面数量。建议根据演示时间控制页数。"
            )
            
            # 添加页数建议提示
            st.markdown("""
                         <div style="background-color: #f0f2f6; padding: 0.5rem; border-radius: 0.25rem; margin-top: 0.5rem;">
             <small>💡 <strong>页数建议：</strong><br>
             • 5分钟演示：3-5页（含标题页）<br>
             • 10分钟演示：5-8页（含标题页）<br>
             • 15分钟演示：8-12页（含标题页）<br>
             • 30分钟演示：15-20页（含标题页）<br>
             • 学术报告：20-30页（含标题页）<br>
             <strong>注：</strong>结尾页使用固定模板，无需计入</small>
             </div>
            """, unsafe_allow_html=True)
        
        with col2:
            if page_split_text:
                char_count = len(page_split_text)
                word_count = len(page_split_text.split())
                st.metric("📊 文本统计", f"{char_count}字符 | {word_count}词")
        
        # 分页处理按钮
        split_button = st.button(
            "🤖 开始AI智能分页",
            type="primary",
            use_container_width=True,
            disabled=not page_split_text.strip(),
            help="AI将分析您的文本结构并智能分页"
        )
        
        # 处理AI分页逻辑
        if split_button and page_split_text.strip():
            from ai_page_splitter import AIPageSplitter, PageContentFormatter
            
            try:
                with st.spinner("🤖 AI正在分析文本结构并进行智能分页..."):
                    # 初始化AI分页器
                    page_splitter = AIPageSplitter(api_key)
                    
                    # 执行智能分页
                    target_page_count = int(target_pages) if target_pages > 0 else None
                    split_result = page_splitter.split_text_to_pages(page_split_text, target_page_count)
                
                if split_result.get('success'):
                    st.markdown('<div class="success-box">✅ AI智能分页完成！</div>', unsafe_allow_html=True)
                    
                    # 显示分析摘要
                    analysis = split_result.get('analysis', {})
                    analysis_summary = PageContentFormatter.format_analysis_summary(analysis)
                    st.markdown(analysis_summary)
                    
                    # 显示分页结果
                    pages = split_result.get('pages', [])
                    
                    if pages:
                        st.markdown("### 📄 分页预览结果")
                        
                        # 使用选项卡显示每一页
                        page_tabs = st.tabs([f"第{page['page_number']}页" for page in pages])
                        
                        for i, (page_tab, page_data) in enumerate(zip(page_tabs, pages)):
                            with page_tab:
                                page_preview = PageContentFormatter.format_page_preview(page_data)
                                st.markdown(page_preview)
                                
                                # 显示原始文本片段
                                with st.expander("📖 查看原始文本片段", expanded=False):
                                    original_segment = page_data.get('original_text_segment', '')
                                    if original_segment:
                                        st.text_area(
                                            "原始文本片段：",
                                            value=original_segment,
                                            height=100,
                                            disabled=True,
                                            key=f"original_text_{i}"
                                        )
                                    else:
                                        st.info("暂无对应的原始文本片段")
                        
                        # 下一步提示
                        st.markdown("---")
                        st.markdown('<div class="info-box">🔄 <strong>下一步开发计划</strong><br>• 扩展模板库：集成200+精美PPT模板（商务、学术、创意等风格）<br>• Dify API集成：智能匹配每页最适合的模板<br>• 批量文本填充：对每一页分别进行智能填充<br>• 多页面整合：生成风格统一的完整PPT<br>• 高级功能：模板风格选择、页面排序调整、批量导出</div>', unsafe_allow_html=True)
                        
                        # 调试信息（可选显示）
                        with st.expander("🔍 查看完整分页数据（开发调试）", expanded=False):
                            st.json(split_result)
                    
                    else:
                        st.warning("⚠️ 分页结果为空，请检查输入文本")
                        
                else:
                    st.error("❌ AI分页失败，请检查您的输入或稍后重试")
                    
                    # 显示错误信息（如果有）
                    if 'error' in split_result:
                        st.error(f"错误详情：{split_result['error']}")
                    
            except Exception as e:
                st.error(f"❌ 处理过程中出现错误：{str(e)}")
                logger.error("AI分页功能错误: %s", str(e))
    
    # 页脚信息
    st.markdown("---")
    st.markdown(
        '<div style="text-align: center; color: #666; padding: 2rem;">'
        '💡 由OpenRouter GPT-4V驱动 | 🎨 专业PPT自动生成'
        '</div>', 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()