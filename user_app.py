#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文本转PPT填充器 - 用户版Web界面
使用OpenAI GPT-4V将文本填入现有PPT文件
集成AI智能分页与Dify-模板桥接功能
"""

import streamlit as st
import os
import sys
import subprocess
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
import json
import re
from config import get_config
from utils import AIProcessor, PPTProcessor, FileManager, PPTAnalyzer
from logger import get_logger, log_user_action, log_file_operation, LogContext

# 依赖检查和安装函数
def check_dependencies_light():
    """轻量级依赖检查（不安装）"""
    try:
        import streamlit
        import pptx
        return True
    except ImportError:
        return False

def check_system_requirements():
    """检查系统要求"""
    print("🔍 检查系统要求...")
    
    # 检查Python版本
    if sys.version_info < (3, 8):
        print("❌ Python版本过低，需要Python 3.8或更高版本")
        return False
    
    print("✅ Python版本检查通过")
    
    # 检查必要的目录和文件
    required_files = [
        'config.py',
        'utils.py',
        'logger.py',
        'ai_page_splitter.py',
        'dify_template_bridge.py',
        'dify_api_client.py'
    ]
    
    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print(f"❌ 缺少必要的文件: {', '.join(missing_files)}")
        return False
    
    print("✅ 必要文件检查通过")
    
    # 检查模板目录
    templates_dir = os.path.join("templates", "ppt_template")
    if not os.path.exists(templates_dir):
        print("❌ 模板目录不存在: templates/ppt_template/")
        return False
    
    template_files = [f for f in os.listdir(templates_dir) if f.startswith("split_presentations_") and f.endswith(".pptx")]
    if len(template_files) == 0:
        print("❌ 模板目录中没有找到可用的PPT模板文件")
        return False
    
    print(f"✅ 模板库检查通过，发现 {len(template_files)} 个模板文件")
    
    return True

def initialize_system():
    """轻量级系统初始化"""
    # 只做基础检查，不执行耗时操作
    if not check_dependencies_light():
        return False
    
    # 检查基础文件存在
    required_files = ['config.py', 'utils.py', 'logger.py']
    for file in required_files:
        if not os.path.exists(file):
            return False
    
    return True

def show_results_section(pages, page_results):
    """显示处理结果部分"""
    # 显示分页和模板匹配结果
    st.markdown("### 📊 生成结果摘要")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("📄 总页数", len(pages))
    
    with col2:
        # 统计实际的Dify API调用次数（排除封面页）
        dify_calls = len([p for p in page_results if not p.get('is_title_page', False)])
        st.metric("🔗 Dify API调用", dify_calls)
    
    with col3:
        # 统计成功匹配数（包括封面页的固定匹配）
        success_count = len([p for p in page_results if p.get('template_number')])
        st.metric("✅ 成功匹配", success_count)
    
    with col4:
        # 统计总耗时（只计算Dify API调用耗时）
        total_time = sum(p.get('processing_time', 0) for p in page_results if not p.get('is_title_page', False))
        st.metric("⏱️ API耗时", f"{total_time:.2f}秒")
    
    # 显示每页详情
    st.markdown("### 📄 页面详情")
    
    for i, page_result in enumerate(page_results):
        # 区分封面页、结尾页和普通页面的显示标题
        if page_result.get('is_title_page', False):
            expander_title = f"第{page_result['page_number']}页 - 📋 封面页(固定模板)"
        elif page_result.get('is_ending_page', False):
            expander_title = f"第{page_result['page_number']}页 - 🔚 结尾页(固定模板)"
        else:
            expander_title = f"第{page_result['page_number']}页 - 模板#{page_result['template_number']}"
        
        with st.expander(expander_title, expanded=i < 3):
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.text(f"📄 页面编号: {page_result['page_number']}")
                if page_result.get('is_title_page', False):
                    st.text(f"📋 页面类型: 封面页")
                    st.text(f"📁 固定模板: {page_result['template_filename']}")
                    st.text(f"⚡ 处理方式: 直接匹配，无需API调用")
                elif page_result.get('is_ending_page', False):
                    st.text(f"🔚 页面类型: 结尾页")
                    st.text(f"📁 固定模板: {page_result['template_filename']}")
                    st.text(f"⚡ 处理方式: 直接匹配，无需API调用")
                else:
                    st.text(f"🔢 模板编号: #{page_result['template_number']}")
                    st.text(f"📁 模板文件: {page_result['template_filename']}")
                    st.text(f"⏱️ 处理时间: {page_result['processing_time']:.2f}秒")
            
            with col2:
                st.text_area(
                    "页面内容:",
                    value=page_result['content'][:200] + "..." if len(page_result['content']) > 200 else page_result['content'],
                    height=100,
                    disabled=True,
                    key=f"page_content_{i}"
                )
            
            if page_result.get('dify_response'):
                response_label = "固定响应:" if page_result.get('is_title_page', False) else "Dify API响应:"
                st.text_area(
                    response_label,
                    value=page_result['dify_response'],
                    height=80,
                    disabled=True,
                    key=f"dify_response_{i}"
                )
    
    # PPT下载区域
    st.markdown("### 📥 下载完整PPT")
    pages_count = len(pages) if pages else len(page_results)
    
    # 初始化session state
    if 'ppt_merge_result' not in st.session_state:
        st.session_state.ppt_merge_result = None
    
    # 检查PPT整合结果
    if st.session_state.ppt_merge_result:
        merge_result = st.session_state.ppt_merge_result
        
        # 显示整合结果
        st.success("🎉 PPT整合成功！")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📄 总页数", merge_result["total_pages"])
        with col2:
            st.metric("✅ 成功页面", merge_result["processed_pages"])
        with col3:
            st.metric("⚠️ 跳过页面", merge_result["skipped_pages"])
        with col4:
            ppt_size_mb = len(merge_result["presentation_bytes"]) / (1024 * 1024)
            st.metric("📦 文件大小", f"{ppt_size_mb:.2f}MB")
        
        # 提供下载
        if merge_result["presentation_bytes"]:
            from datetime import datetime
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"AI智能生成PPT_{timestamp}.pptx"
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.download_button(
                    label="📥 下载完整PPT文件",
                    data=merge_result["presentation_bytes"],
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    key="download_merged_ppt"
                )
            
            st.markdown('<div class="success-box">🎉 <strong>PPT自动生成完成！</strong><br>• ✅ 每页都使用了Dify API推荐的最佳模板<br>• ✅ 所有模板页面已自动整合为完整PPT<br>• ✅ 保持了每个模板的原有设计风格<br>• 📥 点击上方按钮即可下载完整的PPT文件</div>', unsafe_allow_html=True)
        
        # 显示错误信息（如果有）
        if merge_result.get("errors"):
            with st.expander("⚠️ 查看处理警告", expanded=False):
                for error in merge_result["errors"]:
                    st.warning(f"• {error}")
        
    
    else:
        # PPT整合正在进行或失败
        st.info("🔄 PPT正在自动整合中，请稍候...")
        st.markdown('<div class="info-box">📋 <strong>处理状态：</strong><br>• ✅ AI智能分页：成功将长文本分割为 {pages_count} 页<br>• ✅ 封面页处理：第1页自动使用 title_slides.pptx 固定模板<br>• ✅ Dify模板桥接：其他页面通过API获取最适合的模板<br>• 🔄 PPT整合：系统正在自动整合模板页面...<br>• ⏳ 请稍候：整合完成后将自动显示下载按钮</div>'.format(pages_count=pages_count), unsafe_allow_html=True)
    
    # 添加重新开始按钮
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🔄 重新开始", help="清除当前结果，重新输入内容", key="restart_process"):
            # 清除所有相关的session state
            if 'current_page_results' in st.session_state:
                del st.session_state.current_page_results
            if 'current_pages' in st.session_state:
                del st.session_state.current_pages
            if 'ppt_merge_result' in st.session_state:
                del st.session_state.ppt_merge_result
            st.rerun()
    
    # 调试信息
    with st.expander("🔍 查看完整处理数据（调试信息）", expanded=False):
        st.json({
            'pages': pages,
            'page_results': page_results
        })

# 获取配置 - 移除阻塞性初始化
config = get_config()
logger = get_logger()

# 云环境检测
def is_cloud_environment():
    """检测是否在云环境中运行"""
    return (os.getenv('STREAMLIT_CLOUD') or 
            '/home/adminuser/venv' in str(sys.executable) or
            '/mount/src/' in os.getcwd())

# 延迟初始化函数
@st.cache_resource
def lazy_initialize():
    """延迟初始化系统资源"""
    if is_cloud_environment():
        # 云环境只做基础检查
        return True
    else:
        # 本地环境执行完整初始化
        return initialize_system()

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
    # 延迟初始化系统
    if not lazy_initialize():
        st.error("❌ 系统初始化失败，请刷新页面重试")
        return
    
    # 页面标题
    st.markdown('<div class="main-header">🎨 AI PPT助手</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">智能将您的文本内容转换为精美的PPT演示文稿</div>', unsafe_allow_html=True)
    
    # 检查Dify API密钥环境变量
    import os
    # 尝试手动加载.env文件
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass
    
    dify_keys = [os.getenv(f"DIFY_API_KEY_{i}") for i in range(1, 6)]
    valid_dify_keys = [key for key in dify_keys if key]
    
    if len(valid_dify_keys) == 0:
        st.error("⚠️ **Dify API密钥未配置**")
        st.markdown("""
        请配置环境变量 `DIFY_API_KEY_1` 到 `DIFY_API_KEY_5`。
        
        **配置方法：**
        1. 复制 `.env.example` 为 `.env`
        2. 填入实际的API密钥
        3. 重启应用
        
        详细说明请查看 `ENVIRONMENT_SETUP.md`
        """)
        return
    elif len(valid_dify_keys) < 5:
        st.warning(f"⚠️ 当前配置了 {len(valid_dify_keys)}/5 个Dify API密钥，建议配置全部5个以获得最佳性能")
    
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
        else:  # 阿里云通义千问
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
            - Qwen Max：阿里云通义千问Max模型，顶级性能和理解能力
            """)
        
        with col2:
            st.markdown("""
            **第二步：准备API密钥** 🔑
            - 根据选择的模型注册相应平台账号
            - OpenRouter/阿里云获取API密钥
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
    elif api_provider == "阿里云":
        if not api_key.startswith('sk-'):
            st.markdown('<div class="warning-box">⚠️ 阿里云API密钥格式可能不正确，请检查密钥格式</div>', unsafe_allow_html=True)
            return
    
    # 跳过系统默认模板检查，直接使用Dify API和模板库
    # 注释掉原有的模板检查，改为检查模板库是否可用
    templates_dir = os.path.join(os.path.dirname(__file__), "templates", "ppt_template")
    if not os.path.exists(templates_dir):
        st.markdown('<div class="error-box">❌ 模板库文件夹不存在，请检查templates/ppt_template目录</div>', unsafe_allow_html=True)
        return
    
    # 检查模板库中是否有可用的模板文件
    template_files = [f for f in os.listdir(templates_dir) if f.startswith("split_presentations_") and f.endswith(".pptx")]
    if len(template_files) == 0:
        st.markdown('<div class="error-box">❌ 模板库中没有找到可用的PPT模板文件</div>', unsafe_allow_html=True)
        return
    
    st.markdown(f'<div class="success-box">✅ 模板库已就绪！发现 {len(template_files)} 个可用模板</div>', unsafe_allow_html=True)
    
    # 初始化AI处理器（不依赖默认模板）
    try:
        with st.spinner("正在验证API密钥..."):
            # 直接初始化AI处理器用于Dify API调用
            from utils import AIProcessor
            ai_processor = AIProcessor(api_key)
            # 测试API密钥有效性
            ai_processor._ensure_client()
            
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
    
    st.markdown('<div class="success-box">✅ AI助手已准备就绪！可以使用Dify API和模板库功能</div>', unsafe_allow_html=True)
    
    # 功能选择选项卡
    st.markdown("---")
    # 仅保留核心入口，移除“AI智能分页（预览）”和“Dify-模板桥接测试”
    tab1, tab3 = st.tabs(["🎨 智能PPT生成", "🧪 自定义模板测试"])
    
    with tab1:
        # 智能PPT生成功能 - AI分页 + Dify模板桥接
        st.markdown("### 🚀 智能PPT生成 (AI分页 + Dify模板桥接)")
        
        # 检查是否有保存的处理结果
        if 'current_page_results' in st.session_state and 'current_pages' in st.session_state:
            # 显示保存的结果
            pages = st.session_state.current_pages
            page_results = st.session_state.current_page_results
            
            st.markdown('<div class="success-box">🎉 智能PPT生成完成！</div>', unsafe_allow_html=True)
            
            # 跳转到结果显示部分
            show_results_section(pages, page_results)
        else:
            # 显示输入界面
            st.markdown('<div class="info-box">🎯 <strong>完整AI处理流程</strong><br>此功能结合AI智能分页与Dify模板桥接：<br>1. 用户输入长文本<br>2. AI模型智能分页（Qwen Max/GPT-4o）<br>3. 每页内容单独调用Dify API获取对应模板<br>4. 系统自动整合所有模板页面为完整PPT<br>5. 用户直接下载完整的PPT文件</div>', unsafe_allow_html=True)
    
        # 文本输入
        st.markdown("#### 📝 输入您的内容")
        
        user_text = st.text_area(
            "请输入您想要制作成PPT的文本内容：",
            height=250,
            placeholder="""例如：

人工智能的发展历程与未来趋势

人工智能技术的发展经历了多个重要阶段。从1950年代的符号主义开始，强调逻辑推理和知识表示，到1980年代的专家系统兴起，再到近年来深度学习的突破性进展。

技术发展阶段：
- 符号主义时代：基于规则和逻辑推理
- 连接主义时代：神经网络和机器学习
- 深度学习时代：大数据驱动的智能系统
- 大模型时代：通用人工智能的探索

当前，大语言模型如GPT、Claude等展现出了前所未有的能力，能够进行复杂的文本理解、生成和推理。这些技术正在革新各个行业，从教育、医疗到金融、娱乐，都能看到AI的身影。

未来发展趋势：
人工智能将继续向更加智能化、人性化的方向发展，实现更好的人机协作，为人类社会带来更多便利和创新可能性。同时需要关注AI安全和伦理问题。""",
            help="AI将分析文本结构进行智能分页，每页内容调用Dify API获取对应模板"
        )
        
        # 页面数量限制提醒
        st.info("📋 **页面数量限制：**最多生成25页（包括标题页、内容页和结尾页）")

        # 分页选项
        st.markdown("#### ⚙️ 分页选项")
        
        col1, col2 = st.columns(2)
        with col1:
            target_pages = st.number_input(
                "目标页面数量（可选）",
                min_value=0,
                max_value=25,
                value=0,
                help="设置为0时，AI将自动判断最佳页面数量"
            )
            
            # 页数建议
            st.markdown("""
            <div style="background-color: #f0f2f6; padding: 0.5rem; border-radius: 0.25rem; margin-top: 0.5rem;">
            <small>💡 <strong>页数建议：</strong><br>
            • 5分钟演示：3-5页<br>
            • 10分钟演示：5-8页<br>
            • 15分钟演示：8-12页<br>
            • 30分钟演示：15-20页<br>
            • 学术报告：20-25页</small>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            if user_text:
                char_count = len(user_text)
                word_count = len(user_text.split())
                st.metric("📊 文本统计", f"{char_count}字符 | {word_count}词")
        
        # 处理按钮
        st.markdown("#### 🚀 生成PPT")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            process_button = st.button(
                "🚀 开始生成PPT（AI分页 + 模板匹配 + 自动整合）",
                type="primary",
                use_container_width=True,
                disabled=not user_text.strip(),
                help="AI分页 → Dify模板匹配 → 自动整合PPT → 可直接下载"
            )
    
        # 处理逻辑 - AI分页 + Dify模板桥接
        if process_button and user_text.strip():
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # 步骤1：AI智能分页
                status_text.text("🤖 AI正在分析文本结构并进行智能分页...")
                progress_bar.progress(20)
                
                from ai_page_splitter import AIPageSplitter
                page_splitter = AIPageSplitter(api_key)
                target_page_count = int(target_pages) if target_pages > 0 else None
                split_result = page_splitter.split_text_to_pages(user_text.strip(), target_page_count)
                
                if not split_result.get('success'):
                    st.error(f"❌ AI分页失败: {split_result.get('error', '未知错误')}")
                    return
                
                pages = split_result.get('pages', [])
                if not pages:
                    st.error("❌ 分页结果为空，请检查输入文本")
                    return
                
                st.success(f"✅ AI智能分页完成！共生成 {len(pages)} 页")
                
                # 步骤2：为每页内容调用Dify API获取模板
                status_text.text("🔗 正在为每页内容调用Dify API获取对应模板...")
                progress_bar.progress(40)
                
                from dify_template_bridge import sync_test_dify_template_bridge
                page_results = []
                
                for i, page in enumerate(pages):
                    # 获取页面内容，优先使用original_text_segment，如果没有则使用title和key_points组合
                    page_content = page.get('original_text_segment', '')
                    if not page_content:
                        # 如果没有original_text_segment，则组合title和key_points
                        title = page.get('title', '')
                        key_points = page.get('key_points', [])
                        page_content = f"{title}\n\n" + "\n".join(key_points)
                    
                    page_type = page.get('page_type', 'content')
                    page_number = page.get('page_number', i + 1)
                    
                    # 封面页直接使用 title_slides.pptx，不调用Dify API
                    if page_type == 'title' or page_number == 1:
                        title_template_path = os.path.join("templates", "title_slides.pptx")
                        page_results.append({
                            'page_number': page_number,
                            'content': page_content,
                            'template_number': 'title',
                            'template_path': title_template_path,
                            'template_filename': "title_slides.pptx",
                            'dify_response': '封面页使用固定标题模板',
                            'processing_time': 0,
                            'is_title_page': True
                        })
                        st.info(f"📋 第{page_number}页(封面页)：使用固定标题模板 title_slides.pptx")
                    
                    # 结尾页直接使用 ending_slides.pptx，不调用Dify API
                    elif page_type == 'ending' or page.get('skip_dify_api', False):
                        ending_template_path = page.get('template_path', os.path.join("templates", "ending_slides.pptx"))
                        page_results.append({
                            'page_number': page_number,
                            'content': page_content,
                            'template_number': 'ending',
                            'template_path': ending_template_path,
                            'template_filename': "ending_slides.pptx",
                            'dify_response': '结尾页使用固定感谢模板',
                            'processing_time': 0,
                            'is_ending_page': True
                        })
                        st.info(f"🔚 第{page_number}页(结尾页)：使用固定结尾模板 ending_slides.pptx")
                    
                    elif page_content:
                        # 其他页面调用Dify API
                        bridge_result = sync_test_dify_template_bridge(page_content)
                        if bridge_result.get('success'):
                            dify_result = bridge_result["step_1_dify_api"]
                            template_result = bridge_result["step_2_template_lookup"]
                            page_results.append({
                                'page_number': page_number,
                                'content': page_content,
                                'template_number': dify_result.get('template_number'),
                                'template_path': template_result.get('file_path'),
                                'template_filename': template_result.get('filename'),
                                'dify_response': dify_result.get('response_text', ''),
                                'processing_time': bridge_result.get('processing_time', 0),
                                'is_title_page': False
                            })
                        else:
                            st.error(f"❌ 第{page_number}页Dify API调用失败: {bridge_result.get('error')}")
                            st.error("🚫 无法继续处理，请检查Dify API配置或稍后重试")
                            return  # 直接退出，不继续处理
                
                # 步骤3：整合PPT页面
                status_text.text("🔗 正在整合模板页面生成PPT...")
                progress_bar.progress(80)
                
                # 保存页面结果到session state
                st.session_state.current_page_results = page_results
                st.session_state.current_pages = pages
                
                # 自动执行PPT整合
                try:
                    # 使用增强版合并器，自动选择最佳方法
                    from ppt_merger import merge_dify_templates_to_ppt_enhanced
                    status_text.text("🔗 正在整合PPT页面(增强格式保留)...")
                    progress_bar.progress(90)
                    merge_result = merge_dify_templates_to_ppt_enhanced(page_results)
                    
                    # 整合PPT结果处理保持不变
                    
                    if merge_result["success"]:
                        # 保存整合结果
                        st.session_state.ppt_merge_result = merge_result
                        
                        # 完成处理流程
                        progress_bar.progress(100)
                        status_text.text("✅ PPT整合完成，可以下载！")
                        
                        # 清除进度显示
                        progress_bar.empty()
                        status_text.empty()
                        
                        # 刷新页面以显示结果
                        st.rerun()
                    else:
                        progress_bar.empty()
                        status_text.empty()
                        st.error(f"❌ PPT整合失败: {merge_result.get('error', '未知错误')}")
                        
                        if merge_result.get("errors"):
                            with st.expander("🔍 查看详细错误信息", expanded=False):
                                for error in merge_result["errors"]:
                                    st.error(f"• {error}")
                        return
                
                except ImportError:
                    progress_bar.empty()
                    status_text.empty()
                    st.error("❌ PPT整合模块未找到，请检查 ppt_merger.py 文件")
                    return
                except Exception as e:
                    progress_bar.empty()
                    status_text.empty()
                    st.error(f"❌ PPT整合过程中出现异常: {str(e)}")
                    return
                
            except ImportError as e:
                st.error(f"❌ 模块导入失败: {str(e)}")
            except Exception as e:
                progress_bar.empty()
                status_text.empty()
                st.error(f"❌ 智能PPT生成过程中出现异常: {str(e)}")
                logger.error("智能PPT生成异常: %s", str(e))
    
    # 原 tab2（AI智能分页预览）已移除
    # with tab2:
        # AI智能分页 + Dify API增强功能
        st.markdown("### 🚀 AI智能分页 + Dify API增强")
        
        st.markdown('<div class="info-box">🎯 <strong>完整AI处理流程</strong><br>默认启用的完整工作流程：AI智能分页 → 多密钥并发Dify API调用 → 增强结果输出<br><br>⚡ <strong>性能优化：</strong>使用3个Dify API密钥进行负载均衡，处理速度提升3倍，支持高并发处理<br><br>📋 <strong>分页规范：</strong>标题页仅提取标题和日期（其他内容固定），结尾页使用预设模板（无需生成），重点关注中间内容页的智能分割和API增强</div>', unsafe_allow_html=True)
        
        
        
        # 分页处理按钮
        split_button = st.button(
            "🤖 开始AI智能分页",
            type="primary",
            use_container_width=True,
            disabled=not user_text.strip(),
            help="AI将分析您的文本结构并智能分页"
        )
        
        # Dify API选项 - 默认启用完整工作流程
        st.markdown("#### 🔗 完整处理流程 (推荐)")
        
        enable_dify_api = st.checkbox(
            "启用完整AI处理流程：智能分页 + Dify API增强",
            value=True,  # 默认启用完整流程
            help="完整流程：AI分页 → 3个Dify API密钥并发处理 → 增强结果输出"
        )
        
        if enable_dify_api:
            st.success("✅ **完整处理流程已启用** - 将获得最佳处理效果")
            st.markdown("""
            **处理步骤：**
            1. 🤖 AI智能分页：第1页提取标题，第2页开始处理内容
            2. 🚀 Dify API并发调用：3个API密钥同时处理各页内容
            3. 📊 结果整合：显示分页结果和API增强内容
            """)
        else:
            st.warning("⚠️ **仅基础分页模式** - 功能不完整，建议启用完整流程")
            st.markdown("只进行AI文本分页，不调用Dify API进行内容增强")
        
        # 处理AI分页逻辑
        if split_button and user_text.strip():
            from ai_page_splitter import AIPageSplitter, PageContentFormatter
            
            try:
                with st.spinner("🤖 AI正在分析文本结构并进行智能分页..."):
                    # 初始化AI分页器
                    page_splitter = AIPageSplitter(api_key)
                    
                    # 执行智能分页
                    target_page_count = int(target_pages) if target_pages > 0 else None
                    split_result = page_splitter.split_text_to_pages(user_text, target_page_count)
                
                if split_result.get('success'):
                    st.markdown('<div class="success-box">✅ AI智能分页完成！</div>', unsafe_allow_html=True)
                    
                    # 显示分析摘要
                    analysis = split_result.get('analysis', {})
                    analysis_summary = PageContentFormatter.format_analysis_summary(analysis)
                    st.markdown(analysis_summary)
                    
                    # Dify API处理（如果启用）
                    final_result = split_result
                    if enable_dify_api:
                        try:
                            with st.spinner("🔗 正在调用Dify API处理每页内容..."):
                                from dify_api_client import process_pages_with_dify
                                
                                # 调用Dify API处理分页结果
                                dify_result = process_pages_with_dify(split_result)
                                final_result = dify_result
                                
                                if dify_result.get('success'):
                                    st.markdown('<div class="success-box">🚀 Dify API处理完成！</div>', unsafe_allow_html=True)
                                    
                                    # 显示Dify处理摘要
                                    from dify_api_client import DifyIntegrationService
                                    service = DifyIntegrationService()
                                    dify_summary = service.format_results_summary(dify_result)
                                    st.markdown(dify_summary)
                                    
                                else:
                                    st.warning(f"⚠️ Dify API处理失败: {dify_result.get('error', '未知错误')}")
                                    # 即使Dify API失败，仍然显示原始分页结果
                                    final_result = split_result
                                    
                        except ImportError:
                            st.error("❌ Dify API客户端模块未找到，请检查安装")
                            final_result = split_result
                        except Exception as e:
                            st.error(f"❌ Dify API调用异常: {str(e)}")
                            final_result = split_result
                    
                    # 显示分页结果（优先显示增强后的结果）
                    display_pages = final_result.get('enhanced_pages', final_result.get('pages', []))
                    original_pages = split_result.get('pages', [])
                    
                    if display_pages:
                        # 根据是否启用了Dify API显示不同的标题
                        if enable_dify_api and final_result != split_result:
                            st.markdown("### 📄 完整处理结果：AI分页 + Dify API增强")
                        else:
                            st.markdown("### 📄 基础分页结果（未启用完整流程）")
                        
                        # 使用选项卡显示每一页
                        page_tabs = st.tabs([f"第{page['page_number']}页" for page in display_pages])
                        
                        for i, (page_tab, page_data) in enumerate(zip(page_tabs, display_pages)):
                            with page_tab:
                                # 显示基本页面信息
                                page_preview = PageContentFormatter.format_page_preview(page_data)
                                st.markdown(page_preview)
                                
                                # 显示Dify API结果（如果有）
                                if 'dify_response' in page_data:
                                    st.markdown("---")
                                    st.markdown("### 🚀 Dify API 响应结果")
                                    
                                    response_text = page_data.get('dify_response', '')
                                    if response_text:
                                        st.text_area(
                                            "API响应内容：",
                                            value=response_text,
                                            height=150,
                                            disabled=True,
                                            key=f"dify_response_{i}"
                                        )
                                    
                                    # 显示API调用详情
                                    api_result = page_data.get('dify_api_result', {})
                                    if api_result:
                                        col1, col2, col3 = st.columns(3)
                                        with col1:
                                            st.metric("🔄 尝试次数", api_result.get('attempt', 1))
                                        with col2:
                                            st.metric("📊 状态码", api_result.get('api_status', 'N/A'))
                                        with col3:
                                            success_status = "✅ 成功" if api_result.get('success') else "❌ 失败"
                                            st.metric("🎯 调用状态", success_status)
                                
                                # 显示Dify API错误（如果有）
                                elif 'dify_error' in page_data:
                                    st.markdown("---")
                                    st.markdown("### ⚠️ Dify API 调用失败")
                                    st.error(f"错误信息: {page_data.get('dify_error', '未知错误')}")
                                
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
                                
                                # 显示完整的API响应数据（调试用）
                                if enable_dify_api and 'dify_full_response' in page_data:
                                    with st.expander("🔍 查看完整API响应（调试信息）", expanded=False):
                                        st.json(page_data.get('dify_full_response', {}))
                        
                        # 功能状态提示（根据是否启用Dify API显示不同信息）
                        st.markdown("---")
                        if enable_dify_api and final_result != split_result:
                            st.markdown('<div class="info-box">🎉 <strong>完整AI处理流程已完成</strong><br>• ✅ AI智能分页：第1页标题，第2页开始内容<br>• ✅ 多密钥并发：3个Dify API密钥负载均衡<br>• ✅ 性能优化：处理速度提升3倍<br>• ✅ 结果增强：每页都获得API增强内容<br><br>🚀 <strong>技术特性</strong><br>• 轮询负载均衡，确保密钥使用均匀<br>• 自动故障转移，单密钥失败不影响整体<br>• 实时监控API使用统计和响应状态</div>', unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="info-box">⚠️ <strong>基础模式警告</strong><br>当前仅使用基础分页功能，未启用完整的AI处理流程<br><br>💡 <strong>建议操作</strong><br>• 勾选上方"启用完整AI处理流程"选项<br>• 获得AI分页 + Dify API增强的完整体验<br>• 享受3倍处理速度提升和更丰富的结果</div>', unsafe_allow_html=True)
                        
                        # 调试信息（可选显示）
                        with st.expander("🔍 查看完整处理数据（开发调试）", expanded=False):
                            if enable_dify_api and final_result != split_result:
                                st.markdown("**完整处理结果（包含Dify API响应）：**")
                                st.json(final_result)
                            else:
                                st.markdown("**分页处理结果：**")
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
    
    with tab3:
        # 自定义模板测试功能
        st.markdown("### 🧪 自定义模板测试")
        
        st.markdown('<div class="info-box">🎯 <strong>功能说明</strong><br>此功能独立于智能分页和Dify API，专门用于测试您自己的PPT模板。您可以上传自定义模板，输入文本内容，系统将智能填充到您的模板中。</div>', unsafe_allow_html=True)
        
        # 模板上传区域
        st.markdown("#### 📁 上传您的PPT模板")
        
        uploaded_file = st.file_uploader(
            "选择您的PPT模板文件",
            type=['pptx'],
            help="请上传.pptx格式的PPT模板文件，建议文件大小不超过50MB",
            key="custom_template_uploader"
        )
        
        if uploaded_file is not None:
            # 显示上传文件信息
            file_details = {
                "文件名": uploaded_file.name,
                "文件大小": f"{uploaded_file.size / 1024:.1f} KB",
                "文件类型": uploaded_file.type
            }
            
            col1, col2 = st.columns([1, 2])
            with col1:
                st.success("✅ 模板文件已上传")
                for key, value in file_details.items():
                    st.text(f"{key}: {value}")
            
            with col2:
                # 保存上传的文件到临时目录
                import tempfile
                import shutil
                
                try:
                    # 创建临时文件
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        temp_ppt_path = tmp_file.name
                    
                    # 验证PPT文件
                    is_valid, error_msg = FileManager.validate_ppt_file(temp_ppt_path)
                    
                    if is_valid:
                        st.success("✅ 模板文件格式验证通过")
                        
                        # 分析模板结构
                        try:
                            from pptx import Presentation
                            temp_presentation = Presentation(temp_ppt_path)
                            
                            # 基本信息
                            slide_count = len(temp_presentation.slides)
                            st.metric("📑 幻灯片数量", slide_count)
                            
                            # 分析占位符
                            total_placeholders = 0
                            placeholder_info = []
                            
                            for i, slide in enumerate(temp_presentation.slides):
                                slide_placeholders = []
                                for shape in slide.shapes:
                                    if hasattr(shape, 'text') and shape.text:
                                        # 查找占位符模式 {xxx}
                                        import re
                                        placeholders = re.findall(r'\{([^}]+)\}', shape.text)
                                        if placeholders:
                                            slide_placeholders.extend(placeholders)
                                            total_placeholders += len(placeholders)
                                
                                if slide_placeholders:
                                    placeholder_info.append({
                                        'slide_num': i + 1,
                                        'placeholders': slide_placeholders
                                    })
                            
                            st.metric("🎯 发现占位符", total_placeholders)
                            
                            # 显示占位符详情
                            if placeholder_info:
                                with st.expander("🔍 模板结构分析", expanded=False):
                                    for info in placeholder_info[:5]:  # 只显示前5页
                                        st.write(f"**第{info['slide_num']}页：** {', '.join([f'{{{p}}}' for p in info['placeholders']])}")
                                    
                                    if len(placeholder_info) > 5:
                                        st.write(f"... 还有 {len(placeholder_info)-5} 页包含占位符")
                            else:
                                st.warning("⚠️ 未检测到占位符模式 {xxx}，请确保模板中包含要填充的占位符")
                        
                        except Exception as e:
                            st.error(f"❌ 模板分析失败: {str(e)}")
                    else:
                        st.error(f"❌ 模板文件验证失败: {error_msg}")
                        temp_ppt_path = None
                    
                except Exception as e:
                    st.error(f"❌ 文件处理失败: {str(e)}")
                    temp_ppt_path = None
            
            # 如果模板验证通过，显示文本输入和处理区域
            if 'temp_ppt_path' in locals() and temp_ppt_path and is_valid:
                st.markdown("---")
                st.markdown("#### 📝 输入测试内容")
                
                test_text = st.text_area(
                    "请输入要填充到模板中的文本内容：",
                    height=200,
                    placeholder="""例如：

我的自定义PPT测试

这是使用自定义模板的测试内容。AI将分析您的文本结构，并智能地将内容分配到模板中的各个占位符位置。

主要特点：
- 支持自定义PPT模板上传
- 智能文本内容分配
- 保持原有模板设计风格
- 独立于其他功能模块

测试结果将展示AI如何理解您的内容并填充到模板的对应位置。""",
                    help="AI将分析您的文本并智能分配到模板的占位符中",
                    key="custom_template_text"
                )
                
                # 处理选项
                col1, col2 = st.columns(2)
                with col1:
                    # 获取当前模型信息
                    current_model_info = config.get_model_info()
                    supports_vision = current_model_info.get('supports_vision', False)
                    
                    if supports_vision:
                        enable_custom_visual = st.checkbox(
                            "🎨 启用视觉优化",
                            value=False,
                            help="对自定义模板应用AI视觉优化（需要额外时间）",
                            key="custom_visual_opt"
                        )
                    else:
                        enable_custom_visual = False
                        st.info("⚠️ 当前模型不支持视觉优化")
                
                with col2:
                    if test_text:
                        char_count = len(test_text)
                        word_count = len(test_text.split())
                        st.metric("📊 文本统计", f"{char_count}字符 | {word_count}词")
                
                # 处理按钮
                st.markdown("#### 🚀 开始测试")
                
                test_button = st.button(
                    "🧪 测试自定义模板",
                    type="primary",
                    use_container_width=True,
                    disabled=not test_text.strip(),
                    help="使用您的模板和内容进行AI智能填充测试",
                    key="custom_template_test_btn"
                )
                
                # 处理逻辑
                if test_button and test_text.strip():
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    try:
                        # 创建自定义模板生成器
                        status_text.text("🔧 正在初始化自定义模板处理器...")
                        progress_bar.progress(20)
                        
                        custom_generator = UserPPTGenerator(api_key)
                        success, message = custom_generator.load_ppt_from_path(temp_ppt_path)
                        
                        if not success:
                            st.error(f"❌ 自定义模板加载失败: {message}")
                            return
                        
                        # AI分析
                        status_text.text("🤖 AI正在分析您的内容和模板结构...")
                        progress_bar.progress(40)
                        
                        assignments = custom_generator.process_text_with_openai(test_text)
                        
                        # 填充内容
                        status_text.text("📝 正在将内容填入自定义模板...")
                        progress_bar.progress(60)
                        
                        success, results = custom_generator.apply_text_assignments(assignments, test_text)
                        
                        if not success:
                            st.error("❌ 内容填充失败，请检查模板格式")
                            return
                        
                        # 清理占位符
                        status_text.text("🧹 正在清理未使用的占位符...")
                        progress_bar.progress(80)
                        
                        cleanup_results = custom_generator.cleanup_unfilled_placeholders()
                        
                        # 可选的视觉优化
                        if enable_custom_visual:
                            status_text.text("🎨 正在应用视觉优化...")
                            progress_bar.progress(90)
                            
                            optimization_results = custom_generator.apply_visual_optimization(
                                temp_ppt_path, 
                                enable_visual_optimization=True
                            )
                        else:
                            optimization_results = custom_generator.apply_basic_beautification()
                        
                        # 完成处理
                        status_text.text("📦 正在准备下载...")
                        progress_bar.progress(100)
                        
                        # 清除进度显示
                        progress_bar.empty()
                        status_text.empty()
                        
                        # 显示成功信息
                        st.markdown('<div class="success-box">🎉 自定义模板测试完成！</div>', unsafe_allow_html=True)
                        
                        # 显示处理摘要
                        if optimization_results and "error" not in optimization_results:
                            st.markdown("### 📊 处理结果")
                            
                            col1, col2, col3, col4 = st.columns(4)
                            
                            with col1:
                                summary = optimization_results.get('summary', {})
                                final_slide_count = summary.get('final_slide_count', 0)
                                st.metric("📑 最终页数", final_slide_count)
                            
                            with col2:
                                cleanup_count = cleanup_results.get('cleaned_placeholders', 0) if cleanup_results else 0
                                st.metric("🧹 清理占位符", cleanup_count)
                            
                            with col3:
                                removed_empty = summary.get('removed_empty_slides_count', 0)
                                st.metric("🗑️ 删除空页", removed_empty)
                            
                            with col4:
                                reorganized = summary.get('reorganized_slides_count', 0)
                                st.metric("🔄 重新排版", reorganized)
                        
                        # 下载文件
                        st.markdown("### 💾 下载测试结果")
                        
                        try:
                            updated_ppt_bytes = custom_generator.get_ppt_bytes()
                            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                            original_name = uploaded_file.name.rsplit('.', 1)[0]
                            filename = f"{original_name}_测试结果_{timestamp}.pptx"
                            
                            col1, col2, col3 = st.columns([1, 2, 1])
                            with col2:
                                st.download_button(
                                    label="📥 下载测试结果PPT",
                                    data=updated_ppt_bytes,
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                    use_container_width=True,
                                    key="download_custom_result"
                                )
                            
                            st.markdown('<div class="info-box">', unsafe_allow_html=True)
                            st.markdown(f"📁 **文件名：** {filename}")
                            st.markdown("🎯 **测试内容：** 基于您的自定义模板生成")
                            st.markdown("📋 **说明：** 可以在PowerPoint中查看AI填充效果")
                            st.markdown('</div>', unsafe_allow_html=True)
                            
                        except Exception as e:
                            st.error(f"❌ 文件生成失败: {str(e)}")
                        
                        # 清理临时文件
                        try:
                            os.unlink(temp_ppt_path)
                        except:
                            pass
                            
                    except Exception as e:
                        progress_bar.empty()
                        status_text.empty()
                        st.error(f"❌ 测试过程中出现错误: {str(e)}")
                        logger.error("自定义模板测试错误: %s", str(e))
                        
                        # 清理临时文件
                        try:
                            os.unlink(temp_ppt_path)
                        except:
                            pass
        
        else:
            # 未上传文件时的说明
            st.markdown("#### 🎯 使用说明")
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("""
                **📋 模板要求：**
                - 文件格式：.pptx
                - 文件大小：<50MB
                - 包含占位符：{标题}、{内容}等
                - 建议结构清晰的模板设计
                """)
            
            with col2:
                st.markdown("""
                **🔄 处理流程：**
                1. 上传您的PPT模板
                2. 系统验证和分析模板结构
                3. 输入要填充的文本内容
                4. AI智能分配内容到占位符
                5. 下载填充后的PPT文件
                """)
            
            st.markdown("#### ✨ 功能特色")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("""
                **🎨 保持设计风格**
                - 完全保留您的模板样式
                - 不改变颜色、字体、布局
                - 只填充内容到指定位置
                """)
            
            with col2:
                st.markdown("""
                **🤖 智能内容分配**
                - AI理解文本结构和含义
                - 自动匹配最合适的占位符
                - 支持多种内容类型处理
                """)
            
            with col3:
                st.markdown("""
                **🔧 独立测试环境**
                - 不影响其他功能模块
                - 专门用于模板测试验证
                - 支持多次测试和调整
                """)
            
            st.markdown('<div class="warning-box">💡 <strong>提示：</strong> 请确保您的PPT模板中包含形如 {标题}、{内容}、{要点} 等占位符，AI将根据这些占位符的名称智能分配相应的内容。</div>', unsafe_allow_html=True)
    
    # 原 tab4（Dify-模板桥接测试）已移除
        
        st.markdown('<div class="info-box">🎯 <strong>功能说明</strong><br>此功能测试Dify API与模板文件库的桥接流程：<br>1. 用户输入文本内容<br>2. Dify API分析并返回模板编号(1-250)<br>3. 系统根据编号查找对应的PPT模板文件<br>4. 返回匹配的模板文件供下载测试<br><br>⚠️ 注意：此为桥接测试，暂不进行文本填充工作</div>', unsafe_allow_html=True)
        
        # 先显示可用模板概览
        st.markdown("#### 📊 模板库概览")
        
        try:
            from dify_template_bridge import DifyTemplateBridge
            
            # 扫描模板库
            bridge = DifyTemplateBridge()
            templates_info = bridge.scan_available_templates()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📁 模板总数", templates_info["total_count"])
            with col2:
                number_range = templates_info["number_range"]
                if number_range["min"] and number_range["max"]:
                    st.metric("🔢 编号范围", f"{number_range['min']}-{number_range['max']}")
                else:
                    st.metric("🔢 编号范围", "无可用模板")
            with col3:
                st.metric("📂 模板目录", "templates/ppt_template/")
            
            # 显示部分模板列表
            if templates_info["templates"]:
                with st.expander("🔍 查看部分模板文件", expanded=False):
                    # 显示前10个和后10个模板
                    templates = templates_info["templates"]
                    display_templates = templates[:10]
                    if len(templates) > 20:
                        display_templates.extend(templates[-10:])
                    elif len(templates) > 10:
                        display_templates.extend(templates[10:])
                    
                    for template in display_templates:
                        st.text(f"📄 {template['filename']} ({template['file_size_kb']}KB)")
                    
                    if len(templates) > 20:
                        st.text(f"... 还有 {len(templates) - 20} 个模板文件")
        
        except ImportError:
            st.error("❌ Dify桥接模块未找到，请检查 dify_template_bridge.py 文件")
        except Exception as e:
            st.error(f"❌ 模板库扫描失败: {str(e)}")
        
        st.markdown("---")
        
        # 桥接测试区域
        st.markdown("#### 🧪 桥接流程测试")
        
        # 文本输入
        bridge_test_text = st.text_area(
            "请输入测试文本内容：",
            height=150,
            placeholder="""例如：

企业数字化转型战略规划

随着数字技术的快速发展，企业数字化转型已成为提升竞争力的关键。本报告将从战略规划、技术选型、实施路径等方面进行深入分析。

主要内容包括：
- 数字化转型的必要性分析
- 技术架构设计与选型
- 实施计划与风险控制
- 预期效果与投资回报

通过系统化的规划和实施，企业可以实现运营效率提升、客户体验优化和商业模式创新。""",
            help="Dify API将分析此文本内容并返回对应的模板编号",
            key="bridge_test_text"
        )
        
        # 测试选项
        col1, col2 = st.columns(2)
        with col1:
            if bridge_test_text:
                char_count = len(bridge_test_text)
                word_count = len(bridge_test_text.split())
                st.metric("📊 文本统计", f"{char_count}字符 | {word_count}词")
        
        with col2:
            st.markdown("**测试步骤预览：**")
            st.text("1. 🤖 调用Dify API分析文本")
            st.text("2. 🔢 获取模板编号(1-250)")
            st.text("3. 📁 查找对应PPT文件")
            st.text("4. ✅ 返回模板文件信息")
        
        # 测试按钮
        st.markdown("#### 🚀 开始桥接测试")
        
        test_bridge_button = st.button(
            "🔗 测试Dify API桥接",
            type="primary",
            use_container_width=True,
            disabled=not bridge_test_text.strip(),
            help="测试Dify API返回编号与模板文件的对应关系",
            key="test_bridge_btn"
        )
        
        # 执行桥接测试
        if test_bridge_button and bridge_test_text.strip():
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                from dify_template_bridge import sync_test_dify_template_bridge
                
                # 步骤1: 调用Dify API
                status_text.text("🤖 正在调用Dify API分析文本...")
                progress_bar.progress(25)
                
                bridge_result = sync_test_dify_template_bridge(bridge_test_text.strip())
                
                # 步骤2: 处理结果
                status_text.text("📊 正在处理API响应...")
                progress_bar.progress(50)
                
                # 步骤3: 查找模板文件
                status_text.text("📁 正在查找对应模板文件...")
                progress_bar.progress(75)
                
                # 步骤4: 完成
                status_text.text("✅ 桥接测试完成")
                progress_bar.progress(100)
                
                # 清除进度显示
                progress_bar.empty()
                status_text.empty()
                
                # 显示测试结果
                if bridge_result["success"]:
                    st.markdown('<div class="success-box">🎉 Dify API桥接测试成功！</div>', unsafe_allow_html=True)
                    
                    # 显示详细结果
                    st.markdown("### 📋 桥接测试结果")
                    
                    # 基本信息
                    col1, col2, col3, col4 = st.columns(4)
                    
                    dify_result = bridge_result["step_1_dify_api"]
                    template_result = bridge_result["step_2_template_lookup"]
                    
                    with col1:
                        st.metric("🔢 Dify返回编号", dify_result["template_number"])
                    
                    with col2:
                        st.metric("📄 模板文件名", template_result["filename"].replace("split_presentations_", "").replace(".pptx", ""))
                    
                    with col3:
                        st.metric("📦 文件大小", f"{template_result['file_size_kb']}KB")
                    
                    with col4:
                        st.metric("⏱️ 处理耗时", f"{bridge_result['processing_time']:.2f}秒")
                    
                    # Dify API详情
                    st.markdown("#### 🤖 Dify API 调用详情")
                    
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.text(f"🔑 使用API密钥: {dify_result['used_api_key']}")
                        st.text(f"🔄 尝试次数: {dify_result['attempt_count']}")
                        st.text(f"✅ 调用状态: 成功")
                    
                    with col2:
                        if "response_text" in dify_result:
                            st.text_area(
                                "API响应内容:",
                                value=dify_result["response_text"],
                                height=100,
                                disabled=True,
                                key="dify_response_display"
                            )
                    
                    # 模板文件详情
                    st.markdown("#### 📁 模板文件详情")
                    
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.text(f"📂 文件路径: {template_result['filename']}")
                        st.text(f"💾 文件大小: {template_result['file_size']} 字节")
                        st.text(f"✅ 文件状态: 存在且有效")
                    
                    with col2:
                        # 提供模板文件下载
                        try:
                            with open(template_result["file_path"], "rb") as f:
                                template_bytes = f.read()
                            
                            st.download_button(
                                label="📥 下载对应的模板文件",
                                data=template_bytes,
                                file_name=template_result["filename"],
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                key="download_matched_template"
                            )
                            
                            st.markdown('<div class="info-box">💡 <strong>提示：</strong> 下载的模板文件是根据Dify API返回的编号自动匹配的，您可以在PowerPoint中打开查看模板结构。</div>', unsafe_allow_html=True)
                            
                        except Exception as e:
                            st.error(f"❌ 模板文件读取失败: {str(e)}")
                    
                    # 完整响应数据（调试用）
                    with st.expander("🔍 查看完整测试数据（调试信息）", expanded=False):
                        st.json(bridge_result)
                
                else:
                    st.markdown('<div class="error-box">❌ Dify API桥接测试失败</div>', unsafe_allow_html=True)
                    st.error(f"错误信息: {bridge_result['error']}")
                    
                    # 显示失败详情
                    if bridge_result["step_1_dify_api"]:
                        st.markdown("#### 🤖 Dify API 调用详情")
                        dify_result = bridge_result["step_1_dify_api"]
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.text(f"🔑 使用API密钥: {dify_result.get('used_api_key', 'N/A')}")
                            st.text(f"🔄 尝试次数: {dify_result.get('attempt_count', 'N/A')}")
                            st.text(f"❌ 调用状态: 失败")
                        
                        with col2:
                            if dify_result.get("api_response"):
                                st.text_area(
                                    "API响应内容:",
                                    value=str(dify_result["api_response"]),
                                    height=100,
                                    disabled=True,
                                    key="failed_dify_response"
                                )
                    
                    # 调试信息
                    with st.expander("🔍 查看失败详情（调试信息）", expanded=False):
                        st.json(bridge_result)
                
            except ImportError:
                st.error("❌ Dify桥接模块未找到，请检查 dify_template_bridge.py 文件")
            except Exception as e:
                progress_bar.empty()
                status_text.empty()
                st.error(f"❌ 桥接测试过程中出现异常: {str(e)}")
                logger.error("Dify桥接测试异常: %s", str(e))
        
        # 功能说明
        st.markdown("---")
        st.markdown("#### 🎯 测试目标")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            **🔍 验证内容：**
            - Dify API能否正常响应
            - 返回的数字是否在有效范围(1-250)
            - 对应的模板文件是否存在
            - 模板文件格式是否有效
            """)
        
        with col2:
            st.markdown("""
            **📋 后续计划：**
            - 第一阶段：桥接流程验证 ✅
            - 第二阶段：模板内容分析
            - 第三阶段：智能文本填充
            - 第四阶段：完整工作流集成
            """)
        
        st.markdown('<div class="warning-box">⚠️ <strong>重要说明：</strong> 当前功能仅测试Dify API与模板文件的对应关系，不进行实际的文本填充工作。这是分步实现的第一阶段，确保基础桥接流程正常工作。</div>', unsafe_allow_html=True)
    
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