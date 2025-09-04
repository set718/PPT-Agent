#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文本转PPT填充器 - 用户版Web界面
使用OpenAI GPT-5将文本填入现有PPT文件
集成AI智能分页与Dify-模板桥接功能
"""

import streamlit as st
import os
import sys

# 强制设置UTF-8编码
import locale
try:
    locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'C.UTF-8')
    except:
        pass

# 设置环境变量
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys, 'setdefaultencoding'):
    sys.setdefaultencoding('utf-8')
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

def check_dify_api_keys():
    """检查Dify API密钥配置，返回(是否有效, 有效密钥数量, 错误消息)"""
    import os
    
    dify_keys = [os.getenv(f"DIFY_API_KEY_{i}") for i in range(1, 6)]
    valid_dify_keys = [key for key in dify_keys if key]
    
    if len(valid_dify_keys) == 0:
        return False, 0, "⚠️ **Dify API密钥未配置**\n\n请配置环境变量 `DIFY_API_KEY_1` 到 `DIFY_API_KEY_5`。\n\n**配置方法：**\n1. 复制 `.env.example` 为 `.env`\n2. 填入实际的API密钥\n3. 重启应用\n\n详细说明请查看 `ENVIRONMENT_SETUP.md`"
    elif len(valid_dify_keys) < 5:
        return True, len(valid_dify_keys), f"⚠️ 当前配置了 {len(valid_dify_keys)}/5 个Dify API密钥，建议配置全部5个以获得最佳性能"
    else:
        return True, len(valid_dify_keys), None

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
        # 区分封面页、目录页、结尾页和普通页面的显示标题
        if page_result.get('is_title_page', False):
            expander_title = f"第{page_result['page_number']}页 - 📋 封面页(固定模板)"
        elif page_result.get('is_toc_page', False):
            expander_title = f"第{page_result['page_number']}页 - 📑 目录页(内容提取)"
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
                elif page_result.get('is_toc_page', False):
                    st.text(f"📑 页面类型: 目录页")
                    st.text(f"📁 模板文件: {page_result['template_filename']}")
                    st.text(f"⚡ 处理方式: AI分页时提取内容页标题，无需API调用")
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
        # 直接传递给AIProcessor，让它处理内置密钥标识符
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
            
            # 智能清理占位符，只清理未填充的
            cleanup_count = 0
            cleaned_placeholders = []
            
            for slide_idx, slide in enumerate(self.presentation.slides):
                # 获取该页已填充的占位符
                filled_placeholders_in_slide = self.ppt_processor.filled_placeholders.get(slide_idx, set())
                
                for shape in slide.shapes:
                    # 处理普通文本框
                    if hasattr(shape, 'text') and shape.text:
                        original_text = shape.text
                        
                        # 找出文本中的所有占位符 - 识别所有{}格式的占位符
                        import re
                        placeholder_matches = re.findall(r'\{([^}]+)\}', original_text)
                        
                        if placeholder_matches:
                            # 检查哪些占位符未被填充
                            unfilled_placeholders = [
                                p for p in placeholder_matches 
                                if p not in filled_placeholders_in_slide
                            ]
                            
                            # 只移除未填充的占位符
                            if unfilled_placeholders:
                                cleaned_text = original_text
                                for unfilled_placeholder in unfilled_placeholders:
                                    pattern = f"{{{unfilled_placeholder}}}"
                                    cleaned_text = cleaned_text.replace(pattern, "")
                                    cleaned_placeholders.append(f"第{slide_idx+1}页(文本框): {{{unfilled_placeholder}}}")
                                
                                # 清理多余的空白
                                cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()
                                
                                if cleaned_text != original_text:
                                    shape.text = cleaned_text
                                    cleanup_count += 1
                    
                    # 处理表格中的占位符
                    elif hasattr(shape, 'shape_type') and shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE = 19
                        table = shape.table
                        for row_idx, row in enumerate(table.rows):
                            for col_idx, cell in enumerate(row.cells):
                                original_cell_text = cell.text.strip()
                                if original_cell_text:
                                    # 找出表格单元格中的占位符
                                    placeholder_matches = re.findall(r'\{([^}]+)\}', original_cell_text)
                                    
                                    if placeholder_matches:
                                        # 检查哪些占位符未被填充
                                        unfilled_placeholders = [
                                            p for p in placeholder_matches 
                                            if p not in filled_placeholders_in_slide
                                        ]
                                        
                                        # 只移除未填充的占位符
                                        if unfilled_placeholders:
                                            cleaned_cell_text = original_cell_text
                                            for unfilled_placeholder in unfilled_placeholders:
                                                pattern = f"{{{unfilled_placeholder}}}"
                                                cleaned_cell_text = cleaned_cell_text.replace(pattern, "")
                                                cleaned_placeholders.append(f"第{slide_idx+1}页(表格{row_idx+1},{col_idx+1}): {{{unfilled_placeholder}}}")
                                            
                                            # 清理多余的空白
                                            cleaned_cell_text = re.sub(r'\s+', ' ', cleaned_cell_text).strip()
                                            
                                            if cleaned_cell_text != original_cell_text:
                                                cell.text = cleaned_cell_text
                                                cleanup_count += 1
            
            # 使用实际清理的占位符数量，而不是修改的文本框数量
            actual_cleaned_count = len(cleaned_placeholders)
            
            return {
                "success": True,
                "cleaned_placeholders": actual_cleaned_count,
                "cleaned_placeholder_list": cleaned_placeholders,
                "message": f"清理了{actual_cleaned_count}个占位符，涉及{cleanup_count}个文本框和表格单元格"
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
    import os
    # 延迟初始化系统
    if not lazy_initialize():
        st.error("❌ 系统初始化失败，请刷新页面重试")
        return
    
    # 页面标题
    st.markdown('<div class="main-header">🎨 AI PPT助手</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">智能将您的文本内容转换为精美的PPT演示文稿</div>', unsafe_allow_html=True)
    
    # 加载环境变量
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass
    
    # 模型选择区域
    st.markdown("### 🤖 选择AI模型")
    
    available_models = config.available_models
    model_options = {}
    for model_key, model_info in available_models.items():
        # 只对有成本信息的模型显示成本
        if model_info['cost']:
            display_name = f"{model_info['name']} ({model_info['cost']}成本)"
        else:
            display_name = model_info['name']
        model_options[display_name] = model_key
    
    model_col1, model_col2 = st.columns([2, 1])
    with model_col1:
        selected_display = st.selectbox(
            "选择适合您需求的AI模型：",
            options=list(model_options.keys()),
            index=0,
            help="不同模型有不同的功能特点"
        )
        
        selected_model = model_options[selected_display]
        model_info = available_models[selected_model]
        
        # 动态更新配置
        if selected_model != config.ai_model:
            config.set_model(selected_model)
    
    with model_col2:
        st.markdown("**模型对比**")
        if selected_model == "liai-chat":
            st.info("🏢 调用公司融合云AgentOps私有化模型\n🔒 数据安全保障\n✅ 支持视觉分析")
        else:  # DeepSeek V3
            st.success("🚀 火山引擎DeepSeek V3模型\n⚡ 性能优异\n🌐 支持中英文对话")
    

    
    st.markdown("---")
    
    # API密钥输入区域
    st.markdown("### 🔑 开始使用")
    
    # 根据选择的模型动态显示API密钥输入信息
    current_model_info = config.get_model_info()
    api_provider = current_model_info.get('api_provider', 'OpenAI')
    api_key_url = current_model_info.get('api_key_url', 'https://platform.openai.com/api-keys')
    
    col1, col2 = st.columns([2, 1])
    with col1:
        if api_provider == "Liai":
            # Liai自动填充API密钥（从环境变量读取，无需显示任何提示）
            import random
            import os
            
            # 强制重新加载环境变量以确保读取到最新的.env文件
            try:
                from dotenv import load_dotenv
                import os
                
                # 尝试多个可能的路径
                script_dir = os.path.dirname(os.path.abspath(__file__))
                current_work_dir = os.getcwd()
                
                possible_paths = [
                    os.path.join(script_dir, '.env'),
                    os.path.join(current_work_dir, '.env'),
                    '.env'
                ]
                
                
                found_env = False
                for env_path in possible_paths:
                    
                    if os.path.exists(env_path):
                        try:
                            with open(env_path, 'r', encoding='utf-8') as f:
                                content = f.read()
                                liai_lines = [line for line in content.split('\n') if 'LIAI_API_KEY' in line and not line.strip().startswith('#')]
                                
                                if len(liai_lines) > 0:
                                    
                                    load_dotenv(dotenv_path=env_path, override=True, encoding='utf-8')
                                    found_env = True
                                    break
                                    
                        except Exception as e:
                            pass
                
            except ImportError as e:
                pass
            except Exception as e:
                pass
            
            liai_api_keys = []
            for i in range(1, 6):  # 读取LIAI_API_KEY_1到LIAI_API_KEY_5
                key_name = f"LIAI_API_KEY_{i}"
                key = os.getenv(key_name)
                if key:
                    liai_api_keys.append(key)
            
            if not liai_api_keys:
                st.error("❌ 未找到Liai API密钥配置，请检查环境变量")
                return
            
            # 随机选择一个API密钥
            auto_api_key = random.choice(liai_api_keys)
            api_key = auto_api_key  # 直接使用自动选择的密钥
        elif api_provider == "Volces":
            # 火山引擎从环境变量读取（无需显示任何提示）
            import os
            ark_keys = [os.getenv(f"ARK_API_KEY_{i}") for i in range(1, 6)]
            valid_keys = [key for key in ark_keys if key]
            
            if not valid_keys:
                st.error("❌ 未找到火山引擎API密钥配置，请检查环境变量ARK_API_KEY_1到ARK_API_KEY_5")
                return
            
            # 使用第一个可用密钥（实际轮询由ai_page_splitter处理）
            api_key = valid_keys[0]
        else:  # 其他平台需要用户输入
            placeholder_text = "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
            help_text = f"通过{api_provider}平台访问AI模型，API密钥不会被保存"
        
        # 只有需要用户输入密钥的情况才显示密钥输入框
        needs_user_input = api_provider not in ["Liai", "Volces"]
        
        if needs_user_input:
            api_key = st.text_input(
                f"请输入您的{api_provider} API密钥",
                type="password",
                placeholder=placeholder_text,
                help=help_text
            )
    with col2:
        # 只有需要用户输入密钥的情况才显示获取密钥链接
        if needs_user_input:
            st.markdown("**获取API密钥**")
            st.markdown(f"[🔗 {api_provider}平台]({api_key_url})")
        
        # API密钥测试按钮（只有需要用户输入时才显示）
        if needs_user_input and api_key and api_key.strip():
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
                        if hasattr(e, 'status_code'):
                            status_code = e.status_code
                            if status_code == 401:
                                st.error("❌ API认证失败 (401): API密钥无效")
                            elif status_code == 402:
                                st.error("❌ API付费限制 (402): 账户余额不足")
                            elif status_code == 429:
                                st.error("❌ API请求频率限制 (429): 请求过于频繁")
                            else:
                                st.error(f"❌ API错误 ({status_code}): 这是API服务的问题")
                        elif "authentication" in error_msg.lower() or "unauthorized" in error_msg.lower():
                            st.error("❌ API密钥认证失败，请检查密钥是否正确")
                        elif "network" in error_msg.lower() or "connection" in error_msg.lower():
                            st.error("❌ 网络连接异常，请检查网络连接")
                        else:
                            st.error("❌ API调用异常，这不是应用程序的问题")
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
            - DeepSeek V3：火山引擎先进模型，非保密场景推荐
            - Liai Chat：保密信息专用模型，安全可靠
            """)
        
        with col2:
            st.markdown("""
            **第二步：准备API密钥** 🔑
            - 根据选择的模型注册相应平台账号
            - OpenAI/Liai平台获取API密钥
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
    # 只对需要用户输入的API提供商进行格式验证
    if needs_user_input and api_key:
        if not api_key.startswith('sk-'):
            st.markdown('<div class="warning-box">⚠️ API密钥格式可能不正确，通常以"sk-"开头</div>', unsafe_allow_html=True)
            return
    # elif api_provider == "Liai":
    #     # Liai API密钥格式检查已移除，直接通过格式验证
    
    # 跳过系统默认模板检查，直接使用Dify API和模板库
    # 注释掉原有的模板检查，改为检查模板库是否可用
    templates_dir = os.path.join(os.getcwd(), "templates", "ppt_template")
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
    
    st.markdown('<div class="success-box">✅ AI助手已准备就绪！可以使用智能PPT生成功能</div>', unsafe_allow_html=True)
    
    # 功能选择选项卡
    st.markdown("---")
    # 仅保留核心入口功能
    tab1, tab3, tab_format = st.tabs(["🎨 智能PPT生成", "🧪 自定义模板测试", "🔍 PPT格式读取展示"])
    
    with tab1:
        # 智能PPT生成功能 - AI分页 + 模板匹配
        st.markdown("### 🚀 智能PPT生成 (AI分页 + 智能模板匹配)")
        
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
            st.markdown('<div class="info-box">🎯 <strong>完整AI处理流程</strong><br>此功能使用AI智能分页与模板匹配：<br>1. 用户输入长文本<br>2. AI模型智能分页（DeepSeek V3/Liai Chat）<br>3. 每页内容调用AI模型获取对应模板<br>4. 系统自动整合所有模板页面为完整PPT<br>5. 用户直接下载完整的PPT文件</div>', unsafe_allow_html=True)
    
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
            help="AI将分析文本结构进行智能分页，每页内容调用AI模型获取对应模板"
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
                help="设置为0时AI自动判断，手动设置时最少3页（封面+目录+结尾）"
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
                "🚀 开始生成PPT（AI分页 + 智能模板匹配 + 自动整合）",
                type="primary",
                use_container_width=True,
                disabled=not user_text.strip(),
                help="AI分页 → 智能模板匹配 → 自动整合PPT → 可直接下载"
            )
    
        # 处理逻辑 - AI分页 + 智能模板匹配
        if process_button and user_text.strip():
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # 步骤1：AI智能分页
                status_text.text("🤖 AI正在分析文本结构并进行智能分页...")
                progress_bar.progress(20)
                
                from ai_page_splitter import AIPageSplitter
                page_splitter = AIPageSplitter(api_key)
                # 验证页面数设置：手动设置时最少3页（封面+目录+结尾）
                if target_pages > 0 and target_pages < 3:
                    st.error("❌ 页面数量不能少于3页（封面页+目录页+结尾页）")
                    return
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
                
                # 步骤2：为每页内容调用AI模型获取模板
                status_text.text("🔗 正在为每页内容调用AI模型获取对应模板...")
                progress_bar.progress(40)
                
                # 检查Dify API密钥配置
                dify_valid, dify_count, dify_message = check_dify_api_keys()
                if not dify_valid:
                    st.error(dify_message)
                    return
                elif dify_message:  # 有警告消息
                    st.warning(dify_message)
                
                from dify_template_bridge import sync_test_dify_template_bridge
                from dify_api_client import BatchProcessor, DifyAPIConfig
                
                # 检查是否启用分批处理（超过5页时自动启用）
                if len(pages) > 5:
                    # 检查当前API提供商
                    current_model_info = st.session_state.get('selected_model_info', {})
                    api_provider = current_model_info.get('api_provider', 'OpenAI')
                    
                    if api_provider == "Liai":
                        st.info(f"📦 检测到{len(pages)}页内容，自动启用Liai分批处理模式（每批5页，负载均衡5个API密钥）")
                        # 使用Liai分批处理
                        try:
                            # 准备Liai批处理的页面数据
                            liai_pages_data = []
                            for i, page in enumerate(pages):
                                page_content = page.get('original_text_segment', '')
                                if not page_content:
                                    title = page.get('title', '')
                                    key_points = page.get('key_points', [])
                                    page_content = f"{title}\n\n" + "\n".join(key_points)
                                
                                page_type = page.get('page_type', 'content')
                                page_number = page.get('page_number', i + 1)
                                
                                # 跳过特殊页面，只处理需要AI分析的内容页
                                if page_type not in ['title', 'ending'] and page_number != 1 and len(pages) > 1:
                                    liai_pages_data.append({
                                        'page_number': page_number,
                                        'content': page_content,
                                        'ppt_structure': {'slides': [{'placeholders': {}}]},  # 简化的结构
                                        'page_data': page
                                    })
                            
                            # 如果有需要处理的页面，使用Liai批处理
                            if liai_pages_data:
                                st.info(f"🔄 开始Liai分批处理{len(liai_pages_data)}个内容页面，每批5个...")
                                
                                # 创建AI处理器并进行批处理
                                ai_processor = AIProcessor(api_key.strip())
                                batch_results = ai_processor.batch_analyze_pages_for_liai(liai_pages_data, 5)
                                
                                # 处理批处理结果
                                page_results = []
                                for result in batch_results:
                                    if result.get('success'):
                                        st.success(f"✅ 第{result['page_number']}页：Liai分析完成")
                                        page_results.append({
                                            'page_number': result['page_number'],
                                            'content': result['content'],
                                            'template_number': 'liai_analyzed',
                                            'template_path': None,
                                            'template_filename': 'Liai智能分析',
                                            'dify_response': str(result.get('analysis_result', ''))[:200] + '...',
                                            'processing_time': result.get('processing_time', 0),
                                            'is_title_page': False
                                        })
                                    else:
                                        st.error(f"❌ 第{result['page_number']}页失败: {result.get('error')}")
                                        
                                st.success(f"🎉 Liai分批处理完成！成功处理{len([r for r in batch_results if r.get('success')])}页")
                                
                        except Exception as e:
                            st.error(f"Liai分批处理出错: {str(e)}")
                            page_results = []
                    else:
                        st.info(f"📦 检测到{len(pages)}页内容，自动启用分批处理模式（每批5页）")
                    
                    # 使用Dify分批处理（仅当不是Liai时）
                    if api_provider != "Liai":
                        try:
                            dify_config = DifyAPIConfig()
                            dify_config.batch_size = 5  # 每批5个
                            
                            page_results = []
                            batch_index = 0
                            
                            # 准备需要调用API的页面（排除title和ending页）
                            api_pages = []
                            for i, page in enumerate(pages):
                                page_content = page.get('original_text_segment', '')
                                if not page_content:
                                    title = page.get('title', '')
                                    key_points = page.get('key_points', [])
                                    page_content = f"{title}\n\n" + "\n".join(key_points)
                                
                                page_type = page.get('page_type', 'content')
                                page_number = page.get('page_number', i + 1)
                            
                                # 特殊页面处理
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
                                    st.info(f"📋 第{page_number}页(封面页)：使用固定标题模板")
                                elif page_type == 'table_of_contents':
                                    toc_template_path = page.get('template_path', os.path.join("templates", "table_of_contents_slides.pptx"))
                                    page_results.append({
                                        'page_number': page_number,
                                        'content': page_content,
                                        'template_number': 'table_of_contents',
                                        'template_path': toc_template_path,
                                        'template_filename': "table_of_contents_slides.pptx",
                                        'dify_response': '目录页使用提取的内容页标题动态生成',
                                        'processing_time': 0,
                                        'is_toc_page': True
                                    })
                                    st.info(f"📑 第{page_number}页(目录页)：使用提取的内容页标题")
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
                                    st.info(f"🔚 第{page_number}页(结尾页)：使用固定结尾模板")
                                elif page_content:
                                    # 需要调用API的页面
                                    api_pages.append({
                                        'page_index': i,
                                        'page_data': page,
                                        'page_content': page_content,
                                        'page_number': page_number
                                    })
                            
                            # 分批处理API调用
                            if api_pages:
                                st.info(f"🔄 开始分批处理{len(api_pages)}个页面，每批5个...")
                                
                                # 创建进度跟踪
                                batch_progress = st.progress(0)
                                batch_status = st.empty()
                                
                                total_batches = (len(api_pages) + 4) // 5  # 向上取整
                                
                                for batch_start in range(0, len(api_pages), 5):
                                    batch_end = min(batch_start + 5, len(api_pages))
                                    batch_pages = api_pages[batch_start:batch_end]
                                    batch_index += 1
                                    
                                    batch_status.text(f"🔄 处理第{batch_index}/{total_batches}批（{len(batch_pages)}页）...")
                                    
                                    # 处理当前批次
                                    for page_info in batch_pages:
                                        # 合并title和content作为完整输入
                                        page_title = page_info['page_data'].get('title', '')
                                        page_content = page_info['page_content']
                                        full_content = f"标题: {page_title}\n\n{page_content}" if page_title else page_content
                                        
                                        bridge_result = sync_test_dify_template_bridge(full_content)
                                        
                                        # 如果成功且有title，强制添加title占位符填充
                                        if bridge_result.get('success') and page_title:
                                            step_3_result = bridge_result.get('step_3_template_fill', {})
                                            if step_3_result.get('success'):
                                                assignments = step_3_result.get('assignments', {}).get('assignments', [])
                                                # 直接添加title占位符填充（内容页都有title占位符）
                                                assignments.append({
                                                    'action': 'replace_placeholder',
                                                    'slide_index': 0,
                                                    'placeholder': 'title',
                                                    'content': page_title,
                                                    'reason': '自动填充页面标题'
                                                })
                                        
                                        if bridge_result.get('success'):
                                            dify_result = bridge_result["step_1_dify_api"]
                                            template_result = bridge_result["step_2_template_lookup"]
                                            page_results.append({
                                                'page_number': page_info['page_number'],
                                                'content': page_info['page_content'],
                                                'template_number': dify_result.get('template_number'),
                                                'template_path': template_result.get('file_path'),
                                                'template_filename': template_result.get('filename'),
                                                'dify_response': dify_result.get('response_text', ''),
                                                'processing_time': bridge_result.get('processing_time', 0),
                                                'is_title_page': False
                                            })
                                            st.success(f"✅ 第{page_info['page_number']}页：模板{dify_result.get('template_number')}")
                                        else:
                                            st.error(f"❌ 第{page_info['page_number']}页失败: {bridge_result.get('error')}")
                                            page_results.append({
                                                'page_number': page_info['page_number'],
                                                'content': page_info['page_content'],
                                                'template_number': None,
                                                'template_path': None,
                                                'template_filename': None,
                                                'dify_response': f"错误: {bridge_result.get('error')}",
                                                'processing_time': bridge_result.get('processing_time', 0),
                                                'is_title_page': False,
                                                'error': True
                                            })
                                    
                                    # 更新进度
                                    progress = batch_index / total_batches
                                    batch_progress.progress(progress)
                                    
                                    # 批次间延迟
                                    if batch_index < total_batches:
                                        batch_status.text(f"⏳ 批次间等待{dify_config.batch_delay}秒...")
                                        import time
                                        time.sleep(dify_config.batch_delay)
                                
                                # 清理进度显示
                                batch_progress.empty()
                                batch_status.empty()
                                
                                st.success(f"✅ 分批处理完成！共处理{len(api_pages)}个API页面，分{total_batches}批")
                        
                        except Exception as e:
                            st.error(f"❌ 分批处理异常: {str(e)}")
                            st.info("🔄 降级到逐页处理模式...")
                            # 降级到原来的逐页处理
                            page_results = []
                            for i, page in enumerate(pages):
                                # 原来的逐页处理逻辑
                                page_content = page.get('original_text_segment', '')
                                if not page_content:
                                    title = page.get('title', '')
                                    key_points = page.get('key_points', [])
                                    page_content = f"{title}\n\n" + "\n".join(key_points)
                                
                                page_type = page.get('page_type', 'content')
                                page_number = page.get('page_number', i + 1)
                                
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
                                    st.info(f"📋 第{page_number}页(封面页)：使用固定标题模板")
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
                                    st.info(f"🔚 第{page_number}页(结尾页)：使用固定结尾模板")
                                elif page_content:
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
                                        return
                else:
                    # 页面数少于等于5页，使用原来的逐页处理
                    st.info(f"📄 页面数较少（{len(pages)}页），使用标准处理模式")
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
                        
                        # 目录页直接使用 table_of_contents_slides.pptx，不调用Dify API
                        elif page_type == 'table_of_contents':
                            toc_template_path = page.get('template_path', os.path.join("templates", "table_of_contents_slides.pptx"))
                            page_results.append({
                                'page_number': page_number,
                                'content': page_content,
                                'template_number': 'table_of_contents',
                                'template_path': toc_template_path,
                                'template_filename': "table_of_contents_slides.pptx",
                                'dify_response': '目录页使用提取的内容页标题动态生成',
                                'processing_time': 0,
                                'is_toc_page': True
                            })
                            st.info(f"📑 第{page_number}页(目录页)：使用提取的内容页标题")
                        
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
                
                # 步骤3：文本填充（新增）
                status_text.text("📝 正在对每个模板进行智能文本填充...")
                progress_bar.progress(70)
                
                filled_page_results = []
                from pptx import Presentation
                
                # 导入PPT处理器（AIProcessor已在文件顶部导入）
                from utils import PPTProcessor
                
                for i, page_result in enumerate(page_results):
                    try:
                        template_path = page_result.get('template_path')
                        page_content = page_result.get('content', '')
                        page_number = page_result.get('page_number', i+1)
                        
                        if template_path and os.path.exists(template_path):
                            # 加载模板
                            template_prs = Presentation(template_path)
                            
                            # 检查是否为结尾页（只有结尾页完全跳过文本填充）
                            if (page_result.get('is_ending_page') or page_result.get('page_type') == 'ending'):
                                # 结尾页直接使用模板，不进行文本填充
                                fill_results = []
                                print(f"🔍 跳过结尾页文本填充: 第{page_number}页")
                                st.info(f"ℹ️ 第{page_number}页(结尾页)：使用固定模板")
                            else:
                                # 创建PPT处理器并进行文本填充
                                print(f"🔍 开始文本填充: 第{page_number}页 - {page_result.get('page_type', 'content')}")
                                print(f"📄 页面内容长度: {len(page_content)}字")
                                st.info(f"🔄 第{page_number}页：开始AI文本填充分析...")
                                
                                processor = PPTProcessor(template_prs)
                                
                                # 使用完整的文本填充流程（会自动使用当前选择的AI模型）
                                # 1. 创建AI处理器来分析文本并生成分配方案
                                ai_processor = AIProcessor()
                                
                                # 2. 分析PPT结构
                                try:
                                    print(f"🔍 分析PPT结构...")
                                    print(f"📁 模板路径: {template_path}")
                                    print(f"📑 模板slides数量: {len(template_prs.slides)}")
                                    
                                    ppt_structure = PPTAnalyzer.analyze_ppt_structure(template_prs)
                                    # 从slides中收集所有占位符
                                    all_placeholders = {}
                                    for slide in ppt_structure.get('slides', []):
                                        all_placeholders.update(slide.get('placeholders', {}))
                                    
                                    print(f"📊 检测到占位符数量: {len(all_placeholders)}")
                                    if all_placeholders:
                                        print(f"🔍 占位符列表: {list(all_placeholders.keys())}")
                                    else:
                                        print(f"⚠️ 未检测到任何占位符，模板可能没有{{placeholder}}格式的内容")
                                    
                                    # 3. 生成文本分配方案
                                    print(f"🤖 调用AI生成分配方案...")
                                    assignments = ai_processor.analyze_text_for_ppt(page_content, ppt_structure)
                                    print(f"📋 生成分配方案数量: {len(assignments.get('assignments', []))}")
                                    
                                    # 4. 应用分配方案
                                    print(f"✏️ 应用分配方案...")
                                    fill_results = processor.apply_assignments(assignments, page_content)
                                    print(f"✅ 文本填充完成，结果数量: {len(fill_results)}")
                                except Exception as fill_error:
                                    print(f"❌ 文本填充过程异常: {fill_error}")
                                    st.error(f"文本填充过程异常: {fill_error}")
                                    fill_results = []
                            
                            # 更新结果信息（为合并器保存临时文件）
                            filled_result = page_result.copy()
                            filled_result['fill_results'] = fill_results
                            
                            # 为所有页面保存临时文件用于合并（确保合并器能正确处理）
                            import tempfile
                            temp_dir = tempfile.gettempdir()
                            filled_temp_path = os.path.join(temp_dir, f"filled_temp_{page_number}_{os.path.basename(template_path)}")
                            template_prs.save(filled_temp_path)
                            filled_result['template_path'] = filled_temp_path  # 使用处理后的临时文件路径
                            
                            filled_page_results.append(filled_result)
                            
                            st.success(f"✅ 第{page_number}页：文本填充完成")
                        else:
                            # 没有模板的页面直接传递
                            filled_page_results.append(page_result)
                            st.info(f"ℹ️ 第{page_number}页：无需填充")
                            
                    except Exception as e:
                        st.error(f"❌ 第{page_result.get('page_number', i+1)}页文本填充失败: {e}")
                        # 失败时使用原始模板
                        filled_page_results.append(page_result)
                
                # 步骤4：清理未填充的占位符
                status_text.text("🧹 正在清理未填充的占位符...")
                progress_bar.progress(75)
                
                # 对每个填充后的页面调用现有的清理功能
                for filled_result in filled_page_results:
                    if filled_result.get('template_path') and os.path.exists(filled_result['template_path']):
                        try:
                            # 加载已填充的模板并创建临时生成器实例用于清理
                            filled_prs = Presentation(filled_result['template_path'])
                            
                            # 创建临时的UserPPTGenerator实例来调用清理方法
                            temp_generator = UserPPTGenerator(api_key)
                            temp_generator.presentation = filled_prs
                            temp_generator.ppt_processor = PPTProcessor(filled_prs)
                            
                            # 调用清理功能
                            cleanup_results = temp_generator.cleanup_unfilled_placeholders()
                            
                            # 保存清理后的结果
                            filled_prs.save(filled_result['template_path'])
                            
                        except Exception as cleanup_error:
                            print(f"⚠️ 第{filled_result.get('page_number')}页占位符清理失败: {cleanup_error}")
                
                # 步骤5：整合PPT页面
                status_text.text("🔗 正在整合填充后的PPT页面...")
                progress_bar.progress(80)
                
                # 保存填充后的页面结果到session state
                st.session_state.current_page_results = filled_page_results
                st.session_state.current_pages = pages
                
                # 自动执行PPT整合
                try:
                    # 使用增强版合并器，自动选择最佳方法
                    from ppt_merger import merge_dify_templates_to_ppt_enhanced
                    status_text.text("🔗 正在整合PPT页面(增强格式保留)...")
                    progress_bar.progress(90)
                    merge_result = merge_dify_templates_to_ppt_enhanced(filled_page_results)
                    
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
                            
                            # 分析占位符 - 支持文本框和表格中的占位符
                            total_placeholders = 0
                            placeholder_info = []
                            
                            for i, slide in enumerate(temp_presentation.slides):
                                slide_placeholders = []
                                table_placeholders = []
                                
                                for shape in slide.shapes:
                                    # 处理普通文本框中的占位符
                                    if hasattr(shape, 'text') and shape.text:
                                        import re  
                                        placeholders = re.findall(r'\{([^}]+)\}', shape.text)
                                        if placeholders:
                                            slide_placeholders.extend(placeholders)
                                            total_placeholders += len(placeholders)
                                    
                                    # 处理表格中的占位符
                                    elif hasattr(shape, 'shape_type') and shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE = 19
                                        table = shape.table
                                        for row_idx, row in enumerate(table.rows):
                                            for col_idx, cell in enumerate(row.cells):
                                                cell_text = cell.text.strip()
                                                if cell_text:
                                                    placeholders = re.findall(r'\{([^}]+)\}', cell_text)
                                                    if placeholders:
                                                        for placeholder in placeholders:
                                                            table_placeholders.append(f"{placeholder}(表格{row_idx+1},{col_idx+1})")
                                                            total_placeholders += 1
                                
                                # 合并文本框和表格占位符
                                all_slide_placeholders = slide_placeholders + table_placeholders
                                if all_slide_placeholders:
                                    placeholder_info.append({
                                        'slide_num': i + 1,
                                        'placeholders': slide_placeholders,
                                        'table_placeholders': table_placeholders,
                                        'total_count': len(all_slide_placeholders)
                                    })
                            
                            st.metric("🎯 发现占位符", total_placeholders)
                            
                            # 显示占位符详情
                            if placeholder_info:
                                with st.expander("🔍 模板结构分析", expanded=False):
                                    for info in placeholder_info[:5]:  # 只显示前5页
                                        slide_num = info['slide_num']
                                        text_placeholders = info['placeholders']
                                        table_placeholders = info['table_placeholders']
                                        
                                        st.write(f"**第{slide_num}页（共{info['total_count']}个占位符）：**")
                                        
                                        if text_placeholders:
                                            st.write(f"  📝 文本框：{', '.join([f'{{{p}}}' for p in text_placeholders])}")
                                        
                                        if table_placeholders:
                                            st.write(f"  📊 表格：{', '.join([f'{{{p}}}' for p in table_placeholders])}")
                                    
                                    if len(placeholder_info) > 5:
                                        remaining_pages = len(placeholder_info) - 5
                                        remaining_placeholders = sum(info['total_count'] for info in placeholder_info[5:])
                                        st.write(f"... 还有 {remaining_pages} 页包含 {remaining_placeholders} 个占位符（包括表格占位符）")
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

这是使用自定义模板的测试内容。AI将分析您的文本结构，并智能地将内容分配到模板中的各个占位符位置。AI能够理解所有{}格式的占位符含义。

主要特点：
- 支持自定义PPT模板上传
- 智能文本内容分配
- 保持原有模板设计风格
- 独立于其他功能模块

测试结果将展示AI如何理解您的内容并填充到模板的对应位置。""",
                    help="AI将分析您的文本并智能分配到模板的所有{}格式占位符中",
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
                                
                                # 显示清理详情
                                if cleanup_results and cleanup_results.get('cleaned_placeholder_list'):
                                    with st.expander("🔍 查看清理详情", expanded=False):
                                        st.write("**已清理的未填充占位符：**")
                                        for item in cleanup_results['cleaned_placeholder_list']:
                                            st.text(f"• {item}")
                                        st.info("💡 已填充的占位符保持不变")
                                elif cleanup_count == 0:
                                    st.success("✅ 所有占位符都已被正确填充")
                            
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
                            
                            # 添加调试信息展示
                            with st.expander("🔧 调试信息：占位符填充状态", expanded=False):
                                if hasattr(custom_generator.ppt_processor, 'filled_placeholders'):
                                    filled_info = custom_generator.ppt_processor.filled_placeholders
                                    if filled_info:
                                        st.write("**已成功填充的占位符：**")
                                        for slide_idx, placeholders in filled_info.items():
                                            if placeholders:
                                                st.write(f"第{slide_idx+1}页: {', '.join([f'{{{p}}}' for p in placeholders])}")
                                        
                                        # 显示分配方案
                                        if 'assignments' in assignments and assignments['assignments']:
                                            st.write("**AI分配方案：**")
                                            for i, assignment in enumerate(assignments['assignments'][:5]):  # 只显示前5个
                                                slide_num = assignment.get('slide_index', 0) + 1
                                                placeholder = assignment.get('placeholder', '')
                                                content = assignment.get('content', '')[:50]
                                                reason = assignment.get('reason', '')
                                                st.write(f"{i+1}. 第{slide_num}页 `{{{placeholder}}}` → {content}{'...' if len(assignment.get('content', '')) > 50 else ''}")
                                                if reason:
                                                    st.caption(f"   理由: {reason}")
                                            
                                            if len(assignments['assignments']) > 5:
                                                st.write(f"... 还有 {len(assignments['assignments']) - 5} 个分配方案")
                                    else:
                                        st.warning("⚠️ 没有占位符被成功填充，请检查模板格式和内容匹配度")
                                else:
                                    st.error("❌ 无法获取填充状态信息")
                            
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
            
            st.markdown('<div class="warning-box">💡 <strong>提示：</strong> 请确保您的PPT模板中包含形如 {标题}、{内容}、{要点}、{作者}、{日期}、{描述} 等占位符。AI将根据占位符的名称自动理解其含义并智能分配相应的内容。支持所有{}格式的占位符，包括文本框和表格单元格中的占位符。</div>', unsafe_allow_html=True)
    
    with tab_format:
        # PPT格式读取展示功能
        st.markdown("### 🔍 PPT格式读取展示")
        st.markdown("**上传一个PPT文件，查看我们的格式读取功能能识别到什么信息**")
        
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.markdown("#### 📤 上传文件")
            uploaded_file = st.file_uploader(
                "选择PPT文件",
                type=['pptx'],
                help="支持.pptx格式的PowerPoint文件"
            )
            
            if uploaded_file is not None:
                st.success(f"✅ 已上传：{uploaded_file.name}")
                
                # 分析按钮
                if st.button("🔍 开始分析格式", type="primary"):
                    with st.spinner("正在分析PPT格式..."):
                        try:
                            # 保存上传的文件到临时位置
                            import tempfile
                            with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as temp_file:
                                temp_file.write(uploaded_file.getbuffer())
                                temp_path = temp_file.name
                            
                            # 使用现有的PPT分析功能
                            from pptx import Presentation as PptxPresentation
                            presentation = PptxPresentation(temp_path)
                            ppt_structure = PPTAnalyzer.analyze_ppt_structure(presentation)
                            
                            # 将结果存储到session state，包括临时文件路径用于后续格式提取
                            st.session_state.format_analysis_result = {
                                'filename': uploaded_file.name,
                                'structure': ppt_structure,
                                'temp_path': temp_path,
                                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            }
                            
                            # 注意：临时文件暂时保留，用于后续格式提取
                            # 文件会在清除结果或会话结束时清理
                                
                            st.success("🎉 分析完成！")
                            st.rerun()
                            
                        except Exception as e:
                            st.error(f"❌ 分析失败：{e}")
                            # 清理临时文件
                            try:
                                os.remove(temp_path)
                            except:
                                pass
        
        with col2:
            st.markdown("#### 📊 分析结果")
            
            if 'format_analysis_result' in st.session_state:
                result = st.session_state.format_analysis_result
                
                st.markdown(f"**文件名：** {result['filename']}")
                st.markdown(f"**分析时间：** {result['timestamp']}")
                
                structure = result['structure']
                total_slides = structure.get('total_slides', 0)
                total_placeholders = structure.get('total_placeholders', 0)
                
                # 基本统计
                st.markdown("---")
                st.markdown("### 📈 基本统计")
                
                metric_cols = st.columns(3)
                with metric_cols[0]:
                    st.metric("幻灯片数量", total_slides)
                with metric_cols[1]:
                    st.metric("占位符总数", total_placeholders)
                with metric_cols[2]:
                    all_placeholders = []
                    for slide in structure.get('slides', []):
                        all_placeholders.extend(slide.get('placeholders', {}).keys())
                    unique_placeholders = len(set(all_placeholders))
                    st.metric("不同占位符", unique_placeholders)
                
                # 详细信息展开
                st.markdown("---")
                st.markdown("### 🔍 详细分析")
                
                with st.expander("📋 占位符详情"):
                    for i, slide in enumerate(structure.get('slides', [])):
                        placeholders = slide.get('placeholders', {})
                        if placeholders:
                            st.markdown(f"**第 {i+1} 页：**")
                            
                            for placeholder_name, placeholder_info in placeholders.items():
                                st.markdown(f"- **{{{placeholder_name}}}**")
                                
                                # 显示类型信息
                                ph_type = placeholder_info.get('type', 'unknown')
                                st.markdown(f"  - 类型：{ph_type}")
                                
                                # 显示原始文本
                                original_text = placeholder_info.get('original_text', '')
                                if original_text:
                                    st.markdown(f"  - 原始文本：`{original_text[:100]}{'...' if len(original_text) > 100 else ''}`")
                                
                                # 实时提取字体格式信息
                                try:
                                    # 重新加载presentation来提取格式
                                    from pptx import Presentation as PptxPresentation
                                    temp_path = result.get('temp_path')
                                    if not temp_path or not os.path.exists(temp_path):
                                        st.markdown(f"  - 🎨 **格式：** ❌ 临时文件不存在")
                                        continue
                                        
                                    temp_presentation = PptxPresentation(temp_path)
                                    slide_obj = temp_presentation.slides[i]
                                    
                                    # 创建临时的PPTProcessor来提取格式
                                    from utils import PPTProcessor
                                    temp_processor = PPTProcessor(temp_presentation)
                                    
                                    # 获取容器对象
                                    container = placeholder_info.get('shape')
                                    if placeholder_info.get('type') == 'table_cell':
                                        container = placeholder_info.get('cell')
                                    
                                    # 如果没有容器信息，尝试重新查找
                                    if not container:
                                        # 在slide中查找包含这个占位符的shape
                                        placeholder_pattern = f"{{{placeholder_name}}}"
                                        for shape in slide_obj.shapes:
                                            if hasattr(shape, 'text') and placeholder_pattern in shape.text:
                                                container = shape
                                                break
                                            elif hasattr(shape, 'table'):
                                                # 检查表格
                                                for row in shape.table.rows:
                                                    for cell in row.cells:
                                                        if placeholder_pattern in cell.text:
                                                            container = cell
                                                            break
                                                    if container:
                                                        break
                                                if container:
                                                    break
                                    
                                    # 如果找到容器，提取格式信息
                                    if container:
                                        format_info = temp_processor._extract_placeholder_format(container, placeholder_name)
                                        
                                        # 格式化显示
                                        font_details = []
                                        
                                        # 字体名称
                                        font_name = format_info.get('font_name')
                                        if font_name:
                                            font_details.append(f"字体: {font_name}")
                                        else:
                                            font_details.append("字体: None")
                                        
                                        # 字体大小
                                        font_size = format_info.get('font_size')
                                        if font_size:
                                            font_details.append(f"大小: {font_size}pt")
                                        else:
                                            font_details.append("大小: None")
                                        
                                        # 字体颜色
                                        font_color = format_info.get('font_color')
                                        if font_color:
                                            font_details.append(f"颜色: {font_color}")
                                        else:
                                            font_details.append("颜色: None")
                                        
                                        # 粗体和斜体
                                        style_details = []
                                        if format_info.get('font_bold'):
                                            style_details.append("粗体")
                                        if format_info.get('font_italic'):
                                            style_details.append("斜体")
                                        
                                        if style_details:
                                            font_details.append(f"样式: {', '.join(style_details)}")
                                        else:
                                            font_details.append("样式: 普通")
                                        
                                        st.markdown(f"  - 🎨 **格式：** {' | '.join(font_details)}")
                                        
                                        # 如果有问题的格式，用颜色标出
                                        problems = []
                                        if not font_name:
                                            problems.append("字体名称")
                                        if not font_size:
                                            problems.append("字体大小")
                                        if not font_color:
                                            problems.append("字体颜色")
                                        
                                        if problems:
                                            st.markdown(f"    ⚠️ *无法读取: {', '.join(problems)}*")
                                    else:
                                        st.markdown(f"  - 🎨 **格式：** ❌ 无法定位占位符容器")
                                        
                                except Exception as format_error:
                                    st.markdown(f"  - 🎨 **格式：** ❌ 提取失败 ({str(format_error)[:50]})")
                                
                                st.markdown("")
                        else:
                            st.markdown(f"**第 {i+1} 页：** 无占位符")
                
                with st.expander("🗂️ 原始结构数据"):
                    st.json(structure, expanded=False)
                
                # 清除结果按钮
                if st.button("🗑️ 清除结果"):
                    # 清理临时文件
                    temp_path = result.get('temp_path')
                    if temp_path and os.path.exists(temp_path):
                        try:
                            os.remove(temp_path)
                        except:
                            pass
                    del st.session_state.format_analysis_result
                    st.rerun()
                    
            else:
                st.markdown("👆 请先上传PPT文件并点击分析按钮")
                
                # 功能说明
                st.markdown("---")
                st.markdown("#### 💡 功能说明")
                st.markdown("""
                **这个工具会显示：**
                
                1. **📊 基本统计**：幻灯片数量、占位符总数等
                2. **🔍 占位符详情**：每个占位符的类型、位置、原始文本
                3. **🎨 格式信息**：字体名称、大小、颜色等（如果可读取）
                4. **📂 原始数据**：完整的结构分析结果
                
                **支持的格式：**
                - 文本框中的 `{占位符}`
                - 表格单元格中的 `{占位符}`
                - 多种字体格式识别
                
                **用途：**
                - 调试模板兼容性
                - 验证占位符识别准确性
                - 了解格式读取能力的边界
                """)

    
    # 页脚信息
    st.markdown("---")
    st.markdown(
        '<div style="text-align: center; color: #666; padding: 2rem;">'
        '💡 由OpenAI API驱动 | 🎨 专业PPT自动生成'
        '</div>', 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()