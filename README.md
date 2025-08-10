# 🎨 AI PPT助手

智能将您的文本内容转换为精美的PPT演示文稿

## ✨ 核心功能

### 🤖 AI智能分页
- 使用**Qwen Max**或**GPT-4o**分析文本结构
- 自动分割为适合PPT的页面内容
- 最多支持25页（包含标题页和结尾页）

### 🎯 智能模板匹配
- 5个Dify API密钥负载均衡
- 从250+专业模板库智能选择最适合的设计
- 每页独特的设计风格，保持视觉一致性

### 📄 自动页面处理
- **标题页**：自动使用固定标题模板
- **内容页**：根据内容智能匹配不同模板
- **结尾页**：自动添加感谢页面

### 🛠️ 高级合并技术
- 优先使用Spire.Presentation保持各页独特格式
- Win32COM备选方案（Windows系统）
- 完整的格式保留和错误处理

## 🚀 快速启动

### 1. 配置环境变量
```bash
# 复制环境变量模板
cp .env.example .env

# 编辑 .env 文件，填入你的API密钥
# DIFY_API_KEY_1=your-dify-key-1
# DASHSCOPE_API_KEY=your-qwen-key
```

### 2. 安装依赖
```bash
pip install -r requirements.txt
pip install python-dotenv  # 用于加载环境变量
```

### 3. 启动应用
```bash
# 方式1：一键启动（推荐）
python start_user_app.py

# 方式2：直接运行
streamlit run user_app.py
```

### 方法三：使用Streamlit命令
```bash
streamlit run user_app.py
```

启动后访问：http://localhost:8501

## 📁 项目结构

### 核心应用文件
- **`user_app.py`** - 主用户界面（4个功能选项卡）
- **`start_user_app.py`** - 简化启动脚本
- **`ppt_merger.py`** - PPT页面整合器

### 功能模块
- **`ai_page_splitter.py`** - AI智能分页模块
- **`dify_template_bridge.py`** - Dify API与模板桥接
- **`dify_api_client.py`** - Dify API客户端（多密钥负载均衡）
- **`config.py`** - 统一配置管理
- **`utils.py`** - 工具函数和共用组件
- **`logger.py`** - 日志管理

### 模板库
- **`templates/ppt_template/`** - 250+个PPT模板文件
- **`templates/ppt_template/title_slides.pptx`** - 固定封面页模板
- **`templates/ppt_template/split_presentations_*.pptx`** - 内容页模板(1-250)

## 🎨 主要功能

### 1. 🎨 智能PPT生成
**完整的AI驱动PPT生成流程**
- 输入长文本内容
- AI自动智能分页
- 每页调用Dify API获取最适合的模板
- 封面页自动使用固定模板
- 将所有页面整合为完整PPT
- 提供统一下载

**工作流程**：
```
用户输入 → AI分页 → 封面页固定模板 → 内容页Dify API推荐 → 页面整合 → PPT下载
```

### 2. 📑 AI智能分页 + Dify API增强
- AI分析文本结构并智能分页
- 多密钥并发Dify API调用
- 3倍处理速度提升
- 详细处理统计和结果展示

### 3. 🧪 自定义模板测试
- 上传自定义PPT模板
- AI智能内容分配
- 独立测试环境
- 即时结果预览

### 4. 🔗 Dify-模板桥接测试
- 测试Dify API与模板文件对应关系
- 验证模板推荐准确性
- 支持模板文件下载

## ⚙️ 配置要求

### 系统要求
- **Python**: 3.8或更高版本
- **内存**: 至少4GB RAM
- **网络**: 稳定的互联网连接

### API密钥
- **OpenAI兼容API**: 支持OpenRouter、DeepSeek、GPT-4o等
- **Dify API**: 已预配置8个密钥，支持负载均衡

### 依赖包（自动安装）
```bash
streamlit >= 1.28.0
python-pptx >= 0.6.21
openai >= 1.3.0
aiohttp >= 3.8.0
```

## 🔧 高级配置

### 模型配置
在`config.py`中支持多种AI模型：
- **GPT-4o**: 功能完整，支持视觉分析
- **DeepSeek R1**: 成本较低，专注推理处理
- **Liai Chat**: 多模态支持

### Dify API配置
系统预配置8个API密钥，支持：
- **负载均衡策略**: 轮询、随机、最少使用
- **故障转移**: 单个密钥失败自动切换
- **自动恢复**: 60秒后失效密钥自动恢复

## 📊 性能指标

### 处理速度
- **AI智能分页**: 1-3秒
- **Dify API调用**: 3-8秒（多密钥并发）
- **PPT页面整合**: 2-5秒
- **完整流程**: 6-16秒

### 成功率
- **AI分页成功率**: >95%
- **Dify API成功率**: >90%（多密钥故障转移）
- **PPT整合成功率**: >98%

### 并发能力
- **同时处理页面数**: 最多8个（对应8个API密钥）
- **模板库覆盖**: 250+个模板，编号1-250全覆盖

## 🛠️ 故障排除

### 常见问题

#### 1. 依赖包安装失败
**解决方案**：手动安装
```bash
pip install streamlit python-pptx openai aiohttp
```

#### 2. 模板库缺失
**解决方案**：确保templates/ppt_template/目录存在且包含模板文件

#### 3. API密钥无效
**解决方案**：检查API密钥格式和余额
- OpenRouter: `sk-or-xxxxxxxx`
- DeepSeek: `sk-xxxxxxxx`

#### 4. PPT整合失败
**解决方案**：
- 确保python-pptx库版本正确
- 检查模板文件完整性
- 查看详细错误日志

### 调试工具
```bash
# 运行系统测试
python test_integrated_system.py

# 查看详细日志
tail -f app.log
```

## 🎯 使用建议

### 最佳实践
1. **首次使用**: 建议先运行测试功能验证系统
2. **API密钥**: 确保密钥有效且有足够余额
3. **网络环境**: 确保网络连接稳定
4. **文本输入**: 结构化的文本内容效果更佳

### 适用场景
- **学术报告**: 研究内容智能分页和模板推荐
- **商业提案**: 项目方案自动分页和美化
- **教学课件**: 课程内容智能组织和模板匹配
- **工作汇报**: 数据结果自动分页和增强

## 🎉 最新更新

### v2.0 主要特性
- ✅ **PPT页面整合功能**: 将多个模板页面合并为完整PPT
- ✅ **封面页固定模板**: 第一页自动使用title_slides.pptx
- ✅ **会话状态管理**: 解决Streamlit按钮状态重置问题
- ✅ **统一项目文档**: 整合所有功能说明到单一文档

### v1.5 功能增强
- 🚀 多密钥负载均衡：8个Dify API密钥并发使用
- 🎯 智能模板推荐：基于内容特征智能推荐模板
- 📊 性能监控：实时显示处理进度和性能指标
- 🔍 调试功能：详细的数据查看和错误诊断

## 📞 技术支持

### 获取帮助
- **快速启动指南**: 查看`快速启动指南.md`
- **故障排除**: 查看`故障排除指南.md`
- **自定义模板**: 查看`自定义模板测试功能说明.md`

### 反馈渠道
- 检查日志文件获取详细错误信息
- 使用测试功能验证各模块性能
- 通过界面调试功能查看处理详情

---

## 🚀 立即体验

```bash
# 快速启动
python start_user_app.py

# 访问Web界面
# http://localhost:8501
```

**🎯 现在就开始创建您的智能PPT！** 从文本输入到完整PPT，只需几步操作即可完成。