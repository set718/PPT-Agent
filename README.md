# 文本转PPT填充器

使用DeepSeek AI智能将您的文本填入预设PPT模板的自动化工具。

## 功能特点

- **预设模板**：使用指定的PPT模板文件 (`D:\jiayihan\Desktop\ppt format V1_2.pptx`)
- **保持原文**：完全保留您的文本内容，不做任何修改
- **智能分析**：AI分析PPT结构和文本逻辑关系
- **合理分配**：将文本内容填入最适合的幻灯片位置
- **灵活处理**：既可更新现有幻灯片，也可新增幻灯片

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 1. 准备工作

确保您有：
- DeepSeek API密钥（从 [DeepSeek平台](https://platform.deepseek.com/api_keys) 获取）
- PPT模板文件位于指定路径：`D:\jiayihan\Desktop\ppt format V1_2.pptx`

### 2. Web界面（推荐）

启动Web界面：
```bash
python run_app.py
```

或直接运行：
```bash
streamlit run app.py
```

然后：
1. 在浏览器中打开应用（通常是 http://localhost:8501）
2. 在左侧输入您的DeepSeek API密钥
3. 确认PPT模板文件状态（应显示"✅ 模板文件存在"）
4. 在文本框中输入您的内容
5. 点击"开始处理"
6. 下载更新后的PPT文件

### 3. 命令行界面

```bash
# 设置API密钥环境变量
export DEEPSEEK_API_KEY=your_api_key_here

# 运行命令行版本
python text_to_ppt.py
```

## 项目结构

```
AI大赛相关Code/
├── app.py              # Streamlit Web界面
├── text_to_ppt.py      # 命令行版本
├── run_app.py          # Web应用启动脚本
├── config.py           # 配置管理模块
├── utils.py            # 工具函数和共用组件
├── logger.py           # 日志管理模块
├── requirements.txt    # 项目依赖
├── config.json         # 配置文件（可选）
└── README.md           # 项目说明
```

## 模板配置

### 默认模板路径
```
D:\jiayihan\Desktop\ppt format V1_2.pptx
```

### 配置方式
您可以通过以下方式修改模板路径：

1. **修改config.py** (推荐)
```python
# 在config.py中修改default_ppt_template参数
default_ppt_template: str = "您的PPT模板路径"
```

2. **创建config.json配置文件** (推荐)
```json
{
  "default_ppt_template": "D:\\path\\to\\your\\template.pptx",
  "ai_temperature": 0.3,
  "ai_max_tokens": 2000
}
```

3. **环境变量**
```bash
export PPT_TEMPLATE_PATH="D:\path\to\your\template.pptx"
```

### 配置验证
程序启动时会自动验证PPT模板文件的有效性，包括：
- 文件是否存在
- 文件格式是否正确(.pptx)
- 文件是否可以正常打开
- 文件是否包含有效的幻灯片

## 适用场景

- **学术报告**：将研究内容填入学术PPT模板
- **商业计划**：将项目信息填入商业PPT格式
- **教学课件**：将课程内容填入教学PPT框架
- **工作汇报**：将数据结果填入汇报PPT模板

## 注意事项

1. 确保PPT模板文件存在于指定路径
2. 输入的文本内容会完全保持原样，AI只负责智能分配到合适位置
3. 如果模板文件不存在，应用会显示错误提示
4. 生成的PPT文件会保存在 `output/` 目录（命令行版本）或提供下载（Web版本）

## 新功能特性

### 🔧 配置管理
- **统一配置系统**：通过config.py管理所有配置项
- **多种配置方式**：支持代码配置、JSON文件配置、环境变量配置
- **配置验证**：启动时自动验证配置有效性

### 📝 日志系统
- **彩色日志**：终端输出支持彩色显示
- **日志分级**：支持DEBUG、INFO、WARNING、ERROR、CRITICAL
- **日志轮转**：自动管理日志文件大小和数量
- **性能监控**：记录函数执行时间和API调用性能

### 🛠️ 错误处理
- **详细错误信息**：提供更详细的错误描述和解决建议
- **优雅降级**：API调用失败时自动使用备用方案
- **文件验证**：启动前验证PPT文件有效性

### ⚡ 性能优化
- **代码重构**：消除重复代码，提高代码复用性
- **模块化设计**：将功能拆分为独立的工具模块
- **资源管理**：优化临时文件创建和清理

## 故障排除

### 配置问题
- **模板文件不存在**：检查config.py中的default_ppt_template路径设置
- **配置文件错误**：确认config.json格式正确，或删除该文件使用默认配置
- **路径格式**：Windows路径需要使用双反斜杠(\\)或正斜杠(/)

### API调用问题
- **API密钥错误**：检查DeepSeek API密钥是否正确设置
- **网络连接**：确认可以正常访问api.deepseek.com
- **API配额**：查看API使用配额是否充足
- **请求超时**：可以在config.py中调整ai_max_tokens参数

### PPT处理问题
- **文件格式**：确认PPT文件格式为.pptx
- **文件损坏**：尝试用Microsoft PowerPoint打开文件验证
- **占位符格式**：确认PPT中的占位符格式为{placeholder_name}
- **权限问题**：确保程序有读写PPT文件的权限

### 调试模式
```bash
# 启用调试日志
export LOG_LEVEL=DEBUG
python text_to_ppt.py
```

### 日志查看
```bash
# 查看应用日志
tail -f app.log
``` 