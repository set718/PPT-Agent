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
├── requirements.txt    # 项目依赖
└── README.md           # 项目说明
```

## 模板配置

当前配置的PPT模板路径为：
```
D:\jiayihan\Desktop\ppt format V1_2.pptx
```

如需更改模板路径，请修改：
- `app.py` 中的 `PRESET_PPT_PATH` 变量
- `text_to_ppt.py` 中 `main()` 函数里的 `ppt_path` 变量

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

## 故障排除

### 模板文件不存在
- 检查路径 `D:\jiayihan\Desktop\ppt format V1_2.pptx` 是否正确
- 确认文件确实存在于该位置
- 如果需要，可以修改代码中的路径配置

### API调用失败
- 检查DeepSeek API密钥是否正确
- 确认网络连接正常
- 查看API配额是否充足

### PPT处理错误
- 确认PPT文件格式为 `.pptx`
- 检查PPT文件是否损坏
- 尝试用较短的文本进行测试 