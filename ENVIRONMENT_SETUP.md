# 环境变量配置指南

## 概述
为了保护API密钥安全，本项目使用环境变量来存储敏感信息。请按照以下步骤配置：

## 配置步骤

### 1. 创建环境变量文件
复制 `.env.example` 文件并重命名为 `.env`：
```bash
cp .env.example .env
```

### 2. 填入实际API密钥
编辑 `.env` 文件，填入你的实际API密钥：

```env
# Dify API密钥配置
DIFY_API_KEY_1=your-actual-dify-key-1
DIFY_API_KEY_2=your-actual-dify-key-2
DIFY_API_KEY_3=your-actual-dify-key-3
DIFY_API_KEY_4=your-actual-dify-key-4
DIFY_API_KEY_5=your-actual-dify-key-5


```

### 3. 安装python-dotenv（可选）
为了更好地支持.env文件，建议安装：
```bash
pip install python-dotenv
```

### 4. 系统环境变量（替代方案）
如果不想使用.env文件，也可以直接设置系统环境变量：

**Windows:**
```cmd
set DIFY_API_KEY_1=your-key-here

```

**Linux/MacOS:**
```bash
export DIFY_API_KEY_1=your-key-here

```

## 安全说明
- `.env` 文件已添加到 `.gitignore` 中，不会被提交到Git仓库
- 请妥善保管你的API密钥，不要在代码中硬编码
- 如果需要在服务器部署，请使用服务器的环境变量配置

## API密钥获取
- **Dify API密钥**: 从你的Dify平台获取
