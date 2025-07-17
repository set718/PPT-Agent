#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
启动Streamlit应用的便捷脚本
"""

import subprocess
import sys
import os

def main():
    """启动Streamlit应用"""
    print("🚀 正在启动文本转PPT填充器 Web界面...")
    print("📱 程序将在浏览器中自动打开")
    print("🔗 如果没有自动打开，请手动访问: http://localhost:8501")
    print("=" * 50)
    
    try:
        # 检查当前目录是否有app.py文件
        if not os.path.exists("app.py"):
            print("❌ 错误: 未找到app.py文件")
            print("请确保在正确的项目目录中运行此脚本")
            sys.exit(1)
        
        # 启动Streamlit应用
        subprocess.run([sys.executable, "-m", "streamlit", "run", "app.py"], check=True)
        
    except subprocess.CalledProcessError as e:
        print(f"❌ 启动失败: {e}")
        print("请确保已安装streamlit: pip install streamlit")
    except KeyboardInterrupt:
        print("\n👋 应用已停止")
    except Exception as e:
        print(f"❌ 意外错误: {e}")

if __name__ == "__main__":
    main() 