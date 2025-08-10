#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
启动用户界面应用
集成AI智能分页与Dify-模板桥接功能的用户界面
"""

import subprocess
import sys
import os

def main():
    """启动用户界面应用"""
    print("🎯 启动AI智能分页与Dify-模板桥接用户界面")
    print("=" * 50)
    
    # 检查user_app.py是否存在
    app_path = "user_app.py"
    if not os.path.exists(app_path):
        print(f"❌ 错误：找不到 {app_path} 文件")
        return
    
    print("✅ 找到用户界面应用文件")
    print("🚀 正在启动Streamlit应用...")
    print("📱 应用将在浏览器中自动打开")
    print("🔗 默认地址：http://localhost:8501")
    print("=" * 50)
    
    try:
        # 启动Streamlit应用
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", app_path,
            "--server.port", "8501",
            "--server.address", "localhost",
            "--browser.gatherUsageStats", "false"
        ])
    except KeyboardInterrupt:
        print("\n👋 用户界面已关闭")
    except Exception as e:
        print(f"❌ 启动失败：{str(e)}")

if __name__ == "__main__":
    main() 