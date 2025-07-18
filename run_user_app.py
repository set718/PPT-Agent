#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
启动用户版Streamlit应用的便捷脚本
"""

import subprocess
import sys
import os

def main():
    """启动用户版Streamlit应用"""
    print("[INFO] 正在启动AI PPT助手 - 用户界面...")
    print("[INFO] 程序将在浏览器中自动打开")
    print("[INFO] 如果没有自动打开，请手动访问: http://localhost:8502")
    print("=" * 50)
    
    try:
        # 检查当前目录是否有user_app.py文件
        if not os.path.exists("user_app.py"):
            print("[ERROR] 错误: 未找到user_app.py文件")
            print("请确保在正确的项目目录中运行此脚本")
            sys.exit(1)
        
        # 启动Streamlit应用 (使用不同端口避免冲突)
        subprocess.run([sys.executable, "-m", "streamlit", "run", "user_app.py", "--server.port", "8502"], check=True)
        
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] 启动失败: {e}")
        print("请确保已安装streamlit: pip install streamlit")
    except KeyboardInterrupt:
        print("\n[INFO] 应用已停止")
    except Exception as e:
        print(f"[ERROR] 意外错误: {e}")

if __name__ == "__main__":
    main()