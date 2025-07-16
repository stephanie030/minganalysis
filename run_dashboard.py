#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
啟動 Streamlit 儀表板的 Python 腳本
"""

import subprocess
import sys
import os
from pathlib import Path

def install_requirements():
    """安裝必要的套件"""
    required_packages = [
        'streamlit',
        'plotly',
        'pandas',
        'openpyxl'
    ]
    
    print("📦 檢查並安裝必要套件...")
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package} 已安裝")
        except ImportError:
            print(f"📥 正在安裝 {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def main():
    """主程式"""
    print("="*50)
    print("🚀 啟動庫存分析動態儀表板")
    print("="*50)
    
    # 檢查資料檔案
    data_file = "銘宣爬蟲合併_0421_0708.xlsx"
    if not Path(data_file).exists():
        print(f"❌ 找不到資料檔案: {data_file}")
        print("請確認資料檔案在當前目錄中")
        return
    
    # 安裝套件
    install_requirements()
    
    # 啟動 Streamlit
    print("\n🌐 正在啟動 Streamlit 儀表板...")
    print("瀏覽器將自動開啟，如果沒有請手動開啟: http://localhost:8501")
    print("按 Ctrl+C 停止儀表板")
    print("-" * 50)
    
    try:
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", 
            "streamlit_dashboard.py",
            "--server.address", "localhost",
            "--server.port", "8501"
        ])
    except KeyboardInterrupt:
        print("\n👋 儀表板已停止")
    except Exception as e:
        print(f"❌ 啟動失敗: {e}")

if __name__ == "__main__":
    main()
