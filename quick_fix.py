#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速修復腳本 - 解決依賴問題
"""

import subprocess
import sys
import os

def main():
    print("🔧 視頻處理工具 - 快速修復腳本")
    print("=" * 50)
    
    # 檢查當前 Python 環境
    print(f"當前 Python: {sys.executable}")
    print(f"Python 版本: {sys.version}")
    print()
    
    # 建議的解決方案
    print("📋 解決 externally-managed-environment 錯誤:")
    print()
    print("方案 1: 使用 start.sh 腳本 (推薦)")
    print("   chmod +x start.sh")
    print("   ./start.sh")
    print()
    
    print("方案 2: 手動創建虛擬環境")
    print("   python3 -m venv venv")
    print("   source venv/bin/activate")
    print("   pip install opencv-python-headless moviepy scikit-image pillow python-pptx numpy")
    print("   python3 video_audio_processor.py")
    print()
    
    print("方案 3: 使用系統權限安裝 (不推薦)")
    print("   pip3 install --user opencv-python-headless moviepy scikit-image pillow python-pptx numpy")
    print()
    
    choice = input("選擇修復方案 (1/2/3) 或 q 退出: ").strip()
    
    if choice == "1":
        print("\n🚀 執行 start.sh 腳本...")
        try:
            # 確保 start.sh 有執行權限
            subprocess.run(["chmod", "+x", "start.sh"], check=True)
            # 執行 start.sh
            subprocess.run(["./start.sh"], check=True)
        except subprocess.CalledProcessError as e:
            print(f"執行失敗: {e}")
            print("請手動執行: ./start.sh")
            
    elif choice == "2":
        print("\n🔧 創建虛擬環境...")
        
        # 創建虛擬環境
        if not os.path.exists("venv"):
            print("正在創建虛擬環境...")
            try:
                subprocess.run([sys.executable, "-m", "venv", "venv"], check=True)
                print("✅ 虛擬環境創建成功")
            except subprocess.CalledProcessError as e:
                print(f"❌ 虛擬環境創建失敗: {e}")
                return
        else:
            print("✅ 虛擬環境已存在")
        
        print("\n📦 安裝依賴套件...")
        
        # 確定虛擬環境中的 pip 路徑
        if os.name == 'nt':  # Windows
            venv_pip = os.path.join("venv", "Scripts", "pip")
            venv_python = os.path.join("venv", "Scripts", "python")
        else:  # macOS/Linux
            venv_pip = os.path.join("venv", "bin", "pip")
            venv_python = os.path.join("venv", "bin", "python")
        
        packages = [
            "opencv-python-headless", 
            "moviepy", 
            "scikit-image", 
            "pillow", 
            "python-pptx", 
            "numpy"
        ]
        
        try:
            subprocess.run([venv_pip, "install", "--upgrade", "pip"], check=True)
            subprocess.run([venv_pip, "install"] + packages, check=True)
            print("✅ 依賴套件安裝成功")
            print("\n🚀 啟動應用程式...")
            subprocess.run([venv_python, "video_audio_processor.py"])
        except subprocess.CalledProcessError as e:
            print(f"❌ 安裝失敗: {e}")
            print("\n請手動執行:")
            print("source venv/bin/activate")
            print(f"pip install {' '.join(packages)}")
            print("python video_audio_processor.py")
            
    elif choice == "3":
        print("\n⚠️  使用 --user 標誌安裝...")
        packages = [
            "opencv-python-headless", 
            "moviepy", 
            "scikit-image", 
            "pillow", 
            "python-pptx", 
            "numpy"
        ]
        
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "--user"] + packages, check=True)
            print("✅ 套件安裝成功")
            print("\n🚀 啟動應用程式...")
            subprocess.run([sys.executable, "video_audio_processor.py"])
        except subprocess.CalledProcessError as e:
            print(f"❌ 安裝失敗: {e}")
            
    elif choice.lower() == "q":
        print("👋 退出修復腳本")
        return
    else:
        print("❌ 無效選擇")

if __name__ == "__main__":
    main()