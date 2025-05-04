#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Chrome捕獲工具快速啟動腳本
"""

import sys
import tkinter as tk
from tkinter import messagebox
import traceback
import os
import subprocess
import importlib


def check_dependencies():
    """檢查必要依賴是否已安裝"""
    required_packages = [
        "selenium", "opencv-python", "numpy", "pillow", 
        "python-pptx", "scikit-image", "pyautogui", "webdriver-manager"
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            # 處理特殊套件名稱
            import_name = package.replace("-", "_")
            # 處理 opencv-python 特例
            if package == "opencv-python":
                import_name = "cv2"
            # 處理 python-pptx 特例
            elif package == "python-pptx":
                import_name = "pptx"
            # 處理 pillow 特例
            elif package == "pillow":
                import_name = "PIL"
                
            importlib.import_module(import_name)
        except ImportError:
            missing_packages.append(package)
    
    return missing_packages


def install_dependencies(packages):
    """安裝缺失的依賴"""
    print(f"正在安裝必要依賴: {', '.join(packages)}")
    
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + packages)
        return True
    except subprocess.CalledProcessError:
        return False


def main():
    # 檢查依賴
    missing_packages = check_dependencies()
    
    if missing_packages:
        print(f"缺少以下依賴: {', '.join(missing_packages)}")
        choice = input("是否自動安裝這些依賴？(y/n): ")
        
        if choice.lower() == 'y':
            if not install_dependencies(missing_packages):
                print("依賴安裝失敗，請手動執行：")
                print(f"pip install {' '.join(missing_packages)}")
                sys.exit(1)
        else:
            print("請手動安裝以下依賴後再運行:")
            print(f"pip install {' '.join(missing_packages)}")
            sys.exit(1)
    
    try:
        from chrome_capture import ChromeCapture
        
        # 創建並運行應用
        root = tk.Tk()
        app = ChromeCapture(root)
        
        # 顯示使用提示
        messagebox.showinfo(
            "使用說明", 
            "1. 輸入包含視頻的網址並點擊「打開瀏覽器」\n"
            "2. 在打開的頁面中播放視頻\n"
            "3. 框選要監控的投影片區域\n"
            "4. 點擊「開始捕獲」開始監控並截取投影片\n"
            "5. 完成後點擊「停止」並選擇「生成PPT」\n\n"
            "提示：調整相似度閾值可以控制檢測靈敏度，值越低檢測越靈敏"
        )
        
        app.run()
        
    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"啟動失敗: {str(e)}\n{error_msg}")
        messagebox.showerror("啟動錯誤", f"程序啟動失敗:\n{str(e)}")


if __name__ == "__main__":
    main() 