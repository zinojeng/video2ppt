#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
簡單的測試腳本，用於檢查 moviepy 是否可以正確工作
"""

import sys

try:
    print("正在嘗試導入 moviepy...")
    import moviepy
    from moviepy import VideoFileClip
    print("成功導入 VideoFileClip！")
    
    # 顯示 moviepy 的版本
    print(f"moviepy 版本: {moviepy.__version__}")
    
    # 測試 VideoFileClip 類是否可用
    print("VideoFileClip 類資訊:", VideoFileClip.__doc__)
    
    print("moviepy 導入測試成功！")
    
except ImportError as e:
    print(f"導入錯誤: {e}")
    print(f"Python 路徑: {sys.executable}")
    print(f"sys.path: {sys.path}")
    
except Exception as e:
    print(f"發生其他錯誤: {e}")

print("測試完成") 