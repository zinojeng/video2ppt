#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
視窗擷取工具：專門用於擷取特定應用程式的視窗
"""

import os
import time
import numpy as np
import cv2
import pyautogui
from PIL import Image
import platform

# 根據不同操作系統導入適當的模組
system = platform.system()

if system == "Darwin":  # macOS
    try:
        from AppKit import NSWorkspace, NSRunningApplication, NSApplicationActivateIgnoringOtherApps
    except ImportError:
        print("請安裝 pyobjc: pip install pyobjc")
        
elif system == "Windows":
    import win32gui
    import win32con
    import win32ui
    from ctypes import windll

class WindowCapture:
    """視窗擷取類"""
    
    def __init__(self):
        self.window_list = []
        self.update_window_list()
    
    def update_window_list(self):
        """更新可用視窗列表"""
        self.window_list = []
        
        if system == "Darwin":  # macOS
            workspace = NSWorkspace.sharedWorkspace()
            for app in workspace.runningApplications():
                if app.isActive() and app.localizedName():
                    self.window_list.append({
                        'id': app.processIdentifier(),
                        'name': app.localizedName(),
                        'handle': app
                    })
        
        elif system == "Windows":
            def callback(hwnd, windows_list):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                    windows_list.append({
                        'id': hwnd,
                        'name': win32gui.GetWindowText(hwnd),
                        'handle': hwnd
                    })
            win32gui.EnumWindows(callback, self.window_list)
        
        return self.window_list
    
    def get_window_by_name(self, name_fragment):
        """根據名稱片段查找窗口"""
        self.update_window_list()
        for window in self.window_list:
            if name_fragment.lower() in window['name'].lower():
                return window
        return None
    
    def get_window_names(self):
        """獲取所有窗口名稱"""
        self.update_window_list()
        return [window['name'] for window in self.window_list]
    
    def capture_window(self, window_info):
        """擷取指定窗口的畫面"""
        if not window_info:
            return None
            
        if system == "Darwin":  # macOS
            # 在 macOS 上，我們使用 pyautogui 截圖並依靠啟動窗口來獲取焦點
            app = window_info['handle']
            app.activateWithOptions_(NSApplicationActivateIgnoringOtherApps)
            time.sleep(0.5)  # 給窗口時間來獲取焦點
            
            # 由於 macOS 沒有簡單的方式獲取窗口位置，我們截取整個屏幕
            screenshot = pyautogui.screenshot()
            return np.array(screenshot)
            
        elif system == "Windows":
            hwnd = window_info['handle']
            
            # 獲取窗口大小
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            
            # 將窗口帶到前面
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.2)  # 給窗口時間來獲取焦點
            
            # 創建 device context
            hwnd_dc = win32gui.GetWindowDC(hwnd)
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()
            
            # 創建 bitmap 對象
            save_bitmap = win32ui.CreateBitmap()
            save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
            save_dc.SelectObject(save_bitmap)
            
            # 複製窗口內容到 bitmap 中
            result = windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 3)
            
            # 轉換 bitmap 為 numpy array
            bmpinfo = save_bitmap.GetInfo()
            bmpstr = save_bitmap.GetBitmapBits(True)
            img = np.frombuffer(bmpstr, dtype=np.uint8).reshape((height, width, 4))
            
            # 釋放資源
            win32gui.DeleteObject(save_bitmap.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwnd_dc)
            
            # 轉換為 BGR 格式
            return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        
        return None
    
    def get_browser_names(self):
        """返回所有瀏覽器相關的窗口名稱"""
        browser_keywords = ["chrome", "edge", "firefox", "safari", "opera", "瀏覽器", "browser"]
        browser_windows = []
        
        for window in self.window_list:
            for keyword in browser_keywords:
                if keyword.lower() in window['name'].lower():
                    browser_windows.append(window['name'])
                    break
                    
        return browser_windows

# 測試代碼
if __name__ == "__main__":
    capture = WindowCapture()
    
    print("可用視窗:")
    windows = capture.get_window_names()
    for i, name in enumerate(windows):
        print(f"{i+1}. {name}")
    
    print("\n瀏覽器視窗:")
    browsers = capture.get_browser_names()
    for browser in browsers:
        print(f"- {browser}")
    
    if browsers:
        window = capture.get_window_by_name(browsers[0])
        if window:
            print(f"\n嘗試擷取視窗: {window['name']}")
            frame = capture.capture_window(window)
            if frame is not None:
                cv2.imwrite("window_capture_test.png", frame)
                print(f"已保存測試截圖 window_capture_test.png")
    else:
        print("\n未找到瀏覽器視窗") 