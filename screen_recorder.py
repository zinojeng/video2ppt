#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import cv2
import numpy as np
import pyautogui
import time
from datetime import datetime
import os

class ScreenRecorder:
    def __init__(self, output_path="screen_recording.mp4", fps=30.0):
        self.output_path = output_path
        self.fps = fps
        self.is_recording = False
        self.width, self.height = pyautogui.size()
        self.roi = None
        self.writer = None
        
    def set_roi(self, roi):
        """設置錄製區域 (x1, y1, x2, y2)"""
        self.roi = roi
        
    def start_recording(self):
        """開始錄製螢幕"""
        if self.roi:
            x1, y1, x2, y2 = self.roi
            width = x2 - x1
            height = y2 - y1
        else:
            x1, y1 = 0, 0
            width, height = self.width, self.height
            
        # 創建視頻寫入器
        fourcc = cv2.VideoWriter_fourcc(*'mp4v')
        self.writer = cv2.VideoWriter(self.output_path, fourcc, self.fps, (width, height))
        
        self.is_recording = True
        self.record_thread(x1, y1, width, height)
        
    def record_thread(self, x, y, width, height):
        """處理錄製邏輯"""
        start_time = time.time()
        while self.is_recording:
            # 擷取螢幕區域
            img = pyautogui.screenshot(region=(x, y, width, height))
            
            # 轉換為OpenCV格式
            frame = np.array(img)
            frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)
            
            # 寫入視頻
            self.writer.write(frame)
            
            # 保持幀率
            elapsed = time.time() - start_time
            if elapsed < 1.0/self.fps:
                time.sleep(1.0/self.fps - elapsed)
            start_time = time.time()
            
    def stop_recording(self):
        """停止錄製"""
        self.is_recording = False
        if self.writer:
            self.writer.release()
        
    def take_screenshot(self, output_path=None):
        """擷取當前螢幕截圖"""
        if not output_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"screenshot_{timestamp}.png"
            
        if self.roi:
            x1, y1, x2, y2 = self.roi
            width = x2 - x1
            height = y2 - y1
            img = pyautogui.screenshot(region=(x1, y1, width, height))
        else:
            img = pyautogui.screenshot()
            
        img.save(output_path)
        return output_path

def get_screen_dimensions():
    """獲取螢幕尺寸"""
    return pyautogui.size()

def main():
    # 測試用例
    recorder = ScreenRecorder("test_recording.mp4")
    
    # 設置錄製區域 (左上角開始，寬度400，高度300)
    width, height = get_screen_dimensions()
    recorder.set_roi((100, 100, 500, 400))
    
    print("開始錄製 5 秒鐘...")
    recorder.start_recording()
    time.sleep(5)
    recorder.stop_recording()
    print(f"錄製完成，保存在 {recorder.output_path}")
    
    # 擷取截圖
    screenshot_path = recorder.take_screenshot()
    print(f"截圖保存在 {screenshot_path}")

if __name__ == "__main__":
    main() 