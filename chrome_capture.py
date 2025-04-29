#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Chrome瀏覽器捕獲工具：專門用於通過Chrome瀏覽器擷取網頁中的視頻投影片
"""

import os
import time
import json
import argparse
import tkinter as tk
from tkinter import messagebox, Scale, filedialog
import numpy as np
import cv2
import pyautogui
from datetime import datetime
from PIL import Image, ImageTk
from skimage.metrics import structural_similarity as ssim
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import subprocess
import sys


class ChromeCapture:
    def __init__(self, root=None):
        # 創建主窗口
        if not root:
            self.root = tk.Tk()
            self.root.title("Chrome瀏覽器投影片捕獲工具")
            self.root.geometry("1000x700")
        else:
            self.root = root
        
        # 加載配置
        self.load_config()
        
        # 預設設置
        self.is_running = False
        self.is_paused = False
        self.last_frame = None
        self.slide_count = 0
        self.roi = None
        self.roi_selecting = False
        self.roi_start = None
        self.roi_end = None
        self.display_size = None
        self.browser = None
        self.last_change_time = time.time()  # 記錄最後一次檢測到變化的時間
        self.inactivity_timeout = 2 * 60  # 預設2分鐘無變化自動停止（秒）
        self.auto_stop_enabled = True  # 是否啟用自動停止功能
        
        # 確保輸出目錄存在
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
            
        # 設置界面
        self.setup_ui()
    
    def load_config(self):
        """從.cursor.json加載配置"""
        try:
            if os.path.exists('.cursor.json'):
                with open('.cursor.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    
                # 從配置中獲取Chrome捕獲設置
                if ('customSettings' in config and 
                        'chromeCapture' in config['customSettings']):
                    chrome_settings = config['customSettings']['chromeCapture']
                    self.threshold = chrome_settings.get(
                        'defaultThreshold', 0.95)
                    self.interval = chrome_settings.get(
                        'defaultInterval', 0.5)
                    self.output_format = chrome_settings.get(
                        'defaultOutputFormat', 'pptx')
                    self.auto_save = chrome_settings.get(
                        'autoSaveSlides', True)
                    self.chrome_path = chrome_settings.get(
                        'browserExecutablePath', '')
                    self.user_data_dir = chrome_settings.get(
                        'userDataDir', '')
                else:
                    self.threshold = 0.95
                    self.interval = 0.5
                    self.output_format = 'pptx'
                    self.auto_save = True
                    self.chrome_path = ''
                    self.user_data_dir = ''
                
                # 從配置中獲取捕獲設置
                if ('customSettings' in config and 
                        'captureSettings' in config['customSettings']):
                    capture_settings = (
                        config['customSettings']['captureSettings']
                    )
                    self.screenshot_method = capture_settings.get(
                        'preferredScreenShotMethod', 'pyautogui')
                    self.fallback_method = capture_settings.get(
                        'fallbackMethod', 'opencv')
                    self.similarity_algorithm = capture_settings.get(
                        'similarityAlgorithm', 'ssim')
                else:
                    self.screenshot_method = 'pyautogui'
                    self.fallback_method = 'opencv'
                    self.similarity_algorithm = 'ssim'
            else:
                # 默認值
                self.threshold = 0.95
                self.interval = 0.5
                self.output_format = 'pptx'
                self.auto_save = True
                self.chrome_path = ''
                self.user_data_dir = ''
                self.screenshot_method = 'pyautogui'
                self.fallback_method = 'opencv'
                self.similarity_algorithm = 'ssim'
        except Exception as e:
            print(f"加載配置時出錯: {str(e)}")
            # 默認值
            self.threshold = 0.95
            self.interval = 0.5
            self.output_format = 'pptx'
            self.auto_save = True
            self.chrome_path = ''
            self.user_data_dir = ''
            self.screenshot_method = 'pyautogui'
            self.fallback_method = 'opencv'
            self.similarity_algorithm = 'ssim'
            
        # 其他設置
        self.output_folder = "slides"
        self.output_file = f"slides.{self.output_format}"
    
    def setup_ui(self):
        # 上方控制區域
        control_frame = tk.Frame(self.root)
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 第一行 - 網址輸入
        row0 = tk.Frame(control_frame)
        row0.pack(fill=tk.X, pady=5)
        
        tk.Label(row0, text="網址:").pack(side=tk.LEFT, padx=5)
        self.url_entry = tk.Entry(row0, width=50)
        self.url_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.go_btn = tk.Button(
            row0, text="打開瀏覽器", command=self.launch_browser, 
            bg="#2196F3", fg="white", width=12
        )
        self.go_btn.pack(side=tk.LEFT, padx=5)
        
        # 瀏覽器選擇區域按鈕 (初始狀態為隱藏)
        self.browser_select_btn = tk.Button(
            row0, text="在瀏覽器中選擇", command=self.select_area_in_browser,
            bg="#FF9800", fg="white", width=12
        )
        # 開始時不顯示，等瀏覽器啟動後再顯示
        
        # 第二行 - 設置
        row1 = tk.Frame(control_frame)
        row1.pack(fill=tk.X, pady=5)
        
        tk.Label(row1, text="相似度閾值:").pack(side=tk.LEFT, padx=5)
        self.threshold_scale = Scale(
            row1, from_=0.5, to=1.0, resolution=0.01, 
            orient=tk.HORIZONTAL, length=200
        )
        self.threshold_scale.set(self.threshold)
        self.threshold_scale.pack(side=tk.LEFT, padx=5)
        
        tk.Label(row1, text="檢測間隔(秒):").pack(side=tk.LEFT, padx=5)
        self.interval_scale = Scale(
            row1, from_=0.1, to=5.0, resolution=0.1, 
            orient=tk.HORIZONTAL, length=200
        )
        self.interval_scale.set(self.interval)
        self.interval_scale.pack(side=tk.LEFT, padx=5)
        
        # 新增設置 - 無變化自動停止
        row1_5 = tk.Frame(control_frame)
        row1_5.pack(fill=tk.X, pady=5)
        
        tk.Label(row1_5, text="無變化自動停止(分鐘):").pack(side=tk.LEFT, padx=5)
        self.timeout_scale = Scale(
            row1_5, from_=0, to=30, resolution=1, 
            orient=tk.HORIZONTAL, length=200
        )
        self.timeout_scale.set(self.inactivity_timeout // 60)  # 轉換秒為分鐘
        self.timeout_scale.pack(side=tk.LEFT, padx=5)
        
        # 添加自動停止功能啟用/禁用選項
        self.auto_stop_var = tk.BooleanVar(value=self.auto_stop_enabled)
        self.auto_stop_checkbox = tk.Checkbutton(
            row1_5, text="啟用自動停止", variable=self.auto_stop_var,
            command=self.toggle_auto_stop
        )
        self.auto_stop_checkbox.pack(side=tk.LEFT, padx=5)
        
        # 第三行 - 按鈕
        row2 = tk.Frame(control_frame)
        row2.pack(fill=tk.X, pady=5)
        
        self.start_btn = tk.Button(
            row2, text="開始捕獲", command=self.start_capture, 
            bg="#4CAF50", fg="white", width=10
        )
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.pause_btn = tk.Button(
            row2, text="暫停", command=self.pause_capture, 
            state=tk.DISABLED, width=10
        )
        self.pause_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = tk.Button(
            row2, text="停止", command=self.stop_capture, 
            state=tk.DISABLED, width=10
        )
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.reset_roi_btn = tk.Button(
            row2, text="重設區域", command=self.reset_roi, width=10
        )
        self.reset_roi_btn.pack(side=tk.LEFT, padx=5)
        
        self.generate_btn = tk.Button(
            row2, text="生成PPT", command=self.generate_ppt, width=10
        )
        self.generate_btn.pack(side=tk.LEFT, padx=5)
        
        # 右側狀態訊息
        self.status_var = tk.StringVar(value="就緒")
        status_label = tk.Label(
            row2, textvariable=self.status_var, 
            bd=1, relief=tk.SUNKEN, anchor=tk.W
        )
        status_label.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        
        # 預覽區域
        self.canvas_frame = tk.Frame(self.root, bg="black")
        self.canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.canvas = tk.Canvas(self.canvas_frame, bg="black")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # 綁定滑鼠事件
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        
        # 日誌區域
        log_frame = tk.Frame(self.root)
        log_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(log_frame, text="執行日誌:").pack(anchor=tk.W)
        self.log_text = tk.Text(log_frame, height=6, width=50)
        self.log_text.pack(fill=tk.X, expand=True)
        
        scrollbar = tk.Scrollbar(self.log_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
        
        self.log("程序已啟動")
        self.log("請輸入網址並打開瀏覽器，然後框選要監控的投影片區域")
    
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
    
    def launch_browser(self):
        url = self.url_entry.get().strip()
        if not url:
            self.log("請輸入有效的網址")
            return
        
        # 檢查 URL 格式
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, url)
        
        try:
            # 啟動 Chrome
            if not self.browser:
                chrome_options = Options()
                chrome_options.add_argument("--start-maximized")
                self.browser = webdriver.Chrome(options=chrome_options)
            
            # 導航到網址
            self.browser.get(url)
            
            # 顯示瀏覽器預覽
            self.show_browser_preview()
            
            # 顯示瀏覽器選擇按鈕
            self.browser_select_btn.pack(side=tk.LEFT, padx=5)
            
            self.log("瀏覽器已啟動，請點擊「在瀏覽器中選擇」進行精確選擇")
        except Exception as e:
            self.log(f"啟動瀏覽器時出錯: {str(e)}")
    
    def select_area_in_browser(self):
        """在瀏覽器中直接選擇擷取區域"""
        try:
            # 這裡暫時導入，是為了在使用時才檢查是否安裝
            import pynput
        except ImportError:
            self.log("未安裝必要的套件: pynput，正在嘗試安裝...")
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", "pynput"]
                )
                self.log("安裝成功，請重新點擊「在瀏覽器中選擇」按鈕")
                return
            except Exception as e:
                self.log(f"安裝失敗: {str(e)}，請手動安裝 pynput: pip install pynput")
                return
        
        if not self.browser:
            self.log("請先啟動瀏覽器")
            return
        
        # 最小化應用程式視窗
        self.log("正在準備瀏覽器區域選擇模式...")
        self.log("請在瀏覽器視窗上按下滑鼠左鍵並拖曳來選擇區域")
        self.log("完成選擇後，應用程式將自動恢復")
        
        # 給使用者時間準備
        self.root.after(1000, self._start_browser_selection)
    
    def _start_browser_selection(self):
        """開始瀏覽器區域選擇的實際過程"""
        # 最小化應用程式視窗
        self.root.iconify()
        
        # 確保瀏覽器視窗在前景
        self.browser.maximize_window()
        
        try:
            from pynput.mouse import Listener, Button
            
            # 用於儲存選擇的座標
            selection_coords = {
                'start_x': 0, 'start_y': 0, 
                'end_x': 0, 'end_y': 0, 
                'selecting': False
            }
            
            def on_click(x, y, button, pressed):
                if button == Button.left:
                    if pressed:
                        # 開始選擇
                        selection_coords['start_x'] = x
                        selection_coords['start_y'] = y
                        selection_coords['selecting'] = True
                    else:
                        # 結束選擇
                        if selection_coords['selecting']:
                            selection_coords['end_x'] = x
                            selection_coords['end_y'] = y
                            selection_coords['selecting'] = False
                            return False  # 停止監聽
                elif button == Button.right:
                    # 右鍵取消選擇
                    return False
                return True  # 繼續監聽
            
            # 啟動監聽器
            with Listener(on_click=on_click) as listener:
                listener.join()
            
            # 處理選擇結果
            self.root.after(
                500, 
                lambda: self._process_browser_selection(selection_coords)
            )
        except Exception as e:
            self.log(f"選擇過程中發生錯誤: {str(e)}")
            self.root.deiconify()  # 恢復視窗
    
    def _process_browser_selection(self, coords):
        """處理瀏覽器區域選擇的結果"""
        # 恢復應用程式視窗
        self.root.deiconify()
        
        # 確保選擇有效（開始點與結束點不同）
        if (coords['start_x'] == coords['end_x'] or 
                coords['start_y'] == coords['end_y']):
            self.log("選擇無效，請重新選擇有效的區域")
            return
        
        # 確保座標是遞增的
        x1 = min(coords['start_x'], coords['end_x'])
        y1 = min(coords['start_y'], coords['end_y'])
        x2 = max(coords['start_x'], coords['end_x'])
        y2 = max(coords['start_y'], coords['end_y'])
        
        # 設定 ROI
        self.roi = (x1, y1, x2, y2)
        
        # 更新顯示
        self.log(f"已設定選擇區域: ({x1}, {y1}) - ({x2}, {y2})")
        
        # 重新顯示預覽
        self.show_browser_preview()
    
    def show_browser_preview(self):
        """顯示瀏覽器預覽"""
        try:
            # 獲取螢幕截圖
            screenshot = pyautogui.screenshot()
            
            # 轉換為 OpenCV 格式
            frame = np.array(screenshot)
            frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)
            
            # 顯示在畫布上
            self.display_frame(frame)
            
            # 設置幀大小，用於ROI計算
            self.frame_size = (frame.shape[1], frame.shape[0])
        except Exception as e:
            self.log(f"無法獲取瀏覽器預覽: {str(e)}")
            messagebox.showerror("錯誤", f"無法獲取瀏覽器預覽: {str(e)}")
    
    def on_mouse_down(self, event):
        # 檢查是否處於可選擇狀態，不管是否正在運行都可以重設選擇
        self.roi_selecting = True
        self.roi_start = (event.x, event.y)
        self.canvas.delete("roi")
    
    def on_mouse_move(self, event):
        if self.roi_selecting:
            self.roi_end = (event.x, event.y)
            self.canvas.delete("roi_temp")
            self.canvas.create_rectangle(
                self.roi_start[0], self.roi_start[1],
                self.roi_end[0], self.roi_end[1],
                outline="red", width=2, tags="roi_temp"
            )
    
    def on_mouse_up(self, event):
        if self.roi_selecting:
            self.roi_selecting = False
            self.roi_end = (event.x, event.y)
            
            # 計算 ROI
            x1 = min(self.roi_start[0], self.roi_end[0])
            y1 = min(self.roi_start[1], self.roi_end[1])
            x2 = max(self.roi_start[0], self.roi_end[0])
            y2 = max(self.roi_start[1], self.roi_end[1])
            
            # 保存實際像素範圍
            if hasattr(self, 'frame_size') and hasattr(self, 'display_size'):
                # 計算實際像素與顯示像素的比例
                ratio_x = self.frame_size[0] / self.display_size[0]
                ratio_y = self.frame_size[1] / self.display_size[1]
                
                # 計算實際像素位置
                real_x1 = int(x1 * ratio_x)
                real_y1 = int(y1 * ratio_y)
                real_x2 = int(x2 * ratio_x)
                real_y2 = int(y2 * ratio_y)
                
                # 獲取畫布起始偏移量
                canvas_width = self.canvas.winfo_width()
                canvas_height = self.canvas.winfo_height()
                offset_x = (canvas_width - self.display_size[0]) // 2
                offset_y = (canvas_height - self.display_size[1]) // 2
                
                # 調整選擇區域以補償偏移量
                if x1 >= offset_x and y1 >= offset_y:
                    # 調整座標以考慮偏移
                    adj_x1 = int((x1 - offset_x) * ratio_x)
                    adj_y1 = int((y1 - offset_y) * ratio_y)
                    adj_x2 = int((x2 - offset_x) * ratio_x)
                    adj_y2 = int((y2 - offset_y) * ratio_y)
                    self.roi = (adj_x1, adj_y1, adj_x2, adj_y2)
                else:
                    # 不調整，使用原始計算的座標
                    self.roi = (real_x1, real_y1, real_x2, real_y2)
            else:
                # 如果沒有 frame_size 或 display_size，重新獲取螢幕截圖和尺寸
                screenshot = pyautogui.screenshot()
                frame = np.array(screenshot)
                self.frame_size = (frame.shape[1], frame.shape[0])
                
                if hasattr(self, 'display_size'):
                    ratio_x = self.frame_size[0] / self.display_size[0]
                    ratio_y = self.frame_size[1] / self.display_size[1]
                    self.roi = (
                        int(x1 * ratio_x), int(y1 * ratio_y),
                        int(x2 * ratio_x), int(y2 * ratio_y)
                    )
                else:
                    self.roi = (x1, y1, x2, y2)
            
            self.canvas.delete("roi_temp")
            self.canvas.create_rectangle(
                x1, y1, x2, y2,
                outline="green", width=2, tags="roi"
            )
            
            # 顯示實際像素座標而不是畫布座標
            if self.roi:
                rx1, ry1, rx2, ry2 = self.roi
                self.log(f"已選擇區域: ({rx1},{ry1}) - ({rx2},{ry2}) [實際像素]")
    
    def reset_roi(self):
        self.roi = None
        self.roi_start = None
        self.roi_end = None
        self.canvas.delete("roi")
        self.log("已重設選擇區域")
        
        # 重新顯示當前螢幕以便選擇新區域
        if self.browser:
            self.show_browser_preview()
        else:
            # 如果沒有瀏覽器，直接截取全螢幕
            screenshot = pyautogui.screenshot()
            frame = np.array(screenshot)
            frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)
            self.display_frame(frame)
            self.frame_size = (frame.shape[1], frame.shape[0])
    
    def display_frame(self, frame):
        """顯示幀到畫布上"""
        # 調整大小以適應畫布
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width <= 1 or canvas_height <= 1:
            canvas_width = self.canvas_frame.winfo_width()
            canvas_height = self.canvas_frame.winfo_height()
        
        # 保持寬高比
        scale = min(canvas_width/frame.shape[1], canvas_height/frame.shape[0])
        new_width = int(frame.shape[1] * scale)
        new_height = int(frame.shape[0] * scale)
        
        self.display_size = (new_width, new_height)
        
        # 調整大小並轉換為 PIL 圖像
        frame_resized = cv2.resize(frame, (new_width, new_height))
        frame_rgb = cv2.cvtColor(frame_resized, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(frame_rgb)
        
        # 轉換為 Tkinter 可用的格式
        img_tk = ImageTk.PhotoImage(image=pil_img)
        
        # 保持引用以防垃圾回收
        self.current_image = img_tk
        
        # 計算居中偏移量
        offset_x = (canvas_width - new_width) // 2
        offset_y = (canvas_height - new_height) // 2
        
        # 保存當前偏移量以供座標轉換使用
        self.canvas_offset = (offset_x, offset_y)
        
        # 更新畫布
        self.canvas.delete("frame")
        self.canvas.create_image(
            offset_x, offset_y,
            image=img_tk, anchor=tk.NW, tags="frame"
        )
        
        # 如果有 ROI，重新繪製
        if self.roi and not self.roi_selecting:
            x1, y1, x2, y2 = self.roi
            
            # 將實際像素位置轉換為顯示像素位置
            ratio_x = new_width / self.frame_size[0]
            ratio_y = new_height / self.frame_size[1]
            
            display_x1 = int(x1 * ratio_x)
            display_y1 = int(y1 * ratio_y)
            display_x2 = int(x2 * ratio_x)
            display_y2 = int(y2 * ratio_y)
            
            self.canvas.create_rectangle(
                display_x1 + offset_x, display_y1 + offset_y,
                display_x2 + offset_x, display_y2 + offset_y,
                outline="green", width=2, tags="roi"
            )
    
    def start_capture(self):
        """開始捕獲投影片"""
        if not self.browser:
            messagebox.showerror("錯誤", "請先啟動瀏覽器")
            return
        
        if not self.roi:
            messagebox.showerror("錯誤", "請先選擇要監控的區域")
            return
        
        # 獲取當前設置
        self.threshold = self.threshold_scale.get()
        self.interval = self.interval_scale.get()
        self.auto_stop_enabled = self.auto_stop_var.get()
        self.inactivity_timeout = self.timeout_scale.get() * 60  # 將分鐘轉換為秒
        self.last_change_time = time.time()  # 重設最後變化時間
        
        # 確認 ROI 座標在當前螢幕截圖上的有效性
        try:
            # 獲取最新的螢幕截圖
            screenshot = pyautogui.screenshot()
            frame = np.array(screenshot)
            frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)
            
            # 更新螢幕尺寸
            self.frame_size = (frame.shape[1], frame.shape[0])
            
            # 確保 ROI 區域在螢幕範圍內
            x1, y1, x2, y2 = self.roi
            screen_w, screen_h = self.frame_size
            
            if x1 < 0 or y1 < 0 or x2 > screen_w or y2 > screen_h:
                self.log("警告：選擇的區域超出螢幕範圍，已自動調整")
                x1 = max(0, min(x1, screen_w - 1))
                y1 = max(0, min(y1, screen_h - 1))
                x2 = max(0, min(x2, screen_w - 1))
                y2 = max(0, min(y2, screen_h - 1))
                
                # 確保區域有足夠大小
                if x2 - x1 < 10 or y2 - y1 < 10:
                    messagebox.showerror("錯誤", "選擇的區域太小，請重新選擇")
                    return
                
                self.roi = (x1, y1, x2, y2)
            
            # 顯示最新的一幀帶 ROI 的預覽
            preview_frame = frame.copy()
            self.display_current_roi(preview_frame)
            self.root.after(10, lambda: self.display_frame(preview_frame))
            
            # 記錄實際捕獲區域
            self.log(f"實際捕獲區域: ({x1},{y1}) - ({x2},{y2})")
            
        except Exception as e:
            self.log(f"獲取螢幕截圖時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"獲取螢幕截圖時出錯: {str(e)}")
            return
        
        self.is_running = True
        self.is_paused = False
        
        # 更新按鈕狀態
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)
        self.go_btn.config(state=tk.DISABLED)
        
        # 啟動捕獲線程
        import threading
        self.capture_thread = threading.Thread(target=self.capture_loop)
        self.capture_thread.daemon = True
        self.capture_thread.start()
        
        msg = f"開始捕獲投影片 (閾值: {self.threshold}, 間隔: {self.interval}秒"
        if self.auto_stop_enabled and self.inactivity_timeout > 0:
            msg += f", 無變化自動停止: {self.inactivity_timeout//60}分鐘)"
        else:
            msg += ", 自動停止功能已禁用)"
        self.log(msg)
        self.status_var.set("正在捕獲")
    
    def capture_loop(self):
        """捕獲循環"""
        while self.is_running:
            if not self.is_paused:
                try:
                    # 獲取螢幕截圖
                    screenshot = pyautogui.screenshot()
                    
                    # 轉換為 OpenCV 格式
                    frame = np.array(screenshot)
                    frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)
                    
                    # 顯示帶 ROI 的預覽
                    preview_frame = frame.copy()
                    self.display_current_roi(preview_frame)
                    
                    # 剪切 ROI 區域
                    if self.roi:
                        x1, y1, x2, y2 = self.roi
                        # 確保座標不超出範圍
                        h, w = frame.shape[:2]
                        x1 = max(0, min(x1, w-1))
                        y1 = max(0, min(y1, h-1))
                        x2 = max(0, min(x2, w-1))
                        y2 = max(0, min(y2, h-1))
                        roi_frame = frame[y1:y2, x1:x2]
                        
                        # 檢查是否需要保存
                        if self.last_frame is None:
                            # 第一幀直接保存
                            self.save_slide(roi_frame)
                            self.last_frame = roi_frame
                            self.last_change_time = time.time()  # 初始化最後變化時間
                        else:
                            # 計算相似度
                            if self.similarity_algorithm == 'ssim':
                                # 將兩幀調整為相同大小
                                if roi_frame.shape != self.last_frame.shape:
                                    roi_frame = cv2.resize(
                                        roi_frame, 
                                        (self.last_frame.shape[1], 
                                         self.last_frame.shape[0])
                                    )
                                
                                # 轉換為灰度圖
                                gray_current = cv2.cvtColor(
                                    roi_frame, cv2.COLOR_BGR2GRAY
                                )
                                gray_last = cv2.cvtColor(
                                    self.last_frame, cv2.COLOR_BGR2GRAY
                                )
                                
                                # 計算結構相似度
                                score, _ = ssim(
                                    gray_last, gray_current, full=True
                                )
                                
                                # 檢查無變化時間
                                current_time = time.time()
                                time_since_last_change = current_time - self.last_change_time
                                minutes = int(time_since_last_change // 60)
                                seconds = int(time_since_last_change % 60)
                                
                                status_text = f"相似度: {score:.4f} | 無變化時間: {minutes}分{seconds}秒"
                                self.status_var.set(status_text)
                                
                                # 檢查是否超過無變化自動停止時間
                                if (self.auto_stop_enabled and 
                                    self.inactivity_timeout > 0 and 
                                    time_since_last_change > self.inactivity_timeout):
                                    self.log(f"已檢測到 {minutes}分{seconds}秒 無變化，自動停止捕獲")
                                    self.root.after(0, self.stop_capture)
                                    break
                                
                                # 如果幀有明顯變化，保存
                                if score < self.threshold:
                                    self.save_slide(roi_frame)
                                    self.last_frame = roi_frame
                                    self.last_change_time = current_time  # 更新最後變化時間
                    
                    # 在UI中顯示
                    self.root.after(
                        10, lambda: self.display_frame(preview_frame)
                    )
                    
                except Exception as e:
                    self.log(f"捕獲過程中出錯: {str(e)}")
                
                # 等待指定間隔
                time.sleep(self.interval)
    
    def display_current_roi(self, frame):
        """在預覽幀上顯示當前ROI區域"""
        if self.roi:
            x1, y1, x2, y2 = self.roi
            # 確保座標不超出範圍
            h, w = frame.shape[:2]
            x1 = max(0, min(x1, w-1))
            y1 = max(0, min(y1, h-1))
            x2 = max(0, min(x2, w-1))
            y2 = max(0, min(y2, h-1))
            cv2.rectangle(frame, (x1, y1), (x2, y2), (0, 255, 0), 2)
    
    def save_slide(self, frame):
        """保存投影片截圖"""
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        filename = os.path.join(
            self.output_folder, 
            f"slide_{timestamp}_{self.slide_count:03d}.png"
        )
        cv2.imwrite(filename, frame)
        self.log(f"保存投影片 #{self.slide_count+1}: {filename}")
        self.slide_count += 1
    
    def pause_capture(self):
        """暫停捕獲"""
        if self.is_running:
            self.is_paused = not self.is_paused
            if self.is_paused:
                self.pause_btn.config(text="繼續")
                self.status_var.set("已暫停")
                self.log("捕獲已暫停")
            else:
                self.pause_btn.config(text="暫停")
                self.status_var.set("正在捕獲")
                self.log("捕獲已繼續")
    
    def stop_capture(self):
        """停止捕獲"""
        self.is_running = False
        self.is_paused = False
        
        # 更新按鈕狀態
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.pause_btn.config(text="暫停")
        self.stop_btn.config(state=tk.DISABLED)
        self.go_btn.config(state=tk.NORMAL)
        
        self.status_var.set("已停止")
        self.log(f"捕獲已停止，共獲取 {self.slide_count} 張投影片")
        
        # 如果有捕獲投影片，詢問是否生成PPT
        if self.slide_count > 0:
            if messagebox.askyesno("生成PPT", "是否立即生成PowerPoint文件？"):
                self.generate_ppt()
    
    def generate_ppt(self):
        """生成PowerPoint文件"""
        try:
            from pptx import Presentation
            from pptx.util import Inches
            
            # 創建演示文稿對象 - 使用16:9比例
            prs = Presentation()
            
            # 設置幻燈片尺寸為16:9 (寬度10英寸, 高度5.625英寸)
            prs.slide_width = Inches(10)
            prs.slide_height = Inches(5.625)
            
            # 檢查投影片目錄中的PNG文件
            slides = sorted([
                f for f in os.listdir(self.output_folder) 
                if f.endswith('.png')
            ])
            
            if not slides:
                messagebox.showinfo("提示", "沒有可用的投影片圖像")
                return
            
            # 添加投影片
            for i, slide_file in enumerate(slides):
                # 添加空白投影片
                slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白佈局
                
                # 加載投影片圖像
                img_path = os.path.join(self.output_folder, slide_file)
                
                # 獲取圖像尺寸
                img = Image.open(img_path)
                width, height = img.size
                
                # 計算投影片中的位置和大小
                slide_width = prs.slide_width
                slide_height = prs.slide_height
                
                # 保持寬高比例
                ratio = min(slide_width / width, slide_height / height)
                
                # 添加圖片
                left = (slide_width - width * ratio) / 2
                top = (slide_height - height * ratio) / 2
                slide.shapes.add_picture(
                    img_path, left, top, 
                    width=width * ratio, height=height * ratio
                )
            
            # 保存為PowerPoint
            output_path = self.output_file
            prs.save(output_path)
            
            self.log(f"已生成16:9格式PowerPoint文件: {output_path}")
            messagebox.showinfo("成功", f"已生成16:9格式PowerPoint文件: {output_path}")
            
        except Exception as e:
            self.log(f"生成PowerPoint時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"生成PowerPoint時出錯: {str(e)}")
    
    def toggle_auto_stop(self):
        """啟用或禁用自動停止功能"""
        self.auto_stop_enabled = self.auto_stop_var.get()
        self.timeout_scale.config(state=tk.NORMAL if self.auto_stop_enabled else tk.DISABLED)
        self.log(f"自動停止功能已{'啟用' if self.auto_stop_enabled else '禁用'}")
    
    def run(self):
        """運行主循環"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()
    
    def on_closing(self):
        """窗口關閉事件處理"""
        if self.is_running:
            if messagebox.askokcancel("退出", "捕獲正在進行中，確定要退出嗎？"):
                self.is_running = False
                if self.browser:
                    self.browser.quit()
                self.root.destroy()
        else:
            if self.browser:
                self.browser.quit()
            self.root.destroy()


def main():
    parser = argparse.ArgumentParser(description="Chrome瀏覽器投影片捕獲工具")
    parser.add_argument("--url", help="要打開的網址")
    parser.add_argument(
        "--threshold", type=float, default=0.95, help="相似度閾值 (0.5-1.0)"
    )
    parser.add_argument(
        "--interval", type=float, default=0.5, help="檢測間隔(秒)"
    )
    args = parser.parse_args()
    
    app = ChromeCapture()
    
    # 如果命令行指定了URL，自動打開
    if args.url:
        app.url_entry.insert(0, args.url)
        app.launch_browser()
    
    # 設置命令行參數
    app.threshold_scale.set(args.threshold)
    app.interval_scale.set(args.interval)
    
    app.run()


if __name__ == "__main__":
    main() 