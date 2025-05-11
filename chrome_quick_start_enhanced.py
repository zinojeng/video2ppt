#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Chrome捕獲工具增強版快速啟動腳本

此版本在原始功能基礎上增加了圖片文字識別功能，
可以使用 MarkItDown 或 OpenAI 視覺模型來辨識截圖中的文字，
並將結果保存為 Markdown 格式。
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox, filedialog
import tkinter.ttk as ttk
import traceback
from PIL import Image, ImageTk


def check_dependencies():
    """檢查必要依賴是否已安裝"""
    required_packages = [
        "selenium", "opencv-python", "numpy", "pillow", 
        "python-pptx", "scikit-image", "pyautogui", "webdriver-manager",
        "openai"
    ]
    
    # 分開檢查 markitdown，因為它可能需要特殊處理
    optional_packages = ["markitdown", "google-genai"]
    
    missing_packages = []
    missing_optional = []
    
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
            # 處理 scikit-image 特例
            elif package == "scikit-image":
                import_name = "skimage"
                
            __import__(import_name)
        except ImportError:
            missing_packages.append(package)
    
    for package in optional_packages:
        try:
            if package == "google-genai":
                import_name = "google.generativeai"
            else:
                import_name = package
            __import__(import_name)
        except ImportError:
            missing_optional.append(package)
    
    if missing_optional:
        print(f"注意: 未安裝可選套件: {', '.join(missing_optional)}")
        print("如果您想使用 MarkItDown 處理功能，請安裝:")
        print("pip install markitdown>=0.1.1")
        print("如果您想使用 Google Gemini API，請安裝:")
        print("pip install google-genai")
    
    return missing_packages


def install_dependencies(packages):
    """安裝缺失的依賴"""
    import subprocess
    print(f"正在安裝必要依賴: {', '.join(packages)}")
    
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install"] + packages
        )
        return True
    except subprocess.CalledProcessError:
        return False


def process_captured_slides(
    slides_folder, 
    output_format="markdown", 
    api_key=None, 
    model="gpt-4o-mini"
):
    """處理已捕獲的投影片，生成 Markdown 或增強的文字分析"""
    try:
        from markitdown_ppt import (
            process_images_to_ppt, 
            convert_images_to_markdown
        )
        
        if not os.path.exists(slides_folder) or not os.listdir(slides_folder):
            messagebox.showinfo("提示", "找不到投影片或資料夾為空")
            return False
            
        # 獲取所有圖片檔案路徑
        image_files = []
        for filename in sorted(os.listdir(slides_folder)):
            if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                image_files.append(os.path.join(slides_folder, filename))
                
        if not image_files:
            messagebox.showinfo("提示", "未找到圖片檔案")
            return False
        
        if output_format == "markdown" or output_format == "both":
            # 轉換為 Markdown
            output_md = os.path.join(
                os.path.dirname(slides_folder), 
                "slides_analysis.md"
            )
            success, md_file, info = convert_images_to_markdown(
                image_paths=image_files,
                output_file=output_md,
                title="投影片內容分析",
                use_llm=(api_key is not None),
                api_key=api_key,
                model=model
            )
            
            if success:
                messagebox.showinfo("成功", f"已生成 Markdown 檔案: {md_file}")
            else:
                msg = (
                    f"生成 Markdown 時出現問題: "
                    f"{info.get('error', '未知錯誤')}"
                )
                messagebox.showwarning("警告", msg)
        
        if output_format == "pptx" or output_format == "both":
            # 生成 PPT
            output_ppt = os.path.join(
                os.path.dirname(slides_folder), 
                "slides.pptx"
            )
            if process_images_to_ppt(
                image_dir=slides_folder,
                output_ppt=output_ppt,
                title="投影片簡報",
                use_llm=(api_key is not None),
                api_key=api_key,
                model=model
            ):
                messagebox.showinfo("成功", f"已生成 PPT 檔案: {output_ppt}")
            else:
                messagebox.showwarning("警告", "生成 PPT 時出現問題")
        
        return True
        
    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"處理投影片時出錯: {str(e)}\n{error_msg}")
        messagebox.showerror("錯誤", f"處理投影片時出錯:\n{str(e)}")
        return False


def process_with_image_analyzer(
    slides_folder, 
    output_file=None, 
    api_key=None, 
    model="o4-mini",
    provider="openai"
):
    """使用 OpenAI、Gemini 或 DeepSeek 模型分析圖片內容"""
    try:
        from image_analyzer import analyze_image
        
        if not os.path.exists(slides_folder) or not os.listdir(slides_folder):
            messagebox.showinfo("提示", "找不到投影片或資料夾為空")
            return {"success": False, "error": "找不到投影片或資料夾為空"}
            
        # 獲取所有圖片檔案路徑
        image_files = []
        for filename in sorted(os.listdir(slides_folder)):
            if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                image_files.append(os.path.join(slides_folder, filename))
                
        if not image_files:
            messagebox.showinfo("提示", "未找到圖片檔案")
            return {"success": False, "error": "未找到圖片檔案"}
        
        # 如果沒有指定輸出檔案，創建預設檔案名稱
        if not output_file:
            provider_name = provider.lower()
            output_file = os.path.join(
                os.path.dirname(slides_folder), 
                f"slides_{provider_name}_analysis.md"
            )
            
        # 確保 API Key 存在
        api_key_env_vars = {
            "openai": "OPENAI_API_KEY",
            "gemini": "GOOGLE_API_KEY"
        }
        env_var = api_key_env_vars.get(provider.lower(), "OPENAI_API_KEY")
        current_api_key = api_key or os.environ.get(env_var)
            
        if not current_api_key:
            messagebox.showwarning(
                "警告", 
                f"未提供 {provider.upper()} API Key，無法進行圖片分析"
            )
            return {
                "success": False, 
                "error": f"未提供 {provider.upper()} API Key"
            }
            
        # 創建 Markdown 檔案
        with open(output_file, 'w', encoding='utf-8') as f:
            provider_display = {
                "openai": "OpenAI",
                "gemini": "Google Gemini"
            }.get(provider.lower(), provider.upper())
            
            f.write(f"# 投影片內容 {provider_display} 視覺分析\\n\\n")
            
            # 處理每張圖片
            success_count = 0
            for img_path in image_files:
                img_name = os.path.basename(img_path)
                f.write(f"## 投影片：{img_name}\\n\\n")
                f.write(f"![{img_name}]({img_path})\\n\\n")
                
                # 分析圖片
                success, analysis = analyze_image(
                    image_path=img_path,
                    api_key=current_api_key,
                    model=model,
                    provider=provider
                )
                
                if success:
                    f.write(f"{analysis}\\n\\n")
                    success_count += 1
                else:
                    f.write(f"*分析失敗: {analysis}*\\n\\n")
        
        if success_count > 0:
            msg = (
                f"已分析 {success_count}/{len(image_files)} 張圖片"
                f"並生成 Markdown 檔案: {output_file}"
            )
            messagebox.showinfo("成功", msg)
            return {
                "success": True, 
                "total_slides": len(image_files), 
                "analyzed_slides": success_count,
                "output_file": output_file
            }
        else:
            messagebox.showwarning("警告", "所有圖片分析均失敗")
            return {"success": False, "error": "所有圖片分析均失敗"}
            
    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"分析圖片時出錯: {str(e)}\\n{error_msg}")
        messagebox.showerror("錯誤", f"分析圖片時出錯:\\n{str(e)}")
        return {"success": False, "error": str(e), "details": error_msg}


class EnhancedChromeCapture:
    """增強型 Chrome 捕獲應用"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Chrome 投影片捕獲與分析工具")
        self.root.geometry("1000x750")
        
        # 圖片選擇相關變數
        self.image_files = []
        self.selected_images = []
        self.current_image_index = 0
        self.image_label = None
        self.image_preview = None
        self.image_cache = []  # 添加圖片緩存列表，防止圖片被垃圾回收
        
        # 創建頁面框架
        self.setup_ui()
    
    def setup_ui(self):
        """設置使用者介面"""
        # 建立選項卡
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 第一個選項卡：捕獲模式
        self.capture_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.capture_frame, text="捕獲投影片")
        
        # 第二個選項卡：圖片選擇模式
        self.select_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.select_frame, text="選擇投影片")
        
        # 第三個選項卡：分析模式
        self.analyze_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.analyze_frame, text="分析已選擇投影片")
        
        # 設置捕獲模式 UI
        self.setup_capture_ui()
        
        # 設置圖片選擇模式 UI
        self.setup_select_ui()
        
        # 設置分析模式 UI
        self.setup_analyze_ui()
        
        # 添加標籤頁切換事件處理
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
    
    def on_tab_changed(self, event):
        """處理標籤頁切換事件"""
        # 不再需要自動填入 API Key，此方法保留但不執行任何操作
        pass
    
    def setup_capture_ui(self):
        """設置捕獲模式 UI"""
        # 在這裡我們將嵌入原始的 ChromeCapture UI
        tk.Label(self.capture_frame, text="在此選項卡中可以捕獲投影片").pack(pady=10)
        
        # 啟動原始 ChromeCapture
        self.launch_chrome_capture_btn = tk.Button(
            self.capture_frame, text="啟動投影片捕獲工具", 
            command=self.launch_chrome_capture,
            bg="#2196F3", fg="white", height=2, width=20
        )
        self.launch_chrome_capture_btn.pack(pady=20)
    
    def setup_select_ui(self):
        """設置圖片選擇模式 UI"""
        frame = self.select_frame
        
        # 選擇投影片資料夾
        folder_frame = tk.Frame(frame)
        folder_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(folder_frame, text="投影片資料夾:").pack(
            side=tk.LEFT, padx=10
        )
        self.select_folder_entry = tk.Entry(folder_frame, width=40)
        self.select_folder_entry.pack(
            side=tk.LEFT, padx=5, fill=tk.X, expand=True
        )
        self.select_folder_entry.insert(0, "slides")  # 預設資料夾
        
        self.select_browse_btn = tk.Button(
            folder_frame, text="瀏覽...", 
            command=lambda: self.browse_folder(self.select_folder_entry)
        )
        self.select_browse_btn.pack(side=tk.LEFT, padx=5)
        
        self.load_images_btn = tk.Button(
            folder_frame, text="載入圖片", 
            command=self.load_images_for_selection,
            bg="#FF9800", fg="white"
        )
        self.load_images_btn.pack(side=tk.LEFT, padx=5)
        
        # 圖片預覽區域
        preview_frame = tk.Frame(frame)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 創建圖片顯示框架，專門用於放置圖片標籤
        self.image_frame = tk.Frame(preview_frame, bg="black")
        self.image_frame.pack(fill=tk.BOTH, expand=True)
        
        # 創建初始標籤顯示
        self.image_label = tk.Label(self.image_frame, text="尚未載入圖片")
        self.image_label.pack(fill=tk.BOTH, expand=True)
        
        # 圖片選擇控制區域
        control_frame = tk.Frame(frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        self.prev_btn = tk.Button(
            control_frame, text="上一張", 
            command=self.prev_image,
            state=tk.DISABLED
        )
        self.prev_btn.pack(side=tk.LEFT, padx=10)
        
        self.select_btn = tk.Button(
            control_frame, text="選擇此圖", 
            command=self.toggle_select_image,
            state=tk.DISABLED,
            bg="#4CAF50", fg="white"
        )
        self.select_btn.pack(side=tk.LEFT, padx=10)
        
        self.next_btn = tk.Button(
            control_frame, text="下一張", 
            command=self.next_image,
            state=tk.DISABLED
        )
        self.next_btn.pack(side=tk.LEFT, padx=10)
        
        self.status_var = tk.StringVar(value="請先載入圖片")
        status_label = tk.Label(control_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT, padx=20)
        
        # 完成按鈕
        finish_frame = tk.Frame(frame)
        finish_frame.pack(fill=tk.X, pady=20)
        
        self.save_selected_btn = tk.Button(
            finish_frame, text="保存所選圖片", 
            command=self.save_selected_images,
            state=tk.DISABLED,
            bg="#2196F3", fg="white", height=2
        )
        self.save_selected_btn.pack(pady=10)
    
    def setup_analyze_ui(self):
        """設置分析模式 UI"""
        frame = self.analyze_frame
        
        # 選擇投影片資料夾
        folder_frame = tk.Frame(frame)
        folder_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(folder_frame, text="投影片資料夾:").pack(side=tk.LEFT, padx=10)
        self.folder_entry = tk.Entry(folder_frame, width=40)
        self.folder_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.folder_entry.insert(0, "selected_slides")  # 預設資料夾
        
        self.browse_btn = tk.Button(
            folder_frame, text="瀏覽...", 
            command=lambda: self.browse_folder(self.folder_entry)
        )
        self.browse_btn.pack(side=tk.LEFT, padx=5)
        
        # 添加從「選擇投影片」匯入按鈕
        self.import_btn = tk.Button(
            folder_frame, text="從「選擇投影片」匯入", 
            command=self.import_from_selection,
            bg="#FFC107", fg="black"
        )
        self.import_btn.pack(side=tk.LEFT, padx=5)
        
        # API Key 設置
        api_frame = tk.Frame(frame)
        api_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(api_frame, text="API Key:").pack(side=tk.LEFT, padx=10)
        self.api_key_entry = tk.Entry(api_frame, width=40)
        self.api_key_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # API 提供者選擇
        provider_frame = tk.Frame(frame)
        provider_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(provider_frame, text="API 提供者:").pack(side=tk.LEFT, padx=10)
        
        self.provider_var = tk.StringVar(value="gemini")
        
        tk.Radiobutton(
            provider_frame, text="OpenAI", 
            variable=self.provider_var, value="openai",
            command=self.update_model_options
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(
            provider_frame, text="Gemini", 
            variable=self.provider_var, value="gemini",
            command=self.update_model_options
        ).pack(side=tk.LEFT, padx=5)
        
        # 模型選擇
        model_frame = tk.Frame(frame)
        model_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(model_frame, text="使用模型:").pack(side=tk.LEFT, padx=10)
        
        self.model_var = tk.StringVar(value="gpt-4o-mini")
        self.openai_models = ["gpt-4o-mini", "gpt-4o", "o4-mini", "o4"]
        self.gemini_models = [
            "gemini-2.5-pro-exp-03-25",
            "gemini-2.5-flash-preview-04-17"
        ]
        
        # 定義 Gemini 模型列表，並輸出以供調試
        print("初始化 Gemini 模型列表:", self.gemini_models)
        
        self.model_menu = ttk.Combobox(
            model_frame, 
            textvariable=self.model_var, 
            values=self.openai_models, 
            width=20
        )
        self.model_menu.pack(side=tk.LEFT, padx=5)
        
        # 輸出格式選擇
        format_frame = tk.Frame(frame)
        format_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(format_frame, text="輸出格式:").pack(side=tk.LEFT, padx=10)
        
        self.format_var = tk.StringVar(value="both")
        
        tk.Radiobutton(
            format_frame, text="僅 Markdown", 
            variable=self.format_var, value="markdown"
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(
            format_frame, text="僅 PowerPoint", 
            variable=self.format_var, value="pptx"
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(
            format_frame, text="僅 Word (.docx)", 
            variable=self.format_var, value="docx"
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(
            format_frame, text="全部 (MD, PPTX, DOCX)", 
            variable=self.format_var, value="all"
        ).pack(side=tk.LEFT, padx=5)
        
        # 處理方式選擇
        method_frame = tk.Frame(frame)
        method_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(method_frame, text="處理方式:").pack(side=tk.LEFT, padx=10)
        
        self.method_var = tk.StringVar(value="markitdown")
        
        tk.Radiobutton(
            method_frame, text="使用 MarkItDown (推薦)", 
            variable=self.method_var, value="markitdown"
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(
            method_frame, text="使用視覺模型", 
            variable=self.method_var, value="openai"
        ).pack(side=tk.LEFT, padx=5)
        
        # Word 模板選擇
        template_frame = tk.Frame(frame)
        template_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(template_frame, text="Word 模板:").pack(side=tk.LEFT, padx=10)
        self.template_entry = tk.Entry(template_frame, width=40)
        self.template_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.template_browse_btn = tk.Button(
            template_frame, text="瀏覽...", 
            command=lambda: self.browse_docx_template(self.template_entry)
        )
        self.template_browse_btn.pack(side=tk.LEFT, padx=5)
        
        # 按鈕區域
        btn_frame = tk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=20)
        
        self.process_btn = tk.Button(
            btn_frame, text="處理投影片", command=self.process_slides,
            bg="#4CAF50", fg="white", height=2, width=15
        )
        self.process_btn.pack(pady=10)
        
        # 初始化模型選項
        print("初始化模型選項...")
        self.update_model_options()
    
    def update_model_options(self):
        """根據選擇的 API 提供者更新模型選項"""
        provider = self.provider_var.get()
        
        print(f"更新模型選項，選擇的提供者: {provider}")
        
        if provider == "gemini":
            print("Gemini 模型列表:", self.gemini_models)
            self.model_menu.config(values=self.gemini_models)
            self.model_var.set(self.gemini_models[0])
        else:  # OpenAI
            print("OpenAI 模型列表:", self.openai_models)
            self.model_menu.config(values=self.openai_models)
            self.model_var.set(self.openai_models[0])
            
        # 確認 Combobox 的值是否已更新
        print(f"下拉選單當前值: {self.model_menu['values']}")
        print(f"當前選中的模型: {self.model_var.get()}")
        
        # 強制更新下拉選單
        self.model_menu.update()
    
    def browse_folder(self, entry_widget):
        """瀏覽並選擇資料夾"""
        folder = filedialog.askdirectory(
            initialdir=os.path.dirname(entry_widget.get()),
            title="選擇投影片資料夾"
        )
        if folder:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, folder)
    
    def load_images_for_selection(self):
        """載入要選擇的圖片"""
        folder = self.select_folder_entry.get()
        if not os.path.exists(folder):
            messagebox.showwarning("警告", f"資料夾不存在: {folder}")
            return
        
        # 獲取所有圖片檔案路徑
        self.image_files = []
        for filename in sorted(os.listdir(folder)):
            if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                self.image_files.append(os.path.join(folder, filename))
        
        if not self.image_files:
            messagebox.showinfo("提示", "未找到圖片檔案")
            return
        
        # 重置選擇狀態
        self.selected_images = [False] * len(self.image_files)
        self.current_image_index = 0
        self.image_cache = []  # 清空圖片緩存
        
        # 顯示第一張圖片
        self.show_current_image()
        
        # 啟用按鈕
        self.select_btn.config(state=tk.NORMAL)
        self.save_selected_btn.config(state=tk.NORMAL)
        
        if len(self.image_files) > 1:
            self.next_btn.config(state=tk.NORMAL)
        
        self.status_var.set(f"已載入 {len(self.image_files)} 張圖片")
    
    def show_current_image(self):
        """顯示當前索引的圖片"""
        if not self.image_files:
            return
        
        # 讀取圖片
        img_path = self.image_files[self.current_image_index]
        img = Image.open(img_path)
        
        # 調整圖片大小以適應視窗
        window_width = self.select_frame.winfo_width() - 40
        window_height = self.select_frame.winfo_height() - 200
        
        if window_width <= 100 or window_height <= 100:
            # 如果框架還沒有正確的尺寸，使用預設尺寸
            window_width = 800
            window_height = 400
        
        img_width, img_height = img.size
        scale = min(window_width/img_width, window_height/img_height)
        
        new_width = int(img_width * scale)
        new_height = int(img_height * scale)
        
        # 確保至少有一個像素
        new_width = max(1, new_width)
        new_height = max(1, new_height)
        
        # 調整圖片大小
        resized_img = img.resize((new_width, new_height), Image.LANCZOS)
        
        # 清除現有顯示
        if self.image_label:
            self.image_label.destroy()
        
        # 創建新的Photo對象並保存到實例變數中
        photo = ImageTk.PhotoImage(resized_img)
        
        # 保存到緩存中避免被垃圾回收
        self.image_cache.append(photo)
        
        # 限制緩存大小
        if len(self.image_cache) > 30:
            self.image_cache = self.image_cache[-15:]
        
        # 創建新的Label顯示圖片
        self.image_label = tk.Label(self.image_frame, image=photo)
        self.image_label.image = photo  # 保存引用防止被垃圾回收
        self.image_label.pack(pady=10)
        
        # 更新選中狀態
        current_idx = self.current_image_index
        is_selected = self.selected_images[current_idx]
        self.select_btn.config(
            text="取消選擇" if is_selected else "選擇此圖片"
        )
        
        # 更新狀態
        img_name = os.path.basename(img_path)
        selected_count = sum(self.selected_images)
        self.status_var.set(
            f"圖片 {self.current_image_index + 1}/{len(self.image_files)}: "
            f"{img_name} (已選擇: {selected_count}張)"
        )
        
        # 更新按鈕狀態
        self.prev_btn.config(
            state=tk.NORMAL if self.current_image_index > 0 else tk.DISABLED
        )
        self.next_btn.config(
            state=tk.NORMAL 
            if self.current_image_index < len(self.image_files) - 1 
            else tk.DISABLED
        )
    
    def prev_image(self):
        """顯示上一張圖片"""
        if self.current_image_index > 0:
            self.current_image_index -= 1
            self.show_current_image()
    
    def next_image(self):
        """顯示下一張圖片"""
        if self.current_image_index < len(self.image_files) - 1:
            self.current_image_index += 1
            self.show_current_image()
    
    def toggle_select_image(self):
        """選擇或取消選擇當前圖片"""
        if not self.image_files:
            return
            
        # 切換選擇狀態
        current_idx = self.current_image_index
        is_selected = self.selected_images[current_idx]
        self.selected_images[current_idx] = not is_selected
        self.show_current_image()
    
    def save_selected_images(self):
        """保存所選圖片到單獨的資料夾"""
        if not self.image_files or not any(self.selected_images):
            messagebox.showinfo("提示", "未選擇任何圖片")
            return
        
        # 創建目標資料夾
        selected_folder = "selected_slides"
        if not os.path.exists(selected_folder):
            os.makedirs(selected_folder)
        
        # 複製所選圖片
        copied_count = 0
        for i, (img_path, selected) in enumerate(
            zip(self.image_files, self.selected_images)
        ):
            if selected:
                # 保持原始檔名以維持順序
                filename = os.path.basename(img_path)
                # 添加編號確保順序正確
                base, ext = os.path.splitext(filename)
                new_filename = f"{i:03d}_{base}{ext}"
                
                dest_path = os.path.join(selected_folder, new_filename)
                
                try:
                    import shutil
                    shutil.copy2(img_path, dest_path)
                    copied_count += 1
                except Exception as e:
                    print(f"複製圖片 {img_path} 時出錯: {str(e)}")
        
        if copied_count > 0:
            # 更新分析標籤頁的資料夾路徑
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, selected_folder)
            
            # 切換到分析標籤頁
            self.notebook.select(2)
            
            messagebox.showinfo(
                "完成", 
                f"已將 {copied_count} 張選定的圖片複製到 '{selected_folder}' 資料夾"
            )
        else:
            messagebox.showinfo("提示", "未複製任何圖片")
    
    def launch_chrome_capture(self):
        """啟動原始的 Chrome 捕獲工具"""
        try:
            # 隱藏當前視窗，顯示原始的 ChromeCapture
            self.root.withdraw()
            
            # 導入 ChromeCapture
            from chrome_capture import ChromeCapture
            
            # 創建新的 Tkinter 根窗口
            capture_root = tk.Tk()
            app = ChromeCapture(capture_root)
            
            # 顯示使用說明
            messagebox.showinfo(
                "使用說明", 
                "1. 輸入包含視頻的網址並點擊「打開瀏覽器」\n"
                "2. 在打開的頁面中播放視頻\n"
                "3. 框選要監控的投影片區域\n"
                "4. 點擊「開始捕獲」開始監控並截取投影片\n"
                "5. 完成後點擊「停止」\n\n"
                "提示：調整相似度閾值可以控制檢測靈敏度，值越低檢測越靈敏"
            )
            
            def on_capture_close():
                """當原始捕獲工具關閉時"""
                try:
                    # 檢查是否已捕獲投影片（檢查輸出目錄是否有圖片）
                    slides_folder = "slides"
                    capture_has_slides = False
                    if os.path.exists(slides_folder):
                        # 檢查是否有圖片檔案
                        for file in os.listdir(slides_folder):
                            if file.lower().endswith(
                                ('.png', '.jpg', '.jpeg')
                            ):
                                capture_has_slides = True
                                break
                    
                    # 關閉捕獲窗口
                    capture_root.destroy()
                    
                    # 恢復主窗口
                    self.root.deiconify()
                    self.root.update()
                    
                    # 如果有新捕獲的投影片，填充到資料夾輸入框
                    if (os.path.exists(slides_folder) and 
                            os.listdir(slides_folder)):
                        self.select_folder_entry.delete(0, tk.END)
                        self.select_folder_entry.insert(0, slides_folder)
                        # 切換到圖片選擇標籤頁
                        self.notebook.select(1)
                        
                        # 如果有截圖且用戶同意，自動載入圖片
                        if capture_has_slides:
                            if messagebox.askyesno(
                                "載入圖片", 
                                "是否立即載入捕獲的投影片進行選擇？"
                            ):
                                self.load_images_for_selection()
                            else:
                                # 提示使用者載入圖片進行選擇
                                messagebox.showinfo(
                                    "下一步", 
                                    "您可以點擊「載入圖片」按鈕，選擇要保留的投影片圖片。"
                                )
                except Exception as e:
                    print(f"關閉捕獲工具時出錯: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    self.root.deiconify()  # 確保主窗口恢復
                    
            # 設置關閉事件處理
            capture_root.protocol("WM_DELETE_WINDOW", on_capture_close)
            app.run()
            
        except Exception as e:
            error_msg = traceback.format_exc()
            print(f"啟動 Chrome 捕獲工具失敗: {str(e)}\n{error_msg}")
            messagebox.showerror("啟動錯誤", f"工具啟動失敗:\n{str(e)}")
            # 恢復主窗口
            self.root.deiconify()
    
    def import_from_selection(self):
        """從選擇投影片頁面匯入資料夾路徑"""
        if hasattr(self, 'select_folder_entry'):
            selected_folder = self.select_folder_entry.get()
            if selected_folder and os.path.exists(selected_folder):
                self.folder_entry.delete(0, tk.END)
                self.folder_entry.insert(0, selected_folder)
                messagebox.showinfo("成功", f"已匯入資料夾路徑: {selected_folder}")
            else:
                messagebox.showwarning("警告", "選擇投影片頁面未指定有效的資料夾")
        else:
            messagebox.showwarning("警告", "無法獲取選擇投影片頁面資訊")
    
    def browse_docx_template(self, entry_widget):
        """瀏覽並選擇 Word 模板檔案"""
        file_path = filedialog.askopenfilename(
            initialdir=(os.path.dirname(entry_widget.get()) 
                      if entry_widget.get() else os.getcwd()),
            title="選擇 Word 模板檔案",
            filetypes=(("Word 檔案", "*.docx"), ("所有檔案", "*.*"))
        )
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)
    
    def process_slides(self):
        """處理選擇的投影片"""
        # 獲取選項
        folder_path = self.folder_entry.get().strip()
        api_key = self.api_key_entry.get().strip()
        model = self.model_var.get()
        provider = self.provider_var.get()  # 獲取提供者選項
        method = self.method_var.get()  # 獲取處理方式：markitdown 或 openai
        output_format = self.format_var.get()  # 獲取輸出格式
        template_path = self.template_entry.get().strip()  # 獲取 Word 模板路徑
        
        # 檢查輸入
        if not folder_path:
            self.show_message("錯誤", "請輸入投影片資料夾路徑", "error")
            return
        
        # 在當前目錄下建立資料夾
        input_folder = os.path.join(os.getcwd(), folder_path)
        if not os.path.exists(input_folder):
            self.show_message("錯誤", f"找不到資料夾: {input_folder}", "error")
            return

        markdown_file_path = None
        
        if method == "markitdown":
            # 使用 MarkItDown 處理投影片
            # 轉換 UI 選擇的格式為 process_captured_slides 函數所需格式
            pct_format = "markdown"  # 預設生成 Markdown
            if output_format == "pptx":
                pct_format = "pptx"
            elif output_format in ["both", "all"]:
                pct_format = "both"
                
            # 執行 MarkItDown 處理
            from markitdown_ppt import process_captured_slides
            success = process_captured_slides(
                slides_folder=input_folder,
                output_format=pct_format,
                api_key=api_key,
                model=model
            )
            
            if success:
                # 找出產生的 Markdown 文件
                if output_format in ["markdown", "both", "all", "docx"]:
                    expected_md = os.path.join(
                        os.path.dirname(input_folder), 
                        "slides_analysis.md"
                    )
                    if os.path.exists(expected_md):
                        self.show_message(
                            "成功", 
                            f"已生成 Markdown 檔案: {expected_md}",
                            "info"
                        )
                        markdown_file_path = expected_md
                    else:
                        self.show_message(
                            "警告", 
                            "處理成功但找不到 Markdown 檔案",
                            "warning"
                        )
                
                # 檢查 PPT 檔案是否生成
                if output_format in ["pptx", "both", "all"]:
                    expected_ppt = os.path.join(
                        os.path.dirname(input_folder), 
                        "slides.pptx"
                    )
                    if os.path.exists(expected_ppt):
                        self.show_message(
                            "成功", 
                            f"已生成 PowerPoint 檔案: {expected_ppt}",
                            "info"
                        )
                    else:
                        self.show_message(
                            "警告", 
                            "處理成功但找不到 PowerPoint 檔案",
                            "warning"
                        )
            else:
                self.show_message(
                    "錯誤", 
                    "MarkItDown 處理失敗，請檢查日誌了解詳情",
                    "error"
                )
                return
        else:  # method == "openai"
            # 使用視覺模型
            result = process_with_image_analyzer(
                slides_folder=input_folder, 
                api_key=api_key,
                model=model,
                provider=provider,
                output_file=None  # 使用預設輸出檔案
            )
            
            if result.get("success"):
                markdown_file_path = result.get("output_file")
                
                self.show_message(
                    "成功", 
                    f"處理完成!\n\n已處理 {result.get('total_slides', 0)} 張投影片\n"
                    f"Markdown 檔案: {markdown_file_path}",
                    "info"
                )
            else:
                self.show_message(
                    "錯誤", 
                    f"處理失敗: {result.get('error', '未知錯誤')}", 
                    "error"
                )
                return
        
        # 處理 Word (.docx) 轉換
        if markdown_file_path and output_format in ["docx", "all"]:
            try:
                import shutil
                pandoc_path = shutil.which("pandoc")
                
                if not pandoc_path:
                    self.show_message(
                        "錯誤",
                        "未安裝 Pandoc，無法轉換為 Word 格式。\n"
                        "請從 https://pandoc.org/installing.html 安裝 Pandoc。",
                        "error"
                    )
                    return
                    
                try:
                    from markdown_converter import convert_markdown_to_docx
                    docx_path = os.path.splitext(markdown_file_path)[0] + \
                        ".docx"
                    
                    # 檢查是否指定了模板
                    reference_docx = None
                    if template_path and os.path.exists(template_path):
                        reference_docx = template_path
                    
                    if convert_markdown_to_docx(
                        markdown_file_path, 
                        docx_path,
                        reference_docx=reference_docx,
                        apply_styles=True
                    ):
                        self.show_message(
                            "成功",
                            f"已將 Markdown 轉換為 Word 檔案: {docx_path}",
                            "info"
                        )
                    else:
                        self.show_message(
                            "錯誤",
                            "Word 轉換失敗，請檢查錯誤日誌",
                            "error"
                        )
                except ImportError:
                    self.show_message(
                        "錯誤",
                        "找不到 markdown_converter 模組，無法轉換為 Word",
                        "error"
                    )
            except Exception as e:
                self.show_message(
                    "錯誤",
                    f"Word 轉換時出錯: {str(e)}",
                    "error"
                )
    
    def show_message(self, title, message, type="info"):
        """顯示消息框"""
        if type == "error":
            messagebox.showerror(title, message)
        elif type == "warning":
            messagebox.showwarning(title, message)
        else:
            messagebox.showinfo(title, message)
    
    def run(self):
        """運行應用"""
        self.root.mainloop()


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
        # 創建並運行應用
        root = tk.Tk()
        app = EnhancedChromeCapture(root)
        
        app.run()
        
    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"啟動失敗: {str(e)}\n{error_msg}")
        messagebox.showerror("啟動錯誤", f"程序啟動失敗:\n{str(e)}")


if __name__ == "__main__":
    main() 