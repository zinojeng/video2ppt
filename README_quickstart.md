# 視頻音頻處理工具 - 快速開始指南

## 🚀 快速啟動方法

### 方法 1：使用啟動腳本（推薦）

1. **賦予執行權限**
   ```bash
   chmod +x start.sh
   chmod +x start.command
   ```

2. **執行啟動腳本**
   - **終端機執行**：
     ```bash
     ./start.sh
     ```
   
   - **雙擊執行**（macOS）：
     直接雙擊 `start.command` 檔案

   - **使用虛擬環境**：
     ```bash
     ./start.sh --venv
     ```

### 方法 2：手動安裝執行

1. **安裝依賴套件**
   ```bash
   pip install -r requirements.txt
   ```

2. **執行程式**
   ```bash
   python3 video_audio_processor.py
   ```

## 📋 系統需求

- Python 3.8 或更高版本
- macOS、Windows 或 Linux 作業系統
- 至少 4GB RAM（處理大型視頻時建議 8GB+）

## 🔧 環境設定

1. **複製環境變數範例檔案**
   ```bash
   cp env.example .env
   ```

2. **編輯 .env 檔案**
   ```bash
   nano .env  # 或使用任何文字編輯器
   ```

3. **填入您的 API Key**（可選）
   - `OPENAI_API_KEY`：用於 MarkItDown 的 AI 功能
   - `GOOGLE_API_KEY`：用於 Gemini 模型支援

## 🎯 主要功能

1. **音頻提取**
   - 從視頻檔案提取音頻
   - 支援 MP3、WAV、AAC 格式

2. **幻燈片捕獲**
   - 智能偵測視頻中的投影片變化
   - 可調整相似度、穩定性參數
   - 支援區域選擇功能

3. **幻燈片處理**
   - 將捕獲的圖片轉換為 PowerPoint
   - 生成 Markdown 文檔
   - 支援 MarkItDown AI 增強

4. **PPT Slide 製作**
   - 從截圖資料夾直接生成 PPT
   - 支援 16:9 和 4:3 比例
   - 可按檔名或時間排序

## 🛠️ 疑難排解

### OpenCV 安裝問題
如果遇到 OpenCV 安裝失敗：
```bash
# 嘗試安裝無頭版本
pip install opencv-python-headless

# 或升級 pip 後重試
pip install --upgrade pip
pip install opencv-python
```

### 權限問題
如果無法執行腳本：
```bash
# 賦予執行權限
chmod +x start.sh
chmod +x start.command
```

### Python 版本問題
確認 Python 版本：
```bash
python3 --version
```

## 📝 使用提示

1. **首次使用**：建議先執行 `./start.sh` 讓腳本自動檢查並安裝依賴

2. **處理大型視頻**：
   - 使用「選擇幻燈片區域」功能可加快處理速度
   - 調整取樣間隔參數以平衡速度和準確度

3. **最佳實踐**：
   - 視頻畫質越高，捕獲效果越好
   - 對於動畫較多的簡報，使用預設設定中的「動畫較多的簡報」

## 🤝 需要幫助？

如有問題，請檢查：
1. Python 版本是否符合要求
2. 所有依賴套件是否已正確安裝
3. 視頻檔案格式是否支援（MP4、AVI、MKV、MOV）

---

祝您使用愉快！ 🎉 