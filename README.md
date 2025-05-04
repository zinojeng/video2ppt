# Video2PPT

將影片、螢幕錄製或瀏覽器內容轉換為 PowerPoint 簡報的工具集。

## 功能特點

- 支援從多種來源獲取內容：
  - 螢幕錄製
  - 瀏覽器內容
  - 本地影片檔案
- 自動擷取關鍵畫面
- 生成高品質的 PowerPoint 簡報
- 從圖片集轉換為 Markdown 並生成簡報 (新功能)

## 安裝

1. 確保已安裝 Python 3.8 或更高版本
2. 複製此倉庫
```
git clone https://github.com/yourusername/video2ppt.git
cd video2ppt
```

3. 安裝所需套件
```
pip install -r requirements.txt
```

## 使用方法

### 快速開始

```
python quick_start.py
```

### 從螢幕錄製生成簡報

```
python screen_recorder.py
```

### 從網頁內容生成簡報

```
python chrome_capture.py
```

### 從本地影片生成簡報

```
python window_capture.py --video path/to/your/video.mp4
```

### 從圖片轉換為 Markdown 並生成簡報 (新功能)

```
python markitdown_ppt.py path/to/images output.pptx
```

選項說明：
- `--markdown, -m`: 指定輸出 Markdown 檔案路徑 (選填)
- `--title, -t`: 設定簡報標題，預設為「圖片簡報」
- `--no-llm, -n`: 不使用 AI 處理圖片
- `--api-key, -k`: 設定 OpenAI API 金鑰
- `--model, -mo`: 指定使用的 OpenAI 模型，預設為「gpt-4o-mini」
- `--template, -tp`: 指定 PPT 模板檔案路徑
- `--verbose, -v`: 顯示詳細資訊

## 程式說明

### 主要檔案

- **quick_start.py**: 整合所有功能的簡單介面
- **screen_recorder.py**: 螢幕錄製模組
- **chrome_capture.py**: Chrome 瀏覽器擷取模組
- **browser_capture.py**: 瀏覽器內容擷取模組
- **window_capture.py**: 窗口擷取模組
- **ppt_generator.py**: PowerPoint 生成模組
- **markitdown_ppt.py**: 圖片轉 Markdown 並生成 PPT 的模組 (新增)

## 新增功能：圖片轉 Markdown 並生成 PPT

`markitdown_ppt.py` 提供了將資料夾中的圖片轉換為 Markdown 文件，然後生成 PowerPoint 簡報的功能。此功能使用 MarkItDown 工具處理圖片，並可選擇使用 AI 技術增強圖片描述。

### 處理流程：
1. 掃描指定資料夾中的所有圖片檔案
2. 透過 MarkItDown 工具將每張圖片轉換為 Markdown 格式的文字描述
3. 可選擇使用 AI (OpenAI GPT 模型) 增強圖片描述
4. 將所有圖片描述合併為一個 Markdown 文件
5. 將 Markdown 文件轉換為 PowerPoint 簡報

### 使用場景：
- 將會議或活動的照片快速整理為簡報
- 建立圖片集的詳細說明文件和簡報
- 為教學或簡報準備圖片資料

## 注意事項

- 使用 AI 增強功能需要有效的 OpenAI API 金鑰
- 轉換大量圖片可能需要較長處理時間
- 支援的圖片格式包括 jpg、jpeg、png、gif 和 bmp

## 授權條款

MIT
