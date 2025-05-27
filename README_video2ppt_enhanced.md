# Video2PPT 增強版

這是 Video2PPT 工具的增強版，除了原來捕獲視頻中投影片並生成 PowerPoint 檔案的功能外，還增加了使用 MarkItDown 或 OpenAI 視覺模型來辨識圖片中的文字內容，並將結果保存為 Markdown 格式。

## 功能特點

1. **視頻投影片捕獲**：捕獲視頻中的投影片畫面，自動識別畫面變化
2. **PowerPoint 生成**：將捕獲的投影片轉換為 PowerPoint 簡報
3. **文字識別與分析**：對捕獲的投影片進行文字內容識別和分析
4. **Markdown 輸出**：將識別結果以 Markdown 格式保存
5. **兩種處理方式**：支援使用 MarkItDown 或 OpenAI 視覺模型處理圖片

## 需求

- Python 3.7 或更高版本
- Chrome 瀏覽器
- 相關 Python 套件（程式會自動檢查並提示安裝）
- 若使用 OpenAI 視覺模型，需要有效的 OpenAI API Key

## 安裝

1. 確保您已安裝 Python 和 Chrome 瀏覽器
2. 下載本工具的所有檔案
3. 執行以下命令安裝依賴：

```bash
pip install -r requirements.txt
```

或者直接執行增強版工具，它會檢查並提示您安裝缺少的套件：

```bash
python chrome_quick_start_enhanced.py
```

## 使用方法

### 基本使用流程

1. 執行增強版工具：

```bash
python chrome_quick_start_enhanced.py
```

2. 在主介面上選擇「捕獲投影片」或「分析已捕獲投影片」標籤頁

### 捕獲投影片

1. 在「捕獲投影片」標籤頁點擊「啟動投影片捕獲工具」
2. 輸入包含視頻的網址（如 YouTube 影片）並點擊「打開瀏覽器」
3. 在瀏覽器中播放視頻
4. 用滑鼠框選要監控的投影片區域
5. 點擊「開始捕獲」開始監控並截取投影片
6. 觀看視頻時，工具會自動檢測並捕獲投影片畫面
7. 完成後點擊「停止」
8. 捕獲的投影片將存放在 `slides` 資料夾中

### 分析已捕獲投影片

1. 在「分析已捕獲投影片」標籤頁中：
   - 選擇投影片資料夾（預設為 `slides`）
   - 輸入 OpenAI API Key（如需文字識別功能）
   - 選擇使用的模型（如 `gpt-4o-mini` 或 `o4-mini`）
   - 選擇輸出格式（Markdown、PowerPoint 或兩者都要）
   - 選擇處理方式（MarkItDown 或 OpenAI 視覺模型）

2. 點擊「處理投影片」按鈕

3. 處理完成後，可以在以下位置找到輸出檔案：
   - Markdown 檔案：`slides_analysis.md`（使用 MarkItDown 時）或 `slides_openai_analysis.md`（使用 OpenAI 視覺模型時）
   - PowerPoint 檔案：`slides.pptx`

## 處理方式說明

### MarkItDown 處理（推薦）

- 使用 MarkItDown 工具處理圖片，轉換為結構化的 Markdown 文件
- 可同時生成 PowerPoint 檔案
- 適合大多數文件和投影片內容
- 若提供 OpenAI API Key，會使用 OpenAI 模型增強圖片描述

### OpenAI 視覺模型處理

- 直接使用 OpenAI 的視覺模型分析圖片內容
- 提供更詳細的視覺內容描述
- 需要有效的 OpenAI API Key
- 適合需要深度理解圖片內容的場景

## 注意事項

1. 使用 OpenAI 視覺模型功能需要有效的 API Key
2. 高解析度圖片處理可能需要更長時間
3. 使用 API 視覺模型會消耗 API 額度
4. 使用 MarkItDown 處理方式需要安裝 MarkItDown 套件
5. 文字識別效果取決於圖片質量和文字清晰度

## 檔案說明

- `chrome_quick_start_enhanced.py`：增強版主程式
- `chrome_capture.py`：Chrome 瀏覽器捕獲功能
- `markitdown_ppt.py`：圖片轉 Markdown 並生成 PPT 的模組
- `image_analyzer.py`：圖片分析工具，使用 OpenAI API 增強圖片描述
- `slides/`：存放捕獲的投影片圖片

## 故障排除

如果遇到問題，請檢查：

1. 是否已安裝所有必要的依賴套件
2. Chrome 瀏覽器是否正確安裝
3. OpenAI API Key 是否有效
4. 捕獲區域是否正確設置

## 授權

本工具僅供學習和研究使用，請尊重視頻內容的版權。 