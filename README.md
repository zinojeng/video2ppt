# Video2PPT - 影片投影片轉換器

自動從視頻中檢測投影片變化並進行截圖，然後生成PowerPoint文件。

## 功能特點

- 從視頻文件或直播視頻中自動檢測投影片變化
- 通過Chrome瀏覽器捕獲網頁中的視頻投影片 (**新功能**)
- 根據相似度閾值判斷投影片是否變化
- 支持手動選擇截圖區域
- 自動將截圖保存為PowerPoint文件
- 支持暫停/繼續截圖功能

## 系統需求

- Python 3.7+
- 依賴庫：opencv-python, numpy, pillow, python-pptx, scikit-image, pyautogui, selenium, webdriver-manager

## 安裝方法

1. 克隆本專案
```
git clone https://github.com/zinojeng/video2ppt.git
cd video2ppt
```

2. 安裝依賴
```
pip install -r requirements.txt
```

## 使用方法

### 基本版本

1. 運行主程序
```
python video2ppt.py
```

2. 程序啟動後：
   - 選擇視頻文件或輸入直播URL
   - 使用滑鼠選擇截圖區域
   - 設置相似度閾值
   - 點擊"開始"按鈕

3. 程序會自動檢測投影片變化並進行截圖，完成後會生成PowerPoint文件

### Chrome瀏覽器捕獲 (新功能)

1. 運行Chrome捕獲工具
```
python chrome_quick_start.py
```

2. 程序啟動後：
   - 輸入包含視頻的網址並點擊「打開瀏覽器」
   - 在打開的Chrome瀏覽器中播放視頻
   - 使用滑鼠框選要監控的投影片區域
   - 設置相似度閾值和檢測間隔
   - 點擊「開始捕獲」按鈕

3. 工具會自動監控選定區域，當檢測到投影片變化時自動截圖
4. 完成後點擊「停止」並選擇「生成PPT」即可生成PowerPoint文件

## 高級選項

### 命令行參數 (video2ppt.py)

- `--threshold`: 設置相似度閾值（默認：0.95）
- `--interval`: 設置檢測間隔（秒）（默認：0.5）
- `--output`: 設置輸出文件名（默認：slides.pptx）

### 命令行參數 (chrome_capture.py)

- `--url`: 指定要打開的網址
- `--threshold`: 設置相似度閾值（默認：0.95）
- `--interval`: 設置檢測間隔（秒）（默認：0.5）

## 配置文件 (.cursor.json)

可以通過創建 `.cursor.json` 文件來自定義Chrome捕獲工具的設置，例如：

```json
{
  "customSettings": {
    "chromeCapture": {
      "defaultThreshold": 0.95,
      "defaultInterval": 0.5,
      "defaultOutputFormat": "pptx",
      "autoSaveSlides": true,
      "browserExecutablePath": "",
      "userDataDir": ""
    },
    "captureSettings": {
      "preferredScreenShotMethod": "pyautogui",
      "fallbackMethod": "opencv",
      "similarityAlgorithm": "ssim"
    }
  }
}
```

## 參考專案

- [Auto-Screenshot](https://github.com/krmanik/Auto-Screenshot)
- [Video-to-PowerPoint](https://github.com/dplem/Video-to-PowerPoint)
- [presentation_extractor](https://github.com/E-CAM/presentation_extractor)

## 許可證

MIT 