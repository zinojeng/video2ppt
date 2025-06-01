#!/bin/bash

# 視頻音頻處理工具快速啟動腳本
# Video Audio Processor Quick Start Script

echo "========================================="
echo "視頻音頻處理工具 - 快速啟動"
echo "Video Audio Processor - Quick Start"
echo "========================================="

# 檢查 Python 是否已安裝
if ! command -v python3 &> /dev/null; then
    echo "❌ 錯誤：未找到 Python 3"
    echo "請先安裝 Python 3.8 或更高版本"
    echo "下載地址：https://www.python.org/downloads/"
    exit 1
fi

# 顯示 Python 版本
echo "✅ Python 版本："
python3 --version
echo ""

# 檢查並創建虛擬環境（可選）
if [ "$1" == "--venv" ]; then
    echo "📦 創建虛擬環境..."
    if [ ! -d "venv" ]; then
        python3 -m venv venv
        echo "✅ 虛擬環境創建成功"
    fi
    
    # 啟動虛擬環境
    echo "🔄 啟動虛擬環境..."
    source venv/bin/activate
fi

# 檢查必要的依賴套件
echo "🔍 檢查依賴套件..."
REQUIRED_PACKAGES=(
    "opencv-python"
    "numpy"
    "pillow"
    "moviepy"
    "python-pptx"
    "scikit-image"
)

MISSING_PACKAGES=()

for package in "${REQUIRED_PACKAGES[@]}"; do
    if ! python3 -m pip show "$package" &> /dev/null; then
        MISSING_PACKAGES+=("$package")
    fi
done

# 如果有缺少的套件，詢問是否安裝
if [ ${#MISSING_PACKAGES[@]} -gt 0 ]; then
    echo ""
    echo "⚠️  缺少以下套件："
    printf '%s\n' "${MISSING_PACKAGES[@]}"
    echo ""
    read -p "是否要自動安裝這些套件？(y/n) " -n 1 -r
    echo ""
    
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        echo "📥 正在安裝依賴套件..."
        python3 -m pip install --upgrade pip
        python3 -m pip install "${MISSING_PACKAGES[@]}"
        
        # 檢查安裝結果
        if [ $? -eq 0 ]; then
            echo "✅ 所有套件安裝成功！"
        else
            echo "❌ 套件安裝失敗，請手動安裝："
            echo "python3 -m pip install ${MISSING_PACKAGES[@]}"
            exit 1
        fi
    else
        echo "⚠️  請手動安裝缺少的套件後再執行"
        exit 1
    fi
else
    echo "✅ 所有必要套件已安裝"
fi

# 檢查可選套件
echo ""
echo "🔍 檢查可選套件..."
OPTIONAL_PACKAGES=(
    "markitdown"
    "selenium"
    "pyautogui"
    "webdriver-manager"
)

MISSING_OPTIONAL=()

for package in "${OPTIONAL_PACKAGES[@]}"; do
    if ! python3 -m pip show "$package" &> /dev/null; then
        MISSING_OPTIONAL+=("$package")
    fi
done

if [ ${#MISSING_OPTIONAL[@]} -gt 0 ]; then
    echo "💡 以下可選套件未安裝："
    printf '%s\n' "${MISSING_OPTIONAL[@]}"
    echo "如需使用進階功能，可執行："
    echo "python3 -m pip install ${MISSING_OPTIONAL[@]}"
fi

# 設置環境變數（如果有 API Key）
if [ -f ".env" ]; then
    echo ""
    echo "📄 載入環境變數..."
    export $(cat .env | xargs)
fi

# 啟動應用程式
echo ""
echo "🚀 啟動視頻音頻處理工具..."
echo "========================================="
echo ""

# 執行主程式
python3 video_audio_processor.py

# 如果使用虛擬環境，退出時關閉
if [ "$1" == "--venv" ]; then
    deactivate
fi 