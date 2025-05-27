#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
快速 API 密鑰設置工具
簡化版本，用於快速設置和測試 API 密鑰
"""

import os


def quick_setup():
    """快速設置 API 密鑰"""
    print("🚀 快速 API 密鑰設置")
    print("=" * 40)
    
    # 選擇 API 提供者
    print("請選擇 API 提供者:")
    print("1. OpenAI (推薦用於高質量分析)")
    print("2. Google Gemini (免費額度較多)")
    print("3. 兩者都設置")
    
    choice = input("請輸入選擇 (1-3): ").strip()
    
    if choice in ["1", "3"]:
        print("\n=== OpenAI API 密鑰 ===")
        openai_key = input("請輸入 OpenAI API 密鑰 (sk-...): ").strip()
        if openai_key.startswith("sk-"):
            os.environ["OPENAI_API_KEY"] = openai_key
            print("✅ OpenAI API 密鑰已設置")
        else:
            print("❌ 無效的 OpenAI API 密鑰格式")
            return False
    
    if choice in ["2", "3"]:
        print("\n=== Google Gemini API 密鑰 ===")
        google_key = input("請輸入 Google Gemini API 密鑰: ").strip()
        if google_key:
            os.environ["GOOGLE_API_KEY"] = google_key
            print("✅ Google API 密鑰已設置")
        else:
            print("❌ 請輸入有效的 Google API 密鑰")
            return False
    
    return True


def test_apis():
    """測試 API 連接"""
    print("\n🧪 測試 API 連接...")
    
    openai_key = os.environ.get("OPENAI_API_KEY")
    google_key = os.environ.get("GOOGLE_API_KEY")
    
    if openai_key:
        try:
            from openai import OpenAI
            client = OpenAI(api_key=openai_key)
            # 簡單測試
            client.models.list()
            print("✅ OpenAI API 連接成功")
        except Exception as e:
            print(f"❌ OpenAI API 連接失敗: {e}")
    
    if google_key:
        try:
            import google.generativeai as genai
            genai.configure(api_key=google_key)
            list(genai.list_models())
            print("✅ Google Gemini API 連接成功")
        except Exception as e:
            print(f"❌ Google Gemini API 連接失敗: {e}")


def run_image_test():
    """運行圖片分析測試"""
    print("\n📸 運行圖片分析測試...")
    
    # 檢查是否有測試圖片
    test_dirs = ["slides/3", "slides/1", "slides/2"]
    test_image = None
    
    for test_dir in test_dirs:
        if os.path.exists(test_dir):
            for file in os.listdir(test_dir):
                if file.lower().endswith(('.png', '.jpg', '.jpeg')):
                    test_image = os.path.join(test_dir, file)
                    break
            if test_image:
                break
    
    if not test_image:
        print("❌ 找不到測試圖片")
        return
    
    print(f"使用測試圖片: {test_image}")
    
    # 導入圖片分析模組
    try:
        from image_analyzer import analyze_image
        
        # 選擇 API 提供者
        openai_key = os.environ.get("OPENAI_API_KEY")
        google_key = os.environ.get("GOOGLE_API_KEY")
        
        if google_key:
            print("使用 Google Gemini 進行測試...")
            success, result = analyze_image(
                image_path=test_image,
                api_key=google_key,
                model="gemini-2.5-flash-preview-05-20",
                provider="gemini"
            )
        elif openai_key:
            print("使用 OpenAI 進行測試...")
            success, result = analyze_image(
                image_path=test_image,
                api_key=openai_key,
                model="gpt-4o-mini",
                provider="openai"
            )
        else:
            print("❌ 沒有可用的 API 密鑰")
            return
        
        if success:
            print("✅ 圖片分析測試成功！")
            print("分析結果預覽:")
            print("-" * 40)
            print(result[:200] + "..." if len(result) > 200 else result)
            print("-" * 40)
        else:
            print(f"❌ 圖片分析測試失敗: {result}")
            
    except Exception as e:
        print(f"❌ 測試過程出錯: {e}")


def main():
    """主函數"""
    print("歡迎使用圖片分析工具！")
    
    # 快速設置
    if not quick_setup():
        print("設置失敗，請重新運行")
        return
    
    # 測試連接
    test_apis()
    
    # 運行圖片分析測試
    if input("\n是否要運行圖片分析測試？(Y/n): ").lower().strip() != 'n':
        run_image_test()
    
    print("\n🎉 設置完成！您現在可以使用以下命令:")
    print("python test_image_analysis.py  # 測試圖片分析")
    print("python chrome_quick_start_enhanced.py  # 啟動完整工具")


if __name__ == "__main__":
    main() 