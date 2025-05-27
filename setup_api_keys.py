#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
API 密鑰設置工具
用於設置 OpenAI 和 Google API 密鑰
"""

import os
import getpass


def setup_openai_key():
    """設置 OpenAI API 密鑰"""
    print("=== OpenAI API 密鑰設置 ===")
    
    current_key = os.environ.get("OPENAI_API_KEY")
    if current_key:
        print(f"當前 OpenAI API 密鑰: {current_key[:10]}...")
        update = input("是否要更新密鑰？(y/N): ").lower().strip()
        if update != 'y':
            return current_key
    
    print("請輸入您的 OpenAI API 密鑰 (sk-...):")
    api_key = getpass.getpass("OpenAI API Key: ").strip()
    
    if api_key.startswith("sk-"):
        os.environ["OPENAI_API_KEY"] = api_key
        print("✅ OpenAI API 密鑰已設置")
        return api_key
    else:
        print("❌ 無效的 OpenAI API 密鑰格式（應該以 sk- 開始）")
        return None


def setup_google_key():
    """設置 Google API 密鑰"""
    print("\n=== Google Gemini API 密鑰設置 ===")
    
    current_key = os.environ.get("GOOGLE_API_KEY")
    if current_key:
        print(f"當前 Google API 密鑰: {current_key[:10]}...")
        update = input("是否要更新密鑰？(y/N): ").lower().strip()
        if update != 'y':
            return current_key
    
    print("請輸入您的 Google Gemini API 密鑰:")
    api_key = getpass.getpass("Google API Key: ").strip()
    
    if api_key:
        os.environ["GOOGLE_API_KEY"] = api_key
        print("✅ Google API 密鑰已設置")
        return api_key
    else:
        print("❌ 無效的 Google API 密鑰")
        return None


def test_openai_connection(api_key):
    """測試 OpenAI 連接"""
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        models = client.models.list()
        print("✅ OpenAI API 連接測試成功")
        return True
    except Exception as e:
        print(f"❌ OpenAI API 連接測試失敗: {e}")
        return False


def test_google_connection(api_key):
    """測試 Google Gemini 連接"""
    try:
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        models = list(genai.list_models())
        print("✅ Google Gemini API 連接測試成功")
        return True
    except Exception as e:
        print(f"❌ Google Gemini API 連接測試失敗: {e}")
        return False


def save_to_env_file():
    """將密鑰保存到 .env 文件"""
    openai_key = os.environ.get("OPENAI_API_KEY")
    google_key = os.environ.get("GOOGLE_API_KEY")
    
    if not openai_key and not google_key:
        print("沒有設置任何 API 密鑰")
        return
    
    env_content = []
    if openai_key:
        env_content.append(f"OPENAI_API_KEY={openai_key}")
    if google_key:
        env_content.append(f"GOOGLE_API_KEY={google_key}")
    
    try:
        with open(".env", "w", encoding="utf-8") as f:
            f.write("\n".join(env_content) + "\n")
        print("✅ API 密鑰已保存到 .env 文件")
    except Exception as e:
        print(f"❌ 保存 .env 文件失敗: {e}")


def load_from_env_file():
    """從 .env 文件載入密鑰"""
    if not os.path.exists(".env"):
        return
        
    try:
        with open(".env", "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if "=" in line and not line.startswith("#"):
                    key, value = line.split("=", 1)
                    os.environ[key] = value
        print("✅ 已從 .env 文件載入密鑰")
    except Exception as e:
        print(f"❌ 載入 .env 文件失敗: {e}")


def main():
    """主函數"""
    print("🔐 API 密鑰設置工具")
    print("=" * 50)
    
    # 先嘗試從 .env 文件載入
    load_from_env_file()
    
    # 設置 API 密鑰
    openai_key = setup_openai_key()
    google_key = setup_google_key()
    
    # 測試連接
    print("\n=== 連接測試 ===")
    if openai_key:
        test_openai_connection(openai_key)
    if google_key:
        test_google_connection(google_key)
    
    # 保存到 .env 文件
    print("\n=== 保存設置 ===")
    save = input("是否要將密鑰保存到 .env 文件？(Y/n): ").lower().strip()
    if save != 'n':
        save_to_env_file()
    
    print("\n✅ 設置完成！")
    print("您現在可以運行圖片分析工具了。")


if __name__ == "__main__":
    main() 