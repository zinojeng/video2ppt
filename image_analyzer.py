#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
圖片分析工具

此模組用於分析圖片並提供增強的描述。它可以處理 Markdown 文件中的圖片標記，
使用 OpenAI API 分析圖片內容，然後增強圖片描述。
"""

import os
import re
import base64
import logging
import time
from typing import Dict, List, Optional, Tuple, Any
from urllib.parse import urlparse, unquote
from io import BytesIO
from PIL import Image

# 設定日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("image-analyzer")


def is_url(path: str) -> bool:
    """
    檢查路徑是否為 URL
    
    Args:
        path (str): 要檢查的路徑
    
    Returns:
        bool: 是否為 URL
    """
    parsed = urlparse(path)
    return parsed.scheme in ('http', 'https')


def encode_image_to_base64(image_path: str) -> str:
    """將圖片編碼為 base64 格式"""
    try:
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    except Exception as e:
        logger.error(f"編碼圖片失敗: {e}")
        return ""


def compress_image(image_path: str, max_size_mb: float = 1.0) -> str:
    """嘗試壓縮圖片並返回 base64 編碼"""
    try:
        # 檢查原始檔案大小
        file_size = os.path.getsize(image_path) / (1024 * 1024)  # 轉換為 MB
        if file_size <= max_size_mb:
            return ""  # 不需要壓縮
            
        # 載入圖片
        img = Image.open(image_path)
        
        # 計算壓縮比例
        quality = min(90, int(max_size_mb / file_size * 100))
        quality = max(50, quality)  # 確保品質不低於 50
        
        # 壓縮圖片
        buffer = BytesIO()
        img.save(buffer, format="JPEG", quality=quality)
        buffer.seek(0)
        
        # 轉換為 Base64
        base64_str = base64.b64encode(buffer.read()).decode('utf-8')
        
        logger.info(f"已壓縮圖片，品質: {quality}%")
        return base64_str
    except Exception as e:
        logger.warning(f"壓縮圖片失敗: {e}")
        return ""


def analyze_image(
    image_path: str,
    api_key: str,
    model: str = "o4-mini",
    provider: str = "openai"
) -> Tuple[bool, str]:
    """
    使用 OpenAI、Gemini 或 DeepSeek API 分析圖片內容
    
    Args:
        image_path (str): 圖片檔案路徑
        api_key (str): API Key (OpenAI、Google 或 DeepSeek)
        model (str): 使用的模型名稱
        provider (str): API 提供者，可為 'openai'、'gemini' 或 'deepseek'
    
    Returns:
        Tuple[bool, str]: (是否成功, 分析結果)
    """
    try:
        # 檢查圖片是否存在
        if not os.path.exists(image_path):
            logger.error(f"圖片不存在: {image_path}")
            return False, f"圖片不存在: {image_path}"
        
        # 檢查圖片大小
        file_size = os.path.getsize(image_path) / (1024 * 1024)  # MB
        if file_size > 20:
            logger.warning(f"圖片太大 ({file_size:.2f}MB)，超過 API 限制")
            return False, f"圖片太大 ({file_size:.2f}MB)，超過 API 限制"
        
        # 清理 API 金鑰（移除可能的換行符和空白）
        api_key = api_key.strip() if api_key else ""
        
        # 最多重試次數和延遲
        max_retries = 2
        retry_delay = 3  # 重試間隔秒數
        
        # 判斷使用哪個提供者的 API
        if provider.lower() == "gemini":
            # 使用 Google Gemini API
            try:
                import google.generativeai as genai
                
                # 初始化 Gemini API
                genai.configure(api_key=api_key)
                
                # 建立系統提示
                gemini_prompt = (
                    "請詳細分析並描述這張圖片中的內容。\n\n"
                    "請使用繁體中文回應，並包含：\n"
                    "1. 總體概述：簡明扼要說明圖片主要內容\n"
                    "2. 詳細分析：包括主要元素、場景、人物、文字等\n"
                    "3. 若圖片包含投影片，請特別描述投影片中的標題、結構和重點\n"
                    "4. 若圖片包含圖表或數據，請具體說明其呈現的趨勢或資訊\n"
                    "5. 若圖片包含文字，請在描述中引用關鍵文字\n\n"
                    "請以流暢的段落形式提供分析，而非列表。分析需專業、客觀、準確。"
                )
                
                # 選擇適合的模型
                # 直接使用傳入的模型名稱，因為 Gemini 的多模態模型名稱本身已包含視覺能力
                model_name = model
                
                logger.info(f"使用 Gemini 模型: {model_name}")
                
                # 加載圖片
                img = Image.open(image_path)
                
                # 設定 Gemini 模型
                generation_config = {
                    "temperature": 0.4,
                    "top_p": 0.95,
                    "top_k": 0,
                    "max_output_tokens": 2048,
                }
                
                # 初始化模型
                gemini_model = genai.GenerativeModel(
                    model_name=model_name,
                    generation_config=generation_config
                )
                
                # 發送請求
                for attempt in range(max_retries + 1):
                    try:
                        response = gemini_model.generate_content([gemini_prompt, img])
                        analysis = response.text
                        return True, analysis
                    except Exception as e:
                        error_msg = str(e)
                        logger.error(f"Gemini API 嘗試 {attempt+1}/{max_retries+1} 失敗: {error_msg}")
                        
                        if attempt < max_retries:
                            wait_time = retry_delay * (attempt + 1)
                            logger.warning(f"等待 {wait_time} 秒後重試...")
                            time.sleep(wait_time)
                        else:
                            raise
                
            except ImportError:
                logger.error("找不到 google.generativeai 模組，請確保已安裝")
                return False, "缺少 Google Generative AI 模組，請執行 'pip install google-genai'"
                
        elif provider.lower() == "deepseek":
            # DeepSeek API 目前不支援圖片分析功能
            logger.info("DeepSeek 模型不支援圖片分析，使用預設模板")
            
            # 返回一個標準模板，而不嘗試發送圖片
            template = """
這是一個投影片內容分析模板。DeepSeek API 目前不支援視覺分析功能。

標準投影片分析通常包含以下元素：
1. 投影片標題和主題概述 
2. 主要內容和關鍵訊息摘要
3. 圖表、數據或可視化元素的解釋
4. 專業術語或縮寫的解釋
5. 上下文相關的背景資訊

若需進行實際的視覺內容分析，建議：
- 切換至支援視覺分析的 OpenAI 或 Gemini API
- 手動記錄投影片中的關鍵內容
- 使用支援圖像識別的其他服務

此回應為預設模板，因 DeepSeek API 目前僅支援純文本處理。
"""
            return True, template.strip()
        else:
            # 使用 OpenAI API
            try:
                from openai import OpenAI
                
                # 轉換圖片為 Base64
                base64_image = encode_image_to_base64(image_path)
                if not base64_image:
                    return False, "無法轉換圖片為 Base64 格式"
                
                # 如果圖片大於 1MB，嘗試壓縮
                if file_size > 1:
                    compressed = compress_image(image_path)
                    if compressed:
                        base64_image = compressed
                
                # 建立系統提示
                system_prompt = (
                    "你是一個專業的圖片分析和描述專家。你的工作是詳細分析圖片內容，"
                    "提供準確且有深度的描述。請使用繁體中文回應。"
                    "\n\n描述應包含：\n"
                    "1. 總體概述：簡明扼要說明圖片主要內容\n"
                    "2. 詳細分析：包括主要元素、場景、人物、文字等\n"
                    "3. 圖片目的或用途（如果明顯）\n"
                    "4. 若圖片包含文字，請在描述中引用關鍵文字\n\n"
                    "請以流暢的段落形式提供分析，而非列表。分析需專業、客觀、準確。"
                )
                
                client = OpenAI(api_key=api_key)
                
                # 建立請求訊息列表
                messages = [
                    {"role": "system", "content": system_prompt},
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": "請分析並描述這張圖片。"},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": (
                                        f"data:image/jpeg;base64,"
                                        f"{base64_image}"
                                    )
                                }
                            }
                        ]
                    }
                ]
                
                # 動態構建請求參數
                request_kwargs = {
                    "model": model,
                    "messages": messages,
                }
                
                # 根據模型類型添加可選參數
                if model.startswith("gpt-4o"):
                    request_kwargs["temperature"] = 0.5
                    request_kwargs["max_tokens"] = 1000
                
                # 發送請求
                logger.info(f"使用 OpenAI 模型: {model}")
                
                for attempt in range(max_retries + 1):
                    try:
                        response = client.chat.completions.create(**request_kwargs)
                        analysis = response.choices[0].message.content
                        return True, analysis
                    except Exception as e:
                        error_msg = str(e)
                        logger.error(f"OpenAI API 嘗試 {attempt+1}/{max_retries+1} 失敗: {error_msg}")
                        
                        if "429" in error_msg and attempt < max_retries:
                            # 速率限制錯誤，等待後重試
                            wait_time = retry_delay * (attempt + 1)
                            logger.warning(f"API 速率限制，等待 {wait_time} 秒後重試...")
                            time.sleep(wait_time)
                        elif attempt < max_retries:
                            # 其他錯誤，也嘗試重試
                            logger.warning(f"請求失敗，等待 {retry_delay} 秒後重試...")
                            time.sleep(retry_delay)
                        else:
                            # 已達最大重試次數，拋出異常
                            raise
                
            except ImportError:
                logger.error("找不到 openai 模組，請確保已安裝")
                return False, "缺少 OpenAI 模組"
        
        # 不應該到達這裡，但為了安全添加
        return False, "分析失敗，請稍後再試"
        
    except Exception as e:
        logger.error(f"API 調用異常: {str(e)}")
        return False, f"API 調用異常: {str(e)}"


def find_markdown_images(markdown_text: str) -> List[Dict[str, Any]]:
    """
    從 Markdown 文件中找出所有圖片標記
    
    Args:
        markdown_text (str): Markdown 文字內容
    
    Returns:
        List[Dict[str, Any]]: 圖片標記列表，每項包含標記文字、位置和圖片路徑等
    """
    # 圖片標記模式：![alt text](image_url "title")
    image_pattern = r'!\[(.*?)\]\((.*?)(?:\s+"(.*?)")?\)'
    
    images = []
    for match in re.finditer(image_pattern, markdown_text):
        alt_text = match.group(1)
        image_path = match.group(2)
        title = match.group(3) if match.group(3) else ""
        
        # 處理路徑中的特殊字符
        if not is_url(image_path):
            image_path = unquote(image_path)
        
        images.append({
            "match": match.group(0),
            "start": match.start(),
            "end": match.end(),
            "alt_text": alt_text,
            "path": image_path,
            "title": title
        })
    
    return images


def enhance_markdown_with_image_analysis(
    markdown_text: str,
    base_dir: str = ".",
    api_key: Optional[str] = None,
    model: str = "o4-mini"
) -> Tuple[str, Dict[str, int]]:
    """
    增強 Markdown 文件中的圖片描述
    
    Args:
        markdown_text (str): Markdown 文字內容
        base_dir (str): 圖片基礎目錄，用於解析相對路徑
        api_key (Optional[str]): OpenAI API Key
        model (str): 使用的 OpenAI 模型
    
    Returns:
        Tuple[str, Dict[str, int]]: (增強後的 Markdown 文字, 處理統計)
    """
    # 獲取 API 金鑰
    if not api_key:
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            logger.error("找不到 OpenAI API 金鑰")
            return markdown_text, {
                "images_processed": 0,
                "images_analyzed": 0,
                "images_failed": 0
            }
    
    # 找出所有圖片標記
    images = find_markdown_images(markdown_text)
    
    if not images:
        logger.info("Markdown 文件中沒有找到圖片標記")
        return markdown_text, {
            "images_processed": 0,
            "images_analyzed": 0,
            "images_failed": 0
        }
    
    logger.info(f"找到 {len(images)} 個圖片標記")
    
    # 準備處理統計
    stats = {
        "images_processed": len(images),
        "images_analyzed": 0,
        "images_failed": 0
    }
    
    # 從後向前處理圖片，避免因替換而改變後續圖片的位置
    for img_info in reversed(images):
        # 獲取圖片路徑
        img_path = img_info["path"]
        
        # 判斷是絕對路徑還是相對路徑
        if not is_url(img_path) and not os.path.isabs(img_path):
            # 相對路徑，轉換為絕對路徑
            img_path = os.path.join(base_dir, img_path)
        
        # 分析圖片
        if not is_url(img_path) and os.path.exists(img_path):
            logger.info(f"分析圖片: {img_path}")
            success, analysis = analyze_image(
                image_path=img_path,
                api_key=api_key,
                model=model
            )
            
            if success:
                # 增強 Markdown
                enhanced = (
                    f"{img_info['match']}\n\n"
                    f"{analysis}\n"
                )
                
                # 替換原始 Markdown
                markdown_text = (
                    markdown_text[:img_info['start']] +
                    enhanced +
                    markdown_text[img_info['end']:]
                )
                
                stats["images_analyzed"] += 1
                logger.info(f"成功分析圖片: {img_path}")
            else:
                stats["images_failed"] += 1
                logger.warning(f"分析圖片失敗: {img_path} - {analysis}")
        else:
            stats["images_failed"] += 1
            logger.warning(f"無法訪問圖片: {img_path}")
    
    return markdown_text, stats


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(
        description="圖片分析工具 - 分析圖片並提供增強描述"
    )
    
    parser.add_argument(
        "input_file",
        help="輸入 Markdown 檔案路徑"
    )
    parser.add_argument(
        "--output", "-o",
        help="輸出 Markdown 檔案路徑，如果未指定則覆蓋輸入檔案"
    )
    parser.add_argument(
        "--api-key", "-k",
        help="OpenAI API 金鑰，如果未指定則從環境變數獲取"
    )
    parser.add_argument(
        "--model", "-m",
        default="o4-mini",
        help="使用的 OpenAI 模型，預設為 o4-mini"
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="顯示詳細資訊"
    )
    
    args = parser.parse_args()
    
    # 設定日誌級別
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # 檢查輸入檔案
    if not os.path.exists(args.input_file):
        logger.error(f"輸入檔案不存在: {args.input_file}")
        exit(1)
    
    # 讀取輸入檔案
    with open(args.input_file, 'r', encoding='utf-8') as f:
        markdown_text = f.read()
    
    # 設定輸出檔案
    output_file = args.output or args.input_file
    
    # 增強 Markdown
    enhanced_markdown, stats = enhance_markdown_with_image_analysis(
        markdown_text=markdown_text,
        base_dir=os.path.dirname(os.path.abspath(args.input_file)),
        api_key=args.api_key,
        model=args.model
    )
    
    # 寫入輸出檔案
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(enhanced_markdown)
    
    # 顯示結果
    logger.info(f"處理完成，已寫入檔案: {output_file}")
    logger.info(
        f"處理統計: 共處理 {stats['images_processed']} 張圖片，"
        f"成功解析 {stats['images_analyzed']} 張，"
        f"失敗 {stats['images_failed']} 張"
    ) 