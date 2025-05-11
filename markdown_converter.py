#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Markdown 文件轉換工具
使用 Pandoc 將 Markdown 轉換為其他格式，例如 Word (.docx)。
支持自定義樣式、表格樣式和圖片/表格標題。
參考了 nihole/md2docx 項目 (https://github.com/nihole/md2docx)。
"""

import subprocess
import os
import logging
import sys
from typing import Dict

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

# 設定日誌
logger = logging.getLogger("markdown_converter")
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )


def check_pandoc_installed():
    """檢查系統是否已安裝 Pandoc"""
    try:
        subprocess.run(
            ["pandoc", "--version"], check=True, capture_output=True
        )
        logger.info("Pandoc 已安裝。")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        logger.error("錯誤：Pandoc 未安裝或未在系統 PATH 中。")
        logger.error(
            "請先安裝 Pandoc (https://pandoc.org/installing.html) 後再試。"
        )
        return False


def convert_markdown_to_docx(
    markdown_file_path: str, 
    output_docx_path: str
) -> bool:
    """
    將 Markdown 檔案轉換為 Word (.docx) 檔案。

    Args:
        markdown_file_path (str): 輸入的 Markdown 檔案路徑。
        output_docx_path (str): 輸出的 .docx 檔案路徑。

    Returns:
        bool: 如果轉換成功則返回 True，否則返回 False。
    """
    if not os.path.exists(markdown_file_path):
        logger.error(f"Markdown 檔案不存在: {markdown_file_path}")
        return False

    if not check_pandoc_installed():
        return False

    try:
        logger.info(f"開始將 {markdown_file_path} 轉換為 {output_docx_path}...")
        
        # 基本的 pandoc 命令
        command = [
            "pandoc",
            markdown_file_path,
            "-o",
            output_docx_path
        ]
        
        # 執行 pandoc 命令
        process = subprocess.run(
            command, check=True, capture_output=True, text=True
        )
        
        if process.stderr:
            logger.warning(f"Pandoc 轉換過程中有警告或訊息: {process.stderr}")
        
        logger.info(f"成功轉換 Markdown 至 DOCX: {output_docx_path}")
        return True
    except FileNotFoundError:
        logger.error(
            "Pandoc 命令未找到。請確保 Pandoc 已安裝並在系統 PATH 中。"
        )
        return False
    except subprocess.CalledProcessError as e:
        logger.error(f"Pandoc 轉換失敗。返回碼: {e.returncode}")
        logger.error(f"Pandoc 錯誤訊息: {e.stderr}")
        return False
    except Exception as e:
        logger.error(f"轉換過程中發生未知錯誤: {str(e)}")
        return False


def apply_word_styles(docx_path: str) -> bool:
    """
    使用 win32com 應用 Word 樣式。
    僅適用於 Windows 系統。

    Args:
        docx_path (str): Word 文檔路徑

    Returns:
        bool: 成功返回 True，否則返回 False
    """
    if not HAS_WIN32COM or os.name != 'nt':
        logger.warning("Win32COM 未安裝或不在 Windows 系統上，無法應用 Word 樣式")
        return False
    
    try:
        logger.info(f"正在應用 Word 樣式到文檔: {docx_path}")
        
        # 初始化 Word 應用程序
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        try:
            # 打開文檔
            doc = word.Documents.Open(os.path.abspath(docx_path))
            
            # 應用表格樣式
            apply_table_styles(doc)
            
            # 處理表格標題
            process_table_captions(doc)
            
            # 處理圖片標題
            process_figure_captions(doc)
            
            # 更新欄位
            doc.Fields.Update()
            
            # 保存並關閉
            doc.Save()
            doc.Close()
            
            logger.info("Word 樣式應用成功")
            return True
        finally:
            # 確保 Word 應用被關閉
            word.Quit()
    except Exception as e:
        logger.error(f"應用 Word 樣式時發生錯誤: {str(e)}")
        return False


def apply_table_styles(doc):
    """
    向 Word 文檔中的所有表格應用樣式

    Args:
        doc: Word 文檔對象
    """
    try:
        # 獲取表格樣式名稱 (可以根據需要修改)
        table_style = "Table Grid"
        
        # 應用樣式到所有表格
        for i in range(1, doc.Tables.Count + 1):
            table = doc.Tables(i)
            table.Style = table_style
            
            # 其他表格格式設置，例如自動調整列寬等
            table.AutoFitBehavior(1)  # 自動適應內容
    except Exception as e:
        logger.error(f"應用表格樣式時發生錯誤: {str(e)}")


def process_table_captions(doc):
    """
    處理表格標題，將 Markdown 風格的標題轉換為 Word 標題欄位

    Args:
        doc: Word 文檔對象
    """
    try:
        # 搜尋類似 "表 X.X: 標題" 的文本
        caption_pattern = "表"
        
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if text.startswith(caption_pattern) and ":" in text:
                # 獲取原始位置
                original_range = para.Range
                
                # 刪除原始文本
                original_caption = original_range.Text
                original_range.Delete()
                
                # 解析標題
                caption_parts = original_caption.split(":", 1)
                title = ""
                if len(caption_parts) > 1:
                    title = caption_parts[1].strip()
                
                # 插入表格標題欄位
                original_range.InsertCaption(
                    Caption="表格", 
                    Title=title
                )
    except Exception as e:
        logger.error(f"處理表格標題時發生錯誤: {str(e)}")


def process_figure_captions(doc):
    """
    處理圖片標題，將 Markdown 風格的標題轉換為 Word 標題欄位

    Args:
        doc: Word 文檔對象
    """
    try:
        # 搜尋類似 "圖 X.X: 標題" 的文本
        caption_pattern = "圖"
        
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if text.startswith(caption_pattern) and ":" in text:
                # 獲取原始位置
                original_range = para.Range
                
                # 刪除原始文本
                original_caption = original_range.Text
                original_range.Delete()
                
                # 解析標題
                caption_parts = original_caption.split(":", 1)
                title = ""
                if len(caption_parts) > 1:
                    title = caption_parts[1].strip()
                
                # 插入圖片標題欄位
                original_range.InsertCaption(
                    Caption="圖", 
                    Title=title
                )
    except Exception as e:
        logger.error(f"處理圖片標題時發生錯誤: {str(e)}")


def batch_convert_markdown_files(
    markdown_folder: str, 
    output_folder: str = None
) -> Dict[str, bool]:
    """
    批量將資料夾中的 Markdown 檔案轉換為 Word 檔案

    Args:
        markdown_folder (str): 包含 Markdown 檔案的資料夾
        output_folder (str, optional): 輸出 Word 檔案的資料夾，預設為與輸入相同

    Returns:
        Dict[str, bool]: 檔案名稱和成功狀態的字典
    """
    if not os.path.isdir(markdown_folder):
        logger.error(f"輸入資料夾不存在: {markdown_folder}")
        return {}

    if output_folder is None:
        output_folder = markdown_folder
    elif not os.path.exists(output_folder):
        os.makedirs(output_folder)

    results = {}
    for filename in os.listdir(markdown_folder):
        if filename.lower().endswith(('.md', '.markdown')):
            md_path = os.path.join(markdown_folder, filename)
            docx_name = os.path.splitext(filename)[0] + '.docx'
            docx_path = os.path.join(output_folder, docx_name)
            
            success = convert_markdown_to_docx(
                md_path, 
                docx_path
            )
            results[filename] = success
    
    return results


if __name__ == '__main__':
    # 簡單的命令列測試
    if len(sys.argv) < 3:
        print(
            "使用方法: python markdown_converter.py "
            "<輸入Markdown檔案> <輸出Word檔案.docx>"
        )
        sys.exit(1)
    
    input_md = sys.argv[1]
    output_docx = sys.argv[2]
    
    if convert_markdown_to_docx(
        input_md, 
        output_docx
    ):
        print(f"檔案已成功轉換並保存至: {output_docx}")
    else:
        print("檔案轉換失敗。") 