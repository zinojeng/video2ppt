#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Markdown 文件轉換工具
使用 Pandoc 將 Markdown 轉換為其他格式，例如 Word (.docx)。
"""

import subprocess
import os
import logging
import sys  # 添加 sys 導入以供 __main__ 使用


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
    markdown_file_path: str, output_docx_path: str
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
        logger.info(
            f"開始將 {markdown_file_path} 轉換為 {output_docx_path}..."
        )
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
        logger.info(f"成功轉換 Markdown 至 DOCX: {output_docx_path}")
        if process.stderr:
            logger.warning(
                f"Pandoc 轉換過程中有警告或訊息: {process.stderr}"
            )
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
    
    if convert_markdown_to_docx(input_md, output_docx):
        print(f"檔案已成功轉換並保存至: {output_docx}")
    else:
        print("檔案轉換失敗。") 