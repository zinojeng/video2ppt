#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
快速入門示例：自動檢測視頻中投影片變化並截圖
"""

import os
import cv2
import numpy as np
from skimage.metrics import structural_similarity as ssim
from datetime import datetime
import argparse

def capture_slides(video_path, output_folder="slides", threshold=0.95, interval=1):
    """
    從視頻文件中自動捕獲投影片並保存為圖片
    
    參數:
        video_path (str): 視頻文件路徑
        output_folder (str): 輸出文件夾路徑
        threshold (float): 相似度閾值 (0 到 1)
        interval (int): 檢查間隔（每隔多少幀檢查一次）
    
    返回:
        int: 捕獲的投影片數量
    """
    # 確保輸出目錄存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 打開視頻文件
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print(f"無法打開視頻文件: {video_path}")
        return 0
    
    # 獲取視頻屬性
    fps = cap.get(cv2.CAP_PROP_FPS)
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    
    print(f"視頻 FPS: {fps}")
    print(f"總幀數: {total_frames}")
    print(f"相似度閾值: {threshold}")
    print(f"檢查間隔: 每 {interval} 幀")
    
    # 初始化變量
    frame_count = 0
    last_frame = None
    slide_count = 0
    
    print("開始處理視頻...")
    
    while True:
        ret, frame = cap.read()
        if not ret:
            break
            
        frame_count += 1
        
        # 只處理每interval幀
        if frame_count % interval != 0:
            continue
            
        # 轉換為灰度圖像以進行比較
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        
        # 第一幀直接保存
        if last_frame is None:
            last_frame = gray
            slide_path = os.path.join(output_folder, f"slide_{slide_count:03d}.png")
            cv2.imwrite(slide_path, frame)
            slide_count += 1
            print(f"保存幀 {frame_count} 為投影片 {slide_count}")
            continue
            
        # 計算相似度
        similarity = ssim(gray, last_frame)
        
        # 打印進度
        if frame_count % (interval * 10) == 0:
            progress = (frame_count / total_frames) * 100
            print(f"進度: {progress:.2f}% (幀 {frame_count}/{total_frames}) - 相似度: {similarity:.4f}")
        
        # 如果幀間差異足夠大，保存為新投影片
        if similarity < threshold:
            slide_path = os.path.join(output_folder, f"slide_{slide_count:03d}.png")
            cv2.imwrite(slide_path, frame)
            slide_count += 1
            last_frame = gray
            print(f"保存幀 {frame_count} 為投影片 {slide_count} (相似度: {similarity:.4f})")
    
    cap.release()
    print(f"處理完成！總共捕獲 {slide_count} 張投影片")
    return slide_count

def main():
    # 解析命令行參數
    parser = argparse.ArgumentParser(description='自動從視頻中捕獲投影片的快速入門示例')
    parser.add_argument('video_path', help='視頻文件路徑')
    parser.add_argument('--output', '-o', default='slides', help='輸出文件夾路徑')
    parser.add_argument('--threshold', '-t', type=float, default=0.95, help='相似度閾值 (0 到 1)')
    parser.add_argument('--interval', '-i', type=int, default=30, help='檢查間隔（每隔多少幀檢查一次）')
    
    args = parser.parse_args()
    
    # 處理視頻
    capture_slides(args.video_path, args.output, args.threshold, args.interval)

if __name__ == "__main__":
    main() 