# -*- coding: utf-8 -*-
"""
通用视频毫秒级精准抽帧与 4x4 16宫格实证图合成引擎
用途：
1. 从原片视频按精确时间戳提取 16 张核心角色正脸特写（C01–C16）与 16 张关键道具/屏幕文字微距图（P01–P16）
2. 自动合成为高清 16 宫格实证大图，杜绝角色认错与文字幻觉
"""
import os, cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont

def extract_frame_at_sec(video_path: str, sec: float, out_jpg: str):
    cap = cv2.VideoCapture(video_path)
    cap.set(cv2.CAP_PROP_POS_MSEC, sec * 1000.0)
    ok, frame = cap.read()
    cap.release()
    if not ok:
        raise RuntimeError(f"Failed to grab frame at {sec}s from {video_path}")
    os.makedirs(os.path.dirname(os.path.abspath(out_jpg)), exist_ok=True)
    cv2.imwrite(out_jpg, frame, [int(cv2.IMWRITE_JPEG_QUALITY), 95])
    return out_jpg

def build_16_grid(image_label_pairs, out_grid_path: str, cell_w=640, cell_h=360):
    """將 16 张 (img_path, label_text) 合成 4x4 高清 16 宫格比对图"""
    assert len(image_label_pairs) == 16, "Must provide exactly 16 images for 4x4 grid!"
    canvas = Image.new("RGB", (cell_w * 4, cell_h * 4), (20, 20, 20))
    draw = ImageDraw.Draw(canvas)
    for idx, (img_p, label) in enumerate(image_label_pairs):
        r, c = divmod(idx, 4)
        im = Image.open(img_p).convert("RGB").resize((cell_w, cell_h))
        canvas.paste(im, (c * cell_w, r * cell_h))
        # Draw bottom label bar
        x0, y0 = c * cell_w, (r + 1) * cell_h - 36
        draw.rectangle([x0, y0, x0 + cell_w, y0 + 36], fill=(0, 0, 0))
        draw.text((x0 + 10, y0 + 8), label, fill=(255, 255, 255))
    os.makedirs(os.path.dirname(os.path.abspath(out_grid_path)), exist_ok=True)
    canvas.save(out_grid_path, quality=95)
    return out_grid_path
