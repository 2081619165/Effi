"""以采集标签和矩形边界定位扫描图，保留原始像素供人工复核。"""

from io import BytesIO
import re

import cv2
import numpy as np
from PIL import Image


def collection_label(text):
    return bool(re.search(r"(?:印章|签章|签名|个人章).{0,12}(?:采集|填报)", re.sub(r"\s+", "", text)))


class LocalLabels:
    def __init__(self):
        self.engine = None
        self.unavailable = False

    def read(self, image):
        from ocr_runtime import create_ocr, models_are_ready, suppress_library_output
        from ocr_utils import extract_ocr_items
        if self.unavailable or not models_are_ready():
            self.unavailable = True
            return []
        try:
            with suppress_library_output():
                if self.engine is None:
                    self.engine = create_ocr(use_local_models=True)
                items = extract_ocr_items(self.engine.predict(image[:, :, ::-1].copy()), 1)
            return [(item.text, item.bbox) for item in items if item.confidence >= 0.75]
        except Exception:
            self.unavailable = True
            return []


def frame_boxes(image):
    # 与独立验证工具相同，使用边框而不是章的颜色定位。
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                  cv2.THRESH_BINARY_INV, 31, 10)
    contours, _ = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    height, width = gray.shape
    boxes = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        polygon = cv2.approxPolyDP(contour, 0.015 * cv2.arcLength(contour, True), True)
        if len(polygon) != 4 or not cv2.isContourConvex(polygon):
            continue
        if w < width * 0.08 or h < height * 0.035 or w * h > width * height * 0.85:
            continue
        if cv2.contourArea(contour) < w * h * 0.9:
            continue
        box = (x, y, x + w, y + h)
        if not any(max(abs(a - b) for a, b in zip(box, old)) < 10 for old in boxes):
            boxes.append(box)
    return sorted(boxes, key=lambda box: (box[1], box[0]))


def scan_crops(image, labels):
    boxes = frame_boxes(image)
    regions = set()
    for box in boxes:
        x0, y0, x1, y1 = box
        for text, (lx0, ly0, lx1, ly1) in labels:
            if not collection_label(text) or lx1 <= x0 or lx0 >= x1:
                continue
            above = 0 <= y0 - ly1 <= max(1, ly1 - ly0) * 2
            inside = y0 <= ly0 and ly1 < y0 + (y1 - y0) * 0.2
            if above or inside:
                regions.add((x0, max(y0, round(ly1)) if inside else y0, x1, y1))
                if above:
                    # 同一采集行中的并列框一起保留，不任意选择其中一个。
                    for other in boxes:
                        if abs(other[1] - y0) < 8 and abs(other[3] - y1) < 8:
                            regions.add(other)
    result = []
    for x0, y0, x1, y1 in sorted(regions, key=lambda box: (box[1], box[0])):
        if (x1 - x0) * (y1 - y0) > image.shape[0] * image.shape[1] * 0.35:
            continue
        if any(x0 <= a and y0 <= b and x1 >= c and y1 >= d
               and (a, b, c, d) != (x0, y0, x1, y1)
               for a, b, c, d in regions):
            continue
        inset = max(4, round(min(x1 - x0, y1 - y0) * 0.015))
        interior = image[y0 + inset:y1 - inset, x0 + inset:x1 - inset]
        if not interior.size:
            continue
        # 自适应背景差异只用于判断是否空白，不据此裁断浅色笔画。
        light = interior.min(axis=2).astype(np.int16)
        background = np.percentile(light, 95)
        if np.count_nonzero(background - light > 2) < 20:
            continue
        output = BytesIO()
        Image.fromarray(interior).save(output, format="PNG")
        result.append((output.getvalue(), y0 + inset, x0 + inset))
    return result
