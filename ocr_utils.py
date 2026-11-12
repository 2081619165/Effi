"""OCR 结果的通用内存数据结构和解析工具。"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from typing import Any, Iterable


BBox = tuple[float, float, float, float]


@dataclass(frozen=True)
class OCRItem:
    item_id: int
    page_number: int
    text: str
    bbox: BBox
    confidence: float


def normalized_label(text: str) -> str:
    normalized_chars: list[str] = []
    for character in text:
        for normalized_character in unicodedata.normalize("NFKC", character):
            if normalized_character.isalnum():
                normalized_chars.append(normalized_character.lower())
    return "".join(normalized_chars)


def bbox_from_result(data: dict[str, Any], index: int) -> BBox | None:
    boxes = data.get("rec_boxes", [])
    if index < len(boxes):
        values = list(boxes[index])
        if len(values) >= 4:
            x0, y0, x1, y1 = (float(value) for value in values[:4])
            return x0, y0, x1, y1

    polygons = data.get("rec_polys", data.get("dt_polys", []))
    if index < len(polygons):
        points = list(polygons[index])
        if points:
            xs = [float(point[0]) for point in points]
            ys = [float(point[1]) for point in points]
            return min(xs), min(ys), max(xs), max(ys)
    return None


def extract_ocr_items(
    results: Iterable[Any], page_number: int
) -> list[OCRItem]:
    """从 PaddleOCR 3.x 结果中读取文本、坐标和置信度，仅保存在内存。"""
    items: list[OCRItem] = []
    next_item_id = 0

    for result in results:
        payload = result.json
        if callable(payload):
            payload = payload()
        data = payload.get("res", payload)
        texts = data.get("rec_texts", [])
        scores = data.get("rec_scores", [])

        for index, text in enumerate(texts):
            bbox = bbox_from_result(data, index)
            if bbox is None or not text.strip():
                continue
            confidence = float(scores[index]) if index < len(scores) else 0.0
            items.append(
                OCRItem(
                    item_id=next_item_id,
                    page_number=page_number,
                    text=text,
                    bbox=bbox,
                    confidence=confidence,
                )
            )
            next_item_id += 1

    return items
