"""独立印章提取验证；候选图片始终需要人工核对。"""

from __future__ import annotations

import argparse
from io import BytesIO
import json
from pathlib import Path
import re
import sys
import tempfile
from time import perf_counter
import warnings

import cv2
from docx import Document
import numpy as np
from PIL import Image, ImageOps
import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ocr_utils import extract_ocr_items, normalized_label


def is_collection_label(text: str) -> bool:
    text = normalized_label(text)
    return bool(re.search(r"(?:印章|签章|签名).{0,8}(?:采集|填报)", text))


def read_pixels(source) -> np.ndarray:
    # 仅矫正 EXIF 方向并将透明部分置白，不缩放或增强原图。
    with warnings.catch_warnings():
        warnings.simplefilter("error", Image.DecompressionBombWarning)
        with Image.open(source) as image:
            if getattr(image, "n_frames", 1) != 1:
                raise ValueError("multi_frame")
            image.load()
            oriented = ImageOps.exif_transpose(image).convert("RGBA")
            background = Image.new("RGBA", oriented.size, "white")
            return np.array(Image.alpha_composite(background, oriented).convert("RGB"))


def frames(image: np.ndarray) -> list[tuple[int, int, int, int]]:
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)[1]
    contours, _ = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    height, width = gray.shape
    found = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        polygon = cv2.approxPolyDP(contour, cv2.arcLength(contour, True) * 0.015, True)
        if len(polygon) != 4 or not cv2.isContourConvex(polygon):
            continue
        if w < width * 0.12 or h < height * 0.04 or w * h > width * height * 0.85:
            continue
        if cv2.contourArea(contour) < w * h * 0.9:
            continue
        box = (x, y, x + w, y + h)
        # 合并同一边框内外轮廓，保留更靠内的边界。
        duplicate = next((i for i, old in enumerate(found)
                          if max(abs(a - b) for a, b in zip(box, old)) < 10), None)
        if duplicate is None:
            found.append(box)
        elif w * h < (found[duplicate][2] - found[duplicate][0]) * (found[duplicate][3] - found[duplicate][1]):
            found[duplicate] = box
    return sorted(found, key=lambda box: (box[1], box[0]))


def locate_regions(image: np.ndarray, labels: list) -> list:
    regions = set()
    for x0, y0, x1, y1 in frames(image):
        for text, (lx0, ly0, lx1, ly1) in labels:
            if not is_collection_label(text):
                continue
            label_height = max(1, ly1 - ly0)
            aligned = lx1 > x0 and lx0 < x1
            above = 0 <= y0 - ly1 <= label_height * 1.5
            inside_top = y0 <= ly0 and ly1 < y0 + (y1 - y0) * 0.25
            if aligned and (above or inside_top):
                inset = max(4, round(min(x1 - x0, y1 - y0) * 0.015))
                top = max(y0 + inset, round(ly1 + inset)) if inside_top else y0 + inset
                regions.add((x0 + inset, top, x1 - inset, y1 - inset))
    return sorted(regions, key=lambda box: (box[1], box[0]))


def trim_candidate(image: np.ndarray) -> tuple[np.ndarray | None, str]:
    if image.size == 0:
        return None, "采集区为空或尺寸无效"
    # 阈值仅用于计算裁剪边界；输出始终取原像素，保留全部颜色。
    ink = np.min(image, axis=2) < 235
    ys, xs = np.nonzero(ink)
    if len(xs) < 20:
        return None, "采集区为空白或有效内容过少"
    height, width = ink.shape
    left, top, right, bottom = int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1
    padding = max(5, round(max(right - left, bottom - top) * 0.03))
    if min(left, top, width - right, height - bottom) < padding:
        return None, "内容靠近采集区边界，可能包含边框或存在截断，需人工定位"
    return image[top - padding:bottom + padding, left - padding:right + padding].copy(), ""


class Probe:
    def __init__(self, output: Path, local_ocr: bool):
        self.output = output
        self.local_ocr = local_ocr
        self.ocr = None
        self.report = {"region_count": 0, "candidate_count": 0,
                       "human_confirmed": False, "items": []}

    def ocr_labels(self, image: np.ndarray, page: int) -> list:
        if not self.local_ocr:
            return []
        from ocr_runtime import create_ocr, models_are_ready, suppress_library_output
        if not models_are_ready():
            raise ValueError("local_models_missing")
        with suppress_library_output():
            if self.ocr is None:
                self.ocr = create_ocr(use_local_models=True)
            items = extract_ocr_items(self.ocr.predict(image[:, :, ::-1].copy()), page)
        # OCR 文本只留在内存，不写入诊断或终端。
        return [(item.text, item.bbox) for item in items if item.confidence >= 0.8]

    def capture(self, image: np.ndarray, page: int | None, method: str, box=None):
        self.report["region_count"] += 1
        region = image
        if box is not None:
            x0, y0, x1, y1 = box
            region = image[y0:y1, x0:x1]
        crop, reason = trim_candidate(region)
        item = {"page": page, "method": method, "region_found": True,
                "candidate_saved": crop is not None, "reason": reason}
        if crop is not None:
            self.report["candidate_count"] += 1
            name = f"candidate_{self.report['candidate_count']:03d}.png"
            Image.fromarray(crop).save(self.output / name)
            item.update(file=name, width=int(crop.shape[1]), height=int(crop.shape[0]))
        self.report["items"].append(item)

    def raster(self, image: np.ndarray, labels: list, page: int | None, method: str):
        regions = locate_regions(image, labels)
        if not regions:
            self.report["items"].append({"page": page, "method": method,
                "region_found": False, "candidate_saved": False,
                "reason": "未找到明确采集标签与闭合矩形边框的对应关系；可使用本地 OCR 或人工指定区域"})
        for box in regions:
            self.capture(image, page, method, box)

    def pdf(self, path: Path, args):
        with pymupdf.open(path) as document:
            if document.needs_pass:
                raise ValueError("encrypted")
            pages = range(len(document))
            if args.region:
                if args.page is None or not 1 <= args.page <= len(document):
                    raise ValueError("manual_page")
                pages = [args.page - 1]
            for index in pages:
                page = document[index]
                scale = args.dpi / 72
                pixmap = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale),
                                          colorspace=pymupdf.csRGB, alpha=False)
                image = np.frombuffer(pixmap.samples, np.uint8).reshape(pixmap.height, pixmap.width, 3)
                if args.region:
                    self.manual(image, index + 1, args.region)
                    continue
                labels = []
                for block in page.get_text("dict")["blocks"]:
                    for line in block.get("lines", []):
                        text = "".join(span["text"] for span in line["spans"])
                        rect = pymupdf.Rect(line["bbox"]) * page.rotation_matrix
                        labels.append((text, tuple(value * scale for value in rect)))
                if not any(is_collection_label(text) for text, _ in labels):
                    labels = self.ocr_labels(image, index + 1)
                self.raster(image, labels, index + 1, "PDF整页渲染＋标签与边框")

    def manual(self, image: np.ndarray, page: int | None, region):
        x, y, width, height = region
        if x < 0 or y < 0 or width <= 0 or height <= 0 or x + width > image.shape[1] or y + height > image.shape[0]:
            raise ValueError("invalid_region")
        self.capture(image, page, "人工指定采集区", (x, y, x + width, y + height))

    def docx(self, path: Path):
        document = Document(path)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
              "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
              "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
        found = set()
        # 只关联明确带有采集标签的表格行；正文中的任意图片不作为印章。
        for row in document.element.iter("{" + ns["w"] + "}tr"):
            text = "".join(node.text or "" for node in row.iter("{" + ns["w"] + "}t"))
            if not is_collection_label(text):
                continue
            for blip in row.iter("{" + ns["a"] + "}blip"):
                relationship = blip.get("{" + ns["r"] + "}embed")
                if not relationship or relationship in found:
                    continue
                found.add(relationship)
                part = document.part.related_parts.get(relationship)
                if part is None:
                    continue
                try:
                    image = read_pixels(BytesIO(part.blob))
                except Exception:
                    self.report["items"].append({"page": None, "region_found": True,
                        "candidate_saved": False, "method": "DOCX采集行内嵌图片",
                        "reason": "图片编码不受支持，矢量图或绘图对象需人工导出 PDF 后验证"})
                    continue
                self.capture(image, None, "DOCX采集行内嵌原始图片")
        if not found:
            self.report["items"].append({"page": None, "region_found": False,
                "candidate_saved": False, "method": "DOCX表格关联",
                "reason": "未找到采集标签所在表格行中的内嵌图片；浮动图、跨行采集区或绘图对象需人工导出 PDF 后验证"})


def main() -> int:
    parser = argparse.ArgumentParser(description="本地印章候选提取验证，输出必须人工确认")
    parser.add_argument("input", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--local-ocr", action="store_true", help="仅启用项目现有完整本地模型")
    parser.add_argument("--dpi", type=int, choices=(200, 300), default=300)
    parser.add_argument("--page", type=int, help="人工区域对应的 PDF 页码，从 1 开始")
    parser.add_argument("--region", type=int, nargs=4, metavar=("X", "Y", "W", "H"),
                        help="人工指定采集区，使用渲染后/EXIF矫正后图片像素坐标")
    args = parser.parse_args()
    started = perf_counter()
    try:
        args.output_dir.mkdir(parents=True, exist_ok=True)
        output = Path(tempfile.mkdtemp(prefix="seal_probe_", dir=args.output_dir))
    except OSError:
        print("无法创建输出目录，请检查本地目录权限和可用空间。")
        return 2
    probe = Probe(output, args.local_ocr)
    kind = args.input.suffix.lower()
    try:
        if kind == ".pdf":
            probe.pdf(args.input, args)
        elif kind == ".docx":
            if args.region or args.page:
                raise ValueError("docx_region")
            probe.docx(args.input)
        elif kind in (".png", ".jpg", ".jpeg"):
            image = read_pixels(args.input)
            if args.region:
                probe.manual(image, None, args.region)
            else:
                probe.raster(image, probe.ocr_labels(image, 1), None, "图片本地OCR标签＋边框")
        else:
            raise ValueError("unsupported")
    except Exception as exc:
        # 第三方异常可能包含客户路径或内容，统一使用固定诊断文字。
        reasons = {
            "local_models_missing": "现有本地 OCR 模型文件不完整，已停止，未尝试下载模型",
            "encrypted": "PDF 需要密码，无法读取",
            "manual_page": "人工指定 PDF 区域时必须提供有效的 --page 页码",
            "invalid_region": "人工指定区域超出图片边界或尺寸无效",
            "docx_region": "DOCX 不支持像素区域或页码参数，请先人工导出 PDF",
            "unsupported": "文件类型不受支持",
            "multi_frame": "暂不支持多帧图片",
        }
        code = exc.args[0] if isinstance(exc, ValueError) and exc.args and isinstance(exc.args[0], str) else None
        probe.report["error"] = reasons.get(code, "处理未完成。文件无法读取、图片编码不受支持、模型初始化或推理失败。已输出的候选也需人工复核。")
    probe.report.update(file_type=kind, elapsed_seconds=round(perf_counter() - started, 3),
                        status="needs_human_review" if probe.report["candidate_count"] else "failed",
                        note="区域定位和候选输出不代表印章准确或完整；所有候选均需人工确认。")
    if "error" in probe.report:
        probe.report["status"] = "incomplete"
    diagnostic = json.dumps(probe.report, ensure_ascii=False, indent=2)
    try:
        (output / "diagnostics.json").write_text(diagnostic, encoding="utf-8")
    except OSError:
        print("无法保存诊断文件，请检查输出目录权限和可用空间。")
        return 2
    print(diagnostic)
    print("本次输出子目录：" + output.name)
    return 0 if probe.report["status"] == "needs_human_review" else 2


if __name__ == "__main__":
    raise SystemExit(main())
