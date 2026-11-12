"""PDF 仅查找填报页，印章主体由用户在原始渲染图上框选。"""

from dataclasses import dataclass, field
from pathlib import Path
import re
from threading import RLock
from time import perf_counter
from io import BytesIO

import numpy as np
import pymupdf
from PIL import Image

from ocr_runtime import create_ocr, models_are_ready, suppress_library_output
from ocr_utils import extract_ocr_items, normalized_label
from .extraction import candidate_from_bytes


def has_seal_title(parts) -> bool:
    text = "".join(normalized_label(part) for part in parts)
    return bool(re.search(r"(?:电子)?(?:印章|签章)(?:签章|印章)?(?:信息)?(?:填报|采集)表", text))


def file_signature(path):
    stat = path.stat()
    return stat.st_size, stat.st_mtime_ns


class PdfPageError(ValueError):
    """仅携带固定提示，避免暴露 PDF 内容和底层异常。"""


@dataclass(frozen=True)
class PdfSource:
    path: Path
    page_count: int
    matches: tuple[int, ...]
    signature: tuple[int, int]
    notice: str
    thumbnails: dict[int, bytes] = field(default_factory=dict)
    timings: dict[str, float] = field(default_factory=dict)
    pending_pages: tuple[int, ...] = ()  # 尚未执行整页 OCR 的页面，不能作为完整负向结果。


class PdfPages:
    def __init__(self):
        self.engine = None
        # 同一采集器串行使用本地模型，避免重复点击并发推理。
        self.lock = RLock()
        self.cached_source = None
        self.full_items = {}
        self.last_cache_hit = False
        self.title_items = {}  # 既有客户标题区域 OCR，按当前 PDF 页码复用。
        self.cache_key = None

    def prepare(self, path):
        key = (path.resolve(), file_signature(path))
        with self.lock:
            if key != self.cache_key:
                self.cache_key = key
                self.cached_source = None
                self.full_items = {}
                self.title_items = {}

    def find(self, path, progress, cancelled, engine=None, allow_ocr=True):
        with self.lock:
            return self._find(path, progress, cancelled, engine, allow_ocr)

    def _find(self, path, progress, cancelled, engine, allow_ocr):
        self.prepare(path)
        if cancelled():
            return None
        self.last_cache_hit = False
        signature = file_signature(path)
        if engine is not None:
            self.engine = engine
        if (self.cached_source is not None and self.cached_source.path == path
                and self.cached_source.signature == signature
                and (not allow_ocr or not self.cached_source.pending_pages)):
            self.last_cache_hit = True
            return self.cached_source
        previous = self.cached_source
        self.cached_source = None
        started = perf_counter()
        timings = dict(pdf_open=0.0, text=0.0, render=0.0, ocr=0.0, model=0.0,
                       thumbnails=0.0, title_match=0.0, ocr_calls=0, reused=0)
        thumbnails = dict(previous.thumbnails) if previous is not None else {}
        pending = []
        matches = []
        failures = 0
        opened = perf_counter()
        with self.lock, pymupdf.open(path) as document:
            timings["pdf_open"] = perf_counter() - opened
            if document.needs_pass:
                raise PdfPageError("PDF 已加密，请先提供可读取的本地 PDF")
            count = len(document)
            if count == 0:
                raise PdfPageError("PDF 没有可读取页面")
            unavailable = self.engine is None and not models_are_ready()
            for index in range(count):
                if cancelled():
                    return None
                progress(f"正在查找印章填报页：第 {index + 1}/{count} 页")
                try:
                    page = document[index]
                    ocr_pix = None
                    stage = perf_counter()
                    try:
                        matched = has_seal_title([page.get_text("text", sort=True)])
                    except Exception:
                        matched = False
                    timings["text"] += perf_counter() - stage
                    header = self.title_items.get(index + 1)
                    if not matched and header is not None:
                        stage = perf_counter()
                        matched = has_seal_title([item.text for item in header if item.confidence >= 0.6])
                        timings["title_match"] += perf_counter() - stage
                        if matched:
                            timings["reused"] += 1
                            print(f"[印章页复用] page={index + 1} status=header_hit")
                    if not matched:
                        if index + 1 in self.full_items:
                            stage = perf_counter()
                            matched = has_seal_title([item.text for item in self.full_items[index + 1] if item.confidence >= 0.6])
                            timings["title_match"] += perf_counter() - stage
                            timings["reused"] += 1
                            print(f"[印章页复用] page={index + 1} status={'full_hit' if matched else 'full_miss'}")
                            if not matched:
                                continue
                        else:
                            matched = None
                    if matched is None:
                        if not allow_ocr:
                            pending.append(index + 1)
                            print(f"[印章页复用] page={index + 1} status=deferred")
                            continue
                        if unavailable:
                            failures += 1
                            pending.append(index + 1)
                            continue
                        stage = perf_counter()
                        pix = page.get_pixmap(matrix=pymupdf.Matrix(120 / 72, 120 / 72),
                                              colorspace=pymupdf.csRGB, alpha=False)
                        rgb = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3)
                        ocr_pix = pix
                        timings["render"] += perf_counter() - stage
                        progress(f"本地 OCR 查找标题：第 {index + 1}/{count} 页")
                        with suppress_library_output():
                            if self.engine is None:
                                stage = perf_counter()
                                try:
                                    self.engine = create_ocr(use_local_models=True)
                                    timings["model"] += perf_counter() - stage
                                except Exception:
                                    unavailable = True
                                    raise
                            stage = perf_counter()
                            timings["ocr_calls"] += 1
                            try:
                                items = extract_ocr_items(self.engine.predict(rgb[:, :, ::-1].copy()), index + 1)
                            finally:
                                page_ocr_elapsed = perf_counter() - stage
                                timings["ocr"] += page_ocr_elapsed
                        items.sort(key=lambda item: (item.bbox[1], item.bbox[0]))
                        self.full_items[index + 1] = tuple(items)
                        stage = perf_counter()
                        matched = has_seal_title([item.text for item in items if item.confidence >= 0.6])
                        timings["title_match"] += perf_counter() - stage
                        print(f"[印章页OCR] page={index + 1} matched={matched} elapsed={page_ocr_elapsed:.3f}s")
                    if matched:
                        matches.append(index + 1)
                        if index + 1 in thumbnails:
                            continue
                        stage = perf_counter()
                        if ocr_pix is not None:
                            preview = Image.fromarray(rgb)
                            preview.thumbnail((160, 160))
                            buffer = BytesIO()
                            preview.save(buffer, format="PNG")
                            thumbnails[index + 1] = buffer.getvalue()
                        else:
                            thumbnails[index + 1] = page.get_pixmap(matrix=pymupdf.Matrix(0.2, 0.2), colorspace=pymupdf.csRGB, alpha=False).tobytes("png")
                        timings["thumbnails"] += perf_counter() - stage
                except Exception:
                    failures += 1
                    pending.append(index + 1)
                    print(f"[印章页处理] page={index + 1} status=failed")
            if signature != file_signature(path):
                raise PdfPageError("PDF 在查找过程中发生变化，请重新提取")
        notice = f"找到 {len(matches)} 个印章填报页面，请点击查看 / 框选。" if matches else "未自动找到电子印章填报表，请手动选择 PDF 页码。"
        if failures:
            notice += f"有 {failures} 页未完成 OCR 或渲染，本地模型缺失或处理失败时可手动选择页码。"
        if pending:
            notice += f"有 {len(set(pending))} 页尚未完成整页识别；需要时点击提取印章继续 OCR，或直接手动选择页码。"
        timings["total"] = perf_counter() - started
        timings["deferred"] = len(set(pending))
        self.cached_source = PdfSource(path, count, tuple(matches), signature, notice, thumbnails, timings, tuple(sorted(set(pending))))
        return self.cached_source

    def render(self, source, page_number):
        if not 1 <= page_number <= source.page_count:
            raise PdfPageError("页码超出 PDF 范围")
        if file_signature(source.path) != source.signature:
            raise PdfPageError("PDF 已发生变化，请重新提取页面")
        with self.lock, pymupdf.open(source.path) as document:
            if document.needs_pass:
                raise PdfPageError("PDF 已加密，无法渲染")
            pix = document[page_number - 1].get_pixmap(matrix=pymupdf.Matrix(300 / 72, 300 / 72),
                                                     colorspace=pymupdf.csRGB, alpha=False)
            candidate = candidate_from_bytes(pix.tobytes("png"), f"PDF 第 {page_number} 页（需框选）")
            if page_number not in source.thumbnails:
                source.thumbnails[page_number] = document[page_number - 1].get_pixmap(matrix=pymupdf.Matrix(0.2, 0.2), colorspace=pymupdf.csRGB, alpha=False).tobytes("png")
        if file_signature(source.path) != source.signature:
            raise PdfPageError("PDF 在渲染过程中发生变化，请重新提取")
        return candidate
