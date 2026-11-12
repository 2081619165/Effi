"""正式 OCR 运行时：管理本地模型、初始化和 PDF 页面渲染。"""

from __future__ import annotations

import os
import sys
import warnings
from contextlib import contextmanager, redirect_stderr, redirect_stdout
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pymupdf


SOURCE_DIR = Path(__file__).resolve().parent
# 开发环境读取项目模型，单文件 EXE 从 PyInstaller 解包目录读取模型。
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", SOURCE_DIR))
MODEL_CACHE_DIR = RESOURCE_DIR / ".paddlex"
OFFICIAL_MODEL_DIR = MODEL_CACHE_DIR / "official_models"
DETECTION_MODEL_NAME = "PP-OCRv5_mobile_det"
RECOGNITION_MODEL_NAME = "PP-OCRv5_mobile_rec"
DETECTION_MODEL_DIR = OFFICIAL_MODEL_DIR / DETECTION_MODEL_NAME
RECOGNITION_MODEL_DIR = OFFICIAL_MODEL_DIR / RECOGNITION_MODEL_NAME
REQUIRED_MODEL_FILES = ("inference.json", "inference.pdiparams", "inference.yml")
MODEL_SOURCE = (
    "https://paddle-model-ecology.bj.bcebos.com/"
    "paddlex/official_inference_model"
)
RENDER_DPI = 200
IMAGE_SUFFIXES = frozenset((".jpg", ".jpeg", ".png"))
LIBRARY_OUTPUT = open(os.devnull, "w", encoding="utf-8")

# Windows 无控制台 EXE 中标准流可能为 None，统一安全丢弃终端输出。
if sys.stdout is None:
    sys.stdout = LIBRARY_OUTPUT
if sys.stderr is None:
    sys.stderr = LIBRARY_OUTPUT

# 模型固定保存在项目目录；正常识别时跳过联网源探测。
os.environ.setdefault("PADDLE_PDX_CACHE_HOME", str(MODEL_CACHE_DIR))
os.environ.setdefault("PADDLE_PDX_MODEL_SOURCE", "bos")
os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
os.environ.setdefault("GLOG_minloglevel", "2")


@contextmanager
def suppress_library_output() -> Iterable[None]:
    """阻止第三方库把模型日志或推理细节写入终端。"""
    sys.stdout.flush()
    sys.stderr.flush()
    saved_descriptors: dict[int, int] = {}
    try:
        # 同时重定向文件描述符，覆盖第三方子进程和原生库输出。
        for descriptor in (1, 2):
            saved_descriptor: int | None = None
            try:
                saved_descriptor = os.dup(descriptor)
                os.dup2(LIBRARY_OUTPUT.fileno(), descriptor)
                saved_descriptors[descriptor] = saved_descriptor
            except OSError:
                # PyInstaller windowed 模式可能没有对应的系统文件描述符。
                if saved_descriptor is not None:
                    os.close(saved_descriptor)
        with redirect_stdout(LIBRARY_OUTPUT), redirect_stderr(LIBRARY_OUTPUT):
            yield
    finally:
        LIBRARY_OUTPUT.flush()
        for descriptor, saved_descriptor in saved_descriptors.items():
            os.dup2(saved_descriptor, descriptor)
            os.close(saved_descriptor)


def load_paddle_ocr_class() -> Any:
    """延迟导入，避免第三方库在正常报告中输出无关信息。"""
    with suppress_library_output():
        from paddleocr import PaddleOCR

    return PaddleOCR


def create_ocr(use_local_models: bool) -> Any:
    PaddleOCR = load_paddle_ocr_class()
    model_options: dict[str, object] = {
        "text_detection_model_name": DETECTION_MODEL_NAME,
        "text_recognition_model_name": RECOGNITION_MODEL_NAME,
    }
    if use_local_models:
        model_options.update(
            {
                "text_detection_model_dir": str(DETECTION_MODEL_DIR),
                "text_recognition_model_dir": str(RECOGNITION_MODEL_DIR),
            }
        )

    return PaddleOCR(
        device="cpu",
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        **model_options,
    )


def render_page(page: pymupdf.Page) -> np.ndarray:
    """在内存中将 PDF 页面渲染为 200 DPI BGR 图片。"""
    scale = RENDER_DPI / 72
    pixmap = page.get_pixmap(
        matrix=pymupdf.Matrix(scale, scale),
        colorspace=pymupdf.csRGB,
        alpha=False,
    )
    rgb = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
        pixmap.height, pixmap.width, 3
    )
    return rgb[:, :, ::-1].copy()


class ImageInputError(ValueError):
    """只携带固定的安全图片读取提示。"""


def read_image(path: Path) -> np.ndarray:
    """在内存中读取图片，统一为方向正确、白底的 BGR 像素坐标。"""
    from PIL import Image, ImageOps

    if path.suffix.lower() not in IMAGE_SUFFIXES:
        raise ImageInputError("不支持的图片格式，请选择 JPG、JPEG 或 PNG。")
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(path) as source:
                if source.format not in {"JPEG", "PNG"}:
                    raise ImageInputError("图片实际格式不受支持，请选择 JPG、JPEG 或 PNG。")
                width, height = source.size
                if min(width, height) < 480 or max(width, height) < 800:
                    raise ImageInputError("图片尺寸过小，请提供短边至少 480、长边至少 800 像素的清晰申请表。")
                if width * height > 25_000_000:
                    raise ImageInputError("图片尺寸过大，请缩小至 2500 万像素以内后重试。")
                if getattr(source, "n_frames", 1) != 1:
                    raise ImageInputError("不支持多帧图片，请使用单张申请表图片。")
                with ImageOps.exif_transpose(source) as oriented:
                    # 大图适度缩小；所有后续定位与裁剪均使用缩放后的同一坐标系。
                    oriented.thumbnail((2400, 2400), Image.Resampling.LANCZOS)
                    if "A" in oriented.getbands() or "transparency" in oriented.info:
                        with oriented.convert("RGBA") as rgba:
                            with Image.new("RGBA", rgba.size, "white") as background:
                                background.alpha_composite(rgba)
                                with background.convert("RGB") as rgb:
                                    return np.asarray(rgb)[:, :, ::-1].copy()
                    with oriented.convert("RGB") as rgb:
                        return np.asarray(rgb)[:, :, ::-1].copy()
    except ImageInputError:
        raise
    except (Image.DecompressionBombWarning, Image.DecompressionBombError):
        raise ImageInputError("图片尺寸过大，无法安全读取，请缩小后重试。") from None
    except Exception:
        raise ImageInputError("图片无法读取或已损坏，请重新保存为 JPG、JPEG 或 PNG 后重试。") from None


def models_are_ready() -> bool:
    """只有两个本地模型文件完整时才允许离线识别。"""
    return all(
        (model_dir / filename).is_file()
        for model_dir in (DETECTION_MODEL_DIR, RECOGNITION_MODEL_DIR)
        for filename in REQUIRED_MODEL_FILES
    )
