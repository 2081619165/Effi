"""机构申请信息本地提取与人工核对界面。"""

from __future__ import annotations

import os
import math
import re
import sys
import threading
import time
import tkinter as tk
import traceback
import unicodedata
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
from typing import Any, Callable

import cv2
import customtkinter as ctk
import numpy as np
import pymupdf

from ocr_runtime import (
    IMAGE_SUFFIXES,
    ImageInputError,
    RENDER_DPI,
    create_ocr,
    models_are_ready,
    read_image,
    render_page,
    suppress_library_output,
)
from ocr_utils import OCRItem, extract_ocr_items, normalized_label
from seals.panel import SealPanel
from word_template_filler import (
    INSTITUTION_TEMPLATE,
    SUPPORTED_TEMPLATE_NAMES,
    TemplateStructureError,
    fill_word_template,
)


Validator = Callable[[str], bool]
FORM_TITLE = "机构个人数字证书业务申请表"
TITLE_SCAN_DPI = 90
SOURCE_DIR = Path(__file__).resolve().parent
# 开发环境读取源码目录，单文件 EXE 从 PyInstaller 解包目录读取资源。
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", SOURCE_DIR))
PROJECT_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else SOURCE_DIR
)
TEMPLATE_DIR = RESOURCE_DIR / "templates"
WORD_TEMPLATE_NAMES = SUPPORTED_TEMPLATE_NAMES


class TargetPageSelectionRequired(Exception):
    def __init__(self, page_count: int) -> None:
        super().__init__("target page not reliably located")
        self.page_count = page_count


@dataclass(frozen=True)
class ExtractFieldSpec:
    key: str
    name: str
    aliases: tuple[str, ...]
    validator: Validator
    role: str | None = None
    blocked_suffixes: tuple[str, ...] = ()


@dataclass(frozen=True)
class PageOCR:
    items: tuple[OCRItem, ...]
    width: int
    height: int


@dataclass(frozen=True)
class Candidate:
    value: str
    confidence: float
    score: float
    format_valid: bool


@dataclass(frozen=True)
class CandidateRegion:
    page_number: int
    bbox: tuple[float, float, float, float]


@dataclass(frozen=True)
class FieldDiagnostic:
    label_found: bool = False
    region_located: bool = False
    candidate_blocks: int = 0
    format_valid: bool = False
    reason: str = "尚未提取"


@dataclass(frozen=True)
class ExtractionResult:
    value: str = ""
    recognized: bool = False
    diagnostic: FieldDiagnostic = FieldDiagnostic()
    regions: tuple[CandidateRegion, ...] = ()


@dataclass(frozen=True)
class StageMetric:
    calls: int = 0
    elapsed: float = 0.0


@dataclass(frozen=True)
class StageTimings:
    model_initialization: StageMetric = StageMetric()
    pdf_open: StageMetric = StageMetric()
    page_detection: StageMetric = StageMetric()
    page_rendering: StageMetric = StageMetric()
    page_ocr: StageMetric = StageMetric()
    local_supplement: StageMetric = StageMetric()
    field_matching: StageMetric = StageMetric()
    model_reused: bool = False
    image_loading: StageMetric = StageMetric()
    seal_lookup: StageMetric = StageMetric()  # 印章页查找墙钟时间，内部子阶段另行输出。
    title_ocr: StageMetric = StageMetric()  # 客户标题 OCR，包含在页面定位耗时中。
    title_render: StageMetric = StageMetric()
    lock_wait: StageMetric = StageMetric()
    local_render: StageMetric = StageMetric()


def print_performance_report(timings: StageTimings, total_elapsed: float) -> None:
    """仅向终端输出固定阶段名称、调用次数和耗时。"""
    input_stages = (
        (("图片读取与预处理", timings.image_loading),)
        if timings.image_loading.calls else (
        ("PDF 打开", timings.pdf_open),
        ("目标页定位", timings.page_detection),
        ("页面渲染", timings.page_rendering),
        )
    )
    stages = (
        ("OCR 模型初始化", timings.model_initialization),
        *input_stages,
        ("申请表格 OCR（既有精识别）", timings.page_ocr),
        ("局部补充 OCR", timings.local_supplement),
        ("局部图片渲染", timings.local_render),
        ("字段定位与匹配", timings.field_matching),
        ("客户流程内印章查找（自动轻量查找在交付后另行统计）", timings.seal_lookup),
        ("客户标题OCR（页面定位子阶段）", timings.title_ocr),
        ("标题渲染（页面定位子阶段）", timings.title_render),
        ("模型锁等待（工作线程计时外）", timings.lock_wait),
    )
    print("[性能统计] 本轮文件信息提取")
    for name, metric in stages:
        print(f"[性能统计] {name}: 调用 {metric.calls} 次，累计 {metric.elapsed:.3f} 秒")
    print(f"[性能统计] 总耗时: {total_elapsed:.3f} 秒")
    accounted = sum(metric.elapsed for metric in (
        timings.model_initialization, timings.pdf_open, timings.image_loading,
        timings.page_detection, timings.page_rendering, timings.page_ocr,
        timings.local_supplement, timings.local_render, timings.field_matching, timings.seal_lookup))
    print(f"[性能统计] 其他处理及调度: {max(0.0, total_elapsed - accounted):.3f} 秒")
    print("[性能统计] 父子阶段有包含关系，不将各行简单相加作为总耗时")


def _safe_exception_message(exc: Exception, private_path: Path | None) -> str:
    """保留诊断信息，同时隐藏客户路径及常见敏感字段格式。"""
    message = str(exc).replace("\r", " ").replace("\n", " ")
    if private_path is not None:
        for private_value in (
            str(private_path),
            private_path.as_posix(),
            private_path.name,
        ):
            if private_value:
                message = re.sub(
                    re.escape(private_value),
                    "[客户文件路径已隐藏]",
                    message,
                    flags=re.IGNORECASE,
                )
    message = re.sub(r"(?i)(?:[A-Z]:\\|\\\\)[^\s\"']+", "[本地路径已隐藏]", message)
    message = re.sub(
        r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+",
        "[邮箱已隐藏]",
        message,
    )
    message = re.sub(r"(?<!\d)1[3-9]\d{9}(?!\d)", "[手机号已隐藏]", message)
    message = re.sub(r"(?<!\w)\d{17}[0-9Xx](?!\w)", "[证件号已隐藏]", message)
    return message[:500] if message else "无附加错误信息"


def print_extraction_exception(
    stage: str, exc: Exception, private_path: Path | None
) -> None:
    """调试版控制台输出安全异常堆栈，不记录 OCR 内容。"""
    # windowed 正式版没有原始标准错误流，仅在调试控制台输出详情。
    if getattr(sys, "__stderr__", None) is None:
        return

    sensitive_stage = stage.startswith(
        ("申请信息 OCR", "字段定位与匹配", "局部补充 OCR")
    )
    print(f"[提取异常] 阶段：{stage}", file=sys.stderr)

    chain: list[tuple[str, BaseException]] = [("最外层异常", exc)]
    seen = {id(exc)}
    current: BaseException = exc
    while True:
        if current.__cause__ is not None:
            relation = "__cause__（直接原因）"
            next_exception = current.__cause__
        elif current.__context__ is not None and not current.__suppress_context__:
            relation = "__context__（处理期间异常）"
            next_exception = current.__context__
        else:
            break
        if id(next_exception) in seen:
            break
        seen.add(id(next_exception))
        chain.append((relation, next_exception))
        current = next_exception

    print(f"[提取异常] 异常链：{len(chain)} 层", file=sys.stderr)
    for index, (relation, chain_exc) in enumerate(chain, start=1):
        if sensitive_stage and not isinstance(chain_exc, (ImportError, OSError)):
            message = "为保护客户数据，已隐藏该阶段的异常正文"
        else:
            message = _safe_exception_message(chain_exc, private_path)
        exception_type = (
            f"{type(chain_exc).__module__}.{type(chain_exc).__qualname__}"
        )
        print(
            f"[提取异常链 {index}] 关系：{relation}", file=sys.stderr
        )
        print(
            f"[提取异常链 {index}] 类型：{exception_type}",
            file=sys.stderr,
        )
        print(f"[提取异常链 {index}] 信息：{message}", file=sys.stderr)
        print(f"[提取异常链 {index}] 堆栈：", file=sys.stderr)
        for frame in traceback.extract_tb(chain_exc.__traceback__):
            location = "/".join(Path(frame.filename).parts[-3:])
            print(
                f"  {location}:{frame.lineno} in {frame.name}",
                file=sys.stderr,
            )
    sys.stderr.flush()


def extraction_failure_summary(stage: str, exc: Exception) -> str:
    """生成不包含异常原文的界面失败提示。"""
    if isinstance(exc, ImageInputError):
        return str(exc)
    if isinstance(exc, ModuleNotFoundError):
        detail = "缺少运行模块"
    elif isinstance(exc, ImportError):
        detail = "模块或动态库加载失败"
    elif isinstance(exc, FileNotFoundError):
        detail = "本地资源文件缺失"
    elif isinstance(exc, OSError):
        detail = "运行库或文件不可用"
    else:
        detail = "发生运行错误"
    return f"{stage}失败：{detail}（{type(exc).__name__}）"


def display_value(text: str) -> str:
    return unicodedata.normalize("NFKC", text).strip(" \t\r\n:：;；_-—|")


def compact_value(text: str) -> str:
    return re.sub(r"[\s()（）_\-—]", "", display_value(text))


def contains_form_title(parts: list[str] | tuple[str, ...]) -> bool:
    """兼容标题空格、符号和 OCR 分段。"""
    combined = "".join(normalized_label(part) for part in parts)
    return normalized_label(FORM_TITLE) in combined


def find_title_pages_in_text_layer(document: pymupdf.Document) -> list[int]:
    matches: list[int] = []
    for page_number, page in enumerate(document, start=1):
        if contains_form_title([page.get_text("text")]):
            matches.append(page_number)
    return matches


def render_title_scan(page: pymupdf.Page) -> np.ndarray:
    """只低分辨率渲染页面上部，用于扫描件标题筛选。"""
    clip = pymupdf.Rect(
        page.rect.x0,
        page.rect.y0,
        page.rect.x1,
        page.rect.y0 + page.rect.height * 0.35,
    )
    scale = TITLE_SCAN_DPI / 72.0
    pixmap = page.get_pixmap(
        matrix=pymupdf.Matrix(scale, scale),
        clip=clip,
        colorspace=pymupdf.csRGB,
        alpha=False,
    )
    rgb = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
        pixmap.height, pixmap.width, 3
    )
    return rgb[:, :, ::-1].copy()


def find_title_pages_with_low_resolution_ocr(
    document: pymupdf.Document, ocr: Any, cached_items=None, roi_cache=None, metrics=None
) -> list[int]:
    matches: list[int] = []
    for page_number, page in enumerate(document, start=1):
        if cached_items and contains_form_title([item.text for item in cached_items.get(page_number, ())]):
            matches.append(page_number)
            continue
        if roi_cache is not None and page_number in roi_cache:
            items = list(roi_cache[page_number])
        else:
            started = time.perf_counter()
            image = render_title_scan(page)
            if metrics is not None:
                metrics["render"] += time.perf_counter() - started
            started = time.perf_counter()
            if metrics is not None:
                metrics["calls"] += 1
            try:
                with suppress_library_output():
                    results = ocr.predict(image)
                    items = extract_ocr_items(results, page_number)
            finally:
                if metrics is not None:
                    metrics["ocr"] += time.perf_counter() - started
            if roi_cache is not None:
                roi_cache[page_number] = tuple(items)
        items.sort(key=lambda item: (box_center_y(item), item.bbox[0]))
        if contains_form_title([item.text for item in items]):
            matches.append(page_number)
    return matches


def detect_application_table_region(
    image: np.ndarray,
) -> tuple[int, int, int, int] | None:
    """根据密集长横线定位上方申请信息表格，不使用固定像素裁剪。"""
    page_height, page_width = image.shape[:2]
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(
        gray,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        31,
        12,
    )
    kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT, (max(40, page_width // 12), 1)
    )
    horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
    contours, _ = cv2.findContours(
        horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    segments: list[tuple[int, int, int]] = []
    for contour in contours:
        x, y, width, height = cv2.boundingRect(contour)
        if width >= page_width * 0.06:
            segments.append((round(y + height / 2), x, x + width))
    segments.sort()

    # 同一横线在扫描件中可能被竖线切成多段，先按纵坐标合并再判断跨度。
    line_groups: list[list[tuple[int, int, int]]] = []
    for segment in segments:
        if not line_groups or segment[0] - line_groups[-1][-1][0] > 3:
            line_groups.append([segment])
        else:
            line_groups[-1].append(segment)
    lines: list[tuple[int, int, int]] = []
    for group in line_groups:
        left = min(segment[1] for segment in group)
        right = max(segment[2] for segment in group)
        covered_width = sum(segment[2] - segment[1] for segment in group)
        if right - left >= page_width * 0.35 and covered_width >= page_width * 0.25:
            lines.append((round(sum(segment[0] for segment in group) / len(group)), left, right))
    lines.sort()
    if len(lines) < 5:
        return None

    groups: list[list[tuple[int, int, int]]] = []
    maximum_gap = page_height * 0.08
    for line in lines:
        if not groups or line[0] - groups[-1][-1][0] > maximum_gap:
            groups.append([line])
        else:
            groups[-1].append(line)
    candidates = [
        group
        for group in groups
        if len(group) >= 5
        and group[0][0] < page_height * 0.55
        and group[-1][0] - group[0][0] >= page_height * 0.10
    ]
    if not candidates:
        return None
    table = max(
        candidates,
        key=lambda group: (
            len(group),
            group[-1][0] - group[0][0],
            -group[0][0],
        ),
    )
    x0 = max(0, min(line[1] for line in table))
    x1 = min(page_width, max(line[2] for line in table))
    y0 = max(0, table[0][0])
    y1 = min(page_height, table[-1][0])
    if x1 <= x0 or y1 <= y0:
        return None
    return x0, y0, x1, y1


def offset_ocr_items(
    items: list[OCRItem], offset_x: int, offset_y: int
) -> list[OCRItem]:
    return [
        OCRItem(
            item.item_id,
            item.page_number,
            item.text,
            (
                item.bbox[0] + offset_x,
                item.bbox[1] + offset_y,
                item.bbox[2] + offset_x,
                item.bbox[3] + offset_y,
            ),
            item.confidence,
        )
        for item in items
    ]


def validate_company(text: str) -> bool:
    value = display_value(text)
    return 2 <= len(value) <= 100 and bool(
        re.search(r"[A-Za-z\u4e00-\u9fff]", value)
    )


def validate_credit_code(text: str) -> bool:
    value = compact_value(text).upper()
    return re.fullmatch(r"[0-9A-HJ-NPQRTUWXY]{18}", value) is not None


def validate_person_name(text: str) -> bool:
    value = display_value(text)
    chinese_name = re.fullmatch(r"[\u4e00-\u9fff·]{2,20}", value) is not None
    latin_name = re.fullmatch(r"[A-Za-z][A-Za-z .'-]{1,39}", value) is not None
    return chinese_name or latin_name


def validate_mobile(text: str) -> bool:
    value = compact_value(text)
    if value.startswith("+86"):
        value = value[3:]
    return re.fullmatch(r"1[3-9]\d{9}", value) is not None


def validate_email(text: str) -> bool:
    value = display_value(text)
    return (
        re.fullmatch(
            r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
            r"[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+",
            value,
        )
        is not None
    )


EXTRACT_SPECS = (
    ExtractFieldSpec(
        "company_name",
        "单位名称",
        ("申请单位名称", "申请单位全称", "申请单位（全称）", "单位名称"),
        validate_company,
    ),
    ExtractFieldSpec(
        "credit_code",
        "统一社会信用代码",
        ("统一社会信用代码", "社会信用代码", "统一信用代码"),
        validate_credit_code,
    ),
    ExtractFieldSpec(
        "agent_name",
        "经办人姓名",
        ("业务经办人姓名", "经办人姓名", "经办人（姓名）", "经办人"),
        validate_person_name,
        role="agent",
        blocked_suffixes=(
            "身份证",
            "证件",
            "实名",
            "手机",
            "电话",
            "联系电话",
            "邮箱",
            "电子邮件",
            "签名",
            "签字",
        ),
    ),
    ExtractFieldSpec(
        "agent_mobile",
        "经办人手机号",
        (
            "经办人实名手机号码",
            "经办人实名手机号",
            "经办人手机号码",
            "经办人手机号",
            "经办人联系电话",
            "实名手机号码",
            "实名手机号",
        ),
        validate_mobile,
        role="agent",
    ),
    ExtractFieldSpec(
        "agent_email",
        "经办人邮箱",
        ("经办人电子邮箱", "经办人邮箱", "电子邮箱", "电子邮件", "E-mail", "Email"),
        validate_email,
        role="agent",
    ),
    ExtractFieldSpec(
        "legal_name",
        "法定代表人姓名",
        (
            "个人证书姓名/机构证书法定代表人姓名",
            "个人证书姓名/单位证书法定代表人姓名",
            "机构证书法定代表人姓名",
            "单位证书法定代表人姓名",
            "法定代表人姓名",
            "法人代表姓名",
            "法人姓名",
            "法定代表人",
            "法人代表",
        ),
        validate_person_name,
        role="legal",
        blocked_suffixes=(
            "身份证", "证件", "手机", "电话", "联系电话", "邮箱", "签名", "签字"
        ),
    ),
    ExtractFieldSpec(
        "personal_mobile",
        "个人实名手机号",
        (
            "个人证书本人实名手机号/（办理机构证书填写法人实名手机号）",
            "个人证书本人实名手机号码",
            "个人证书本人实名手机号",
            "本人实名手机号",
            "办理机构证书填写法人实名手机号",
            "法定代表人联系电话",
        ),
        validate_mobile,
    ),
)

SPEC_BY_KEY = {spec.key: spec for spec in EXTRACT_SPECS}
MANUAL_KEYS = ("company_address", "legal_id", "agent_id")
ALL_LABEL_ALIASES = tuple(
    # 个人手机号使用独立单元格定位，不改变其他字段的既有标签边界。
    alias for spec in EXTRACT_SPECS if spec.key != "personal_mobile"
    for alias in spec.aliases
) + (
    "身份证号码",
    "身份证号",
    "证件号码",
)

FIELD_SECTIONS = (
    (
        "单位基本信息",
        (
            ("company_name", "单位名称"),
            ("credit_code", "统一社会信用代码"),
            ("company_address", "单位法定地址"),
        ),
    ),
    (
        "法定代表人信息",
        (
            ("legal_name", "法定代表人姓名"),
            ("legal_id", "法定代表人身份证号码"),
            ("personal_mobile", "个人实名手机号"),
        ),
    ),
    (
        "经办人信息",
        (
            ("agent_name", "经办人姓名"),
            ("agent_id", "经办人身份证号码"),
            ("agent_mobile", "经办人手机号"),
            ("agent_email", "经办人邮箱"),
        ),
    ),
)
FIELD_KEYS = tuple(
    key for _, fields in FIELD_SECTIONS for key, _ in fields
)

REJECTED_NAME_TOKENS = tuple(
    normalized_label(value)
    for value in ("签名", "签字", "盖章", "日期", "机构证书", "个人证书", "证书类型")
)
LEGAL_NAME_REJECTED_TOKENS = tuple(
    normalized_label(value)
    for value in ("身份证", "申请材料", "证明材料", "材料说明", "经办人身份证")
)
REJECTED_NAME_LABELS = frozenset(
    normalized_label(value)
    for value in (
        "联系电话", "手机号", "手机号码", "身份证号", "身份证号码", "证件号码",
        "电子邮箱", "经办人联系电话", "法定代表人联系电话",
        "经办人手机号", "法定代表人手机号",
    )
)


def find_alias_match(text: str, spec: ExtractFieldSpec) -> tuple[int, str] | None:
    normalized_text = normalized_label(text)
    normalized_positions: list[int] = []
    for index, character in enumerate(text):
        for normalized_character in unicodedata.normalize("NFKC", character):
            if normalized_character.isalnum():
                normalized_positions.append(index)

    matches: list[tuple[int, int, str]] = []
    for alias in spec.aliases:
        normalized_alias = normalized_label(alias)
        start = normalized_text.find(normalized_alias)
        if start < 0 or start > 2:
            continue
        normalized_end = start + len(normalized_alias)
        original_end = normalized_positions[normalized_end - 1] + 1
        suffix = normalized_text[normalized_end:]
        if any(
            suffix.startswith(normalized_label(blocked))
            for blocked in spec.blocked_suffixes
        ):
            continue
        matches.append((len(normalized_alias), original_end, alias))
    if not matches:
        return None
    _, original_end, alias = max(matches, key=lambda item: item[0])
    return original_end, alias


def contains_known_label(text: str) -> bool:
    normalized_text = normalized_label(text)
    return any(
        normalized_text.startswith(normalized_label(alias))
        for alias in ALL_LABEL_ALIASES
    )


def box_center_y(item: OCRItem) -> float:
    return (item.bbox[1] + item.bbox[3]) / 2


def box_height(item: OCRItem) -> float:
    return max(item.bbox[3] - item.bbox[1], 1.0)


def role_context_matches(
    label_item: OCRItem,
    matched_alias: str,
    role: str | None,
    items: tuple[OCRItem, ...],
    page_height: int,
) -> bool:
    if role is None:
        return True
    normalized_alias = normalized_label(matched_alias)
    if role == "agent" and "经办人" in normalized_alias:
        return True
    if role == "legal" and (
        "法定代表人" in normalized_alias or "法人" in normalized_alias
    ):
        return True

    role_terms = {
        "agent": ("经办人", "业务经办人"),
        "legal": ("法定代表人", "法人代表"),
    }
    opposite = "legal" if role == "agent" else "agent"

    def nearest_distance(terms: tuple[str, ...]) -> float:
        distances = []
        for item in items:
            text = normalized_label(item.text)
            if not any(text.startswith(normalized_label(term)) for term in terms):
                continue
            vertical_distance = abs(box_center_y(item) - box_center_y(label_item))
            if vertical_distance <= page_height * 0.18:
                distances.append(vertical_distance)
        return min(distances, default=float("inf"))

    own_distance = nearest_distance(role_terms[role])
    opposite_distance = nearest_distance(role_terms[opposite])
    return own_distance < float("inf") and own_distance < opposite_distance


def same_row_right_limit(
    label: OCRItem, items: tuple[OCRItem, ...], page_width: int
) -> float:
    limits = []
    for item in items:
        if item.item_id == label.item_id or not contains_known_label(item.text):
            continue
        if item.bbox[0] <= label.bbox[2]:
            continue
        if abs(box_center_y(item) - box_center_y(label)) <= max(
            box_height(item), box_height(label)
        ):
            limits.append(item.bbox[0])
    return min(limits, default=float(page_width))


def candidate_is_rejected(spec: ExtractFieldSpec, value: str) -> bool:
    if spec.key not in {"agent_name", "legal_name"}:
        return False
    normalized = normalized_label(value)
    # 仅排除明确的字段标签，不缩短真实姓名的长度范围。
    if normalized in REJECTED_NAME_LABELS:
        return True
    tokens = REJECTED_NAME_TOKENS
    if spec.key == "legal_name":
        tokens += LEGAL_NAME_REJECTED_TOKENS
    return any(token in normalized for token in tokens)


def locate_candidate_region(
    spec: ExtractFieldSpec,
    label: OCRItem,
    items: tuple[OCRItem, ...],
    page_width: int,
    page_height: int,
) -> CandidateRegion | None:
    """用当前标签与同行下一个标签限定填写单元格，避免搜索到签名区。"""
    label_height = box_height(label)
    x0 = max(0.0, label.bbox[2] + label_height * 0.05)
    default_right = min(
        page_width * 0.97,
        label.bbox[2] + max(page_width * 0.32, label_height * 18),
    )
    x1 = min(default_right, same_row_right_limit(label, items, page_width))
    if spec.key == "legal_name":
        # 法定代表人姓名只取标签同一表格行，避免跨到上方证书类型区域。
        y0 = max(0.0, label.bbox[1] - label_height * 0.15)
        y1 = min(float(page_height), label.bbox[3] + label_height * 0.35)
    else:
        y0 = max(0.0, label.bbox[1] - label_height * 0.70)
        y1 = min(float(page_height), label.bbox[3] + label_height * 0.70)
    if x1 - x0 < label_height * 1.5 or y1 <= y0:
        return None
    return CandidateRegion(label.page_number, (x0, y0, x1, y1))


def item_center_in_region(item: OCRItem, region: CandidateRegion) -> bool:
    x0, y0, x1, y1 = region.bbox
    center_x = (item.bbox[0] + item.bbox[2]) / 2
    center_y = box_center_y(item)
    return x0 <= center_x <= x1 and y0 <= center_y <= y1


def merge_line_positions(positions: list[int], tolerance: int = 3) -> list[int]:
    merged: list[list[int]] = []
    for position in sorted(positions):
        if not merged or position - merged[-1][-1] > tolerance:
            merged.append([position])
        else:
            merged[-1].append(position)
    return [round(sum(group) / len(group)) for group in merged]


def detect_value_cell_from_table(
    image: np.ndarray,
    label: OCRItem,
    target_x: float | None = None,
) -> CandidateRegion | None:
    """用标签所在行的横线、竖线确定右侧完整填写单元格。"""
    page_height, page_width = image.shape[:2]
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )
    label_height = box_height(label)
    center_y = box_center_y(label)

    horizontal_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT, (max(30, page_width // 16), 1)
    )
    horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)
    contours, _ = cv2.findContours(
        horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    horizontal_positions: list[int] = []
    for contour in contours:
        x, y, width, height = cv2.boundingRect(contour)
        horizontal_target = target_x if target_x is not None else label.bbox[2]
        if width >= page_width * 0.06 and x <= horizontal_target <= x + width:
            horizontal_positions.append(round(y + height / 2))
    horizontal_positions = merge_line_positions(horizontal_positions)
    top_candidates = [
        y for y in horizontal_positions if y <= label.bbox[1] + label_height * 0.30
    ]
    bottom_candidates = [
        y for y in horizontal_positions if y >= label.bbox[3] - label_height * 0.30
    ]
    if not top_candidates or not bottom_candidates:
        return None
    top = max(top_candidates)
    bottom = min(bottom_candidates)
    row_height = bottom - top
    if not (label_height * 1.10 <= row_height <= page_height * 0.10):
        return None

    row_binary = binary[max(0, top - 2) : min(page_height, bottom + 3), :]
    vertical_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT, (1, max(8, round(row_height * 0.58)))
    )
    vertical = cv2.morphologyEx(row_binary, cv2.MORPH_OPEN, vertical_kernel)
    contours, _ = cv2.findContours(
        vertical, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    vertical_positions = merge_line_positions(
        [
            round(x + width / 2)
            for contour in contours
            for x, _, width, height in (cv2.boundingRect(contour),)
            if height >= row_height * 0.58
        ]
    )
    if target_x is None:
        boundaries = [
            x
            for x in vertical_positions
            if x >= label.bbox[2] - label_height * 0.60
        ]
        if len(boundaries) < 2:
            return None
        left, right = boundaries[:2]
    else:
        left_candidates = [x for x in vertical_positions if x < target_x]
        right_candidates = [x for x in vertical_positions if x > target_x]
        if not left_candidates or not right_candidates:
            return None
        left = max(left_candidates)
        right = min(right_candidates)
    maximum_cell_width = page_width * (0.80 if target_x is None else 0.30)
    if not (label_height * 2 <= right - left <= maximum_cell_width):
        return None
    padding = max(2.0, label_height * 0.08)
    return CandidateRegion(
        label.page_number,
        (left + padding, top + padding, right - padding, bottom - padding),
    )


def refine_legal_name_result(
    pages: tuple[PageOCR, ...],
    page_images: dict[int, np.ndarray],
    current: ExtractionResult,
) -> ExtractionResult:
    """按经办人姓名同列关系重建法人姓名区域，并先复用整页 OCR。"""
    if current.recognized:
        return current
    legal_spec = SPEC_BY_KEY["legal_name"]
    agent_spec = SPEC_BY_KEY["agent_name"]
    refined_regions: list[CandidateRegion] = []
    all_candidates: list[Candidate] = []
    candidate_blocks = 0

    for page in pages:
        legal_labels = [
            item
            for item in page.items
            if find_alias_match(item.text, legal_spec) is not None
        ]
        for legal_label in legal_labels:
            agent_labels = [
                item
                for item in page.items
                if box_center_y(item) > box_center_y(legal_label)
                and box_center_y(item) - box_center_y(legal_label) <= page.height * 0.15
                and find_alias_match(item.text, agent_spec) is not None
            ]
            if not agent_labels:
                continue
            agent_label = min(
                agent_labels,
                key=lambda item: box_center_y(item) - box_center_y(legal_label),
            )
            agent_region = locate_candidate_region(
                agent_spec, agent_label, page.items, page.width, page.height
            )
            if agent_region is None:
                continue
            agent_values = [
                item
                for item in page.items
                if item.item_id != agent_label.item_id
                and item_center_in_region(item, agent_region)
                and validate_person_name(display_value(item.text))
                and not candidate_is_rejected(agent_spec, item.text)
                and not contains_known_label(item.text)
            ]
            if agent_values:
                agent_value = min(
                    agent_values,
                    key=lambda item: (
                        abs(box_center_y(item) - box_center_y(agent_label)),
                        item.bbox[0],
                    ),
                )
                target_x = (agent_value.bbox[0] + agent_value.bbox[2]) / 2
            else:
                target_x = (agent_region.bbox[0] + agent_region.bbox[2]) / 2

            image = page_images.get(legal_label.page_number)
            region = (
                detect_value_cell_from_table(image, legal_label, target_x)
                if image is not None
                else None
            )
            if region is None:
                # 无法确认表格单元格边界时留空，禁止扩大到其他行取值。
                continue
            refined_regions.append(region)

            usable = [
                item
                for item in page.items
                if item.item_id not in {legal_label.item_id, agent_label.item_id}
                and item_center_in_region(item, region)
                and not contains_known_label(item.text)
                and not candidate_is_rejected(legal_spec, item.text)
                and display_value(item.text)
            ]
            usable.sort(key=lambda item: item.bbox[0])
            candidate_blocks += len(usable)
            if usable:
                combined = "".join(display_value(item.text) for item in usable)
                all_candidates.append(
                    Candidate(
                        combined,
                        min(item.confidence for item in usable),
                        108.0,
                        validate_person_name(combined)
                        and not candidate_is_rejected(legal_spec, combined),
                    )
                )

    unique_regions = tuple(dict.fromkeys(refined_regions))
    selected, format_valid, reason = rank_candidates(all_candidates)
    diagnostic = FieldDiagnostic(
        label_found=current.diagnostic.label_found or bool(refined_regions),
        region_located=bool(refined_regions),
        candidate_blocks=candidate_blocks,
        format_valid=format_valid,
        reason=("已复用整页 OCR 中的法人姓名，待人工核对" if selected else reason),
    )
    if selected is None:
        return ExtractionResult(diagnostic=diagnostic, regions=unique_regions)
    return ExtractionResult(
        value=selected.value,
        recognized=True,
        diagnostic=diagnostic,
        regions=unique_regions,
    )


def refine_centered_table_field(
    spec: ExtractFieldSpec,
    pages: tuple[PageOCR, ...],
    page_images: dict[int, np.ndarray],
    current: ExtractionResult,
) -> ExtractionResult:
    """在完整右侧单元格内重新查找居中填写的字段值。"""
    if current.recognized:
        return current
    regions: list[CandidateRegion] = []
    candidates: list[Candidate] = []
    candidate_blocks = 0
    rejected_tokens = tuple(
        normalized_label(value)
        for value in ("证书类型", "机构证书", "个人证书", "身份证号", "申请材料说明")
    )

    for page in pages:
        image = page_images.get(page.items[0].page_number) if page.items else None
        if image is None:
            continue
        for label in page.items:
            match = find_alias_match(label.text, spec)
            if match is None:
                continue
            _, matched_alias = match
            if not role_context_matches(
                label, matched_alias, spec.role, page.items, page.height
            ):
                continue
            region = detect_value_cell_from_table(image, label)
            if region is None:
                continue
            regions.append(region)
            usable = [
                item
                for item in page.items
                if item.item_id != label.item_id
                and item_center_in_region(item, region)
                and not contains_known_label(item.text)
                and display_value(item.text)
            ]
            usable.sort(key=lambda item: (box_center_y(item), item.bbox[0]))
            candidate_blocks += len(usable)
            if not usable:
                continue
            value = "".join(display_value(item.text) for item in usable)
            normalized_value = normalized_label(value)
            safe_value = not any(token in normalized_value for token in rejected_tokens)
            candidates.append(
                Candidate(
                    value,
                    min(item.confidence for item in usable),
                    108.0,
                    safe_value and spec.validator(value),
                )
            )

    if not regions:
        return current
    unique_regions = tuple(dict.fromkeys(regions))
    selected, format_valid, reason = rank_candidates(candidates)
    diagnostic = FieldDiagnostic(
        label_found=True,
        region_located=True,
        candidate_blocks=candidate_blocks,
        format_valid=format_valid,
        reason=("已在完整表格单元格中提取，待人工核对" if selected else reason),
    )
    if selected is None:
        return ExtractionResult(diagnostic=diagnostic, regions=unique_regions)
    return ExtractionResult(
        value=selected.value,
        recognized=True,
        diagnostic=diagnostic,
        regions=unique_regions,
    )


def candidates_for_label(
    spec: ExtractFieldSpec,
    label_item: OCRItem,
    label_end: int,
    items: tuple[OCRItem, ...],
    page_width: int,
    page_height: int,
) -> tuple[list[Candidate], int, CandidateRegion | None]:
    candidates: list[Candidate] = []
    block_ids: set[int] = set()
    inline_value = display_value(label_item.text[label_end:])
    if inline_value:
        format_valid = spec.validator(inline_value) and not candidate_is_rejected(
            spec, inline_value
        )
        candidates.append(
            Candidate(inline_value, label_item.confidence, 110.0, format_valid)
        )
        block_ids.add(label_item.item_id)

    region = locate_candidate_region(spec, label_item, items, page_width, page_height)
    if region is None:
        return candidates, len(block_ids), None

    related: list[OCRItem] = []
    for item in items:
        if item.item_id == label_item.item_id or contains_known_label(item.text):
            continue
        if not item_center_in_region(item, region):
            continue
        block_ids.add(item.item_id)
        related.append(item)

    related.sort(key=lambda item: item.bbox[0])
    structured = spec.key in {
        "credit_code",
        "agent_mobile",
        "agent_email",
    }
    if structured and related:
        combined = "".join(display_value(item.text) for item in related)
        format_valid = spec.validator(combined)
        confidence = min(
            label_item.confidence, *(item.confidence for item in related)
        )
        distance = related[0].bbox[0] - label_item.bbox[2]
        score = 100.0 - max(distance, 0.0) / max(page_width, 1) * 20
        candidates.append(Candidate(combined, confidence, score, format_valid))
    elif not structured:
        for item in related:
            value = display_value(item.text)
            if not value:
                continue
            format_valid = spec.validator(value) and not candidate_is_rejected(
                spec, value
            )
            distance = item.bbox[0] - label_item.bbox[2]
            center_distance = abs(box_center_y(item) - box_center_y(label_item))
            score = 96.0 - max(distance, 0.0) / max(page_width, 1) * 20
            score -= center_distance / max(box_height(label_item), 1) * 5
            candidates.append(
                Candidate(
                    value,
                    min(label_item.confidence, item.confidence),
                    score,
                    format_valid,
                )
            )
    return candidates, len(block_ids), region


def rank_candidates(
    candidates: list[Candidate],
) -> tuple[Candidate | None, bool, str]:
    unique: dict[str, Candidate] = {}
    for candidate in candidates:
        canonical = compact_value(candidate.value).lower()
        if not canonical:
            continue
        existing = unique.get(canonical)
        if existing is None or (candidate.score, candidate.confidence) > (
            existing.score,
            existing.confidence,
        ):
            unique[canonical] = candidate

    format_valid = any(candidate.format_valid for candidate in unique.values())
    ranked = sorted(
        (candidate for candidate in unique.values() if candidate.format_valid),
        key=lambda item: (item.score, item.confidence),
        reverse=True,
    )
    if not ranked:
        reason = (
            "候选内容未通过格式检查" if unique else "候选区域未识别到文字"
        )
        return None, format_valid, reason
    if ranked[0].confidence < 0.75:
        return None, True, "候选 OCR 置信度不足"
    if len(ranked) > 1 and ranked[0].score - ranked[1].score < 8.0:
        return None, True, "存在多个相近候选，无法确定"
    return ranked[0], True, "已提取，待人工核对"


def agent_mobile_label_items(page: PageOCR) -> tuple[OCRItem, ...]:
    """仅为经办人手机号组合相邻标签块，不组合号码或其他字段。"""
    spec = SPEC_BY_KEY["agent_mobile"]
    aliases = tuple(
        normalized_label(alias) for alias in spec.aliases if alias.startswith("经办人")
    )
    combined: list[OCRItem] = []
    for anchor in page.items:
        prefix = normalized_label(anchor.text)
        if (
            not prefix.startswith("经办人")
            or find_alias_match(anchor.text, spec) is not None
            or not any(alias.startswith(prefix) for alias in aliases)
        ):
            continue
        pending = [(anchor,)]
        while pending:
            parts = pending.pop()
            last = parts[-1]
            text = normalized_label("".join(item.text for item in parts))
            if text in aliases:
                combined.append(OCRItem(
                    anchor.item_id, anchor.page_number,
                    "".join(item.text for item in parts),
                    (
                        min(item.bbox[0] for item in parts),
                        min(item.bbox[1] for item in parts),
                        max(item.bbox[2] for item in parts),
                        max(item.bbox[3] for item in parts),
                    ),
                    min(item.confidence for item in parts),
                ))
                continue
            if len(parts) >= 3:
                continue
            for item in page.items:
                if item in parts:
                    continue
                joined = text + normalized_label(item.text)
                if joined == text or not any(alias.startswith(joined) for alias in aliases):
                    continue
                height = max(box_height(last), box_height(item))
                same_line = (
                    0 <= item.bbox[0] - last.bbox[2] <= height
                    and abs(box_center_y(item) - box_center_y(last)) <= height * 0.45
                )
                next_line = (
                    -height * 0.15 <= item.bbox[1] - last.bbox[3] <= height * 0.60
                    and box_center_y(item) > box_center_y(last)
                    and abs(item.bbox[0] - anchor.bbox[0]) <= height
                )
                if same_line or next_line:
                    pending.append(parts + (item,))
    return page.items + tuple(combined)


def extract_field(spec: ExtractFieldSpec, pages: tuple[PageOCR, ...]) -> ExtractionResult:
    candidates: list[Candidate] = []
    regions: list[CandidateRegion] = []
    label_found = False
    context_matched = False
    candidate_blocks = 0
    for page in pages:
        label_items = agent_mobile_label_items(page) if spec.key == "agent_mobile" else page.items
        for item in label_items:
            match = find_alias_match(item.text, spec)
            if match is None:
                continue
            label_found = True
            label_end, matched_alias = match
            if not role_context_matches(
                item,
                matched_alias,
                spec.role,
                page.items,
                page.height,
            ):
                continue
            context_matched = True
            found, block_count, region = candidates_for_label(
                spec,
                item,
                label_end,
                page.items,
                page.width,
                page.height,
            )
            candidates.extend(found)
            candidate_blocks += block_count
            if region is not None:
                regions.append(region)

    region_located = bool(regions) or any(
        candidate.score >= 110 for candidate in candidates
    )
    if not label_found:
        return ExtractionResult(
            diagnostic=FieldDiagnostic(reason="未找到对应字段标签")
        )
    if not context_matched:
        return ExtractionResult(
            diagnostic=FieldDiagnostic(
                label_found=True,
                reason="标签所在角色区域不明确",
            )
        )
    if not region_located:
        return ExtractionResult(
            diagnostic=FieldDiagnostic(
                label_found=True,
                reason="未可靠定位候选值区域",
            )
        )

    selected, format_valid, reason = rank_candidates(candidates)
    diagnostic = FieldDiagnostic(
        label_found=True,
        region_located=True,
        candidate_blocks=candidate_blocks,
        format_valid=format_valid,
        reason=reason,
    )
    unique_regions = tuple(dict.fromkeys(regions))
    if selected is None:
        return ExtractionResult(diagnostic=diagnostic, regions=unique_regions)
    return ExtractionResult(
        value=selected.value,
        recognized=True,
        diagnostic=diagnostic,
        regions=unique_regions,
    )


def local_candidates(
    spec: ExtractFieldSpec, items: list[OCRItem]
) -> list[Candidate]:
    usable = [
        item
        for item in items
        if not contains_known_label(item.text)
        and not candidate_is_rejected(spec, item.text)
        and display_value(item.text)
    ]
    usable.sort(key=lambda item: (box_center_y(item), item.bbox[0]))
    if not usable:
        return []

    if spec.key == "legal_name":
        usable.sort(key=lambda item: item.bbox[0])
        value = "".join(display_value(item.text) for item in usable)
        return [
            Candidate(
                value,
                min(item.confidence for item in usable),
                105.0,
                spec.validator(value) and not candidate_is_rejected(spec, value),
            )
        ]

    if spec.key in {
        "credit_code",
        "agent_mobile",
        "agent_email",
        "personal_mobile",
    }:
        value = "".join(display_value(item.text) for item in usable)
        return [
            Candidate(
                value,
                min(item.confidence for item in usable),
                105.0,
                spec.validator(value),
            )
        ]

    return [
        Candidate(
            value,
            item.confidence,
            105.0 - index,
            spec.validator(value) and not candidate_is_rejected(spec, value),
        )
        for index, item in enumerate(usable)
        if (value := display_value(item.text))
    ]


def extract_application_fields(
    pages: tuple[PageOCR, ...]
) -> dict[str, ExtractionResult]:
    return {
        spec.key: (
            ExtractionResult() if spec.key == "personal_mobile"
            else extract_field(spec, pages)
        )
        for spec in EXTRACT_SPECS
    }


def extract_personal_mobile(
    pages: tuple[PageOCR, ...], page_images: dict[int, np.ndarray]
) -> ExtractionResult:
    """独立确认标签与物理单元格，不扩大到其他行搜索号码。"""
    regions: list[CandidateRegion] = []
    candidates: list[Candidate] = []
    label_found = False
    block_count = 0

    def label_matches(text: str) -> bool:
        # 新版联系方式标签使用明确匹配，不放宽到通用“联系电话”。
        if normalized_label("法定代表人联系电话") in text:
            return True
        # 允许一个漏字，仍须包含完整角色、实名及手机字段特征。
        for expected in (
            "个人证书本人实名手机号",
            "办理机构证书填写法人实名手机号",
        ):
            if expected in text or any(
                expected[:index] + expected[index + 1:] in text
                for index in range(len(expected))
            ):
                return True
        return False

    for page in pages:
        if not page.items:
            continue
        image = page_images.get(page.items[0].page_number)
        binary = None
        contours = []
        if image is not None:
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            _, binary = cv2.threshold(
                gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
            )
            horizontal = cv2.morphologyEx(
                binary, cv2.MORPH_OPEN,
                cv2.getStructuringElement(
                    cv2.MORPH_RECT, (max(30, page.width // 16), 1)
                ),
            )
            contours, _ = cv2.findContours(
                horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
            )
        lines = [
            (round(y + height / 2), x, x + width)
            for contour in contours
            for x, y, width, height in (cv2.boundingRect(contour),)
            if width >= page.width * 0.06
        ]
        agent_labels = [
            item for item in agent_mobile_label_items(page)
            if "经办人" in normalized_label(item.text)
            and find_alias_match(item.text, SPEC_BY_KEY["agent_mobile"]) is not None
        ]
        email_labels = [
            item for item in page.items
            if find_alias_match(item.text, SPEC_BY_KEY["agent_email"]) is not None
        ]
        for anchor in page.items:
            text = normalized_label(anchor.text)
            if not any(
                token in text
                for token in (
                    "个人证书", "本人实名", "办理机构证书", "法人实名", "法定代表人"
                )
            ):
                continue
            label_height = box_height(anchor)
            horizontal_positions = merge_line_positions([
                y for y, left, right in lines
                if left <= anchor.bbox[2] <= right
            ])
            tops = [
                y for y in horizontal_positions
                if y <= anchor.bbox[1] + label_height * 0.30
            ]
            bottoms = [
                y for y in horizontal_positions
                if y >= anchor.bbox[3] - label_height * 0.30
            ]
            top = max(tops) if tops else None
            bottom = min(bottoms) if bottoms else None
            row_reliable = (
                top is not None and bottom is not None
                and label_height * 1.10 <= bottom - top <= page.height * 0.10
            )
            aligned_labels = [
                item for item in agent_labels
                if 0 < box_center_y(anchor) - box_center_y(item) <= label_height * 10
                and abs(item.bbox[0] - anchor.bbox[0]) <= label_height * 8
            ]
            label_left = min(
                [anchor.bbox[0]] + [item.bbox[0] for item in aligned_labels]
            )
            # 标签先独立确认，缺失表格线时继续检查相邻字段。
            label_parts = [
                item for item in page.items
                if item.bbox[0] >= label_left - label_height
                and item.bbox[2] <= label_left + max(
                    anchor.bbox[2] - anchor.bbox[0], label_height * 16
                )
                and (
                    top <= box_center_y(item) <= bottom if row_reliable
                    else abs(box_center_y(item) - box_center_y(anchor)) <= label_height * 3.2
                )
                and "经办人" not in normalized_label(item.text)
                and not any(character.isdigit() or character == "@" for character in item.text)
            ]
            # 同行文字块先从左到右拼接，再组合上下各行。
            label_rows: list[list[OCRItem]] = []
            for item in sorted(label_parts, key=box_center_y):
                if (
                    not label_rows
                    or abs(box_center_y(item) - box_center_y(label_rows[-1][0]))
                    > label_height * 0.5
                ):
                    label_rows.append([item])
                else:
                    label_rows[-1].append(item)
            label_parts = [
                item for row in label_rows
                for item in sorted(row, key=lambda part: part.bbox[0])
            ]
            label_text = normalized_label("".join(item.text for item in label_parts))
            if not label_matches(label_text):
                continue
            label_found = True
            label_top = min(item.bbox[1] for item in label_parts)
            label_bottom = max(item.bbox[3] for item in label_parts)
            label_right = max(item.bbox[2] for item in label_parts)
            above = [item for item in aligned_labels if item.bbox[3] < label_top]
            agent = max(above, key=box_center_y) if above else None
            emails = [
                item for item in email_labels
                if item.bbox[2] < label_left
                and label_top - label_height <= box_center_y(item) <= label_bottom + label_height
            ]
            email_same_row = bool(emails)
            following_labels = [
                item for item in page.items
                if label_bottom < item.bbox[1] <= label_bottom + label_height * 6
                and (
                    contains_known_label(item.text)
                    or any(term in normalized_label(item.text) for term in (
                        "申请提供", "证明材料", "用户须知", "声明", "签名", "签字"
                    ))
                )
            ]
            # 横线缺失时，以上一角色行、同排邮箱及下一标签约束行范围。
            top_line = (
                top is not None and 0 <= label_top - top <= label_height * 2
                and (agent is None or top > box_center_y(agent))
            )
            bottom_line = (
                bottom is not None and 0 <= bottom - label_bottom <= label_height * 2
            )
            row_top = min([label_top] + [item.bbox[1] for item in emails])
            row_bottom = max([label_bottom] + [item.bbox[3] for item in emails])
            if not top_line:
                top = max(0, round(row_top - label_height))
                if agent is not None:
                    top = max(
                        top, round((agent.bbox[3] + row_top) / 2),
                        round((agent.bbox[3] + label_top + label_height * 0.25) / 2),
                        round(agent.bbox[3] + box_height(agent) * 0.60),
                    )
            if not bottom_line:
                bottom = min(page.height, round(row_bottom + label_height))
                if following_labels:
                    bottom = min(bottom, round(
                        (row_bottom + min(item.bbox[1] for item in following_labels)) / 2
                    ))
            if bottom <= top:
                continue
            if any(top < box_center_y(item) < bottom for item in agent_labels):
                continue
            row_height = bottom - top
            contours = []
            if binary is not None:
                vertical = cv2.morphologyEx(
                    binary[max(0, top - 2):min(page.height, bottom + 3), :],
                    cv2.MORPH_OPEN,
                    cv2.getStructuringElement(
                        cv2.MORPH_RECT, (1, max(8, round(row_height * 0.45)))
                    ),
                )
                contours, _ = cv2.findContours(
                    vertical, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
                )
            boundaries = merge_line_positions([
                round(x + width / 2)
                for contour in contours
                for x, _, width, height in (cv2.boundingRect(contour),)
                if height >= row_height * 0.45
                and x + width / 2 >= label_right - label_height * 0.60
            ])
            # 相邻经办人单元格只提供列边界，绝不提供个人手机号的值。
            agent_cell = (
                detect_value_cell_from_table(image, agent)
                if image is not None and agent is not None else None
            )
            left_line = bool(boundaries and boundaries[0] - label_right <= label_height * 2)
            left = boundaries[0] if left_line else label_right
            if not left_line and agent_cell is not None:
                if abs(agent_cell.bbox[0] - label_right) <= label_height * 2:
                    left = agent_cell.bbox[0]
            right_edges = [
                x1 for y, x0, x1 in lines
                if top - label_height <= y <= bottom + label_height
                and x0 <= left and x1 > left + label_height * 2
            ]
            if left_line and len(boundaries) >= 2:
                right = boundaries[1]
            elif agent_cell is not None and agent_cell.bbox[2] > left + label_height * 2:
                right = agent_cell.bbox[2]
            elif right_edges:
                right = max(right_edges)
            else:
                right = page.width
            next_columns = [
                item.bbox[0] for item in page.items
                if item.bbox[0] > left and top < box_center_y(item) < bottom
                and contains_known_label(item.text)
            ]
            if next_columns:
                right = min(right, min(next_columns))
            structure_evidence = top_line and bottom_line and left_line
            context_evidence = agent is not None and email_same_row
            mixed_evidence = (
                (agent is not None or email_same_row)
                and (left_line or agent_cell is not None)
                and (top_line or bottom_line)
            )
            if not (structure_evidence or context_evidence or mixed_evidence):
                continue
            if right - left < label_height * 2:
                continue
            # 使用实际边界或多项上下文推断的边界，统一合并近重复区域。
            region = CandidateRegion(anchor.page_number, (left, top, right, bottom))
            tolerance = max(3.0, label_height * 0.15)
            if any(
                existing.page_number == region.page_number
                and all(abs(a - b) <= tolerance for a, b in zip(existing.bbox, region.bbox))
                for existing in regions
            ):
                continue
            regions.append(region)
            usable: list[OCRItem] = []
            for item in page.items:
                if (
                    not item_center_in_region(item, region)
                    or contains_known_label(item.text) or not display_value(item.text)
                ):
                    continue
                if agent is not None and (
                    box_center_y(item) <= agent.bbox[3]
                    or abs(box_center_y(item) - box_center_y(agent))
                    <= box_height(agent) * 0.60
                ):
                    continue
                if agent is not None and not top_line:
                    center_y = box_center_y(item)
                    personal_distance = max(label_top - center_y, center_y - label_bottom, 0)
                    agent_distance = max(agent.bbox[1] - center_y, center_y - agent.bbox[3], 0)
                    if agent_distance <= personal_distance + label_height * 0.25:
                        continue
                x0, y0, x1, y1 = item.bbox
                overlap = max(0.0, min(x1, right) - max(x0, left)) * max(
                    0.0, min(y1, bottom) - max(y0, top)
                )
                area = max((x1 - x0) * (y1 - y0), 1.0)
                if overlap / area >= 0.65:
                    usable.append(item)
            usable.sort(key=lambda item: item.bbox[0])
            block_count += len(usable)
            if usable:
                value = "".join(display_value(item.text) for item in usable)
                candidates.append(Candidate(
                    value, min(item.confidence for item in usable),
                    108.0, validate_mobile(value),
                ))
    selected, format_valid, reason = rank_candidates(candidates)
    if len(regions) > 1:
        selected = None
        reason = "存在多个不同填写单元格，无法确认个人手机号归属"
    elif not label_found:
        reason = "未找到对应字段标签"
    elif not regions:
        reason = "标签已确认，但相邻行列及表格线证据不足，需人工补充"
    diagnostic = FieldDiagnostic(
        label_found=label_found,
        region_located=bool(regions),
        candidate_blocks=block_count,
        format_valid=format_valid,
        reason=reason,
    )
    return ExtractionResult(
        value=selected.value if selected else "",
        recognized=selected is not None,
        diagnostic=diagnostic,
        regions=tuple(regions),
    )


def _id_grid_state(
    cells: list[list[tuple[str, float]]], uncertain: set[int],
    extra_low: set[int] | None = None,
) -> dict[str, Any]:
    """共用格位统计及验证；含字符的内存状态禁止直接输出。"""
    conflicts = [index for index, candidates in enumerate(cells) if len(candidates) > 1]
    low = {
        index for index, candidates in enumerate(cells)
        if any(not np.isfinite(score) or score < 0.75 for _, score in candidates)
    } | (extra_low or set())
    missing = [index for index, candidates in enumerate(cells) if not candidates and index not in uncertain]
    determined = [index for index, candidates in enumerate(cells)
                  if len(candidates) == 1 and index not in uncertain and index not in low]
    candidate = "".join(cells[index][0][0] for index in determined)
    complete = len(determined) == 18
    format_ok = complete and re.fullmatch(r"[1-9][0-9]{16}[0-9X]", candidate) is not None
    birth_result = checksum_result = "未执行（候选不完整或格式不通过）"
    if format_ok:
        try:
            birth = date(int(candidate[6:10]), int(candidate[10:12]), int(candidate[12:14]))
            birth_result = "通过" if birth <= date.today() else "不通过"
        except ValueError:
            birth_result = "不通过"
        weights = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
        remainder = sum(int(c) * w for c, w in zip(candidate[:17], weights)) % 11
        checksum_result = "通过" if "10X98765432"[remainder] == candidate[-1] else "不通过"
    return {
        "determined": len(determined), "missing": len(missing), "conflicts": len(conflicts),
        "complete": complete, "format_ok": format_ok, "birth": birth_result,
        "checksum": checksum_result,
        "valid": format_ok and birth_result == "通过" and checksum_result == "通过",
        "candidate": candidate if complete else "", "length": len(candidate),
        "cells": tuple(tuple(values) for values in cells), "uncertain": frozenset(uncertain),
        "determined_indices": tuple(determined), "missing_indices": tuple(missing),
        "conflict_indices": tuple(conflicts), "low_indices": tuple(sorted(low)),
    }


def diagnose_id_grid_assignment(
    items: tuple[OCRItem, ...], label_id: int, region: CandidateRegion,
    boundaries: list[int], grid_evidence: bool, original_length: int, prefix: str,
) -> dict[str, Any] | None:
    """按实测格位分析字符，校验结果仅在诊断内存中使用，不回填字段。"""
    started = time.perf_counter()
    print(prefix + f"格位优化前字符数={original_length}")
    if not grid_evidence:
        print(prefix + "格位优化跳过：未可靠确认19条边界，不进行平均切格")
        print(prefix + f"格位处理额外耗时={time.perf_counter() - started:.3f}秒")
        return
    left, top, right, bottom = region.bbox
    cells: list[list[tuple[str, float]]] = [[] for _ in range(18)]
    uncertain: set[int] = set()
    ambiguous_blocks = 0
    multi_blocks = 0
    outside_row_blocks = 0
    for item in items:
        if item.item_id == label_id or item.page_number != region.page_number:
            continue
        x0, y0, x1, y1 = item.bbox
        width, height = x1 - x0, y1 - y0
        if width <= 0 or height <= 0:
            continue
        vertical_overlap = max(0.0, min(y1, bottom) - max(y0, top)) / height
        overlaps = [
            max(0.0, min(x1, b) - max(x0, a))
            for a, b in zip(boundaries, boundaries[1:])
        ]
        touched = [index for index, overlap in enumerate(overlaps) if overlap > 0]
        if not touched or vertical_overlap == 0:
            continue
        # 邻行文字框的小幅越线不应污染本角色；不把它移入当前行。
        if not top <= box_center_y(item) <= bottom:
            outside_row_blocks += 1
            continue
        # 当前行的边缘块保留疑点，不直接当作缺失字符。
        if vertical_overlap < 0.50:
            uncertain.update(touched)
            ambiguous_blocks += 1
            continue
        text = re.sub(r"\s+", "", unicodedata.normalize("NFKC", item.text)).upper()
        if not text:
            continue
        if re.fullmatch(r"[0-9X]+", text) is None:
            uncertain.update(touched)
            ambiguous_blocks += 1
            continue
        assignments: list[tuple[int, str]] = []
        if len(text) == 1:
            ranked = sorted(touched, key=lambda index: overlaps[index], reverse=True)
            best = ranked[0]
            runner = overlaps[ranked[1]] / width if len(ranked) > 1 else 0.0
            share = overlaps[best] / width
            center = (x0 + x1) / 2
            in_cell = boundaries[best] <= center <= boundaries[best + 1]
            near_line = min(abs(center - edge) for edge in boundaries) <= (
                boundaries[best + 1] - boundaries[best]
            ) * 0.08
            # 横跨格线时须有明显的多数重叠，不能仅凭框中心落点选择。
            if in_cell and share >= (0.75 if near_line else 0.60) and share - runner >= 0.30:
                assignments = [(best, text)]
        else:
            spans = [
                index for index in touched
                if overlaps[index] / (boundaries[index + 1] - boundaries[index]) >= 0.30
            ]
            # 依靠实际格线对应字符顺序，不构造等宽字符框。
            if (
                len(spans) == len(text)
                and spans == list(range(spans[0], spans[-1] + 1))
                and all(x0 <= (boundaries[i] + boundaries[i + 1]) / 2 <= x1 for i in spans)
                and x0 >= boundaries[spans[0]] - (boundaries[spans[0] + 1] - boundaries[spans[0]]) * 0.15
                and x1 <= boundaries[spans[-1] + 1] + (boundaries[spans[-1] + 1] - boundaries[spans[-1]]) * 0.15
                and all(overlaps[i] / width <= 0.10 for i in touched if i not in spans)
            ):
                assignments = list(zip(spans, text))
                multi_blocks += 1
        if not assignments:
            uncertain.update(touched)
            ambiguous_blocks += 1
            continue
        for index, character in assignments:
            if character == "X" and index != 17:
                uncertain.add(index)
            else:
                cells[index].append((character, item.confidence))

    # 不根据邻格是否为空来挪动字符；重复/相邻块的竞争在格位内统一保留。
    state = _id_grid_state(cells, uncertain)
    numbers = lambda indices: [index + 1 for index in sorted(indices)]
    print(prefix + f"优化后确定格子数={state['determined']}；缺失格子数={state['missing']}；缺失序号={numbers(state['missing_indices'])}")
    print(prefix + f"冲突序号={numbers(state['conflict_indices'])}；低置信度序号={numbers(state['low_indices'])}；归属/字符不确定序号={numbers(uncertain)}")
    print(prefix + f"多字符块格位对应数={multi_blocks}；不确定文字块数={ambiguous_blocks}；排除邻行越线块数={outside_row_blocks}")
    print(prefix + f"最终候选长度={state['length']}；完整18位={state['complete']}；格式通过={state['format_ok']}；出生日期={state['birth']}；校验码={state['checksum']}")
    print(prefix + "缺失格指无可用OCR字符证据，不证明原表未填；不完整字符不作为号码使用，仍需人工核对")
    print(prefix + f"格位处理额外耗时={time.perf_counter() - started:.3f}秒；新增OCR调用=0")
    return state


def diagnose_id_row_ocr(
    image: np.ndarray, ocr: Any, region: CandidateRegion, boundaries: list[int],
    baseline: dict[str, Any], prefix: str, call_budget: int,
) -> int:
    """整行补充实验：每角色最多3次，不合并版本，不向业务字段返回号码。"""
    started = time.perf_counter()
    calls = 0
    if baseline["valid"]:
        print(prefix + "局部OCR跳过：已有结果完整且全部校验通过；调用次数=0")
        return calls
    x0, y0, x1, y1 = region.bbox
    left, top = max(0, int(np.floor(x0))), max(0, int(np.floor(y0)))
    right = min(image.shape[1], int(np.ceil(x1)))
    bottom = min(image.shape[0], int(np.ceil(y1)))
    crop = image[top:bottom, left:right].copy()
    if crop.size == 0 or call_budget <= 0:
        print(prefix + "局部OCR跳过：裁剪为空或本任务调用额度已用完")
        return calls
    complete_versions: set[str] = set()
    if baseline["complete"]:
        complete_versions.add(baseline["candidate"])
    obtained_valid = False
    succeeded = 0
    for variant in range(min(3, call_budget)):
        if variant == 0:
            name, prepared = "原始整行裁剪", crop.copy()
        elif variant == 1:
            name = "整行2倍放大"
            prepared = cv2.resize(crop, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        else:
            # 只淡化上下外横线的无邻接笔画部分；保留内竖线和数字接线处。
            gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
            _, ink = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
            line_mask = cv2.morphologyEx(
                ink, cv2.MORPH_OPEN,
                cv2.getStructuringElement(cv2.MORPH_RECT, (max(12, crop.shape[1] // 3), 1)),
            )
            edge_band = np.zeros_like(ink)
            edge_band[:2, :] = 255
            edge_band[-2:, :] = 255
            line_mask = cv2.bitwise_and(line_mask, edge_band)
            long_rows = np.count_nonzero(line_mask, axis=1) >= crop.shape[1] * 0.80
            line_mask[~long_rows, :] = 0
            other_ink = cv2.bitwise_and(ink, cv2.bitwise_not(line_mask))
            protected = cv2.dilate(other_ink, np.ones((3, 3), dtype=np.uint8))
            mask = (line_mask > 0) & (protected == 0)
            if not np.any(mask):
                print(prefix + "横边线淡化未执行：没有可安全处理的长横线段；不增加OCR调用")
                break
            cleaned = crop.copy()
            cleaned[mask] = (cleaned[mask].astype(np.float32) * 0.35 + 255 * 0.65).astype(np.uint8)
            prepared = cv2.resize(cleaned, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
            name = "整行2倍放大+外横边线保守淡化"
        scale_x = prepared.shape[1] / crop.shape[1]
        scale_y = prepared.shape[0] / crop.shape[0]
        # 留白在裁剪之外，不删除边缘像素；回映坐标时扣除留白。
        padding = 8
        prepared = cv2.copyMakeBorder(
            prepared, padding, padding, padding, padding,
            cv2.BORDER_CONSTANT, value=(255, 255, 255),
        )
        call_started = time.perf_counter()
        calls += 1
        try:
            with suppress_library_output():
                results = ocr.predict(prepared)
                local_items = extract_ocr_items(results, region.page_number)
            mapped = tuple(
                OCRItem(
                    -100 - index, region.page_number, item.text,
                    (
                        left + (item.bbox[0] - padding) / scale_x,
                        top + (item.bbox[1] - padding) / scale_y,
                        left + (item.bbox[2] - padding) / scale_x,
                        top + (item.bbox[3] - padding) / scale_y,
                    ),
                    item.confidence,
                )
                for index, item in enumerate(local_items)
            )
            call_elapsed = time.perf_counter() - call_started
        except Exception:
            print(prefix + f"局部OCR方式={name}；调用序号={calls}；执行成功=False；耗时={time.perf_counter()-call_started:.3f}秒；异常正文已隐藏")
            continue
        finally:
            del prepared
        succeeded += 1
        print(prefix + f"局部OCR方式={name}；调用序号={calls}；执行成功=True；耗时={call_elapsed:.3f}秒；文字块数={len(mapped)}")
        raw_length = sum(
            len(text) for item in mapped
            if re.fullmatch(r"[0-9X]+", (text := re.sub(r"\s+", "", unicodedata.normalize("NFKC", item.text)).upper()))
        )
        summary = diagnose_id_grid_assignment(
            mapped, -1, region, boundaries, True, raw_length,
            prefix + f"处理={name}；",
        )
        if summary is None:
            continue
        print(prefix + f"局部OCR前后对比（{name}）：确定格数={baseline['determined']}→{summary['determined']}；"
              + f"缺失格数={baseline['missing']}→{summary['missing']}；冲突格数={baseline['conflicts']}→{summary['conflicts']}")
        if summary["complete"]:
            complete_versions.add(summary["candidate"])
        obtained_valid = obtained_valid or summary["valid"]
        del mapped, local_items, results
        if summary["valid"]:
            print(prefix + "本版本已完整并通过全部校验，停止后续尝试；仍不自动填入")
            break
    print(prefix + f"局部整行OCR调用次数={calls}；成功调用数={succeeded}；存在完整18位候选={bool(complete_versions)}；"
          + f"存在全部校验通过候选={obtained_valid}；完整版本内容冲突={len(complete_versions) > 1}")
    print(prefix + "各版本独立验证，不跨版本拼接或按校验结果改字；版本冲突或不完整均待人工核对")
    print(prefix + f"局部整行OCR阶段总额外耗时={time.perf_counter() - started:.3f}秒")
    return calls


def diagnose_id_batch_ocr(
    image: np.ndarray, ocr: Any, region: CandidateRegion, boundaries: list[int],
    baseline: dict[str, Any], role: str, prefix: str,
) -> tuple[int, bool]:
    """仅补充不确定格；每角色一次批量提交，返回调用数和接口可用状态。"""
    started = time.perf_counter()
    calls, batch_size = 0, 6
    jobs: list[tuple[str, int, tuple[int, int, int, int]]] = []
    crops: list[np.ndarray] = []
    try:
        try:
            recognizer = getattr(getattr(ocr, "paddlex_pipeline", None), "text_rec_model", None)
            predict = getattr(recognizer, "predict", None)
            sampler = getattr(recognizer, "batch_sampler", None)
            original_batch = getattr(sampler, "batch_size", None)
            if not callable(predict) or not isinstance(original_batch, int) or original_batch < 1:
                print(prefix + "批量单字符跳过：识别接口或可恢复的批大小状态不可用")
                return 0, False
        except Exception:
            print(prefix + "批量单字符跳过：当前内部接口不兼容，未创建额外模型")
            return 0, False
        if len(boundaries) != 19 or any(b <= a for a, b in zip(boundaries, boundaries[1:])):
            print(prefix + "批量单字符跳过：实测18格边界不可靠")
            return 0, True
        selected = [index for index in range(18) if index not in baseline["determined_indices"]]
        if not selected:
            print(prefix + "批量单字符跳过：没有不确定格；即使整号校验失败也不改动已确定字符")
            return 0, True
        top = max(0, int(np.floor(region.bbox[1])))
        bottom = min(image.shape[0], int(np.ceil(region.bbox[3])))
        for index in selected:
            left = max(0, int(np.floor(boundaries[index])))
            right = min(image.shape[1], int(np.ceil(boundaries[index + 1])))
            if right <= left or bottom <= top:
                continue
            # 保留完整格内像素，仅在图像副本外围加白边，不拉伸、不强行内缩。
            crop = image[top:bottom, left:right].copy()
            padding = max(2, round(min(crop.shape[:2]) * 0.10))
            crops.append(cv2.copyMakeBorder(
                crop, padding, padding, padding, padding,
                cv2.BORDER_CONSTANT, value=(255, 255, 255),
            ))
            jobs.append((role, index, (left, top, right, bottom)))
        if not jobs:
            print(prefix + "批量单字符跳过：没有有效裁剪；原格位状态保留")
            return 0, True
        compatible = True
        inference_started = time.perf_counter()
        calls = 1
        try:
            with suppress_library_output():
                try:
                    results = list(predict(input=crops, batch_size=batch_size))
                finally:
                    try:
                        sampler.batch_size = original_batch
                        compatible = sampler.batch_size == original_batch
                    except Exception:
                        compatible = False
        except Exception:
            print(prefix + "批量单字符调用失败，已隐藏异常正文；原格位状态保留")
            return calls, compatible
        finally:
            print(prefix + f"批量识别调用耗时={time.perf_counter() - inference_started:.3f}秒；批大小恢复成功={compatible}")
        if not compatible:
            print(prefix + "批量结果弃用：批大小未能恢复，停止后续批量实验")
            return calls, False
        if len(results) != len(jobs):
            print(prefix + f"批量结果弃用：输入数={len(jobs)}；输出数={len(results)}，不按缺项顺移匹配")
            return calls, True
        try:
            parsed: list[tuple[str, float]] = []
            for result in results:
                text, score = result["rec_text"], float(result["rec_score"])
                if not isinstance(text, str):
                    raise ValueError("unexpected recognition result")
                parsed.append((unicodedata.normalize("NFKC", text).strip().upper(), score))
        except Exception:
            print(prefix + "批量结果弃用：结果字段或类型不兼容，原格位状态保留")
            return calls, True
        cells = [list(values) for values in baseline["cells"]]
        uncertain = set(baseline["uncertain"])
        extra_low: set[int] = set()
        for (job_role, index, _bbox), (text, score) in zip(jobs, parsed):
            if job_role != role or index in baseline["determined_indices"]:
                print(prefix + "批量结果弃用：角色或格号映射异常")
                return calls, True
            valid_character = re.fullmatch(r"[0-9X]" if index == 17 else r"[0-9]", text) is not None
            if not np.isfinite(score) or score < 0.75 or score > 1:
                extra_low.add(index)
                uncertain.add(index)
                continue
            if not valid_character:
                uncertain.add(index)
                continue
            previous = {character for character, _ in cells[index]}
            if previous and previous != {text}:
                # 不用新结果覆盖不同的旧候选，更不以校验码选择字符。
                cells[index].append((text, score))
                uncertain.add(index)
            else:
                confidence = max([score] + [value for _, value in cells[index] if np.isfinite(value)])
                cells[index] = [(text, confidence)]
                uncertain.discard(index)
        summary = _id_grid_state(cells, uncertain, extra_low)
        print(prefix + f"批量前后确定格数={baseline['determined']}→{summary['determined']}；"
              + f"缺失格数={baseline['missing']}→{summary['missing']}；冲突格数={baseline['conflicts']}→{summary['conflicts']}；"
              + f"低置信度格数={len(baseline['low_indices'])}→{len(summary['low_indices'])}；归属/字符不确定格数={len(summary['uncertain'])}")
        print(prefix + f"批量后候选长度={summary['length']}；完整18位={summary['complete']}；格式通过={summary['format_ok']}；"
              + f"出生日期={summary['birth']}；校验码={summary['checksum']}；全部校验通过={summary['valid']}")
        print(prefix + "批量结果仅用于诊断；保留已有确定字符，冲突/无效结果待人工核对，不自动填写")
        return calls, True
    finally:
        print(prefix + f"实际裁剪格数={len(jobs)}；批量API调用次数={calls}；批大小={batch_size}；"
              + f"预计内部批次数={(len(jobs) + batch_size - 1) // batch_size if calls else 0}；"
              + f"单字符实验总额外耗时={time.perf_counter() - started:.3f}秒")


def diagnose_id_numbers(
    pages: tuple[PageOCR, ...], page_images: dict[int, np.ndarray], ocr: Any | None = None,
) -> None:
    """输出结构统计并受限补充整行 OCR，不返回或回填字段值。"""
    remaining_calls = 6
    row_ocr_elapsed = 0.0
    batch_calls = 0
    batch_available = True
    batch_elapsed = 0.0
    mode = os.environ.get("EFFI_DEBUG_ID_MODE", "batch").strip().lower()
    if mode not in {"batch", "row", "off"}:
        print("[身份证诊断] 实验模式无效，只统计已有OCR；可选batch、row、off")
        mode = "off"
    processed_roles: set[str] = set()
    print(f"[身份证诊断] 自动填写=关闭；实验模式={mode}；整行与单字符实验互斥")
    print("[身份证诊断] row每角色最多3次完整OCR；batch每角色最多1次仅识别提交，最多18格；off不额外推理")
    print("[身份证诊断] 坐标为预处理后整张图像像素；法人栏沿用复合姓名标签角色")
    print("[身份证诊断] 未保留的检测框/空识别结果不在本次统计中，无法仅凭此证明漏字")
    for page in pages:
        if not page.items:
            print("[身份证诊断] 主 OCR 无可用文字块，无法定位或判断是否漏识别")
            continue
        page_number = page.items[0].page_number
        image = page_images.get(page_number)
        if image is None:
            print(f"[身份证诊断] 页={page_number}；图像不可用，跳过区域诊断")
            continue
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        horizontal = cv2.morphologyEx(
            binary, cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (max(30, page.width // 16), 1)),
        )
        contours, _ = cv2.findContours(horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        lines = [
            (y + height / 2, x, x + width)
            for contour in contours
            for x, y, width, height in (cv2.boundingRect(contour),)
            if width >= page.width * 0.06
        ]
        role_labels = {
            key: [item for item in page.items if find_alias_match(item.text, SPEC_BY_KEY[key])]
            for key in ("legal_name", "agent_name")
        }
        matched: dict[str, list[OCRItem]] = {key: [] for key in role_labels}
        unassigned = 0
        for label in page.items:
            if not normalized_label(label.text).startswith(("身份证号", "证件号码")):
                continue
            roles = [
                key for key, anchors in role_labels.items()
                if any(
                    anchor.bbox[0] < label.bbox[0]
                    and min(anchor.bbox[3], label.bbox[3])
                    - max(anchor.bbox[1], label.bbox[1])
                    >= min(box_height(anchor), box_height(label)) * 0.30
                    and abs(box_center_y(anchor) - box_center_y(label))
                    <= max(box_height(anchor), box_height(label))
                    for anchor in anchors
                )
            ]
            if len(roles) == 1:
                matched[roles[0]].append(label)
            else:
                unassigned += 1
        print(f"[身份证诊断] 页={page_number}；角色不明或冲突的证件标签数={unassigned}")
        for key, role_name in (("legal_name", "法人"), ("agent_name", "经办人")):
            prefix = f"[身份证诊断] 页={page_number}；角色={role_name}；"
            labels = matched[key]
            print(prefix + f"姓名标签数={len(role_labels[key])}；同排证件标签数={len(labels)}")
            if len(labels) != 1:
                print(prefix + "区域未确认：缺少唯一的同排角色/证件标签，不猜测上下顺序")
                continue
            label = labels[0]
            print(prefix + f"证件标签坐标={tuple(round(v, 1) for v in label.bbox)}")
            height = box_height(label)
            tops = [line for line in lines if line[0] <= label.bbox[1] + height * 0.30
                    and line[1] <= label.bbox[2] <= line[2]]
            bottoms = [line for line in lines if line[0] >= label.bbox[3] - height * 0.30
                       and line[1] <= label.bbox[2] <= line[2]]
            if not tops or not bottoms:
                print(prefix + "区域未确认：上下横线不完整，不扩大到其他行")
                continue
            upper, lower = max(tops), min(bottoms)
            top, bottom = int(round(upper[0])), int(round(lower[0]))
            if not height * 1.10 <= bottom - top <= page.height * 0.10:
                print(prefix + "区域未确认：横线间距不符合标签所在单行")
                continue
            opposite = "agent_name" if key == "legal_name" else "legal_name"
            if any(top < box_center_y(item) < bottom for item in role_labels[opposite]):
                print(prefix + "区域未确认：范围包含另一角色行，不分析混合区域")
                continue
            vertical = cv2.morphologyEx(
                binary[max(0, top - 2):min(page.height, bottom + 3), :], cv2.MORPH_OPEN,
                cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(8, round((bottom - top) * 0.58)))),
            )
            contours, _ = cv2.findContours(vertical, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            positions = merge_line_positions([
                round(x + width / 2) for contour in contours
                for x, _, width, line_height in (cv2.boundingRect(contour),)
                if line_height >= (bottom - top) * 0.58
                and x + width / 2 >= label.bbox[2] - height * 0.60
            ])
            left_line = bool(positions and positions[0] - label.bbox[2] <= height * 2)
            left = positions[0] if left_line else label.bbox[2]
            right = min(upper[2], lower[2])
            if right - left < height * 4:
                print(prefix + "区域未确认：标签右侧范围不足")
                continue
            region = CandidateRegion(page_number, (left, top, right, bottom))
            boundaries = [x for x in positions if left - 3 <= x <= right + 3]
            gaps = np.diff(boundaries)
            grid_evidence = (
                len(boundaries) == 19 and len(gaps) == 18
                and abs(boundaries[0] - left) <= 3 and abs(boundaries[-1] - right) <= 3
                and float(np.mean(gaps)) > 0
                and float(np.std(gaps) / np.mean(gaps)) <= 0.25
            )
            print(prefix + "定位依据=同排姓名及证件标签+上下横线端点；"
                  + f"左边线检出={left_line}；区域坐标={tuple(round(v, 1) for v in region.bbox)}；"
                  + f"竖线数={len(boundaries)}；18格结构证据={grid_evidence}")
            blocks: list[OCRItem] = []
            edge_blocks = 0
            for item in page.items:
                if item.item_id == label.item_id:
                    continue
                x0, y0, x1, y1 = item.bbox
                overlap = max(0.0, min(x1, right) - max(x0, left)) * max(
                    0.0, min(y1, bottom) - max(y0, top)
                )
                if not overlap:
                    continue
                if item_center_in_region(item, region) and overlap / max((x1-x0)*(y1-y0), 1) >= 0.50:
                    blocks.append(item)
                else:
                    edge_blocks += 1
            numeric_blocks: list[tuple[OCRItem, str]] = []
            print(prefix + f"区域文字块数={len(blocks)}；边缘归属不确定块数={edge_blocks}")
            for index, item in enumerate(blocks, start=1):
                text = re.sub(r"\s+", "", unicodedata.normalize("NFKC", item.text)).upper()
                number_like = re.fullmatch(r"[0-9X]+", text) is not None
                print(prefix + f"块序号={index}；坐标={tuple(round(v, 1) for v in item.bbox)}；"
                      + f"字符数={len(text)}；类型={'数字或X' if number_like else '其他/混合'}；"
                      + f"置信度={item.confidence:.3f}")
                if number_like:
                    numeric_blocks.append((item, text))
            ordered = sorted(numeric_blocks, key=lambda part: part[0].bbox[0])
            order_changed = [item.item_id for item, _ in numeric_blocks] != [item.item_id for item, _ in ordered]
            overlap_pairs = sum(
                min(a.bbox[2], b.bbox[2]) > max(a.bbox[0], b.bbox[0])
                and min(a.bbox[3], b.bbox[3]) > max(a.bbox[1], b.bbox[1])
                for i, (a, _) in enumerate(ordered) for b, _ in ordered[i + 1:]
            )
            # 只在诊断内存中拼接与验证，不写回字段、不修正或补齐字符。
            candidate = "".join(text for _, text in ordered)
            length_ok = len(candidate) == 18
            format_ok = re.fullmatch(r"[1-9][0-9]{16}[0-9X]", candidate) is not None
            checksum = "未执行"
            if format_ok:
                weights = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
                remainder = sum(int(c) * w for c, w in zip(candidate[:17], weights)) % 11
                checksum = "通过" if "10X98765432"[remainder] == candidate[-1] else "不通过"
            print(prefix + f"数字块数={len(ordered)}；多块拆分迹象={len(ordered) > 1}；"
                  + f"原始块顺序需调整={order_changed}；文字框重叠对数={overlap_pairs}")
            print(prefix + f"诊断候选字符数={len(candidate)}；长度通过={length_ok}；"
                  + f"字符格式通过={format_ok}；校验位={checksum}；"
                  + f"最低置信度通过={bool(ordered) and min(item.confidence for item, _ in ordered) >= 0.75}")
            print(prefix + "不足18位仅提示可能漏识别/未填/区域不全；校验通过也不证明号码或角色正确")
            print(prefix + f"主OCR存在数字块={bool(ordered)}；提取阶段状态=人工字段未接入自动提取")
            baseline = diagnose_id_grid_assignment(
                page.items, label.item_id, region, boundaries, grid_evidence,
                len(candidate), prefix,
            )
            if baseline is not None and ocr is not None and key not in processed_roles:
                processed_roles.add(key)
                if mode == "row":
                    row_started = time.perf_counter()
                    remaining_calls -= diagnose_id_row_ocr(
                        image, ocr, region, boundaries, baseline, prefix, remaining_calls,
                    )
                    row_ocr_elapsed += time.perf_counter() - row_started
                elif mode == "batch" and batch_available:
                    batch_started = time.perf_counter()
                    calls, batch_available = diagnose_id_batch_ocr(
                        image, ocr, region, boundaries, baseline, key, prefix,
                    )
                    batch_calls += calls
                    batch_elapsed += time.perf_counter() - batch_started
    print(f"[身份证诊断] 本任务身份证整行补充OCR总调用次数={6 - remaining_calls}；上限=6")
    print(f"[身份证诊断] 本任务局部整行阶段累计额外耗时={row_ocr_elapsed:.3f}秒")
    print(f"[身份证诊断] 本任务单字符批量API调用次数={batch_calls}；上限=2；累计额外耗时={batch_elapsed:.3f}秒")


def render_local_region(
    document: pymupdf.Document | None, region: CandidateRegion,
    source_image: np.ndarray | None = None,
) -> np.ndarray:
    """按标签定位区域重新高分辨率渲染，图像只保存在内存。"""
    if document is None:
        if source_image is None:
            raise ValueError("source image unavailable")
        height, width = source_image.shape[:2]
        x0, y0, x1, y1 = region.bbox
        left, top = max(0, int(np.floor(x0))), max(0, int(np.floor(y0)))
        right, bottom = min(width, int(np.ceil(x1))), min(height, int(np.ceil(y1)))
        if right <= left or bottom <= top:
            raise ValueError("empty image region")
        crop = source_image[top:bottom, left:right]
        return cv2.resize(
            crop, None, fx=360.0 / RENDER_DPI, fy=360.0 / RENDER_DPI,
            interpolation=cv2.INTER_CUBIC,
        )
    page = document.load_page(region.page_number - 1)
    source_scale = 72.0 / RENDER_DPI
    x0, y0, x1, y1 = region.bbox
    clip = pymupdf.Rect(
        x0 * source_scale,
        y0 * source_scale,
        x1 * source_scale,
        y1 * source_scale,
    ) & page.rect
    if clip.is_empty or clip.width <= 0 or clip.height <= 0:
        raise ValueError("empty region")
    local_scale = 360.0 / 72.0
    pixmap = page.get_pixmap(
        matrix=pymupdf.Matrix(local_scale, local_scale),
        clip=clip,
        colorspace=pymupdf.csRGB,
        alpha=False,
    )
    rgb = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
        pixmap.height, pixmap.width, 3
    )
    return rgb[:, :, ::-1].copy()


def supplement_with_local_ocr(
    spec: ExtractFieldSpec,
    result: ExtractionResult,
    document: pymupdf.Document | None,
    ocr: Any,
    page_images: dict[int, np.ndarray] | None = None,
    metrics=None,
) -> ExtractionResult:
    if result.recognized or not result.regions:
        return result

    # 多个候选区域无法证明属于同一字段，避免重复 OCR 和误填。
    if len(result.regions) != 1:
        diagnostic = FieldDiagnostic(
            label_found=result.diagnostic.label_found,
            region_located=result.diagnostic.region_located,
            candidate_blocks=result.diagnostic.candidate_blocks,
            format_valid=result.diagnostic.format_valid,
            reason="存在多个候选区域，未执行局部补充识别",
        )
        return ExtractionResult(diagnostic=diagnostic, regions=result.regions)

    candidates: list[Candidate] = []
    local_block_count = 0
    region = result.regions[0]
    try:
        source_image = page_images.get(region.page_number) if page_images else None
        stage = time.perf_counter()
        try:
            image = render_local_region(document, region, source_image)
        finally:
            if metrics is not None:
                metrics["render"] += time.perf_counter() - stage
                metrics["render_calls"] += 1
        stage = time.perf_counter()
        if metrics is not None:
            metrics["calls"] += 1
        try:
            with suppress_library_output():
                ocr_results = ocr.predict(image)
                items = extract_ocr_items(ocr_results, region.page_number)
        finally:
            if metrics is not None:
                metrics["ocr"] += time.perf_counter() - stage
        local_block_count += len(items)
        candidates.extend(local_candidates(spec, items))
        del image, items, ocr_results
    except Exception:
        pass

    selected, format_valid, reason = rank_candidates(candidates)
    diagnostic = FieldDiagnostic(
        label_found=result.diagnostic.label_found,
        region_located=result.diagnostic.region_located,
        candidate_blocks=result.diagnostic.candidate_blocks + local_block_count,
        format_valid=result.diagnostic.format_valid or format_valid,
        reason=(
            "局部区域补充识别成功，待人工核对"
            if selected is not None
            else reason
        ),
    )
    if selected is None:
        return ExtractionResult(diagnostic=diagnostic, regions=result.regions)
    return ExtractionResult(
        value=selected.value,
        recognized=True,
        diagnostic=diagnostic,
        regions=result.regions,
    )


class TemplateComboBox(ctk.CTkComboBox):
    """保留原生矩形边框，分离白色按钮背景与外边线颜色。"""

    def _restore_border_color(self) -> None:
        color = self._apply_appearance_mode(self.cget("border_color"))
        # 只调整原生已有边框；不创建额外图元或控件。
        self._canvas.itemconfigure("border_parts", fill=color, outline=color)

    def _draw(self, no_color_updates: bool = False) -> None:
        super()._draw(no_color_updates)
        self._restore_border_color()
        # 原生组件未提供独立箭头颜色，重绘后只调整该图元。
        if self.cget("state") != tk.DISABLED:
            self._canvas.itemconfigure("dropdown_arrow", fill="#526477")

    def _on_leave(self, event=0) -> None:
        super()._on_leave(event)
        # 原生离开事件会将右侧边线重新设为按钮背景色。
        self._restore_border_color()


class ApplicationExtractor:
    def __init__(self, root: ctk.CTk) -> None:
        self.root = root
        self.root.title("机构申请信息提取与核对")
        # CTk 的窗口尺寸使用逻辑像素，避免高 DPI 下窗口超出屏幕。
        dpi_scale = max(1.0, self.root.winfo_fpixels("1i") / 96.0)
        window_width = min(1180, max(640, int((self.root.winfo_screenwidth() - 64) / dpi_scale)))
        window_height = min(820, max(360, int((self.root.winfo_screenheight() - 100) / dpi_scale)))
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.minsize(min(960, window_width), min(600, window_height))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)
        self.root.option_add("*Font", ("Microsoft YaHei UI", 10))
        self.ui_colors = {
            "window": "#F5F5F5", "surface": "#FFFFFF", "input": "#FAFAFA",
            "border": "#E3E5E8", "text": "#242424", "muted": "#737373",
            "accent": "#426EA8", "accent_hover": "#365D90",
            "subtle": "#EEEEEE", "hover": "#E5E7EB",
            "scroll": "#D3D5D9", "scroll_hover": "#B8BDC5",
            "warning": "#896239", "valid": "#496E60",
            "template_input": "#FFFFFF", "template_border": "#B8CCE4",
            "template_hover": "#E9EEF5", "template_menu": "#FFFFFF",
            "template_text": "#243247",
        }
        self.window_background = self.ui_colors["window"]
        self.panel_background = self.ui_colors["surface"]
        self.root.configure(fg_color=self.window_background)
        self.ui_font = ctk.CTkFont(family="Microsoft YaHei UI", size=13)
        self.heading_font = ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")

        self.pdf_path: Path | None = None
        self.ocr_engine: Any | None = None
        self.busy = False
        self.closed = False
        self.request_number = 0
        self.updating_fields = False
        self.extraction_started_at: float | None = None
        self.context_entry: tk.Entry | None = None
        # 当前客户的统一字段数据区；界面和 Word 生成均与这里同步。
        self.field_data: dict[str, str] = {key: "" for key in FIELD_KEYS}
        self.field_vars: dict[str, tk.StringVar] = {}
        self.status_vars: dict[str, tk.StringVar] = {}
        self.format_vars: dict[str, tk.StringVar] = {}
        self.field_widgets: dict[str, tuple[tk.Widget, ...]] = {}

        self.file_var = tk.StringVar(value="尚未选择文件")
        self.template_var = tk.StringVar(value="")
        self.template_status_var = tk.StringVar(value="未选择 Word 模板")
        self.progress_var = tk.StringVar(value="等待选择文件")
        self.top_status_var = tk.StringVar(value="请选择 PDF 或图片文件")
        self.status_expanded = False
        self.status_layout_pending = False
        self.status_canvas_width: int | None = None
        self.status_summary_var = tk.StringVar()
        self.elapsed_var = tk.StringVar(value="耗时：-")
        self.stage_timing_var = tk.StringVar(
            value=(
                "阶段耗时：模型 -；PDF 打开 -；页面定位 -；页面渲染 -；整页 OCR -；"
                "局部 OCR -；字段匹配 -"
            )
        )
        self.review_var = tk.StringVar(value="需人工确认：-")

        self._build_interface()
        for variable in (self.top_status_var, self.elapsed_var):
            variable.trace_add("write", lambda *_args: self._update_status_summary())
        self._update_status_summary()
        self._reset_fields()
        self._update_controls()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.after_idle(self._set_minimum_layout_size)

    def _build_interface(self) -> None:
        colors = self.ui_colors
        self.entry_menu = tk.Menu(self.root, tearoff=False)
        for label, action in (
            ("剪切", "<<Cut>>"), ("复制", "<<Copy>>"),
            ("粘贴", "<<Paste>>"), ("全选", "select_all"),
        ):
            self.entry_menu.add_command(
                label=label, command=lambda command=action: self._entry_menu_action(command)
            )
        # 同时覆盖手动选页等现有对话框使用的 Tk 输入框。
        self.root.bind_class("Entry", "<Button-3>", self._show_entry_menu)
        self.root.bind_class("Entry", "<Control-a>", self._select_entry_all)
        self.root.bind_class("Entry", "<Control-A>", self._select_entry_all)

        button_options = {"height": 30, "corner_radius": 7, "font": self.ui_font}
        card_options = {
            "fg_color": self.panel_background, "corner_radius": 10,
            "border_width": 1, "border_color": colors["border"], "height": 1,
        }
        file_frame = ctk.CTkFrame(
            self.root, fg_color=self.window_background, corner_radius=0, height=1
        )
        self.file_frame = file_frame
        file_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 4))
        file_frame.columnconfigure(2, weight=1)
        self.select_button = ctk.CTkButton(
            file_frame, text="选择文件", command=self.select_pdf, width=94,
            fg_color=colors["subtle"], hover_color=colors["hover"], text_color=colors["text"],
            **button_options,
        )
        self.select_button.grid(row=0, column=0, sticky="w")
        self.start_button = ctk.CTkButton(
            file_frame, text="开始提取", command=self.start_extraction, width=100,
            fg_color=colors["accent"], hover_color=colors["accent_hover"],
            text_color=colors["surface"], **button_options,
        )
        self.start_button.grid(row=0, column=1, padx=(8, 12), sticky="w")
        ctk.CTkLabel(
            file_frame, textvariable=self.top_status_var, font=self.ui_font,
            text_color=colors["muted"], height=26, anchor="w",
        ).grid(
            row=0, column=2, sticky="w"
        )
        file_label = ctk.CTkLabel(
            file_frame, textvariable=self.file_var, font=self.ui_font,
            text_color=colors["muted"], height=20, anchor="w", justify="left",
        )
        file_label.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(3, 0))
        last_file_width = 0

        def resize_file_label(event: tk.Event) -> None:
            nonlocal last_file_width
            if event.width > 1 and event.width != last_file_width:
                last_file_width = event.width
                file_label.configure(wraplength=max(200, event.width - 24))

        tk.Misc.bind(file_frame,
            "<Configure>",
            resize_file_label,
            add="+",
        )

        # 固定模板只从项目 templates 目录读取，原始文件始终保持不变。
        template_frame = ctk.CTkFrame(self.root, **card_options)
        self.template_frame = template_frame
        template_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 6))
        template_frame.columnconfigure(3, weight=1)
        ctk.CTkLabel(
            template_frame, text="Word 模板", font=self.ui_font,
            text_color=colors["text"], height=28,
        ).grid(row=0, column=0, padx=(12, 8), pady=8)
        # 使用原生细边框，不叠加外层控件模拟边线。
        self.template_combo = TemplateComboBox(
            template_frame,
            values=list(WORD_TEMPLATE_NAMES),
            state="readonly", command=self._template_selected,
            width=300, height=30, corner_radius=0, border_width=1,
            border_color=colors["template_border"],
            fg_color=colors["template_input"], text_color=colors["template_text"],
            button_color=colors["template_input"], button_hover_color=colors["template_input"],
            dropdown_fg_color=colors["template_menu"],
            dropdown_hover_color=colors["template_hover"],
            dropdown_text_color=colors["template_text"], font=self.ui_font, dropdown_font=self.ui_font,
        )
        self.template_combo.grid(row=0, column=1, padx=(0, 10), pady=8)
        self.template_combo.bind("<Button-1>", self.template_combo._clicked)
        self.template_combo._canvas.tag_bind(
            "left_parts", "<Button-1>", self.template_combo._clicked
        )
        self.template_combo.set(self.template_var.get() or "请选择 Word 模板")
        self.generate_button = ctk.CTkButton(
            template_frame, text="生成 Word", command=self.generate_word, width=106,
            fg_color=colors["accent"], hover_color=colors["accent_hover"],
            text_color=colors["surface"], **button_options,
        )
        self.generate_button.grid(row=0, column=2, padx=(0, 10), pady=8)
        template_status_label = ctk.CTkLabel(
            template_frame, textvariable=self.template_status_var, font=self.ui_font,
            text_color=colors["muted"], height=20, anchor="w", justify="left",
        )
        template_status_label.grid(row=0, column=3, sticky="ew", padx=(0, 12))
        last_template_width = 0

        def resize_template_status(event: tk.Event) -> None:
            nonlocal last_template_width
            if event.width > 1 and event.width != last_template_width:
                last_template_width = event.width
                template_status_label.configure(
                    wraplength=max(120, event.width - self.template_combo.winfo_width() - 250)
                )

        tk.Misc.bind(
            template_frame, "<Configure>",
            resize_template_status, add="+",
        )

        # 客户信息按内容固定高度，只有下方印章列表随窗口伸缩。
        content = tk.Frame(self.root, background=self.window_background)
        self.fields_content = content
        content.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 4))
        self.root.bind("<MouseWheel>", self._scroll_status_details, add="+")
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)
        content.grid_columnconfigure((0, 1), uniform="customer_columns")
        self.customer_cards = []
        for column in range(2):
            card = ctk.CTkFrame(content, **card_options)
            card.grid(row=0, column=column, sticky="nsew", padx=2, pady=(0, 3))
            card.columnconfigure(0, weight=1)
            self.customer_cards.append(card)

        row = 0
        for section_name, fields in FIELD_SECTIONS:
            section = ctk.CTkFrame(self.customer_cards[0 if row < 2 else 1],
                                   fg_color="transparent", corner_radius=0, border_width=0, height=1)
            section.grid(row=row if row < 2 else 0, column=0, sticky="new", padx=2, pady=(0, 3))
            section.columnconfigure(1, weight=1, minsize=90)
            heading = ctk.CTkLabel(
                section, text=section_name, font=self.heading_font,
                text_color=colors["text"], height=22, anchor="w",
            )
            heading.grid(row=0, column=0, columnspan=4, sticky="ew", padx=6, pady=(3, 1))
            for field_row, (key, label) in enumerate(fields, start=1):
                value_var = tk.StringVar()
                status_var = tk.StringVar()
                format_var = tk.StringVar()
                self.field_vars[key] = value_var
                self.status_vars[key] = status_var
                self.format_vars[key] = format_var
                field_label = ctk.CTkLabel(
                    section, text=label.replace("法定代表人", "").replace("经办人", ""), width=120, height=25, anchor="w",
                    font=self.ui_font, text_color=colors["text"],
                )
                field_label.grid(
                    row=field_row, column=0, sticky="w", padx=(6, 0), pady=1
                )
                entry = ctk.CTkEntry(
                    section, textvariable=value_var, exportselection=False,
                    height=25, corner_radius=6, border_width=1, font=self.ui_font,
                    fg_color=colors["input"], border_color=colors["border"],
                    text_color=colors["text"],
                )
                entry.grid(row=field_row, column=1, sticky="ew", padx=(4, 4), pady=1)
                self._bind_entry_menu(entry)
                status_label = ctk.CTkLabel(
                    section, textvariable=status_var, width=85, height=25,
                    anchor="w", font=ctk.CTkFont(size=11), text_color=colors["muted"],
                )
                status_label.grid(
                    row=field_row, column=2, sticky="w", pady=1
                )
                format_label = ctk.CTkLabel(
                    section, textvariable=format_var, width=75, height=25,
                    anchor="w", font=ctk.CTkFont(size=11), text_color=colors["muted"],
                )
                format_label.grid(
                    row=field_row, column=3, sticky="w", padx=(0, 6), pady=1
                )
                self.field_widgets[key] = (
                    field_label,
                    entry,
                    status_label,
                    format_label,
                )
                value_var.trace_add(
                    "write", lambda *_args, field_key=key: self._field_changed(field_key)
                )
                for variable in (status_var, format_var):
                    variable.trace_add(
                        "write", lambda *_args, field_key=key: self._update_field_colors(field_key)
                    )
            row += 1
            section.grid_rowconfigure(len(fields) + 1, minsize=2)

        self.seal_panel = SealPanel(self.root, current_file=lambda: self.pdf_path, **card_options)
        self.seal_panel.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 4))

        status_frame = ctk.CTkFrame(self.root, **card_options)
        self.status_frame = status_frame
        status_frame.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 6))
        status_frame.columnconfigure(0, weight=1)
        self.status_toggle = ctk.CTkButton(
            status_frame, textvariable=self.status_summary_var,
            command=self._toggle_status_details, anchor="w", height=26,
            corner_radius=6, fg_color=self.panel_background, hover_color=colors["subtle"],
            text_color=colors["muted"], font=self.ui_font,
        )
        self.status_toggle.grid(row=0, column=0, sticky="ew", padx=6, pady=2)
        self.status_details = ctk.CTkFrame(
            status_frame, fg_color=self.panel_background, corner_radius=0, height=1
        )
        self.status_details.columnconfigure(0, weight=1)
        self.status_canvas = tk.Canvas(
            self.status_details, height=1, highlightthickness=0, borderwidth=0,
            background=self.panel_background,
        )
        self.status_canvas.grid(row=0, column=0, sticky="ew")
        status_scrollbar = ctk.CTkScrollbar(
            self.status_details, orientation="vertical", command=self.status_canvas.yview,
            # 滚动条不能用默认 200 像素请求高度，否则会撑大详情区。
            width=12, height=1, fg_color=self.panel_background,
            button_color=colors["scroll"], button_hover_color=colors["scroll_hover"],
        )
        status_scrollbar.grid(row=0, column=1, sticky="ns")
        self.status_canvas.configure(yscrollcommand=status_scrollbar.set)
        self.status_content = tk.Frame(
            self.status_canvas, background=self.panel_background, padx=2, pady=0
        )
        self.status_content_window = self.status_canvas.create_window(
            (0, 0), window=self.status_content, anchor="nw"
        )
        self.status_labels = []
        for variable in (
            self.progress_var, self.elapsed_var, self.stage_timing_var, self.review_var
        ):
            label = ctk.CTkLabel(
                self.status_content, textvariable=variable, wraplength=800,
                font=self.ui_font, text_color=colors["muted"], height=20,
                anchor="w", justify="left",
            )
            label.pack(anchor="w", fill=tk.X, pady=(0, 2))
            self.status_labels.append(label)
        self.status_content.bind("<Configure>", self._status_content_changed)
        self.status_canvas.bind("<Configure>", self._status_canvas_resized)
        content.bind("<Configure>", self._queue_status_layout, add="+")
        self.root.bind("<Configure>", self._queue_status_layout, add="+")

    def _update_field_colors(self, key: str) -> None:
        colors = self.ui_colors
        status = self.status_vars[key].get()
        if "未识别" in status:
            state_color = colors["warning"]
        elif "已识别" in status or "人工修改" in status:
            state_color = colors["accent"]
        else:
            state_color = colors["muted"]
        format_text = self.format_vars[key].get()
        format_color = colors["muted"]
        if "未通过" in format_text:
            format_color = colors["warning"]
        elif "通过" in format_text:
            format_color = colors["valid"]
        self.field_widgets[key][2].configure(text_color=state_color)
        self.field_widgets[key][3].configure(text_color=format_color)

    def _bind_entry_menu(self, entry: ctk.CTkEntry) -> None:
        entry.bind("<Button-3>", self._show_entry_menu)
        entry.bind("<Control-a>", self._select_entry_all)
        entry.bind("<Control-A>", self._select_entry_all)

    def _select_entry_all(self, event: tk.Event) -> str:
        entry = event.widget
        entry.selection_range(0, tk.END)
        entry.icursor(tk.END)
        return "break"

    def _entry_is_editable(self, entry: tk.Entry) -> bool:
        return str(entry.cget("state")) == tk.NORMAL

    def _show_entry_menu(self, event: tk.Event) -> str:
        entry = event.widget
        if not self._entry_is_editable(entry):
            return "break"
        self.context_entry = entry
        entry.configure(exportselection=False)
        entry.focus_set()
        if not entry.selection_present():
            entry.icursor(entry.index(f"@{event.x}"))
        selected = entry.selection_present()
        for index in (0, 1):
            self.entry_menu.entryconfigure(index, state=tk.NORMAL if selected else tk.DISABLED)
        try:
            can_paste = bool(self.root.clipboard_get())
        except tk.TclError:
            can_paste = False
        self.entry_menu.entryconfigure(2, state=tk.NORMAL if can_paste else tk.DISABLED)
        self.entry_menu.entryconfigure(3, state=tk.NORMAL if entry.get() else tk.DISABLED)
        try:
            self.entry_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.entry_menu.grab_release()
        return "break"

    def _entry_menu_action(self, action: str) -> None:
        entry = self.context_entry
        if (
            entry is None or not entry.winfo_exists()
            or not self._entry_is_editable(entry)
        ):
            return
        entry.focus_set()
        if action == "select_all":
            entry.selection_range(0, tk.END)
            entry.icursor(tk.END)
        else:
            # 使用 Tk 原生编辑事件，继续触发 StringVar 的数据同步回调。
            entry.event_generate(action)

    def _update_status_summary(self) -> None:
        status = self.top_status_var.get().split("，用时", 1)[0]
        elapsed = self.elapsed_var.get()
        parts = ["提取状态：" + status]
        if elapsed != "耗时：-":
            parts.append(elapsed.replace("耗时：", "用时 "))
        parts.append("收起详情" if self.status_expanded else "查看详情")
        self.status_summary_var.set(" · ".join(parts))

    def _toggle_status_details(self) -> None:
        self.status_expanded = not self.status_expanded
        if self.status_expanded:
            self.status_details.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 3))
            self._queue_status_layout()
        else:
            # 同时清除 CTk 的布局记录，防止 DPI 刷新重新显示详情。
            self.status_details.grid_forget()
            self.status_canvas.configure(height=1)
            self.status_details.configure(height=1)
            self.status_frame.grid_rowconfigure(1, minsize=0, weight=0)
            self.status_frame.configure(height=1)
        self._update_status_summary()

    def _status_content_changed(self, _event: tk.Event) -> None:
        self.status_canvas.configure(scrollregion=self.status_canvas.bbox("all"))
        self._queue_status_layout()

    def _status_canvas_resized(self, event: tk.Event) -> None:
        # 高度变化也会触发 Configure，但不需要重新计算文字换行。
        if event.width > 1 and event.width != self.status_canvas_width:
            self.status_canvas_width = event.width
            self.status_canvas.itemconfigure(self.status_content_window, width=event.width)
            for label in self.status_labels:
                label.configure(wraplength=max(160, event.width - 12))
        self._queue_status_layout()

    def _queue_status_layout(self, event: tk.Event | None = None) -> None:
        if self.closed or not self.status_expanded:
            return
        if event is not None and event.widget not in (self.root, self.fields_content):
            return
        if not self.status_layout_pending:
            self.status_layout_pending = True
            self.root.after_idle(self._resize_status_details)

    def _resize_status_details(self) -> None:
        self.status_layout_pending = False
        if self.closed or not self.status_expanded:
            return
        # 按内容决定高度；只在超过可用空间或 160 像素上限时滚动。
        overhead = (
            self.file_frame.winfo_reqheight() + self.template_frame.winfo_reqheight()
            + self.status_toggle.winfo_reqheight() + 70
        )
        available = max(1, self.root.winfo_height() - overhead)
        spare = available - self.fields_content.winfo_reqheight()
        height = min(
            160, max(24, spare - 8), max(24, round(available * 0.25)),
            max(1, self.status_content.winfo_reqheight()),
        )
        if int(self.status_canvas.cget("height")) != height:
            self.status_canvas.configure(height=height)
            self.root.after_idle(self._set_minimum_layout_size)

    def _scroll_status_details(self, event: tk.Event) -> None:
        widget_path = str(event.widget)
        if (
            event.delta and self.status_expanded
            and (widget_path == str(self.status_canvas)
                 or widget_path.startswith(str(self.status_content) + ".")
                 or widget_path == str(self.status_content))
        ):
            if self.status_content.winfo_height() > self.status_canvas.winfo_height():
                self.status_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
            return "break"

    def _set_minimum_layout_size(self):
        if self.closed:
            return
        scaling = self.root._get_window_scaling()
        # 最小窗口容纳全部客户字段、固定操作栏和一段可用的候选列表。
        fixed = (self.file_frame.winfo_reqheight() + self.template_frame.winfo_reqheight()
                 + self.fields_content.winfo_reqheight() + self.status_frame.winfo_reqheight())
        seal_controls = sum(self.seal_panel.grid_bbox(0, row, 0, row)[3] for row in range(3))
        minimum_height = math.ceil((fixed + seal_controls + 100 + 40) / scaling)
        minimum_width = max(960, math.ceil((self.fields_content.winfo_reqwidth() + 24) / scaling))
        self.root.minsize(minimum_width, minimum_height)

    def _extraction_elapsed(self, fallback: float) -> float:
        if self.extraction_started_at is None:
            return fallback
        return time.perf_counter() - self.extraction_started_at

    def _set_field_value(self, key: str, value: str) -> None:
        """同步更新当前客户数据和对应输入框。"""
        self.field_data[key] = value
        if self.field_vars[key].get() != value:
            self.field_vars[key].set(value)

    def _reset_fields(self) -> None:
        self.updating_fields = True
        try:
            for key in FIELD_KEYS:
                self._set_field_value(key, "")
                self.status_vars[key].set(
                    "人工填写" if key in MANUAL_KEYS else "未识别，需补充"
                )
                self.format_vars[key].set("")
        finally:
            self.updating_fields = False
        self._update_format_status("credit_code")
        self._update_format_status("agent_mobile")
        self._update_format_status("personal_mobile")
        self._update_review_summary()

    def _field_changed(self, key: str) -> None:
        if self.updating_fields:
            return
        # Tkinter 写入回调是同步的，人工修改立即进入统一字段数据区。
        self.field_data[key] = self.field_vars[key].get()
        value = self.field_data[key].strip()
        if key in MANUAL_KEYS:
            self.status_vars[key].set("人工填写")
        elif value:
            self.status_vars[key].set("人工修改，待核对")
        else:
            self.status_vars[key].set("未识别，需补充")
        self._update_format_status(key)
        self._update_review_summary()

    def _update_format_status(self, key: str) -> None:
        if key not in {"credit_code", "agent_mobile", "personal_mobile"}:
            return
        value = self.field_data[key].strip()
        if not value:
            status = "格式：待填写"
        else:
            valid = SPEC_BY_KEY[key].validator(value)
            status = f"格式：{'通过' if valid else '未通过'}"
        self.format_vars[key].set(status)

    def _update_review_summary(self) -> None:
        needs_review = []
        for _, fields in FIELD_SECTIONS:
            for key, label in fields:
                if (
                    key == "company_address"
                    and self.template_var.get().strip() != INSTITUTION_TEMPLATE
                ):
                    continue
                status = self.status_vars[key].get()
                value = self.field_data[key].strip()
                if key in MANUAL_KEYS:
                    if not value:
                        needs_review.append(label)
                elif "待核对" in status or "需补充" in status:
                    needs_review.append(label)
        self.review_var.set(
            "需人工确认：" + ("、".join(needs_review) if needs_review else "无")
        )

    def _update_controls(self) -> None:
        self.select_button.configure(state=tk.DISABLED if self.busy else tk.NORMAL)
        self.template_combo.configure(
            state=tk.DISABLED if self.busy else "readonly"
        )
        template_name = self.template_var.get().strip()
        template_ready = (
            template_name in WORD_TEMPLATE_NAMES
            and (TEMPLATE_DIR / template_name).is_file()
        )
        self.generate_button.configure(
            state=(tk.NORMAL if template_ready and not self.busy else tk.DISABLED)
        )
        self.start_button.configure(
            state=(
                tk.NORMAL
                if self.pdf_path is not None and not self.busy
                else tk.DISABLED
            )
        )

    def _template_selected(self, _choice: str | None = None) -> None:
        """仅检查固定模板是否已放入本地目录。"""
        # 未选择提示仅用于显示，真实模板状态仍保持空值。
        if _choice is not None:
            self.template_var.set(_choice)
        template_name = self.template_var.get().strip()
        self._update_review_summary()
        if not template_name:
            self.template_status_var.set("未选择 Word 模板")
            self._update_controls()
            return
        template_path = TEMPLATE_DIR / template_name
        self.template_status_var.set(
            "模板已就绪"
            if template_path.is_file()
            else "模板文件尚未放入 templates 目录"
        )
        self._update_controls()

    def generate_word(self) -> None:
        """使用界面当前值生成新 Word，不覆盖内置模板。"""
        template_name = self.template_var.get().strip()
        if template_name not in WORD_TEMPLATE_NAMES:
            messagebox.showerror("无法生成", "请先选择 Word 模板。")
            return
        template_path = TEMPLATE_DIR / template_name
        if not template_path.is_file():
            messagebox.showerror("无法生成", "未找到所选 Word 模板。")
            return

        seals = self.seal_panel.prepare_export()
        if seals is None:
            return

        output = filedialog.asksaveasfilename(
            title="保存生成的 Word 文件",
            initialdir=str(self.pdf_path.parent if self.pdf_path else PROJECT_DIR),
            initialfile=f"{template_path.stem}_已填写.docx",
            defaultextension=".docx",
            filetypes=(("Word 文档", "*.docx"),),
        )
        if not output:
            return
        output_path = Path(output)
        if output_path.suffix.lower() != ".docx":
            messagebox.showerror("无法生成", "输出文件必须使用 .docx 扩展名。")
            return

        value_keys = [
            "company_name",
            "credit_code",
            "agent_name",
            "agent_id",
            "agent_mobile",
            "agent_email",
        ]
        # 个人模板不读取单位法定地址。
        if template_name == INSTITUTION_TEMPLATE:
            value_keys.insert(2, "company_address")
        else:
            value_keys.extend(("legal_name", "legal_id", "personal_mobile"))
        values = {
            key: self.field_data[key].strip()
            for key in value_keys
        }
        safe_stem = re.sub(r"[^0-9A-Za-z]", "", output_path.stem).lower()
        for key in (
            "credit_code", "agent_id", "agent_mobile", "legal_id", "personal_mobile"
        ):
            sensitive = re.sub(r"[^0-9A-Za-z]", "", values.get(key, "")).lower()
            if len(sensitive) >= 7 and sensitive in safe_stem:
                messagebox.showerror(
                    "文件名不安全",
                    "文件名不得包含证件号码、手机号或统一社会信用代码。",
                )
                return

        try:
            fill_word_template(template_path, output_path, values, seals=seals)
        except TemplateStructureError as exc:
            messagebox.showerror("模板结构不匹配", str(exc))
            return
        except Exception:
            messagebox.showerror("生成失败", "Word 文件生成失败，请检查保存位置。")
            return

        self.template_status_var.set(f"已生成：{output_path.name}")
        messagebox.showinfo("生成成功", f"Word 文件已保存至：\n{output_path}")

    def select_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="选择申请表文件",
            filetypes=(
                ("申请表文件", "*.pdf *.jpg *.jpeg *.png"),
                ("PDF 文件", "*.pdf"), ("图片文件", "*.jpg *.jpeg *.png"),
            ),
        )
        if not selected:
            return
        path = Path(selected)
        if path.suffix.lower() not in IMAGE_SUFFIXES | {".pdf"} or not path.is_file():
            messagebox.showerror("选择失败", "请选择有效的本地 PDF、JPG、JPEG 或 PNG 文件。")
            return
        self.pdf_path = path
        self.seal_panel.reset_customer()
        self.file_var.set(path.name)
        self.extraction_started_at = None
        self.top_status_var.set("准备就绪")
        self.progress_var.set("文件已选择，等待开始提取")
        self.elapsed_var.set("耗时：-")
        self.stage_timing_var.set(
            ("阶段耗时：模型 -；图片读取与预处理 -；整页 OCR -；"
             if path.suffix.lower() in IMAGE_SUFFIXES else
             "阶段耗时：模型 -；PDF 打开 -；页面定位 -；页面渲染 -；整页 OCR -；")
            +
            "局部 OCR -；字段匹配 -"
        )
        self.status_canvas.yview_moveto(0)
        self._reset_fields()
        self._update_controls()

    def start_extraction(self) -> None:
        if self.busy or self.pdf_path is None:
            return
        self.extraction_started_at = time.perf_counter()
        self.top_status_var.set("正在提取中，请稍等…")
        self.elapsed_var.set("耗时：-")
        self.stage_timing_var.set(
            ("阶段耗时：模型 -；图片读取与预处理 -；整页 OCR -；"
             if self.pdf_path.suffix.lower() in IMAGE_SUFFIXES else
             "阶段耗时：模型 -；PDF 打开 -；页面定位 -；页面渲染 -；整页 OCR -；")
            +
            "局部 OCR -；字段匹配 -"
        )
        self.status_canvas.yview_moveto(0)
        if not models_are_ready():
            self.top_status_var.set("提取失败，请检查提示")
            self.progress_var.set("本地 OCR 模型不可用")
            self.extraction_started_at = None
            messagebox.showerror(
                "本地模型不可用",
                "未找到已下载的本地 OCR 模型，请先使用 ocr_test.py 准备模型。",
            )
            return
        self.busy = True
        if self.pdf_path.suffix.lower() == ".pdf":
            self.seal_panel.use_current_pdf()
            self.seal_request_token = self.seal_panel.generation
        self.request_number += 1
        request_number = self.request_number
        path = self.pdf_path
        self.progress_var.set("正在初始化本地 OCR……")
        self._reset_fields()
        self._update_controls()
        worker = threading.Thread(
            target=self._extract_worker,
            args=(path, request_number),
            daemon=True,
        )
        worker.start()

    def _extract_worker(
        self, path: Path, request_number: int, forced_page_number: int | None = None,
    ) -> None:
        # 与独立印章查找共用可重入锁，同一模型始终串行推理。
        waiting = time.perf_counter()
        with self.seal_panel.pdf_reader.lock:
            self.ocr_lock_wait = time.perf_counter() - waiting
            if self.ocr_engine is None:
                self.ocr_engine = self.seal_panel.pdf_reader.engine
            self._extract_worker_serial(path, request_number, forced_page_number)

    def _extract_worker_serial(
        self,
        path: Path,
        request_number: int,
        forced_page_number: int | None = None,
    ) -> None:
        start = time.perf_counter()
        failure_stage = "提取准备"
        pages: list[PageOCR] = []
        page_images: dict[int, np.ndarray] = {}
        document: pymupdf.Document | None = None
        model_initialization = StageMetric()
        pdf_open = StageMetric()
        image_loading = StageMetric()
        is_image = path.suffix.lower() in IMAGE_SUFFIXES
        page_detection_calls = 0
        page_detection_elapsed = 0.0
        page_rendering_calls = 0
        page_rendering_elapsed = 0.0
        page_ocr_calls = 0
        page_ocr_elapsed = 0.0
        field_matching = StageMetric()
        local_supplement = StageMetric()
        model_reused = self.ocr_engine is not None
        seal_lookup = StageMetric()
        title_metrics = dict(render=0.0, ocr=0.0, calls=0)
        local_metrics = dict(render=0.0, render_calls=0, ocr=0.0, calls=0)

        def current_timings() -> StageTimings:
            return StageTimings(
                model_initialization=model_initialization,
                pdf_open=pdf_open,
                page_detection=StageMetric(
                    page_detection_calls, page_detection_elapsed
                ),
                page_rendering=StageMetric(
                    page_rendering_calls, page_rendering_elapsed
                ),
                page_ocr=StageMetric(page_ocr_calls, page_ocr_elapsed),
                local_supplement=local_supplement,
                field_matching=field_matching,
                model_reused=model_reused,
                image_loading=image_loading,
                seal_lookup=seal_lookup,
                title_ocr=StageMetric(title_metrics['calls'], title_metrics['ocr']),
                title_render=StageMetric(title_metrics['calls'], title_metrics['render']),
                lock_wait=StageMetric(1, self.ocr_lock_wait),
                local_render=StageMetric(local_metrics["render_calls"], local_metrics["render"]),
            )

        panel_token = self.seal_request_token if not is_image else None

        def locate_seals():
            nonlocal seal_lookup
            if self.closed or request_number != self.request_number or panel_token != self.seal_panel.generation:
                return
            seal_started = time.perf_counter()
            try:
                seal_source = self.seal_panel.pdf_reader.find(path,
                    lambda text: self.root.after(0, self.seal_panel.pdf_progress, panel_token, text),
                    lambda: self.closed or request_number != self.request_number or panel_token != self.seal_panel.generation,
                    engine=self.ocr_engine, allow_ocr=False)
            except Exception:
                seal_source = None
                if not self.closed:
                    self.root.after(0, self.seal_panel.pdf_progress, panel_token,
                                    "印章页面查找失败，可稍后单独重新查找；客户信息继续提取")
            seal_lookup = StageMetric(1, time.perf_counter() - seal_started)
            if seal_source is not None and not self.closed:
                cached = self.seal_panel.pdf_reader.last_cache_hit
                print(f"[印章自动轻量查找] elapsed={time.perf_counter() - seal_started:.3f}s cached={cached} pages={len(seal_source.matches)} "
                          + ("" if cached else " ".join(
                              f"{key}={int(value)}" if key in {"ocr_calls", "reused", "deferred"} else f"{key}={value:.3f}s"
                              for key, value in seal_source.timings.items())))
                self.root.after(0, self.seal_panel.accept_pdf_source, panel_token, seal_source)

        try:
            if is_image:
                failure_stage = "图片读取与预处理"
                stage_start = time.perf_counter()
                try:
                    image = read_image(path)
                finally:
                    image_loading = StageMetric(1, time.perf_counter() - stage_start)
                target_page_number = page_count = 1
            else:
                failure_stage = "PDF 打开"
                stage_start = time.perf_counter()
                document = pymupdf.open(path)
                pdf_open = StageMetric(1, time.perf_counter() - stage_start)
                if document.needs_pass or document.page_count == 0:
                    raise ValueError("unsupported PDF")

                page_count = document.page_count
                if self.ocr_engine is None:
                    stage_start = time.perf_counter()
                    with suppress_library_output():
                        self.ocr_engine = create_ocr(use_local_models=True)
                    model_initialization = StageMetric(1, time.perf_counter() - stage_start)
                self.seal_panel.pdf_reader.prepare(path)
                self.seal_panel.pdf_reader.engine = self.ocr_engine
                if forced_page_number is not None:
                    if not 1 <= forced_page_number <= page_count:
                        raise ValueError("invalid page number")
                    target_page_number = forced_page_number
                else:
                    failure_stage = "目标页文字层定位"
                    stage_start = time.perf_counter()
                    text_matches = find_title_pages_in_text_layer(document)
                    page_detection_calls += page_count
                    page_detection_elapsed += time.perf_counter() - stage_start
                    if len(text_matches) == 1:
                        target_page_number = text_matches[0]
                    elif len(text_matches) > 1:
                        raise TargetPageSelectionRequired(page_count)
                    else:
                        if self.ocr_engine is None:
                            failure_stage = "OCR 模型初始化"
                            stage_start = time.perf_counter()
                            with suppress_library_output():
                                self.ocr_engine = create_ocr(use_local_models=True)
                            model_initialization = StageMetric(
                                1, time.perf_counter() - stage_start
                            )
                        failure_stage = "目标页标题 OCR"
                        stage_start = time.perf_counter()
                        ocr_matches = find_title_pages_with_low_resolution_ocr(
                            document, self.ocr_engine, self.seal_panel.pdf_reader.full_items,
                            self.seal_panel.pdf_reader.title_items, title_metrics
                        )
                        page_detection_calls += page_count
                        page_detection_elapsed += time.perf_counter() - stage_start
                        if len(ocr_matches) != 1:
                            raise TargetPageSelectionRequired(page_count)
                        target_page_number = ocr_matches[0]

            if self.ocr_engine is None:
                failure_stage = "OCR 模型初始化"
                stage_start = time.perf_counter()
                with suppress_library_output():
                    self.ocr_engine = create_ocr(use_local_models=True)
                model_initialization = StageMetric(
                    1, time.perf_counter() - stage_start
                )

            if self.closed:
                return
            self.seal_panel.pdf_reader.engine = self.ocr_engine
            self.root.after(
                0,
                self._update_progress,
                request_number,
                target_page_number,
                page_count,
            )
            if not is_image:
                failure_stage = "目标页渲染"
                target_page = document.load_page(target_page_number - 1)
                stage_start = time.perf_counter()
                page_rendering_calls = 1
                image = render_page(target_page)
                page_rendering_elapsed = time.perf_counter() - stage_start
            failure_stage = "申请信息表格定位"
            table_region = detect_application_table_region(image)
            if table_region is None:
                if is_image:
                    raise ImageInputError("未能定位申请信息表格，请选择清晰、端正且包含完整表格的图片。")
                raise ValueError("application table not reliably located")
            x0, y0, x1, y1 = table_region
            table_image = image[y0:y1, x0:x1]
            if table_image.size == 0:
                raise ValueError("empty application table")

            failure_stage = "申请信息 OCR"
            stage_start = time.perf_counter()
            page_ocr_calls = 1
            try:
                with suppress_library_output():
                    ocr_results = self.ocr_engine.predict(table_image)
                    cropped_items = extract_ocr_items(ocr_results, target_page_number)
            finally:
                page_ocr_elapsed = time.perf_counter() - stage_start
            items = offset_ocr_items(cropped_items, x0, y0)
            pages.append(PageOCR(tuple(items), image.shape[1], image.shape[0]))
            page_images[target_page_number] = image
            del table_image, cropped_items, items, ocr_results

            failure_stage = "字段定位与匹配"
            stage_start = time.perf_counter()
            extracted = extract_application_fields(tuple(pages))
            local_region_keys: set[str] = set()
            for key in ("company_name", "credit_code"):
                original = extracted[key]
                refined = refine_centered_table_field(
                    SPEC_BY_KEY[key], tuple(pages), page_images, original
                )
                extracted[key] = refined
                if refined is not original:
                    local_region_keys.add(key)
            original_legal = extracted["legal_name"]
            extracted["legal_name"] = refine_legal_name_result(
                tuple(pages), page_images, original_legal
            )
            if extracted["legal_name"] is not original_legal:
                local_region_keys.add("legal_name")
            extracted["personal_mobile"] = extract_personal_mobile(
                tuple(pages), page_images
            )
            if extracted["personal_mobile"].regions:
                local_region_keys.add("personal_mobile")
            field_matching = StageMetric(1, time.perf_counter() - stage_start)

            # 主表 OCR 缺失时，每个重点字段最多补充识别一次。
            local_elapsed = 0.0
            for key in (
                "company_name",
                "credit_code",
                "legal_name",
                "personal_mobile",
            ):
                if key not in local_region_keys:
                    continue
                spec = SPEC_BY_KEY[key]
                result = extracted[key]
                if result.recognized or len(result.regions) != 1:
                    continue
                self.root.after(
                    0,
                    self._update_local_progress,
                    request_number,
                    spec.name,
                )
                failure_stage = f"局部补充 OCR（{spec.name}）"
                local_started = time.perf_counter()
                extracted[key] = supplement_with_local_ocr(
                    spec,
                    result,
                    document,
                    self.ocr_engine,
                    page_images if is_image else None,
                    local_metrics,
                )
                local_elapsed += time.perf_counter() - local_started
            local_supplement = StageMetric(local_metrics["calls"], local_metrics["ocr"])
            field_matching = StageMetric(field_matching.calls, field_matching.elapsed
                                        + max(0.0, local_elapsed - local_metrics["render"] - local_metrics["ocr"]))
            elapsed = time.perf_counter() - start
            timings = current_timings()
            print_performance_report(timings, elapsed)
            if os.environ.get("EFFI_DEBUG_ID", "").strip().lower() in {"1", "true", "yes"}:
                diagnostic_started = time.perf_counter()
                try:
                    diagnose_id_numbers(tuple(pages), page_images, self.ocr_engine)
                except Exception:
                    print("[身份证诊断] 诊断未完成，已隐藏异常正文；不影响正常字段结果")
                print(f"[身份证诊断] 独立诊断耗时={time.perf_counter() - diagnostic_started:.3f}秒")
            pages.clear()
            page_images.clear()
            if not self.closed:
                self.root.after(
                    0,
                    self._finish_extraction,
                    request_number,
                    extracted,
                    elapsed,
                    timings,
                    True,
                )
                if not is_image:
                    locate_seals()
        except TargetPageSelectionRequired as exc:
            elapsed = time.perf_counter() - start
            timings = current_timings()
            print_performance_report(timings, elapsed)
            pages.clear()
            page_images.clear()
            if not self.closed:
                self.root.after(
                    0,
                    self._request_manual_target_page,
                    path,
                    request_number,
                    exc.page_count,
                    elapsed,
                    timings,
                )
                if not is_image:
                    locate_seals()
        except Exception as exc:
            elapsed = time.perf_counter() - start
            timings = current_timings()
            print_performance_report(timings, elapsed)
            print_extraction_exception(failure_stage, exc, path)
            failure_reason = extraction_failure_summary(failure_stage, exc)
            pages.clear()
            page_images.clear()
            if not self.closed:
                self.root.after(
                    0,
                    self._finish_extraction,
                    request_number,
                    {},
                    elapsed,
                    timings,
                    False,
                    failure_reason,
                )
        finally:
            if document is not None:
                document.close()

    def _request_manual_target_page(
        self,
        path: Path,
        request_number: int,
        page_count: int,
        elapsed: float,
        timings: StageTimings,
    ) -> None:
        if self.closed or request_number != self.request_number:
            return
        page_number = simpledialog.askinteger(
            "选择申请表页面",
            "未能可靠定位“机构/个人数字证书业务申请表”。\n"
            f"请输入申请表所在页码（1～{page_count}）：",
            parent=self.root,
            minvalue=1,
            maxvalue=page_count,
        )
        if page_number is None:
            self.busy = False
            elapsed = self._extraction_elapsed(elapsed)
            self.extraction_started_at = None
            self.top_status_var.set("已取消提取")
            self.progress_var.set("未选择目标页面，提取已停止")
            self.elapsed_var.set(f"耗时：{elapsed:.2f} 秒")
            self._show_stage_timings(timings)
            self._update_controls()
            return
        self.progress_var.set(f"正在处理人工选择的第 {page_number} 页……")
        worker = threading.Thread(
            target=self._extract_worker,
            args=(path, request_number, page_number),
            daemon=True,
        )
        worker.start()

    def _update_progress(
        self, request_number: int, page_number: int, page_count: int
    ) -> None:
        if request_number == self.request_number:
            self.progress_var.set(
                "正在识别申请表图片……"
                if self.pdf_path and self.pdf_path.suffix.lower() in IMAGE_SUFFIXES
                else f"正在识别第 {page_number} / {page_count} 页……"
            )

    def _update_local_progress(self, request_number: int, field_name: str) -> None:
        if request_number == self.request_number:
            self.progress_var.set(f"正在补充识别已定位区域：{field_name}……")

    def _show_stage_timings(self, timings: StageTimings) -> None:
        model_note = "（已复用）" if timings.model_reused else ""
        metric_text = lambda metric: f"{metric.calls}次/{metric.elapsed:.3f}秒"
        input_summary = (
            f"图片读取与预处理 {metric_text(timings.image_loading)}；"
            if timings.image_loading.calls else
            f"PDF 打开 {metric_text(timings.pdf_open)}；"
            f"页面定位 {metric_text(timings.page_detection)}；"
            f"页面渲染 {metric_text(timings.page_rendering)}；"
        )
        self.stage_timing_var.set(
            "阶段耗时："
            f"模型 {metric_text(timings.model_initialization)}{model_note}；"
            + input_summary +
            f"申请表格 OCR {metric_text(timings.page_ocr)}；"
            f"印章页查找 {metric_text(timings.seal_lookup)}；"
            f"局部 OCR {metric_text(timings.local_supplement)}；"
            f"字段匹配 {metric_text(timings.field_matching)}"
        )

    def _finish_extraction(
        self,
        request_number: int,
        extracted: dict[str, ExtractionResult],
        elapsed: float,
        timings: StageTimings,
        success: bool,
        failure_reason: str = "",
    ) -> None:
        if self.closed or request_number != self.request_number:
            return
        ui_started = time.perf_counter()
        overall_started = self.extraction_started_at
        elapsed = self._extraction_elapsed(elapsed)
        self.extraction_started_at = None
        self.busy = False
        self.updating_fields = True
        try:
            for spec in EXTRACT_SPECS:
                result = extracted.get(spec.key)
                if result is None:
                    result = ExtractionResult(
                        diagnostic=FieldDiagnostic(
                            reason=failure_reason or "提取过程未完成"
                        )
                    )
                self._set_field_value(
                    spec.key, result.value if result.recognized else ""
                )
                self.status_vars[spec.key].set(
                    "已识别，待核对" if result.recognized else "未识别，需补充"
                )
        finally:
            self.updating_fields = False
        extracted.clear()
        self._update_format_status("credit_code")
        self._update_format_status("agent_mobile")
        self._update_format_status("personal_mobile")
        self._update_review_summary()
        self.elapsed_var.set(f"耗时：{elapsed:.2f} 秒")
        self._show_stage_timings(timings)
        self.progress_var.set(
            "提取完成"
            if success
            else f"{failure_reason or '提取失败'}，请人工填写"
        )
        self.top_status_var.set(
            f"提取完成，用时 {elapsed:.1f} 秒"
            if success else "提取失败，请检查提示"
        )
        self._update_controls()
        ui_elapsed = time.perf_counter() - ui_started
        wall_elapsed = time.perf_counter() - overall_started if overall_started is not None else elapsed + ui_elapsed
        self.elapsed_var.set(f"耗时：{wall_elapsed:.2f} 秒")
        print(f"[性能统计] UI数据更新={ui_elapsed:.3f}s 最终墙钟={wall_elapsed:.3f}s（含锁等待、队列及人工等待）")
        if not success:
            messagebox.showerror(
                "提取失败",
                f"{failure_reason or '提取过程未完成'}。\n"
                "请运行调试版查看详细异常信息。",
            )

    def close(self) -> None:
        self.closed = True
        self.request_number += 1
        self.pdf_path = None
        self.ocr_engine = None
        self.updating_fields = True
        for key in FIELD_KEYS:
            self._set_field_value(key, "")
        self.root.destroy()


def main() -> None:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    ApplicationExtractor(root)
    root.mainloop()


if __name__ == "__main__":
    main()
