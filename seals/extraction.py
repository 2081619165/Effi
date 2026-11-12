"""读取原始图片字节，候选是否为印章由客服确认。"""

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import warnings

from docx import Document
from docx.image.image import Image as DocxImage
from docx.oxml.ns import qn
from PIL import Image
from .regions import collection_label


@dataclass(frozen=True)
class SealCandidate:
    data: bytes
    width: int
    height: int
    context: str
    selectable: bool = True
    position: tuple[int, int, int] = (0, 0, 0)
    source_data: bytes | None = None
    pdf_page: int | None = None


def candidate_from_bytes(data: bytes, context: str) -> SealCandidate:
    # 只验证编码和尺寸，插入 Word 时沿用原始字节。
    with warnings.catch_warnings():
        warnings.simplefilter("error", Image.DecompressionBombWarning)
        with Image.open(BytesIO(data)) as image:
            if getattr(image, "n_frames", 1) != 1:
                raise ValueError("不支持多帧图片")
            image.load()
            width, height = image.size
    DocxImage.from_blob(data)
    return SealCandidate(data, width, height, context)


def extract_docx_images(path: Path) -> tuple[list[SealCandidate], str]:
    document = Document(path)
    candidates = []
    unsupported = 0
    source_order = 0
    # 延续验证工具的表格标签关联方式，同时展示区域外图片供人工核对。
    for node in document.element.iter():
        if node.tag in {qn("w:drawing"), qn("w:pict")}:
            has_image = any(child.tag in {qn("a:blip"), "{urn:schemas-microsoft-com:vml}imagedata"}
                            for child in node.iter())
            if not has_image:
                unsupported += 1
        if node.tag not in {qn("a:blip"), "{urn:schemas-microsoft-com:vml}imagedata"}:
            if node.tag in {qn("w:object"), "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}wsp"}:
                unsupported += 1
            continue
        relationship = node.get(qn("r:embed")) or node.get(qn("r:id"))
        part = document.part.related_parts.get(relationship)
        if part is None:
            unsupported += 1
            continue
        source_order += 1
        in_collection = False
        parent = next(node.iterancestors(qn("w:tr")), None)
        if parent is not None:
            # 仅关联最近的表格行或该表格标题，避免外层整表标签污染其他区域。
            header = parent.getparent().find(qn("w:tr"))
            row_text = "".join(item.text or "" for item in parent.iter(qn("w:t")))
            header_text = "".join(item.text or "" for item in header.iter(qn("w:t"))) if header is not None else ""
            in_collection = collection_label(row_text) or collection_label(header_text)
        try:
            candidate = candidate_from_bytes(part.blob, "原始图片")
            context = "DOCX 采集区原图 · 待框选" if in_collection else "DOCX 其他原图 · 待框选，请排除 Logo 等"
            candidates.append(SealCandidate(candidate.data, candidate.width, candidate.height,
                                            context, False, (source_order, 0, 0)))
        except Exception:
            unsupported += 1
    notice = f"找到 {len(candidates)} 张待框选原图，请框选并保存所需印章后确认填充。"
    if unsupported:
        notice += f"有 {unsupported} 处特殊对象、链接图片或不支持的编码未能提取，请核对原文件。"
    if not candidates:
        notice += "未找到可读取的内嵌图片；可将特殊对象人工另存为 PNG/JPG。"
    return candidates, notice
