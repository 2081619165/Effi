"""按已确认的固定结构填充本地 Word 模板。"""

from __future__ import annotations

import re
import unicodedata
from datetime import date
from io import BytesIO
from pathlib import Path
from typing import Mapping

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image as DocxImage
from docx.shared import Pt, Twips
from docx.text.paragraph import Paragraph
from docx.text.run import Run


INSTITUTION_TEMPLATE = "机构证章申请表.docx"
PERSONAL_TEMPLATE = "个人证书及电子印章申请表.docx"
SUPPORTED_TEMPLATE_NAMES = (INSTITUTION_TEMPLATE, PERSONAL_TEMPLATE)
_DATE_PATTERN = r"([0-9]{4})\s*年\s*([0-9]{1,2})\s*月\s*([0-9]{1,2})\s*日"


class TemplateStructureError(ValueError):
    """模板结构与已确认版本不一致。"""


def _compact(text: str) -> str:
    return re.sub(r"\s+", "", text)


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise TemplateStructureError(message)


def _main_table(document: Document):
    _require(len(document.tables) == 1, "模板主表格数量不是 1")
    table = document.tables[0]
    _require(len(table.rows) == 4, "模板主表格行数不是 4")
    _require(len(table.columns) == 2, "模板主表格列数不是 2")
    for row_index in (1, 2, 3):
        _require(
            table.cell(row_index, 0)._tc is table.cell(row_index, 1)._tc,
            f"模板主表格第 {row_index + 1} 行合并关系不一致",
        )
    return table


def _underline_groups(paragraph: Paragraph) -> list[list[Run]]:
    """查找连续的下划线空白运行，避免覆盖同段落内其他标签。"""
    groups: list[list[Run]] = []
    current: list[Run] = []
    for run in paragraph.runs:
        is_placeholder = bool(run.underline) and not run.text.strip()
        if is_placeholder:
            current.append(run)
        elif current:
            groups.append(current)
            current = []
    if current:
        groups.append(current)
    return groups


def _fill_underline_group(group: list[Run], value: str) -> None:
    if not value:
        return
    _require(bool(group), "未找到下划线填写位置")
    group[0].text = value
    for run in group[1:]:
        run.text = ""


def _fill_institution(document: Document, values: Mapping[str, str]) -> None:
    table = _main_table(document)
    basic_cell = table.cell(1, 0)
    agent_cell = table.cell(2, 0)
    _require(len(basic_cell.paragraphs) == 2, "单位基本信息段落数量不一致")
    _require(len(agent_cell.paragraphs) == 3, "经办人信息段落数量不一致")

    basic_line = basic_cell.paragraphs[0]
    _require(
        "单位名称（全称）" in basic_line.text
        and "统一社会信用代码" in basic_line.text,
        "单位基本信息标签不一致",
    )
    basic_groups = _underline_groups(basic_line)
    _require(len(basic_groups) == 2, "单位基本信息填写位置数量不是 2")

    address_line = basic_cell.paragraphs[1]
    _require("单位法定地址" in address_line.text, "单位法定地址标签不一致")
    address_groups = _underline_groups(address_line)
    _require(len(address_groups) == 1, "单位法定地址填写位置数量不是 1")

    agent_name_line = agent_cell.paragraphs[1]
    _require(
        "姓名" in agent_name_line.text and "身份证号码" in agent_name_line.text,
        "经办人姓名或身份证标签不一致",
    )
    agent_name_groups = _underline_groups(agent_name_line)
    _require(len(agent_name_groups) == 2, "经办人姓名行填写位置数量不是 2")

    agent_contact_line = agent_cell.paragraphs[2]
    _require(
        "手机号码" in agent_contact_line.text and "邮箱地址" in agent_contact_line.text,
        "经办人手机号或邮箱标签不一致",
    )
    agent_contact_groups = _underline_groups(agent_contact_line)
    _require(len(agent_contact_groups) == 2, "经办人联系方式填写位置数量不是 2")

    _fill_underline_group(basic_groups[0], values.get("company_name", ""))
    _fill_underline_group(basic_groups[1], values.get("credit_code", ""))
    _fill_underline_group(address_groups[0], values.get("company_address", ""))
    _fill_underline_group(agent_name_groups[0], values.get("agent_name", ""))
    _fill_underline_group(agent_name_groups[1], values.get("agent_id", ""))
    _fill_underline_group(agent_contact_groups[0], values.get("agent_mobile", ""))
    _fill_underline_group(agent_contact_groups[1], values.get("agent_email", ""))
    for paragraph, left_group, label in (
        (basic_line, basic_groups[0], "统一社会信用代码"),
        (agent_name_line, agent_name_groups[0], "身份证号码"),
        (agent_contact_line, agent_contact_groups[0], "邮箱地址"),
    ):
        start = 0
        for run in paragraph.runs:
            start += len(run.text)
            if run._r is left_group[-1]._r:
                break
        end = paragraph.text.index(label, start)
        _require(bool(paragraph.text[start:end]) and not paragraph.text[start:end].strip(),
                 "机构字段分隔位置不一致")
        left_text = paragraph.text[:start].split("\n")[-1]
        # 原字段为 10 磅宋体，按全角/半角预留安全宽度；长值另起一行。
        estimated_width = sum(10 if unicodedata.east_asian_width(char) in {"W", "F"} else 5
                              for char in left_text)
        separator = "\n" if estimated_width > 250 else "\t"
        paragraph.paragraph_format.tab_stops.add_tab_stop(Pt(270))
        _replace_run_spans(paragraph, [(start, end, separator)], "机构字段分隔")


def _fill_personal(document: Document, values: Mapping[str, str]) -> None:
    table = _main_table(document)
    for row_index, prefix, keys in (
        (1, "personal", ("legal_name", "legal_id", "personal_mobile", "agent_email")),
        (2, "agent", ("agent_name", "agent_id", "agent_mobile", "agent_email")),
    ):
        cell = table.cell(row_index, 0)
        _require(len(cell.tables) == 1, "基本信息区域缺少字段表格")
        fields = dict(zip(("name", "id", "mobile", "email"), keys))
        found = set()
        for row in cell.tables[0].rows:
            for field_cell in row.cells:
                for paragraph in field_cell.paragraphs:
                    replacements = []
                    pattern = r"\{\{" + prefix + r"_(name|id|mobile|email)\}\}"
                    for match in re.finditer(pattern, paragraph.text):
                        field = match.group(1)
                        _require(field not in found, "基本信息字段占位符重复")
                        found.add(field)
                        replacements.append((*match.span(), values.get(fields[field], "")))
                    _replace_run_spans(paragraph, replacements, "基本信息")
        _require(found == set(fields), "基本信息字段占位符缺失")


def _fill_personal_declaration(document: Document, values: Mapping[str, str]) -> None:
    cell = _main_table(document).cell(3, 0)
    paragraph = cell.paragraphs[2]
    match = re.match(
        r"本人姓名[：:](?P<name> *(?:（申请人） *)?)身份证号码[：:]"
        r"(?P<id> *(?:（申请人身份证） *)?)，现委托经办人[：:]"
        r"(?P<agent> *(?:（经办人姓名） *)?)为本次",
        paragraph.text,
    )
    _require(match is not None, "申请人委托申明填写位置不一致")
    applicant = values.get("legal_name", "")
    agent = values.get("agent_name", "")
    replacements = []
    for group, value in (
        ("name", applicant),
        ("id", values.get("legal_id", "")),
        ("agent", agent),
    ):
        start, end = match.span(group)
        _require(start < end, "申请人委托申明缺少填写留白")
        blank = match.group(group)
        if value or blank.strip():
            replacements.append((start, end, f" {value} " if value else "  "))
    _replace_run_spans(paragraph, replacements, "申请人委托申明")

    signature = cell.paragraphs[-2]
    _require(
        signature.text == "经办人签字：\t申请人签字：",
        "个人签字填写位置不一致",
    )
    replacements = []
    for match, value in zip(re.finditer("：", signature.text), (agent, applicant)):
        replacements.append((*match.span(), "：" + value))
    _replace_run_spans(signature, replacements, "个人签字")


def _replace_run_spans(
    paragraph: Paragraph, replacements: list[tuple[int, int, str]], field_name: str
) -> None:
    """按原始字符位置倒序写入，保留各 run 的字符格式。"""
    run_ranges: list[tuple[int, int, Run]] = []
    offset = 0
    for run in paragraph.runs:
        run_ranges.append((offset, offset + len(run.text), run))
        offset += len(run.text)
    _require(offset == len(paragraph.text), f"{field_name}包含不支持的文本结构")
    for start, end, value in reversed(replacements):
        targets = [
            (left, right, run) for left, right, run in run_ranges
            if left < end and right > start
        ]
        _require(bool(targets), f"{field_name}填写位置缺失")
        for index, (left, right, run) in enumerate(targets):
            local_start = max(start, left) - left
            local_end = min(end, right) - left
            run.text = (
                run.text[:local_start]
                + (value if index == 0 else "")
                + run.text[local_end:]
            )


def _date_replacements(match: re.Match[str], today: date) -> list[tuple[int, int, str]]:
    replacements = []
    for group, value in enumerate((today.year, today.month, today.day), start=1):
        start, end = match.span(group)
        replacements.append((start, end, str(value).zfill(end - start)))
    return replacements


def _fill_personal_delegation_start_date(document: Document, today: date) -> None:
    paragraph = _main_table(document).cell(3, 0).paragraphs[2]
    match = re.search(rf"委托期限[：:]\s*自\s*{_DATE_PATTERN}\s*起\s*至", paragraph.text)
    _require(match is not None, "委托期限起始日期格式不一致")
    _replace_run_spans(paragraph, _date_replacements(match, today), "委托期限起始日期")


def _fill_application_dates(
    document: Document, template_name: str, today: date
) -> None:
    """只更新委托申明签章下方的申请日期，保留原运行格式。"""
    cell = _main_table(document).cell(3, 0)
    institution = template_name == INSTITUTION_TEMPLATE
    paragraph_count = 5 if institution else 7
    heading = "申请单位委托申明" if institution else "申请人委托申明"
    _require(len(cell.paragraphs) == paragraph_count, "委托申明段落数量不一致")
    _require(_compact(cell.paragraphs[1].text) == heading, "委托申明标题不一致")
    signature = _compact(cell.paragraphs[-2].text)
    _require(
        signature == ("单位公章：" if institution else "经办人签字：申请人签字："),
        "申请日期上方签章标签不一致",
    )
    paragraph = cell.paragraphs[-1]
    date_pattern = rf"申请日期\s*[：:]\s*{_DATE_PATTERN}"
    count = 1 if institution else 2
    _require(
        re.fullmatch(rf"\s*(?:{date_pattern}\s*){{{count}}}", paragraph.text) is not None,
        "申请日期格式或数量不一致",
    )
    replacements: list[tuple[int, int, str]] = []
    for match in re.finditer(date_pattern, paragraph.text):
        replacements.extend(_date_replacements(match, today))

    _replace_run_spans(paragraph, replacements, "申请日期")


def _fill_seals(document: Document, seals: tuple[bytes, ...], heading: str) -> None:
    """通过采集标题和三格结构定位，按固定行高留出边框安全距离。"""
    _require(len(seals) <= 3, "印章最多三张")
    if not seals:
        return
    cell = _main_table(document).cell(3, 0)
    tables = [table for table in cell.tables
              if heading in _compact(table.cell(0, 0).text)]
    _require(len(tables) == 1, "未找到唯一的三样本采集表格")
    table = tables[0]
    _require(len(table.rows) == 2 and len(table.columns) == 3, "印章采集框结构不一致")
    row = table.rows[1]
    _require(len({id(target._tc) for target in row.cells}) == 3, "印章采集框存在合并")
    _require(row.height is not None, "印章采集框缺少固定高度")
    _require(row.height_rule == WD_ROW_HEIGHT_RULE.EXACTLY and not table.autofit,
             "印章采集框必须使用固定行高和列宽")
    for target, data in zip(row.cells, seals):
        _require(not target.text.strip() and not target._tc.xpath('.//w:drawing|.//w:pict'), "印章采集框不是空白")
        _require(target.width is not None, "印章采集框缺少宽度")
        image = DocxImage.from_blob(data)
        # 四周预留 12 磅，容纳单元格内边距和图片行基线。
        width = target.width - Twips(480)
        height = row.height - Twips(480)
        _require(width > 0 and height > 0, "印章采集框可用尺寸不足")
        ratio = min(width / image.px_width, height / image.px_height)
        paragraph = target.paragraphs[0]
        for extra in target.paragraphs[1:]:
            target._tc.remove(extra._p)
        paragraph.clear()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.first_line_indent = Pt(0)
        paragraph.paragraph_format.line_spacing = 1
        target.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        paragraph.add_run().add_picture(BytesIO(data), width=int(image.px_width * ratio),
                                        height=int(image.px_height * ratio))


def fill_word_template(
    template_path: Path,
    output_path: Path,
    values: Mapping[str, str],
    seals: tuple[bytes, ...] = (),
) -> None:
    """从固定模板生成新文件；不会覆盖模板本身。"""
    template_path = template_path.resolve()
    output_path = output_path.resolve()
    _require(template_path.is_file(), "未找到所选 Word 模板")
    _require(template_path.suffix.lower() == ".docx", "模板不是 DOCX 文件")
    _require(output_path.suffix.lower() == ".docx", "输出文件必须是 DOCX")
    _require(
        template_path.name in SUPPORTED_TEMPLATE_NAMES,
        "所选模板不在固定模板清单中",
    )
    protected_templates = {
        (template_path.parent / name).resolve()
        for name in SUPPORTED_TEMPLATE_NAMES
    }
    _require(output_path not in protected_templates, "输出路径不能覆盖内置模板")

    today = date.today()
    document = Document(template_path)
    cleaned_values = {
        key: str(value).strip()
        for key, value in values.items()
        if value is not None
    }
    if template_path.name == INSTITUTION_TEMPLATE:
        _fill_institution(document, cleaned_values)
    elif template_path.name == PERSONAL_TEMPLATE:
        _fill_personal(document, cleaned_values)
    else:  # pragma: no cover - 上面的固定清单校验已经拦截
        raise TemplateStructureError("无法识别模板类型")
    _fill_application_dates(document, template_path.name, today)
    if template_path.name == PERSONAL_TEMPLATE:
        _fill_personal_delegation_start_date(document, today)
        _fill_personal_declaration(document, cleaned_values)
    heading = "单位公章采集处" if template_path.name == INSTITUTION_TEMPLATE else "个人章采集处"
    _fill_seals(document, seals, heading)
    document.save(output_path)
