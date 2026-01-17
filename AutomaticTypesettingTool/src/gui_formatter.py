from __future__ import annotations

import re
from typing import Iterable


def format_result_payload(payload: object) -> str:
    if not isinstance(payload, dict):
        return "解析结果格式异常"
    roles_payload = payload.get("roles") if isinstance(payload.get("roles"), dict) else None
    if roles_payload is None:
        roles_payload = payload
    if not isinstance(roles_payload, dict):
        return "解析结果格式异常"

    section_labels = _collect_section_labels(payload)

    ordered_roles: list[str] = []
    used_roles: set[str] = set()
    synthetic_roles: dict[str, dict[str, object]] = {}
    if "document_title" in roles_payload:
        ordered_roles.append("document_title")
        used_roles.add("document_title")
    if "document_title_en" in roles_payload:
        ordered_roles.append("document_title_en")
        used_roles.add("document_title_en")
    levels = _collect_levels(roles_payload)
    common_body_token = None
    if levels:
        body_levels = [level for level in levels if f"body_L{level}" in roles_payload]
        if len(body_levels) > 1 and _roles_identical(roles_payload, body_levels):
            key = f"::body_common::{','.join(str(level) for level in body_levels)}"
            common_body_token = key
            sample_role = f"body_L{body_levels[0]}"
            synthetic_roles[key] = {
                "label": _common_body_label(levels or body_levels),
                "data": roles_payload.get(sample_role, {}),
            }
    if levels:
        ordered_roles.append("::group::多级标题/正文")
        for level in levels:
            title_role = f"title_L{level}"
            body_role = f"body_L{level}"
            if title_role in roles_payload:
                ordered_roles.append(title_role)
                used_roles.add(title_role)
            if body_role in roles_payload:
                if common_body_token is not None:
                    if common_body_token not in ordered_roles:
                        ordered_roles.append(common_body_token)
                    used_roles.add(body_role)
                else:
                    ordered_roles.append(body_role)
                    used_roles.add(body_role)

    for label, roles in (
        ("摘要", ["abstract_title", "abstract_body", "abstract_en_title", "abstract_en_body", "keyword_line"]),
        ("首页", ["cover_title", "cover_info"]),
        ("目录", ["toc_title", "toc_body_L1", "toc_body_L2", "toc_body_L3", "toc_body"]),
        ("参考文献", ["reference_title", "reference_body"]),
        ("图表", ["figure_body", "figure_caption", "figure_note", "table_body", "table_caption", "table_note"]),
        ("脚注", ["footnote_reference", "footnote_text"]),
    ):
        group_roles = [role for role in roles if role in roles_payload]
        if not group_roles:
            continue
        ordered_roles.append(f"::group::{label}")
        for role in group_roles:
            ordered_roles.append(role)
            used_roles.add(role)

    remaining_roles = [role for role in roles_payload if role not in used_roles]
    if remaining_roles:
        ordered_roles.append("::group::其他角色")
        ordered_roles.extend(sorted(remaining_roles))

    lines: list[str] = []
    for role in ordered_roles:
        if role.startswith("::group::"):
            lines.append(f"【{role.split('::')[-1]}】")
            continue
        label_override = _section_role_label(role, section_labels)
        if role in synthetic_roles:
            info = synthetic_roles[role]
            data = info.get("data")
            if not isinstance(data, dict):
                continue
            lines.extend(_format_role(role, data, label_override=str(info.get("label", ""))))
            lines.append("")
            continue
        data = roles_payload.get(role)
        if not isinstance(data, dict):
            continue
        lines.extend(_format_role(role, data, label_override=label_override))
        lines.append("")
    meta_lines = _format_meta_sections(payload)
    if meta_lines:
        if lines and lines[-1] == "":
            lines.pop()
        if lines:
            lines.append("")
        lines.extend(meta_lines)
    if lines and lines[-1] == "":
        lines.pop()
    return "\n".join(lines) if lines else "未解析到可展示的结果"


def _collect_levels(roles_payload: dict[str, object]) -> list[int]:
    levels: set[int] = set()
    for role in roles_payload:
        for match in (
            re.match(r"^title_L(\d+)$", role, re.IGNORECASE),
            re.match(r"^body_L(\d+)$", role, re.IGNORECASE),
        ):
            if match:
                levels.add(int(match.group(1)))
                break
    return sorted(levels)


_SECTION_ROLE_PATTERN = re.compile(r"^section_(.+)_(title|body)$", re.IGNORECASE)


def _collect_section_labels(payload: dict[str, object]) -> dict[str, str]:
    meta = payload.get("meta")
    if not isinstance(meta, dict):
        return {}
    rules = meta.get("section_rules")
    if not isinstance(rules, list):
        return {}
    mapping: dict[str, str] = {}
    for rule in rules:
        if not isinstance(rule, dict):
            continue
        key = rule.get("key")
        display = rule.get("display_name")
        if not isinstance(key, str) or not isinstance(display, str):
            continue
        key = key.strip().lower()
        display = display.strip()
        if key and display:
            mapping[key] = display
    return mapping


def _section_role_label(role: str, section_labels: dict[str, str]) -> str | None:
    match = _SECTION_ROLE_PATTERN.match(role)
    if not match:
        return None
    key = match.group(1).strip().lower()
    label = section_labels.get(key)
    if not label:
        return None
    suffix = "标题" if match.group(2).lower() == "title" else "正文"
    return f"{label}{suffix}"


def _format_role(role: str, data: dict[str, object], label_override: str | None = None) -> list[str]:
    role_label = label_override or _role_label(role)
    lines = [f"【{role_label}】"]
    lines.append(f"字体：{_format_text(data.get('font_name'))}")
    detail_fonts = [
        ("东亚", data.get("font_name_eastAsia")),
        ("西文", data.get("font_name_hAnsi")),
        ("ASCII", data.get("font_name_ascii")),
    ]
    if any(value is not None for _, value in detail_fonts):
        detail_text = "，".join(
            f"{label}={_format_text(value, empty='-')}"
            for label, value in detail_fonts
        )
        lines.append(f"字体细分：{detail_text}")
    lines.append(
        "字号："
        + _format_font_size(
            data.get("font_size_name"),
            data.get("font_size_pt"),
        )
    )
    lines.append(f"粗体：{_format_bool(data.get('bold'))}")
    lines.append(f"对齐：{_format_alignment(data.get('alignment'))}")
    lines.append(
        "行距："
        + _format_line_spacing(
            data.get("line_spacing_rule"),
            data.get("line_spacing_value"),
            data.get("line_spacing_unit"),
        )
    )
    lines.append(
        "段前："
        + _format_spacing_with_unit(
            data.get("space_before_value"),
            data.get("space_before_unit"),
            data.get("space_before_pt"),
        )
    )
    lines.append(
        "段后："
        + _format_spacing_with_unit(
            data.get("space_after_value"),
            data.get("space_after_unit"),
            data.get("space_after_pt"),
        )
    )
    lines.append(
        "缩进："
        + _format_indent(
            data.get("indent_left_pt"),
            data.get("indent_right_pt"),
            data.get("indent_first_line_pt"),
            data.get("indent_hanging_pt"),
        )
    )
    missing_fields = data.get("missing_fields")
    if isinstance(missing_fields, list) and missing_fields:
        lines.append("缺失字段：" + ", ".join(str(item) for item in missing_fields))
    else:
        lines.append("缺失字段：无")
    return lines


def _role_label(role: str) -> str:
    match = re.match(r"^title_L(\d+)$", role, re.IGNORECASE)
    if match:
        return f"{int(match.group(1))}级标题"
    match = re.match(r"^body_L(\d+)$", role, re.IGNORECASE)
    if match:
        return f"{int(match.group(1))}级正文"
    match = re.match(r"^toc_body_L(\d+)$", role, re.IGNORECASE)
    if match:
        return f"目录正文 L{int(match.group(1))}"
    mapping = {
        "document_title": "文章标题",
        "document_title_en": "文章英文标题",
        "abstract_title": "摘要标题",
        "abstract_body": "摘要正文",
        "abstract_en_title": "英文摘要标题",
        "abstract_en_body": "英文摘要正文",
        "keyword_line": "关键词行",
        "cover_title": "首页文章标题",
        "cover_info": "首页信息",
        "toc_title": "目录标题",
        "toc_body": "目录正文",
        "reference_title": "参考文献标题",
        "reference_body": "参考文献正文",
        "figure_body": "图片样式",
        "figure_caption": "图标题",
        "figure_note": "图注释",
        "table_body": "表内容",
        "table_caption": "表标题",
        "table_note": "表注释",
        "footnote_reference": "脚注序号样式",
        "footnote_text": "脚注正文样式",
        "formula_block": "公式段落样式",
        "formula_inline": "行内公式样式",
        "formula_number": "公式编号样式",
        "superscript": "上标样式",
        "subscript": "下标样式",
    }
    return mapping.get(role, role)


def _roles_identical(roles_payload: dict[str, object], levels: list[int]) -> bool:
    if not levels:
        return False
    baseline = roles_payload.get(f"body_L{levels[0]}")
    if not isinstance(baseline, dict):
        return False
    keys = [
        "font_name",
        "font_size_pt",
        "bold",
        "alignment",
        "line_spacing_rule",
        "line_spacing_value",
        "line_spacing_unit",
        "space_before_pt",
        "space_after_pt",
        "indent_left_pt",
        "indent_right_pt",
        "indent_first_line_pt",
        "indent_hanging_pt",
    ]
    baseline_view = {key: baseline.get(key) for key in keys}
    for level in levels[1:]:
        data = roles_payload.get(f"body_L{level}")
        if not isinstance(data, dict):
            return False
        data_view = {key: data.get(key) for key in keys}
        if data_view != baseline_view:
            return False
    return True


def _common_body_label(levels: list[int]) -> str:
    if not levels:
        return "正文（通用）"
    sorted_levels = sorted(levels)
    if len(sorted_levels) >= 2 and sorted_levels == list(range(sorted_levels[0], sorted_levels[-1] + 1)):
        level_text = f"{sorted_levels[0]}-{sorted_levels[-1]}"
    else:
        level_text = ",".join(str(level) for level in sorted_levels)
    return f"正文（通用，适用于 {level_text} 级）"


def _format_text(value: object, empty: str = "未提供") -> str:
    if value is None:
        return empty
    return str(value)


def _format_number(value: object) -> str:
    if value is None:
        return "未提供"
    if isinstance(value, (int, float)):
        return f"{value:g}"
    return str(value)


def _format_font_size(name: object, pt: object) -> str:
    name_text = str(name).strip() if name else ""
    if isinstance(pt, (int, float)):
        pt_text = f"{pt:g} pt"
    else:
        pt_text = ""
    if name_text and pt_text:
        return f"{name_text} ({pt_text})"
    if name_text:
        return name_text
    if pt_text:
        return pt_text
    return "未提供"


def _format_spacing_with_unit(
    value: object,
    unit: object,
    fallback_pt: object,
) -> str:
    unit_text = str(unit).upper() if unit is not None else ""
    if unit_text == "LINE":
        return f"{_format_number(value)} 行" if value is not None else "未提供"
    if unit_text == "PT":
        if value is None:
            value = fallback_pt
        return f"{_format_number(value)} pt" if value is not None else "未提供"
    if value is not None:
        return str(value)
    if fallback_pt is not None:
        return f"{_format_number(fallback_pt)} pt"
    return "未提供"


def _format_bool(value: object) -> str:
    if value is None:
        return "未提供"
    return "是" if bool(value) else "否"


def _format_alignment(value: object) -> str:
    if value is None:
        return "未提供"
    mapping = {
        "LEFT": "左对齐",
        "CENTER": "居中",
        "RIGHT": "右对齐",
        "JUSTIFY": "两端对齐",
    }
    return mapping.get(str(value), str(value))


def _format_line_spacing(
    rule: object,
    value: object,
    unit: object,
) -> str:
    if rule is None and value is None:
        return "未提供"
    rule_text = str(rule) if rule is not None else ""
    if rule_text == "SINGLE":
        return "单倍"
    if rule_text == "ONE_POINT_FIVE":
        return "1.5 倍"
    if rule_text == "DOUBLE":
        return "2 倍"
    if rule_text == "MULTIPLE":
        if isinstance(value, (int, float)):
            return f"{value:g} 倍"
        return "多倍行距"
    if rule_text == "EXACTLY":
        if isinstance(value, (int, float)):
            return f"固定 {value:g} pt"
        return "固定行距"
    if rule_text == "AT_LEAST":
        if isinstance(value, (int, float)):
            return f"最小 {value:g} pt"
        return "最小行距"
    unit_text = str(unit) if unit is not None else ""
    value_text = _format_number(value) if value is not None else ""
    parts = [part for part in (rule_text, value_text, unit_text) if part]
    return " ".join(parts) if parts else "未提供"


def _format_indent(
    left: object,
    right: object,
    first_line: object,
    hanging: object,
) -> str:
    def _num(value: object) -> float | None:
        if isinstance(value, (int, float)):
            return float(value)
        return None

    left_pt = _num(left)
    right_pt = _num(right)
    first_pt = _num(first_line)
    hanging_pt = _num(hanging)

    parts: list[str] = []
    if first_pt is not None and first_pt > 0:
        parts.append(f"首行 {first_pt:g} pt")
    if hanging_pt is not None and hanging_pt > 0:
        parts.append(f"悬挂 {hanging_pt:g} pt")
    if left_pt is not None and left_pt > 0:
        parts.append(f"左 {left_pt:g} pt")
    if right_pt is not None and right_pt > 0:
        parts.append(f"右 {right_pt:g} pt")
    if parts:
        return "，".join(parts)
    if any(value is not None for value in (left, right, first_line, hanging)):
        return "0"
    return "未提供"


def _format_meta_sections(payload: dict[str, object]) -> list[str]:
    meta = payload.get("meta")
    if not isinstance(meta, dict):
        return []
    groups = [
        _format_page_margins(meta.get("page_margins")),
        _format_table_borders(meta.get("table_borders")),
        _format_title_spacing(meta.get("title_spacing")),
        _format_footnote_numbering(meta.get("footnote_numbering")),
        _format_header_footer(meta.get("header_footer")),
    ]
    lines: list[str] = []
    for group in groups:
        if not group:
            continue
        if lines:
            lines.append("")
        lines.extend(group)
    return lines


def _format_page_margins(value: object) -> list[str]:
    lines = ["【页面边距】"]
    sections = None
    summary = None
    if isinstance(value, dict):
        sections = value.get("sections")
        summary = value.get("summary")
    elif isinstance(value, list):
        sections = value
    if not sections:
        lines.append("未识别")
        return lines
    if summary:
        lines.append(f"摘要：{_format_summary_items(summary)}")
    for idx, section in enumerate(sections, start=1):
        if not isinstance(section, dict):
            continue
        label = _format_section_label(section, idx)
        margins = section.get("margins") if isinstance(section.get("margins"), dict) else section
        margin_text = _format_margin_values(margins if isinstance(margins, dict) else {})
        lines.append(f"{label}：{margin_text}")
    return lines


def _format_table_borders(value: object) -> list[str]:
    lines = ["【表格边框】"]
    if not value:
        lines.append("未识别")
        return lines
    summary = None
    tables = None
    if isinstance(value, dict):
        summary = value.get("summary")
        tables = value.get("tables")
    if summary:
        lines.append(f"统计：{_format_summary_items(summary)}")
    if isinstance(tables, list) and tables:
        lines.append(f"表格数量：{len(tables)}")
    if not summary and not tables:
        lines.append("未识别")
    return lines


def _format_title_spacing(value: object) -> list[str]:
    lines = ["【标题空行】"]
    if not isinstance(value, dict) or not value:
        lines.append("未识别")
        return lines
    order = ["abstract_title", "abstract_en_title", "toc_title"]
    keys = order + [key for key in sorted(value.keys()) if key not in order]
    for key in keys:
        spacing = value.get(key)
        if spacing is None:
            continue
        label = _role_label(str(key))
        lines.append(f"{label}：{_format_spacing_value(spacing)}")
    if len(lines) == 1:
        lines.append("未识别")
    return lines


def _format_footnote_numbering(value: object) -> list[str]:
    lines = ["【脚注编号】"]
    if not isinstance(value, dict) or not value:
        lines.append("未识别")
        return lines
    parts: list[str] = []
    fmt = value.get("format")
    start = value.get("start")
    restart = value.get("restart")
    if fmt is not None:
        parts.append(f"格式：{fmt}")
    if start is not None:
        parts.append(f"起始：{start}")
    if restart is not None:
        parts.append(f"重启：{restart}")
    lines.append("，".join(parts) if parts else "未识别")
    return lines


def _format_header_footer(value: object) -> list[str]:
    lines = ["【页眉页脚】"]
    if not isinstance(value, dict):
        lines.append("未识别")
        return lines
    sections = value.get("sections")
    if not isinstance(sections, list) or not sections:
        lines.append("未识别")
        return lines
    summary = value.get("summary")
    if summary:
        lines.append(f"摘要：{_format_summary_items(summary)}")
    for idx, section in enumerate(sections, start=1):
        if not isinstance(section, dict):
            continue
        label = _format_section_label(section, idx)
        headers = _format_header_footer_items(section.get("headers"), "页眉")
        footers = _format_header_footer_items(section.get("footers"), "页脚")
        parts = [item for item in (headers, footers) if item]
        if parts:
            lines.append(f"{label}：{'；'.join(parts)}")
    if len(lines) == 1:
        lines.append("未识别")
    return lines


def _format_header_footer_items(value: object, label: str) -> str:
    if not isinstance(value, list) or not value:
        return ""
    parts: list[str] = []
    for item in value:
        if not isinstance(item, dict):
            continue
        variant = item.get("type") or "default"
        style = item.get("style")
        style_text = _format_header_footer_style(style)
        parts.append(f"{label}({variant})：{style_text}")
    return "，".join(parts)


def _format_header_footer_style(style: object) -> str:
    if not isinstance(style, dict):
        return "未识别"
    parts: list[str] = []
    font_name = style.get("font_name")
    if font_name:
        parts.append(f"字体={font_name}")
    size_text = _format_font_size(
        style.get("font_size_name"),
        style.get("font_size_pt"),
    )
    if size_text != "未提供":
        parts.append(f"字号={size_text}")
    alignment = style.get("alignment")
    if alignment:
        parts.append(f"对齐={_format_alignment(alignment)}")
    return "，".join(parts) if parts else "未识别"


def _format_section_label(section: dict[str, object], idx: int) -> str:
    index = section.get("index")
    if index is None:
        index = section.get("section_index")
    label = f"分节{index}" if isinstance(index, int) else f"分节{idx}"
    logical_part = section.get("logical_part")
    if logical_part:
        label = f"{label}（{_format_logical_part(logical_part)}）"
    return label


def _format_logical_part(value: object) -> str:
    mapping = {
        "cover": "封面",
        "statement": "声明",
        "main": "正文",
        "back": "后置",
        "unknown": "未知",
    }
    return mapping.get(str(value), str(value))


def _format_margin_values(margins: dict[str, object]) -> str:
    def _read(keys: tuple[str, ...]) -> object | None:
        for key in keys:
            if key in margins:
                return margins.get(key)
        return None

    def _fmt(value: object) -> str:
        if isinstance(value, (int, float)):
            return f"{value:g} pt"
        return str(value)

    parts: list[str] = []
    for label, keys in (
        ("上", ("top", "top_pt")),
        ("下", ("bottom", "bottom_pt")),
        ("左", ("left", "left_pt")),
        ("右", ("right", "right_pt")),
        ("页眉", ("header", "header_pt")),
        ("页脚", ("footer", "footer_pt")),
        ("装订线", ("gutter", "gutter_pt")),
    ):
        value = _read(keys)
        if value is None:
            continue
        parts.append(f"{label} {_fmt(value)}")
    return "，".join(parts) if parts else "未提供"


def _format_summary_items(value: object) -> str:
    if isinstance(value, dict):
        if not value:
            return "未识别"
        items = []
        for key, count in value.items():
            label = _format_summary_key(key)
            if count is None:
                items.append(str(label))
            else:
                items.append(f"{label}={count}")
        return "，".join(items)
    if isinstance(value, list):
        return "，".join(str(item) for item in value)
    return str(value)


def _format_summary_key(key: object) -> str:
    mapping = {
        "grid": "内外全边框",
        "outer_only": "仅外边框",
        "inner_only": "仅内边框",
        "none": "无边框",
        "mixed": "混合",
    }
    return mapping.get(str(key), str(key))


def _format_spacing_value(value: object) -> str:
    if isinstance(value, dict):
        before = value.get("before")
        after = value.get("after")
        parts: list[str] = []
        if before is not None:
            parts.append(f"上 {_format_number(before)} 行")
        if after is not None:
            parts.append(f"下 {_format_number(after)} 行")
        return "，".join(parts) if parts else "未识别"
    if isinstance(value, (int, float)):
        return f"{value:g} 行"
    if value is None:
        return "未识别"
    return str(value)
