from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from . import config
from .section_rules import (
    BodyRangeRule,
    DEFAULT_SECTION_RULES,
    SectionPosition,
    SectionRule,
    validate_section_rules,
)


@dataclass(frozen=True)
class TemplateType:
    key: str
    display_name: str
    section_rules: tuple[SectionRule, ...]

    def validate(self) -> None:
        if not self.key.strip():
            raise ValueError("template type key must be non-empty")
        if not self.display_name.strip():
            raise ValueError("template type display_name must be non-empty")
        validate_section_rules(self.section_rules)


DEFAULT_TEMPLATE_TYPES: tuple[TemplateType, ...] = (
    TemplateType(
        key="auto",
        display_name="\u81ea\u52a8\u8bc6\u522b",
        section_rules=DEFAULT_SECTION_RULES,
    ),
    TemplateType(
        key="generic",
        display_name="\u901a\u7528\u6a21\u677f",
        section_rules=DEFAULT_SECTION_RULES,
    ),
    TemplateType(
        key="school_a",
        display_name="\u5b66\u6821A\u6a21\u677f",
        section_rules=(
            SectionRule(
                key="original_statement",
                display_name="\u539f\u521b\u58f0\u660e",
                title_keywords=(
                    "\u539f\u521b\u58f0\u660e",
                    "\u539f\u521b\u6027\u7533\u660e",
                    "\u5b66\u672f\u8bda\u4fe1\u58f0\u660e",
                    "\u72ec\u521b\u6027\u58f0\u660e",
                    "\u5b66\u4f4d\u8bba\u6587\u539f\u521b\u6027\u7533\u660e",
                ),
                position=SectionPosition.FRONT,
                body_range=BodyRangeRule.UNTIL_NEXT_TITLE,
            ),
            SectionRule(
                key="authorization_statement",
                display_name="\u6388\u6743\u58f0\u660e",
                title_keywords=(
                    "\u6388\u6743\u58f0\u660e",
                    "\u7248\u6743\u58f0\u660e",
                ),
                position=SectionPosition.FRONT,
                body_range=BodyRangeRule.UNTIL_BLANK,
            ),
            SectionRule(
                key="acknowledgement",
                display_name="\u81f4\u8c22",
                title_keywords=(
                    "\u81f4\u8c22",
                    "\u611f\u8c22",
                ),
                position=SectionPosition.BACK,
                body_range=BodyRangeRule.UNTIL_NEXT_TITLE,
            ),
        ),
    ),
    TemplateType(
        key="school_b",
        display_name="\u5b66\u6821B\u6a21\u677f",
        section_rules=(
            SectionRule(
                key="original_statement",
                display_name="\u539f\u521b\u58f0\u660e",
                title_keywords=(
                    "\u539f\u521b\u58f0\u660e",
                    "\u539f\u521b\u6027\u7533\u660e",
                    "\u5b66\u4f4d\u8bba\u6587\u72ec\u521b\u6027\u58f0\u660e",
                ),
                position=SectionPosition.FRONT,
                body_range=BodyRangeRule.UNTIL_NEXT_TITLE,
            ),
            SectionRule(
                key="acknowledgement",
                display_name="\u81f4\u8c22",
                title_keywords=(
                    "\u81f4\u8c22",
                    "\u9e23\u8c22",
                ),
                position=SectionPosition.BACK,
                body_range=BodyRangeRule.UNTIL_NEXT_TITLE,
            ),
        ),
    ),
)

CUSTOM_TEMPLATE_SCHEMA_VERSION = "1.0"


def iter_template_types() -> Iterable[TemplateType]:
    custom = load_custom_template_types()
    custom_map = {template.key.lower(): template for template in custom}
    merged: list[TemplateType] = []
    for template in DEFAULT_TEMPLATE_TYPES:
        override = custom_map.pop(template.key.lower(), None)
        if override is not None:
            merged.append(override)
        else:
            merged.append(template)
    if custom_map:
        merged.extend(sorted(custom_map.values(), key=lambda item: item.key.lower()))
    return merged


def get_template_type_choices() -> list[tuple[str, str]]:
    return [(template.key, template.display_name) for template in iter_template_types()]


def resolve_template_type(key: str | None) -> TemplateType:
    normalized = (key or "").strip().lower()
    for template in iter_template_types():
        if template.key.lower() == normalized:
            return template
    for template in iter_template_types():
        if template.key == "generic":
            return template
    return list(iter_template_types())[0]


def detect_template_type(template_path: str | Path) -> TemplateType:
    texts = _extract_paragraph_texts(Path(template_path))
    if not texts:
        return resolve_template_type("generic")
    candidates = [
        template
        for template in iter_template_types()
        if template.key not in {"auto", "generic"}
    ]
    if not candidates:
        return resolve_template_type("generic")
    front_limit = _front_limit(len(texts))
    back_start = max(0, len(texts) - front_limit)
    scored: list[tuple[int, TemplateType]] = []
    for template in candidates:
        score = _score_template_type(template, texts, front_limit, back_start)
        scored.append((score, template))
    scored.sort(key=lambda item: (-item[0], item[1].key))
    if not scored or scored[0][0] <= 0:
        return resolve_template_type("generic")
    if len(scored) > 1 and scored[0][0] == scored[1][0]:
        return resolve_template_type("generic")
    return scored[0][1]


def load_custom_template_types() -> list[TemplateType]:
    path = config.TEMPLATE_TYPES_PATH
    if not path.exists():
        return []
    if not path.is_file():
        return []
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []
    if isinstance(raw, dict):
        items = raw.get("templates")
    else:
        items = raw
    if not isinstance(items, list):
        return []
    templates: list[TemplateType] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        template = _template_type_from_dict(item)
        if template is None:
            continue
        try:
            template.validate()
        except ValueError:
            continue
        templates.append(template)
    return templates


def save_custom_template_types(templates: Iterable[TemplateType]) -> None:
    output = {
        "schema_version": CUSTOM_TEMPLATE_SCHEMA_VERSION,
        "templates": [
            {
                "key": template.key,
                "display_name": template.display_name,
                "section_rules": [rule.to_dict() for rule in template.section_rules],
            }
            for template in templates
        ],
    }
    config.ensure_base_dirs()
    config.TEMPLATE_TYPES_PATH.write_text(
        json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def is_builtin_template_type(key: str) -> bool:
    normalized = (key or "").strip().lower()
    return any(template.key.lower() == normalized for template in DEFAULT_TEMPLATE_TYPES)


def _template_type_from_dict(data: dict[str, object]) -> TemplateType | None:
    key = data.get("key")
    display_name = data.get("display_name")
    if not isinstance(key, str) or not key.strip():
        return None
    if not isinstance(display_name, str) or not display_name.strip():
        return None
    rules_data = data.get("section_rules", [])
    if not isinstance(rules_data, list):
        rules_data = []
    rules: list[SectionRule] = []
    for rule in rules_data:
        if not isinstance(rule, dict):
            continue
        parsed = _section_rule_from_dict(rule)
        if parsed is None:
            continue
        rules.append(parsed)
    return TemplateType(
        key=key.strip(),
        display_name=display_name.strip(),
        section_rules=tuple(rules),
    )


def _section_rule_from_dict(data: dict[str, object]) -> SectionRule | None:
    key = data.get("key")
    display_name = data.get("display_name")
    title_keywords = data.get("title_keywords")
    if not isinstance(key, str) or not key.strip():
        return None
    if not isinstance(display_name, str) or not display_name.strip():
        return None
    if not isinstance(title_keywords, list):
        title_keywords = []
    keywords = tuple(str(item).strip() for item in title_keywords if str(item).strip())
    if not keywords:
        keywords = ()
    content_keywords_raw = data.get("content_keywords", [])
    if not isinstance(content_keywords_raw, list):
        content_keywords_raw = []
    content_keywords = tuple(
        str(item).strip() for item in content_keywords_raw if str(item).strip()
    )
    if not keywords and not content_keywords:
        return None
    style_names_raw = data.get("title_style_names", [])
    if not isinstance(style_names_raw, list):
        style_names_raw = []
    style_names = tuple(
        str(item).strip() for item in style_names_raw if str(item).strip()
    )
    position_raw = str(data.get("position") or SectionPosition.BODY.value)
    body_range_raw = str(data.get("body_range") or BodyRangeRule.UNTIL_NEXT_TITLE.value)
    position = SectionPosition(position_raw) if position_raw in SectionPosition._value2member_map_ else SectionPosition.BODY
    body_range = BodyRangeRule(body_range_raw) if body_range_raw in BodyRangeRule._value2member_map_ else BodyRangeRule.UNTIL_NEXT_TITLE
    body_paragraph_limit = data.get("body_paragraph_limit")
    limit = None
    if isinstance(body_paragraph_limit, int):
        limit = body_paragraph_limit
    elif isinstance(body_paragraph_limit, str):
        try:
            limit = int(body_paragraph_limit)
        except ValueError:
            limit = None
    try:
        rule = SectionRule(
            key=key.strip(),
            display_name=display_name.strip(),
            title_keywords=keywords,
            content_keywords=content_keywords,
            title_style_names=style_names,
            position=position,
            body_range=body_range,
            body_paragraph_limit=limit,
        )
        rule.validate()
        return rule
    except ValueError:
        return None


def _front_limit(total: int) -> int:
    if total <= 0:
        return 0
    base = max(5, total // 3)
    return min(30, base)


def _score_template_type(
    template: TemplateType,
    texts: list[str],
    front_limit: int,
    back_start: int,
) -> int:
    matched: set[str] = set()
    total = len(texts)
    for rule in template.section_rules:
        for index, text in enumerate(texts):
            if not _position_matches(rule.position, index, total, front_limit, back_start):
                continue
            if _contains_any(text, rule.title_keywords + rule.content_keywords):
                matched.add(rule.key)
                break
    return len(matched)


def _position_matches(
    position: SectionPosition,
    index: int,
    total: int,
    front_limit: int,
    back_start: int,
) -> bool:
    if total <= 0:
        return False
    if position == SectionPosition.FRONT:
        return index < front_limit
    if position == SectionPosition.FIRST_PAGE:
        return index < front_limit
    if position == SectionPosition.BACK:
        return index >= back_start
    if position == SectionPosition.LAST_PAGE:
        return index >= back_start
    return True


def _contains_any(text: str, keywords: tuple[str, ...]) -> bool:
    lowered = text.lower()
    return any(keyword.lower() in lowered for keyword in keywords if keyword)


def _extract_paragraph_texts(template_path: Path) -> list[str]:
    try:
        from docx import Document
    except ImportError:
        return []
    try:
        document = Document(str(template_path))
    except Exception:
        return []
    texts: list[str] = []
    for paragraph in _iter_paragraphs(document):
        text = getattr(paragraph, "text", "") or ""
        text = text.strip()
        if text:
            texts.append(text)
    return texts


def _iter_paragraphs(document: object) -> Iterable[object]:
    for paragraph in getattr(document, "paragraphs", []):
        yield paragraph
    for table in getattr(document, "tables", []):
        yield from _iter_table_paragraphs(table)


def _iter_table_paragraphs(table: object) -> Iterable[object]:
    for row in getattr(table, "rows", []):
        for cell in getattr(row, "cells", []):
            for paragraph in getattr(cell, "paragraphs", []):
                yield paragraph
            for nested in getattr(cell, "tables", []):
                yield from _iter_table_paragraphs(nested)
