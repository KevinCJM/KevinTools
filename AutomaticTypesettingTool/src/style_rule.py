from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

ALIGNMENTS = {"LEFT", "CENTER", "RIGHT", "JUSTIFY"}
LINE_SPACING_RULES = {
    "SINGLE",
    "ONE_POINT_FIVE",
    "DOUBLE",
    "MULTIPLE",
    "EXACTLY",
    "AT_LEAST",
}
LINE_SPACING_UNITS = {"MULTIPLE", "PT"}
REQUIRED_FIELDS = (
    "font_name",
    "font_size_pt",
    "bold",
    "alignment",
    "line_spacing_rule",
    "line_spacing_value",
    "line_spacing_unit",
    "space_before_pt",
    "space_after_pt",
)


@dataclass
class StyleRule:
    role: str
    font_name: str | None = None
    font_name_ascii: str | None = None
    font_name_eastAsia: str | None = None
    font_name_hAnsi: str | None = None
    font_size_pt: float | None = None
    bold: bool | None = None
    alignment: str | None = None
    line_spacing_rule: str | None = None
    line_spacing_value: float | None = None
    line_spacing_unit: str | None = None
    space_before_pt: float | None = None
    space_after_pt: float | None = None
    space_before_value: float | None = None
    space_before_unit: str | None = None
    space_after_value: float | None = None
    space_after_unit: str | None = None
    indent_left_pt: float | None = None
    indent_right_pt: float | None = None
    indent_first_line_pt: float | None = None
    indent_hanging_pt: float | None = None
    missing_fields: list[str] | None = None

    def validate(self) -> None:
        _validate_enum("alignment", self.alignment, ALIGNMENTS)
        _validate_enum("line_spacing_rule", self.line_spacing_rule, LINE_SPACING_RULES)
        _validate_enum("line_spacing_unit", self.line_spacing_unit, LINE_SPACING_UNITS)

    def to_dict(
        self,
        include_missing_fields: bool,
        missing_fields: list[str] | None = None,
    ) -> dict[str, object | None]:
        font_size_name = _font_size_name_from_pt(self.font_size_pt)
        data: dict[str, object | None] = {
            "font_name": self.font_name,
            "font_name_ascii": self.font_name_ascii,
            "font_name_eastAsia": self.font_name_eastAsia,
            "font_name_hAnsi": self.font_name_hAnsi,
            "font_size_pt": self.font_size_pt,
            "font_size_name": font_size_name,
            "bold": self.bold,
            "alignment": self.alignment,
            "line_spacing_rule": self.line_spacing_rule,
            "line_spacing_value": self.line_spacing_value,
            "line_spacing_unit": self.line_spacing_unit,
            "space_before_pt": self.space_before_pt,
            "space_after_pt": self.space_after_pt,
            "space_before_value": self.space_before_value,
            "space_before_unit": self.space_before_unit,
            "space_after_value": self.space_after_value,
            "space_after_unit": self.space_after_unit,
            "indent_left_pt": self.indent_left_pt,
            "indent_right_pt": self.indent_right_pt,
            "indent_first_line_pt": self.indent_first_line_pt,
            "indent_hanging_pt": self.indent_hanging_pt,
        }
        if include_missing_fields:
            if missing_fields is None:
                missing_fields = self.missing_fields
            data["missing_fields"] = list(missing_fields) if missing_fields else None
        return data


def serialize_style_rules(
    rules: dict[str, StyleRule],
    allow_fallback: bool,
    strict: bool,
    schema_version: str | None = None,
    role_links: Iterable[dict[str, object]] | None = None,
    meta: dict[str, object] | None = None,
) -> dict[str, object]:
    roles_payload = _serialize_roles(rules, allow_fallback=allow_fallback, strict=strict)
    if schema_version is None:
        if role_links is not None or meta is not None:
            raise ValueError("schema_version is required when role_links or meta is provided")
        return roles_payload
    if schema_version != "2.0":
        raise ValueError(f"schema_version must be '2.0', got {schema_version!r}")
    payload: dict[str, object] = {
        "schema_version": schema_version,
        "roles": roles_payload,
        "role_links": _filter_role_links(role_links, roles_payload),
        "meta": dict(meta) if meta is not None else {},
    }
    return payload


def _serialize_roles(
    rules: dict[str, StyleRule],
    allow_fallback: bool,
    strict: bool,
) -> dict[str, dict[str, object | None]]:
    include_missing_fields = (not allow_fallback) and (not strict)
    output: dict[str, dict[str, object | None]] = {}
    for role, rule in rules.items():
        rule.validate()
        missing_required = _missing_required_fields(rule)
        if strict and missing_required:
            missing_text = ", ".join(missing_required)
            raise ValueError(
                f"strict mode forbids missing fields for role {role!r}: {missing_text}"
            )
        output[role] = rule.to_dict(
            include_missing_fields=include_missing_fields,
            missing_fields=missing_required,
        )
    return output


def _filter_role_links(
    role_links: Iterable[dict[str, object]] | None,
    roles_payload: dict[str, object],
) -> list[dict[str, object]]:
    if not role_links:
        return []
    filtered: list[dict[str, object]] = []
    for link in role_links:
        if not isinstance(link, dict):
            continue
        title_role = link.get("title_role")
        body_role = link.get("body_role")
        if not isinstance(title_role, str) or not isinstance(body_role, str):
            continue
        if title_role not in roles_payload or body_role not in roles_payload:
            continue
        cleaned: dict[str, object] = {
            "title_role": title_role,
            "body_role": body_role,
        }
        level = link.get("level") if "level" in link else None
        section = link.get("section") if "section" in link else None
        if level is None and section is None:
            continue
        if level is not None:
            cleaned["level"] = level
        elif section is not None:
            cleaned["section"] = section
        filtered.append(cleaned)
    return filtered


def _validate_enum(name: str, value: str | None, allowed: set[str]) -> None:
    if value is None:
        return
    if value not in allowed:
        allowed_list = ", ".join(sorted(allowed))
        raise ValueError(f"{name} must be one of {allowed_list}, got {value!r}")


def _missing_required_fields(rule: StyleRule) -> list[str]:
    missing: list[str] = []
    for field in REQUIRED_FIELDS:
        if getattr(rule, field) is None:
            missing.append(field)
    return missing


def _font_size_name_from_pt(value: float | None) -> str | None:
    if value is None:
        return None
    mapping = [
        (42.0, "初号"),
        (36.0, "小初"),
        (26.0, "一号"),
        (24.0, "小一"),
        (22.0, "二号"),
        (18.0, "小二"),
        (16.0, "三号"),
        (15.0, "小三"),
        (14.0, "四号"),
        (12.0, "小四"),
        (10.5, "五号"),
        (9.0, "小五"),
        (7.5, "六号"),
        (6.5, "小六"),
        (5.5, "七号"),
        (5.0, "八号"),
    ]
    tolerance = 0.25
    for pt, name in mapping:
        if abs(value - pt) <= tolerance:
            return name
    return None
