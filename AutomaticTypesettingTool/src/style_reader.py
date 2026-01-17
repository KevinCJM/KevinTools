from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


@dataclass
class FontSpec:
    ascii: str | None = None
    hAnsi: str | None = None
    eastAsia: str | None = None
    ascii_theme: str | None = None
    hAnsi_theme: str | None = None
    eastAsia_theme: str | None = None

    def preferred_name(self) -> str | None:
        return self.eastAsia or self.hAnsi or self.ascii

    def apply_theme(self, theme_map: dict[str, str] | None) -> "FontSpec":
        if not theme_map:
            return self
        return FontSpec(
            ascii=self.ascii or _resolve_theme_font(theme_map, self.ascii_theme),
            hAnsi=self.hAnsi or _resolve_theme_font(theme_map, self.hAnsi_theme),
            eastAsia=self.eastAsia or _resolve_theme_font(theme_map, self.eastAsia_theme),
            ascii_theme=self.ascii_theme,
            hAnsi_theme=self.hAnsi_theme,
            eastAsia_theme=self.eastAsia_theme,
        )


@dataclass
class StyleDefaults:
    fonts: FontSpec
    font_size_pt: float | None
    bold: bool | None
    alignment: str | None
    space_before_pt: float | None
    space_after_pt: float | None
    line_rule: str | None
    line_twips: int | None
    outline_level: int | None


@dataclass
class StyleDefinition:
    style_id: str
    name: str | None
    based_on: str | None
    fonts: FontSpec
    font_size_pt: float | None
    bold: bool | None
    alignment: str | None
    space_before_pt: float | None
    space_after_pt: float | None
    line_rule: str | None
    line_twips: int | None
    outline_level: int | None


@dataclass
class ResolvedStyle:
    style_id: str
    name: str | None
    fonts: FontSpec
    font_name: str | None
    font_size_pt: float | None
    bold: bool | None
    alignment: str | None
    space_before_pt: float | None
    space_after_pt: float | None
    line_rule: str | None
    line_twips: int | None
    outline_level: int | None


def parse_styles_xml(
    xml_path: Path,
) -> tuple[dict[str, StyleDefinition], StyleDefaults]:
    tree = etree.parse(str(xml_path))
    root = tree.getroot()
    defaults = _parse_doc_defaults(root)
    styles: dict[str, StyleDefinition] = {}
    for style in root.findall("w:style", namespaces=NS):
        style_type = style.get(_attr_name("type"))
        if style_type == "paragraph":
            pass
        elif style_type == "character":
            pass
        elif style_type is None:
            if style.find("w:pPr", namespaces=NS) is None and style.find("w:rPr", namespaces=NS) is None:
                continue
        else:
            continue
        style_id = style.get(_attr_name("styleId"))
        if not style_id:
            continue
        name_elem = style.find("w:name", namespaces=NS)
        name = _attr(name_elem, "val")
        based_on_elem = style.find("w:basedOn", namespaces=NS)
        based_on = _attr(based_on_elem, "val")
        fonts, font_size_pt, bold = _parse_run_properties(style.find("w:rPr", NS))
        (
            alignment,
            space_before_pt,
            space_after_pt,
            line_rule,
            line_twips,
            outline_level,
        ) = _parse_paragraph_properties(style.find("w:pPr", NS))
        styles[style_id] = StyleDefinition(
            style_id=style_id,
            name=name,
            based_on=based_on,
            fonts=fonts,
            font_size_pt=font_size_pt,
            bold=bold,
            alignment=alignment,
            space_before_pt=space_before_pt,
            space_after_pt=space_after_pt,
            line_rule=line_rule,
            line_twips=line_twips,
            outline_level=outline_level,
        )
    return styles, defaults


def resolve_style(
    style_id: str,
    styles: dict[str, StyleDefinition],
    defaults: StyleDefaults,
    theme_map: dict[str, str] | None = None,
) -> ResolvedStyle:
    if style_id not in styles:
        raise KeyError(f"unknown style_id: {style_id}")
    chain = list(_collect_style_chain(styles, style_id))
    fonts = FontSpec()
    font_size_pt = defaults.font_size_pt
    bold = defaults.bold
    alignment = defaults.alignment
    space_before_pt = defaults.space_before_pt
    space_after_pt = defaults.space_after_pt
    line_rule = defaults.line_rule
    line_twips = defaults.line_twips
    outline_level = defaults.outline_level
    for style in chain:
        fonts = _merge_fonts(fonts, style.fonts)
        font_size_pt = style.font_size_pt if style.font_size_pt is not None else font_size_pt
        bold = style.bold if style.bold is not None else bold
        alignment = style.alignment if style.alignment is not None else alignment
        if style.space_before_pt is not None:
            space_before_pt = style.space_before_pt
        if style.space_after_pt is not None:
            space_after_pt = style.space_after_pt
        line_rule = style.line_rule if style.line_rule is not None else line_rule
        line_twips = style.line_twips if style.line_twips is not None else line_twips
        if style.outline_level is not None:
            outline_level = style.outline_level
    resolved_fonts = fonts.apply_theme(theme_map)
    default_fonts = defaults.fonts.apply_theme(theme_map)
    resolved_fonts = _merge_fonts(default_fonts, resolved_fonts)
    target = styles[style_id]
    return ResolvedStyle(
        style_id=style_id,
        name=target.name,
        fonts=resolved_fonts,
        font_name=resolved_fonts.preferred_name(),
        font_size_pt=font_size_pt,
        bold=bold,
        alignment=alignment,
        space_before_pt=space_before_pt,
        space_after_pt=space_after_pt,
        line_rule=line_rule,
        line_twips=line_twips,
        outline_level=outline_level,
    )


def _collect_style_chain(
    styles: dict[str, StyleDefinition],
    style_id: str,
) -> Iterable[StyleDefinition]:
    visited: set[str] = set()
    chain: list[StyleDefinition] = []
    current_id: str | None = style_id
    while current_id is not None:
        if current_id in visited:
            break
        visited.add(current_id)
        current = styles.get(current_id)
        if current is None:
            break
        chain.append(current)
        current_id = current.based_on
    chain.reverse()
    return chain


def _parse_doc_defaults(root: etree._Element) -> StyleDefaults:
    r_pr_default = root.find("w:docDefaults/w:rPrDefault/w:rPr", namespaces=NS)
    p_pr_default = root.find("w:docDefaults/w:pPrDefault/w:pPr", namespaces=NS)
    fonts, font_size_pt, bold = _parse_run_properties(r_pr_default)
    (
        alignment,
        space_before_pt,
        space_after_pt,
        line_rule,
        line_twips,
        outline_level,
    ) = _parse_paragraph_properties(p_pr_default)
    return StyleDefaults(
        fonts=fonts,
        font_size_pt=font_size_pt,
        bold=bold,
        alignment=alignment,
        space_before_pt=space_before_pt,
        space_after_pt=space_after_pt,
        line_rule=line_rule,
        line_twips=line_twips,
        outline_level=outline_level,
    )


def _parse_run_properties(
    r_pr: etree._Element | None,
) -> tuple[FontSpec, float | None, bool | None]:
    fonts = FontSpec()
    if r_pr is None:
        return fonts, None, None
    r_fonts = r_pr.find("w:rFonts", namespaces=NS)
    if r_fonts is not None:
        fonts.ascii = _attr(r_fonts, "ascii")
        fonts.hAnsi = _attr(r_fonts, "hAnsi")
        fonts.eastAsia = _attr(r_fonts, "eastAsia")
        fonts.ascii_theme = _attr(r_fonts, "asciiTheme")
        fonts.hAnsi_theme = _attr(r_fonts, "hAnsiTheme")
        fonts.eastAsia_theme = _attr(r_fonts, "eastAsiaTheme")
    font_size_pt = None
    sz_elem = r_pr.find("w:sz", namespaces=NS)
    sz_val = _attr(sz_elem, "val")
    if sz_val:
        try:
            font_size_pt = float(sz_val) / 2
        except ValueError:
            font_size_pt = None
    bold = None
    bold_elem = r_pr.find("w:b", namespaces=NS)
    if bold_elem is not None:
        bold = _parse_on_off(bold_elem)
    return fonts, font_size_pt, bold


def _parse_paragraph_properties(
    p_pr: etree._Element | None,
) -> tuple[str | None, float | None, float | None, str | None, int | None, int | None]:
    if p_pr is None:
        return None, None, None, None, None, None
    alignment = None
    jc_elem = p_pr.find("w:jc", namespaces=NS)
    if jc_elem is not None:
        alignment = _map_alignment(_attr(jc_elem, "val"))
    space_before_pt = None
    space_after_pt = None
    line_rule = None
    line_twips = None
    spacing_elem = p_pr.find("w:spacing", namespaces=NS)
    if spacing_elem is not None:
        before_val = _attr(spacing_elem, "before")
        after_val = _attr(spacing_elem, "after")
        line_val = _attr(spacing_elem, "line")
        line_rule = _attr(spacing_elem, "lineRule")
        space_before_pt = _twips_to_pt(before_val)
        space_after_pt = _twips_to_pt(after_val)
        line_twips = _parse_int(line_val)
    before_elem = p_pr.find("w:before", namespaces=NS)
    if space_before_pt is None and before_elem is not None:
        space_before_pt = _twips_to_pt(_attr(before_elem, "val"))
    after_elem = p_pr.find("w:after", namespaces=NS)
    if space_after_pt is None and after_elem is not None:
        space_after_pt = _twips_to_pt(_attr(after_elem, "val"))
    outline_level = None
    outline_elem = p_pr.find("w:outlineLvl", namespaces=NS)
    if outline_elem is not None:
        outline_level = _parse_int(_attr(outline_elem, "val"))
    return alignment, space_before_pt, space_after_pt, line_rule, line_twips, outline_level


def _attr(elem: etree._Element | None, name: str) -> str | None:
    if elem is None:
        return None
    return elem.get(_attr_name(name))


def _attr_name(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _parse_int(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _twips_to_pt(value: str | None) -> float | None:
    numeric = _parse_int(value)
    if numeric is None:
        return None
    return numeric / 20.0


def _parse_on_off(elem: etree._Element) -> bool:
    val = _attr(elem, "val")
    if val is None:
        return True
    val_lower = val.lower()
    if val_lower in {"0", "false", "off"}:
        return False
    return True


def _map_alignment(val: str | None) -> str | None:
    if val is None:
        return None
    val_lower = val.lower()
    if val_lower in {"left", "start"}:
        return "LEFT"
    if val_lower in {"center", "centercontinuous"}:
        return "CENTER"
    if val_lower in {"right", "end"}:
        return "RIGHT"
    if val_lower in {"both", "distribute", "distributed", "justify"}:
        return "JUSTIFY"
    return "JUSTIFY"


def _merge_fonts(base: FontSpec, override: FontSpec) -> FontSpec:
    return FontSpec(
        ascii=override.ascii or base.ascii,
        hAnsi=override.hAnsi or base.hAnsi,
        eastAsia=override.eastAsia or base.eastAsia,
        ascii_theme=override.ascii_theme or base.ascii_theme,
        hAnsi_theme=override.hAnsi_theme or base.hAnsi_theme,
        eastAsia_theme=override.eastAsia_theme or base.eastAsia_theme,
    )


def _resolve_theme_font(theme_map: dict[str, str], token: str | None) -> str | None:
    if token is None:
        return None
    return theme_map.get(token)
