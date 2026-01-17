from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Iterable


class SectionPosition(str, Enum):
    FIRST_PAGE = "first_page"
    FRONT = "front"
    BODY = "body"
    BACK = "back"
    LAST_PAGE = "last_page"


class BodyRangeRule(str, Enum):
    UNTIL_NEXT_TITLE = "until_next_title"
    UNTIL_BLANK = "until_blank"
    FIXED_PARAGRAPHS = "fixed_paragraphs"


@dataclass(frozen=True)
class SectionRule:
    key: str
    display_name: str
    title_keywords: tuple[str, ...]
    content_keywords: tuple[str, ...] = ()
    title_style_names: tuple[str, ...] = ()
    position: SectionPosition = SectionPosition.BODY
    body_range: BodyRangeRule = BodyRangeRule.UNTIL_NEXT_TITLE
    body_paragraph_limit: int | None = None

    def validate(self) -> None:
        if not self.key.strip():
            raise ValueError("section key must be non-empty")
        if not self.display_name.strip():
            raise ValueError("section display_name must be non-empty")
        if not self.title_keywords and not self.content_keywords:
            raise ValueError("section keywords must be non-empty")
        if self.body_range == BodyRangeRule.FIXED_PARAGRAPHS:
            if self.body_paragraph_limit is None or self.body_paragraph_limit <= 0:
                raise ValueError("fixed paragraph range requires positive limit")
        elif self.body_paragraph_limit is not None:
            raise ValueError("body_paragraph_limit only applies to fixed paragraph range")

    def to_dict(self) -> dict[str, object]:
        return {
            "key": self.key,
            "display_name": self.display_name,
            "title_keywords": list(self.title_keywords),
            "content_keywords": list(self.content_keywords),
            "title_style_names": list(self.title_style_names),
            "position": self.position.value,
            "body_range": self.body_range.value,
            "body_paragraph_limit": self.body_paragraph_limit,
        }


DEFAULT_SECTION_RULES: tuple[SectionRule, ...] = (
    SectionRule(
        key="original_statement",
        display_name="\u539f\u521b\u58f0\u660e",
        title_keywords=(
            "\u539f\u521b\u58f0\u660e",
            "\u539f\u521b\u6027\u58f0\u660e",
            "\u539f\u521b\u6027\u7533\u660e",
            "\u5b66\u672f\u8bda\u4fe1\u58f0\u660e",
            "\u72ec\u521b\u6027\u58f0\u660e",
            "\u5b66\u4f4d\u8bba\u6587\u539f\u521b\u6027\u58f0\u660e",
            "\u5b66\u4f4d\u8bba\u6587\u539f\u521b\u6027\u7533\u660e",
            "\u5b66\u4f4d\u8bba\u6587\u72ec\u521b\u6027\u58f0\u660e",
            "\u8bda\u4fe1\u58f0\u660e",
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
            "\u4f7f\u7528\u6388\u6743\u58f0\u660e",
            "\u5b66\u4f4d\u8bba\u6587\u7248\u6743\u4f7f\u7528\u6388\u6743\u4e66",
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
            "\u9e23\u8c22",
        ),
        position=SectionPosition.BACK,
        body_range=BodyRangeRule.UNTIL_NEXT_TITLE,
    ),
)


def iter_default_section_rules() -> Iterable[SectionRule]:
    return DEFAULT_SECTION_RULES


def validate_section_rules(rules: Iterable[SectionRule]) -> None:
    for rule in rules:
        rule.validate()


def serialize_section_rules(rules: Iterable[SectionRule]) -> list[dict[str, object]]:
    return [rule.to_dict() for rule in rules]
