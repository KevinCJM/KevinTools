import unittest
from datetime import datetime
from pathlib import Path

from src import config
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt

from src.template_parser import (
    ParseLogState,
    TemplateParser,
    _collect_paragraph_samples,
    _extract_paragraph_line_spacing,
)


FIXTURES_DIR = Path(__file__).resolve().parents[1] / "tests" / "fixtures"


def _fixture(name: str) -> Path:
    return FIXTURES_DIR / name


def _latest_log_text() -> str:
    logs = sorted(config.LOG_DIR.glob("template_parser_*.log"))
    if not logs:
        return ""
    return logs[-1].read_text(encoding="utf-8")


class TemplateParserFixtureTests(unittest.TestCase):
    def setUp(self) -> None:
        config.ensure_base_dirs()
        for log_file in config.LOG_DIR.glob("template_parser_*.log"):
            log_file.unlink()

    def test_parse_docdefaults(self) -> None:
        parser = TemplateParser(allow_fallback=False)
        rules = parser.parse_roles(str(_fixture("TPL_DOCDEFAULTS.docx")))
        body = rules["body_L1"]
        self.assertEqual(body.font_name_eastAsia, "宋体")
        self.assertEqual(body.font_size_pt, 12.0)
        self.assertEqual(body.alignment, "CENTER")
        self.assertEqual(body.space_before_pt, 6.0)
        self.assertEqual(body.space_after_pt, 12.0)
        self.assertEqual(body.line_spacing_rule, "MULTIPLE")
        self.assertAlmostEqual(body.line_spacing_value or 0, 1.5, places=2)

    def test_parse_no_sample(self) -> None:
        parser = TemplateParser(allow_fallback=False)
        rules = parser.parse_roles(str(_fixture("TPL_NO_SAMPLE.docx")))
        self.assertIn("body_L1", rules)
        self.assertIn("title_L1", rules)

    def test_line_spacing_inference_warning(self) -> None:
        class DummyFormat:
            line_spacing_rule = None
            line_spacing = Pt(16)

        class DummyElement:
            pPr = None

        class DummyParagraph:
            paragraph_format = DummyFormat()
            _element = DummyElement()

        log_state = ParseLogState(
            template_path=_fixture("TPL_BASIC.docx"),
            start_time=datetime.now(),
        )
        spacing = _extract_paragraph_line_spacing(
            DummyParagraph(),
            WD_LINE_SPACING,
            log_state,
            "DummyStyle",
            1,
        )
        self.assertEqual(spacing[0], "EXACTLY")
        self.assertTrue(log_state.warnings)
        self.assertEqual(log_state.warnings[0].rule, "line_spacing")

    def test_parse_role_map_ignores_bare_style_id(self) -> None:
        parser = TemplateParser(
            role_map={"Heading1": "body_L1"},
            allow_fallback=True,
        )
        parser.parse_roles(str(_fixture("TPL_BASIC.docx")))
        log_text = _latest_log_text()
        self.assertIn("rule=role_map", log_text)

    def test_parse_role_map_allows_special_role(self) -> None:
        parser = TemplateParser(
            role_map={"Normal": "abstract_title"},
            allow_fallback=False,
        )
        rules = parser.parse_roles(str(_fixture("TPL_BASIC.docx")))
        self.assertIn("abstract_title", rules)

    def test_parse_special_role_detection(self) -> None:
        parser = TemplateParser(allow_fallback=False)
        rules = parser.parse_roles(str(_fixture("TPL_SPECIAL.docx")))
        self.assertIn("abstract_title", rules)
        self.assertIn("reference_title", rules)
        self.assertIn("figure_caption", rules)
        self.assertIn("table_caption", rules)
        self.assertNotIn("figure_note", rules)
        self.assertNotIn("table_note", rules)

    def test_parse_body_stack_and_special_bodies(self) -> None:
        parser = TemplateParser(allow_fallback=False)
        rules = parser.parse_roles(str(_fixture("TPL_BODY_STACK.docx")))
        self.assertIn("body_L1", rules)
        self.assertIn("body_L2", rules)
        self.assertIn("abstract_body", rules)
        self.assertIn("reference_body", rules)
        self.assertIn("figure_note", rules)
        self.assertIn("table_note", rules)

    def test_parse_log_contains_role_sources(self) -> None:
        parser = TemplateParser(allow_fallback=True)
        parser.parse_roles(str(_fixture("TPL_BASIC.docx")))
        log_text = _latest_log_text()
        self.assertIn("template_path:", log_text)
        self.assertIn("styles_count:", log_text)
        self.assertIn("role_source[title_L1]:", log_text)
        self.assertIn("role_source[body_L1]:", log_text)
        self.assertIn("role_candidate[title_L1]:", log_text)
        self.assertIn("warnings_count:", log_text)

    def test_collect_paragraph_samples(self) -> None:
        stats = _collect_paragraph_samples(
            _fixture("TPL_ALIGN.docx"),
            log_state=None,
        )
        self.assertTrue(stats)


if __name__ == "__main__":
    unittest.main()
