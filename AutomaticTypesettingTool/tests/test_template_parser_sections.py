import tempfile
import unittest
from pathlib import Path

from src.template_parser import TemplateParser
from src.section_rules import DEFAULT_SECTION_RULES


class TemplateParserSectionTests(unittest.TestCase):
    def _build_doc(self, root: Path, paragraphs: list[str]) -> Path:
        try:
            from docx import Document
        except ImportError as exc:  # pragma: no cover - dependency required
            raise unittest.SkipTest("python-docx not installed") from exc
        path = root / "sections.docx"
        doc = Document()
        for text in paragraphs:
            doc.add_paragraph(text)
        doc.save(str(path))
        return path

    def test_section_title_and_body_detected(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = self._build_doc(
                Path(tmpdir),
                ["学位论文原创声明", "这是正文", "致谢与展望", "感谢内容"],
            )
            parser = TemplateParser(section_rules=DEFAULT_SECTION_RULES)
            result = parser.parse(str(path))
            roles = result.roles
            self.assertIn("section_original_statement_title", roles)
            self.assertIn("section_original_statement_body", roles)
            self.assertIn("section_acknowledgement_title", roles)
            self.assertIn("section_acknowledgement_body", roles)

    def test_section_body_until_next_title_allows_blank(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = self._build_doc(
                Path(tmpdir),
                ["原创声明", "", "声明正文", "致谢"],
            )
            parser = TemplateParser(section_rules=DEFAULT_SECTION_RULES)
            result = parser.parse(str(path))
            roles = result.roles
            self.assertIn("section_original_statement_body", roles)

    def test_section_body_until_blank_stops(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = self._build_doc(
                Path(tmpdir),
                ["授权声明", "", "授权内容"],
            )
            parser = TemplateParser(section_rules=DEFAULT_SECTION_RULES)
            result = parser.parse(str(path))
            roles = result.roles
            self.assertIn("section_authorization_statement_title", roles)
            self.assertNotIn("section_authorization_statement_body", roles)

    def test_section_role_links(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = self._build_doc(
                Path(tmpdir),
                ["原创声明", "正文"],
            )
            parser = TemplateParser(section_rules=DEFAULT_SECTION_RULES)
            result = parser.parse(str(path))
            links = result.role_links
            link_keys = {
                (
                    link.get("title_role"),
                    link.get("body_role"),
                    link.get("section"),
                )
                for link in links
            }
            self.assertIn(
                (
                    "section_original_statement_title",
                    "section_original_statement_body",
                    "original_statement",
                ),
                link_keys,
            )


if __name__ == "__main__":
    unittest.main()
