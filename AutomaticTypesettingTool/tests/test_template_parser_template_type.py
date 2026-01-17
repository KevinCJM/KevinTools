import unittest
import tempfile
from pathlib import Path

from src.template_parser import TemplateParser
from src import template_types


class TemplateParserTemplateTypeTests(unittest.TestCase):
    def test_template_type_meta(self) -> None:
        fixture = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "TPL_BASIC.docx"
        template = template_types.resolve_template_type("school_a")
        parser = TemplateParser(
            template_type=template.key,
            section_rules=template.section_rules,
        )
        result = parser.parse(str(fixture))
        meta = result.meta
        self.assertEqual(meta["template_type"], "school_a")
        self.assertIsInstance(meta["section_rules"], list)
        self.assertTrue(meta["section_rules"])
        self.assertIn("key", meta["section_rules"][0])
        self.assertIn("body_range", meta["section_rules"][0])

    def test_template_type_auto_detection(self) -> None:
        try:
            from docx import Document
        except ImportError as exc:  # pragma: no cover - dependency required
            raise unittest.SkipTest("python-docx not installed") from exc
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "auto.docx"
            doc = Document()
            doc.add_paragraph("鸣谢")
            doc.save(str(path))
            parser = TemplateParser(template_type="auto")
            result = parser.parse(str(path))
            self.assertEqual(result.meta["template_type"], "school_b")


if __name__ == "__main__":
    unittest.main()
