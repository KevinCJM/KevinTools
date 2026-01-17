import tempfile
import unittest
from pathlib import Path

from src import template_types


class TemplateTypeDetectionTests(unittest.TestCase):
    def _build_doc(self, root: Path, paragraphs: list[str]) -> Path:
        try:
            from docx import Document
        except ImportError as exc:  # pragma: no cover - dependency required
            raise unittest.SkipTest("python-docx not installed") from exc
        path = root / "tpl.docx"
        doc = Document()
        for text in paragraphs:
            doc.add_paragraph(text)
        doc.save(str(path))
        return Path(path)

    def test_detect_school_a(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = self._build_doc(Path(tmpdir), ["授权声明", "致谢"])
            resolved = template_types.detect_template_type(path)
            self.assertEqual(resolved.key, "school_a")

    def test_detect_school_b(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = self._build_doc(Path(tmpdir), ["鸣谢"])
            resolved = template_types.detect_template_type(path)
            self.assertEqual(resolved.key, "school_b")

    def test_detect_fallback_generic(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = self._build_doc(Path(tmpdir), ["前言", "摘要"])
            resolved = template_types.detect_template_type(path)
            self.assertEqual(resolved.key, "generic")


if __name__ == "__main__":
    unittest.main()
