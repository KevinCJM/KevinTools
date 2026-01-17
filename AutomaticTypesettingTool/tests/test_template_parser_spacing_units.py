import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from lxml import etree

from src.template_parser import TemplateParser, _W_NS


class TemplateParserSpacingUnitsTests(unittest.TestCase):
    def test_parse_spacing_units_lines(self) -> None:
        document = Document()
        document.add_paragraph("测试段前段后")

        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "spacing.docx"
            document.save(str(path))
            xml_bytes = _update_paragraph_spacing_lines(path)
            _rewrite_document_xml(path, xml_bytes)

            parser = TemplateParser()
            result = parser.parse(str(path))

        body = result.roles["body_L1"]
        self.assertEqual(body.space_before_unit, "LINE")
        self.assertEqual(body.space_after_unit, "LINE")
        self.assertAlmostEqual(body.space_before_value or 0.0, 1.0)
        self.assertAlmostEqual(body.space_after_value or 0.0, 2.0)


def _update_paragraph_spacing_lines(path: Path) -> bytes:
    with ZipFile(path) as archive:
        document_bytes = archive.read("word/document.xml")
    root = etree.fromstring(document_bytes)
    paragraph = root.find(".//w:p", namespaces={"w": _W_NS})
    if paragraph is None:
        return document_bytes
    p_pr = paragraph.find(f"{{{_W_NS}}}pPr")
    if p_pr is None:
        p_pr = etree.SubElement(paragraph, f"{{{_W_NS}}}pPr")
    spacing = p_pr.find(f"{{{_W_NS}}}spacing")
    if spacing is None:
        spacing = etree.SubElement(p_pr, f"{{{_W_NS}}}spacing")
    spacing.set(f"{{{_W_NS}}}beforeLines", "100")
    spacing.set(f"{{{_W_NS}}}afterLines", "200")
    return etree.tostring(root, xml_declaration=True, encoding="utf-8")


def _rewrite_document_xml(path: Path, xml_bytes: bytes) -> None:
    with ZipFile(path) as archive:
        data = {name: archive.read(name) for name in archive.namelist()}
    data["word/document.xml"] = xml_bytes
    with ZipFile(path, "w") as archive:
        for name, content in data.items():
            archive.writestr(name, content)


if __name__ == "__main__":
    unittest.main()
