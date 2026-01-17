import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from docx import Document

from src.template_parser import TemplateParser, _W_NS


class TemplateParserFootnotesTests(unittest.TestCase):
    def test_parse_footnote_roles_and_numbering(self) -> None:
        document = Document()
        document.add_paragraph("正文")
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "footnotes.docx"
            document.save(str(path))
            _inject_footnote_parts(path)

            parser = TemplateParser()
            result = parser.parse(str(path))

        self.assertIn("footnote_text", result.roles)
        self.assertIn("footnote_reference", result.roles)
        footnote_text = result.roles["footnote_text"]
        footnote_reference = result.roles["footnote_reference"]
        self.assertEqual(footnote_text.font_name, "宋体")
        self.assertEqual(footnote_reference.font_name, "Times New Roman")
        numbering = result.meta.get("footnote_numbering", {})
        self.assertEqual(numbering.get("format"), "decimal")
        self.assertEqual(numbering.get("start"), 1)
        self.assertEqual(numbering.get("restart"), "eachPage")


def _inject_footnote_parts(path: Path) -> None:
    footnotes_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes xmlns:w="{_W_NS}">'
        f'<w:footnote w:id="1">'
        f'<w:p>'
        f'<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rStyle w:val="FootnoteReference"/>'
        f'<w:rFonts w:ascii="Times New Roman"/>'
        f'<w:sz w:val="20"/>'
        f'<w:b/>'
        f'</w:rPr>'
        f'<w:footnoteRef/>'
        f'</w:r>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:eastAsia="宋体"/>'
        f'<w:sz w:val="24"/>'
        f'</w:rPr>'
        f'<w:t>示例脚注</w:t>'
        f'</w:r>'
        f'</w:p>'
        f'</w:footnote>'
        f'</w:footnotes>'
    )
    settings_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{_W_NS}">'
        f'<w:footnotePr>'
        f'<w:numFmt w:val="decimal"/>'
        f'<w:numStart w:val="1"/>'
        f'<w:numRestart w:val="eachPage"/>'
        f'</w:footnotePr>'
        f'</w:settings>'
    )
    with ZipFile(path) as archive:
        data = {name: archive.read(name) for name in archive.namelist()}
    data["word/footnotes.xml"] = footnotes_xml.encode("utf-8")
    data["word/settings.xml"] = settings_xml.encode("utf-8")
    with ZipFile(path, "w") as archive:
        for name, content in data.items():
            archive.writestr(name, content)


if __name__ == "__main__":
    unittest.main()
