import tempfile
import unittest
from pathlib import Path

from src.style_reader import parse_styles_xml, resolve_style


STYLE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="SimSun"/>
        <w:sz w:val="24"/>
        <w:b w:val="1"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>
        <w:jc w:val="center"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"
        w:asciiTheme="minorAscii" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ThemeOnly">
    <w:name w:val="ThemeOnly"/>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorAscii" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="宋体"/>
      <w:sz w:val="32"/>
      <w:b/>
    </w:rPr>
    <w:pPr>
      <w:jc w:val="right"/>
      <w:spacing w:before="200" w:after="100" w:line="480" w:lineRule="auto"/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Weird">
    <w:name w:val="Weird"/>
    <w:rPr>
      <w:b w:val="0"/>
    </w:rPr>
    <w:pPr>
      <w:jc w:val="foo"/>
    </w:pPr>
  </w:style>
</w:styles>
"""


class StyleReaderTests(unittest.TestCase):
    def _write_styles(self, root: Path) -> Path:
        path = root / "styles.xml"
        path.write_text(STYLE_XML, encoding="utf-8")
        return path

    def test_parse_doc_defaults(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            xml_path = self._write_styles(Path(tmpdir))
            styles, defaults = parse_styles_xml(xml_path)

        self.assertIn("Normal", styles)
        self.assertEqual(defaults.fonts.eastAsia, "SimSun")
        self.assertEqual(defaults.fonts.ascii, "Times New Roman")
        self.assertEqual(defaults.font_size_pt, 12.0)
        self.assertTrue(defaults.bold)
        self.assertEqual(defaults.alignment, "CENTER")
        self.assertEqual(defaults.space_before_pt, 6.0)
        self.assertEqual(defaults.space_after_pt, 12.0)
        self.assertEqual(defaults.line_rule, "auto")
        self.assertEqual(defaults.line_twips, 360)
        self.assertEqual(styles["Normal"].fonts.ascii_theme, "minorAscii")
        self.assertEqual(styles["Normal"].fonts.hAnsi_theme, "minorHAnsi")
        self.assertEqual(styles["Normal"].fonts.eastAsia_theme, "minorEastAsia")

    def test_resolve_style_chain(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            xml_path = self._write_styles(Path(tmpdir))
            styles, defaults = parse_styles_xml(xml_path)

        resolved = resolve_style("Heading1", styles, defaults)
        self.assertEqual(resolved.font_name, "宋体")
        self.assertEqual(resolved.font_size_pt, 16.0)
        self.assertTrue(resolved.bold)
        self.assertEqual(resolved.alignment, "RIGHT")
        self.assertEqual(resolved.space_before_pt, 10.0)
        self.assertEqual(resolved.space_after_pt, 5.0)
        self.assertEqual(resolved.line_rule, "auto")
        self.assertEqual(resolved.line_twips, 480)

        normal = resolve_style("Normal", styles, defaults)
        self.assertEqual(normal.font_name, "SimSun")
        self.assertEqual(normal.font_size_pt, 12.0)
        self.assertTrue(normal.bold)
        self.assertEqual(normal.alignment, "CENTER")

    def test_alignment_and_bold_mapping(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            xml_path = self._write_styles(Path(tmpdir))
            styles, defaults = parse_styles_xml(xml_path)

        resolved = resolve_style("Weird", styles, defaults)
        self.assertEqual(resolved.alignment, "JUSTIFY")
        self.assertFalse(resolved.bold)

    def test_theme_font_resolution(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            xml_path = self._write_styles(Path(tmpdir))
            styles, defaults = parse_styles_xml(xml_path)

        theme_map = {
            "majorAscii": "ThemeAscii",
            "majorHAnsi": "ThemeHAnsi",
            "majorEastAsia": "ThemeEastAsia",
        }
        resolved = resolve_style("ThemeOnly", styles, defaults, theme_map=theme_map)
        self.assertEqual(resolved.font_name, "ThemeEastAsia")
        self.assertEqual(resolved.fonts.ascii, "ThemeAscii")


if __name__ == "__main__":
    unittest.main()
