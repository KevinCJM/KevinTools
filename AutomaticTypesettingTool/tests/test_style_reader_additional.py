import tempfile
import unittest
from pathlib import Path

from lxml import etree

from src.style_reader import (
    FontSpec,
    StyleDefaults,
    StyleDefinition,
    _collect_style_chain,
    _map_alignment,
    _parse_int,
    parse_styles_xml,
    resolve_style,
)


STYLE_XML_EXTRA = """<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:styleId="NoTypeNoPpr">
    <w:name w:val="NoTypeNoPpr"/>
  </w:style>
  <w:style w:styleId="NoTypeWithPpr">
    <w:name w:val="NoTypeWithPpr"/>
    <w:pPr>
      <w:jc w:val="left"/>
    </w:pPr>
  </w:style>
  <w:style w:type="character" w:styleId="CharStyle">
    <w:name w:val="CharStyle"/>
  </w:style>
  <w:style w:type="paragraph">
    <w:name w:val="MissingId"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="BadSize">
    <w:name w:val="BadSize"/>
    <w:rPr>
      <w:sz w:val="bad"/>
      <w:b w:val="false"/>
    </w:rPr>
    <w:pPr>
      <w:before w:val="120"/>
      <w:after w:val="240"/>
    </w:pPr>
  </w:style>
</w:styles>
"""


class StyleReaderAdditionalTests(unittest.TestCase):
    def _write_styles(self, root: Path, xml_text: str) -> Path:
        path = root / "styles.xml"
        path.write_text(xml_text, encoding="utf-8")
        return path

    def test_parse_styles_skips_invalid_entries(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            xml_path = self._write_styles(Path(tmpdir), STYLE_XML_EXTRA)
            styles, _ = parse_styles_xml(xml_path)

        self.assertNotIn("NoTypeNoPpr", styles)
        self.assertIn("NoTypeWithPpr", styles)
        self.assertNotIn("CharStyle", styles)
        self.assertNotIn("MissingId", styles)

    def test_parse_styles_invalid_size_and_spacing(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            xml_path = self._write_styles(Path(tmpdir), STYLE_XML_EXTRA)
            styles, _ = parse_styles_xml(xml_path)

        bad = styles["BadSize"]
        self.assertIsNone(bad.font_size_pt)
        self.assertFalse(bad.bold)
        self.assertEqual(bad.space_before_pt, 6.0)
        self.assertEqual(bad.space_after_pt, 12.0)

    def test_collect_style_chain_breaks_on_cycle_and_missing(self) -> None:
        fonts = FontSpec()
        styles = {
            "A": StyleDefinition(
                style_id="A",
                name="A",
                based_on="B",
                fonts=fonts,
                font_size_pt=None,
                bold=None,
                alignment=None,
                space_before_pt=None,
                space_after_pt=None,
                line_rule=None,
                line_twips=None,
                outline_level=None,
            ),
            "B": StyleDefinition(
                style_id="B",
                name="B",
                based_on="A",
                fonts=fonts,
                font_size_pt=None,
                bold=None,
                alignment=None,
                space_before_pt=None,
                space_after_pt=None,
                line_rule=None,
                line_twips=None,
                outline_level=None,
            ),
        }
        chain = list(_collect_style_chain(styles, "A"))
        self.assertTrue(chain)
        missing_chain = list(_collect_style_chain({}, "Missing"))
        self.assertEqual(missing_chain, [])

    def test_resolve_style_unknown(self) -> None:
        defaults = StyleDefaults(
            fonts=FontSpec(),
            font_size_pt=None,
            bold=None,
            alignment=None,
            space_before_pt=None,
            space_after_pt=None,
            line_rule=None,
            line_twips=None,
            outline_level=None,
        )
        with self.assertRaises(KeyError):
            resolve_style("Missing", {}, defaults)

    def test_parse_int_and_alignment_helpers(self) -> None:
        self.assertIsNone(_parse_int("x"))
        self.assertIsNone(_map_alignment(None))
        self.assertEqual(_map_alignment("left"), "LEFT")
        self.assertEqual(_map_alignment("both"), "JUSTIFY")
        self.assertEqual(_map_alignment("start"), "LEFT")
        self.assertEqual(_map_alignment("end"), "RIGHT")
        self.assertEqual(_map_alignment("centerContinuous"), "CENTER")
        self.assertEqual(_map_alignment("distribute"), "JUSTIFY")
        self.assertEqual(_map_alignment("distributed"), "JUSTIFY")


if __name__ == "__main__":
    unittest.main()
