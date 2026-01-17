import json
import tempfile
import unittest
from pathlib import Path

from src.template_parser import TemplateParser


class TemplateParserSchemaV2Tests(unittest.TestCase):
    def test_export_json_schema_v2(self) -> None:
        fixture = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "TPL_BASIC.docx"
        parser = TemplateParser()
        result = parser.parse(str(fixture))
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / "rules.json"
            parser.export_json(result, output_path=str(out_path))
            payload = json.loads(out_path.read_text(encoding="utf-8"))
        self.assertEqual(payload.get("schema_version"), "2.0")
        self.assertIn("roles", payload)
        self.assertIn("meta", payload)
        meta = payload["meta"]
        self.assertEqual(meta["max_heading_level"], parser.max_heading_level)
        self.assertEqual(meta["required_roles"], parser.required_roles)
        self.assertEqual(meta["required_on_presence_map"], parser.required_on_presence_map)
        self.assertEqual(meta["allow_fallback"], parser.allow_fallback)
        self.assertEqual(meta["strict"], parser.strict)
        self.assertIn(1, meta["detected_heading_levels"])
        self.assertEqual(meta["detected_heading_levels_overflow"], [])
        self.assertIn("page_margins", meta)
        self.assertIn("table_borders", meta)
        self.assertIn("title_spacing", meta)
        self.assertIn("footnote_numbering", meta)
        self.assertIn("toc_levels", meta)

    def test_detect_heading_levels_from_styles_xml(self) -> None:
        fixture = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "TPL_HEADINGS_1_4.docx"
        parser = TemplateParser(max_heading_level=4)
        result = parser.parse(str(fixture))
        self.assertEqual(result.meta["detected_heading_levels"], [1, 2, 3, 4])

    def test_export_json_optional_roles_not_output(self) -> None:
        fixture = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "TPL_BASIC.docx"
        parser = TemplateParser()
        result = parser.parse(str(fixture))
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / "rules.json"
            parser.export_json(result, output_path=str(out_path))
            payload = json.loads(out_path.read_text(encoding="utf-8"))
        roles = payload.get("roles", {})
        self.assertNotIn("abstract_title", roles)
        self.assertNotIn("abstract_body", roles)
        self.assertNotIn("reference_title", roles)
        self.assertNotIn("reference_body", roles)

    def test_export_json_role_links_missing_optional(self) -> None:
        fixture = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "TPL_BASIC.docx"
        parser = TemplateParser()
        result = parser.parse(str(fixture))
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / "rules.json"
            parser.export_json(result, output_path=str(out_path))
            payload = json.loads(out_path.read_text(encoding="utf-8"))
        links = payload.get("role_links", [])
        self.assertEqual(len(links), 1)
        self.assertEqual(links[0].get("level"), 1)
        self.assertNotIn("section", links[0])

    def test_explicit_mapping_overflow_role_output(self) -> None:
        fixture = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "TPL_BASIC.docx"
        parser = TemplateParser(
            role_map={"Heading 1": "title_L7"},
            max_heading_level=3,
        )
        result = parser.parse(str(fixture))
        self.assertIn("title_L7", result.roles)
        self.assertIn(7, result.meta["detected_heading_levels_overflow"])

    def test_export_json_role_links(self) -> None:
        fixture = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "TPL_BODY_STACK.docx"
        parser = TemplateParser()
        result = parser.parse(str(fixture))
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / "rules.json"
            parser.export_json(result, output_path=str(out_path))
            payload = json.loads(out_path.read_text(encoding="utf-8"))
        links = payload.get("role_links", [])
        link_keys = {
            (
                link.get("title_role"),
                link.get("body_role"),
                link.get("level"),
                link.get("section"),
            )
            for link in links
        }
        self.assertIn(("title_L1", "body_L1", 1, None), link_keys)
        self.assertIn(("title_L2", "body_L2", 2, None), link_keys)
        self.assertIn(("abstract_title", "abstract_body", None, "abstract"), link_keys)
        self.assertIn(("reference_title", "reference_body", None, "reference"), link_keys)
        self.assertIn(("figure_caption", "figure_note", None, "figure"), link_keys)
        self.assertIn(("table_caption", "table_note", None, "table"), link_keys)


if __name__ == "__main__":
    unittest.main()
