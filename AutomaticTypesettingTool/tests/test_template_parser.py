import json
import tempfile
import unittest
from pathlib import Path

from src import config
from src.style_rule import StyleRule
from src.template_parser import TemplateParser, _apply_fallbacks, _parse_theme_map, _validate_strict


class DummyParser(TemplateParser):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.seen_path: Path | None = None

    def _parse_template(self, template_path: Path) -> dict[str, StyleRule]:
        self.seen_path = template_path
        return {
            "title_L1": StyleRule(role="title_L1"),
            "body_L1": StyleRule(role="body_L1"),
        }


def _complete_rule(role: str) -> StyleRule:
    return StyleRule(
        role=role,
        font_name="Arial",
        font_size_pt=12.0,
        bold=False,
        alignment="LEFT",
        line_spacing_rule="SINGLE",
        line_spacing_value=1.0,
        line_spacing_unit="MULTIPLE",
        space_before_pt=0.0,
        space_after_pt=0.0,
    )


class DummyParserMissing(TemplateParser):
    def _parse_template(self, template_path: Path) -> dict[str, StyleRule]:
        rules = {"body_L1": _complete_rule("body_L1")}
        rules = _apply_fallbacks(
            rules,
            allow_fallback=self.allow_fallback,
            strict=self.strict,
            log_state=None,
            required_roles=self.required_roles,
            required_on_presence_map=self.required_on_presence_map,
            global_body_rule=None,
        )
        if self.strict:
            _validate_strict(
                rules,
                required_roles=self.required_roles,
                required_on_presence_map=self.required_on_presence_map,
            )
        return rules


class TemplateParserTests(unittest.TestCase):
    def test_parse_missing_file(self) -> None:
        parser = DummyParser()
        with self.assertRaises(FileNotFoundError):
            parser.parse("/tmp/not-exist.docx")

    def test_parse_resolves_default_role_map_path(self) -> None:
        parser = DummyParser()
        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "tpl.docx"
            template_path.write_text("stub", encoding="utf-8")
            parser.parse(str(template_path))
            self.assertEqual(parser.seen_path, template_path)
            self.assertEqual(
                parser._effective_role_map_path,
                template_path.parent / "role_mapping.json",
            )

    def test_parse_role_map_path_override(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "tpl.docx"
            template_path.write_text("stub", encoding="utf-8")
            override_path = Path(tmpdir) / "role_mapping.json"
            parser = DummyParser(role_map_path=str(override_path))
            parser.parse(str(template_path))
            self.assertEqual(parser._effective_role_map_path, override_path)

    def test_parse_role_map_precedence(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "tpl.docx"
            template_path.write_text("stub", encoding="utf-8")
            parser = DummyParser(role_map={"Normal": "body_L1"}, role_map_path="ignored")
            parser.parse(str(template_path))
            self.assertIsNone(parser._effective_role_map_path)

    def test_parse_missing_role_allow_fallback_false(self) -> None:
        parser = DummyParserMissing(allow_fallback=False, strict=False)
        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "tpl.docx"
            template_path.write_text("stub", encoding="utf-8")
            rules = parser.parse_roles(str(template_path))
            self.assertIn("title_L1", rules)
            self.assertIsNone(rules["title_L1"].font_name)
            self.assertEqual(rules["body_L1"].font_name, "Arial")
            out_path = Path(tmpdir) / "rules.json"
            parser.export_json(rules, output_path=str(out_path))
            with out_path.open("r", encoding="utf-8") as handle:
                data = json.load(handle)
            roles = data["roles"]
            self.assertIn("missing_fields", roles["title_L1"])
            self.assertIsNone(roles["body_L1"]["missing_fields"])

    def test_parse_missing_role_allow_fallback_true(self) -> None:
        parser = DummyParserMissing(allow_fallback=True, strict=False)
        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "tpl.docx"
            template_path.write_text("stub", encoding="utf-8")
            rules = parser.parse_roles(str(template_path))
            self.assertEqual(rules["title_L1"].font_name, "宋体")
            self.assertEqual(rules["body_L1"].font_name, "Arial")

    def test_parser_defaults_from_config(self) -> None:
        parser = TemplateParser()
        self.assertEqual(parser.max_heading_level, config.DEFAULT_MAX_HEADING_LEVEL)
        self.assertEqual(parser.required_roles, config.DEFAULT_REQUIRED_ROLES)
        self.assertEqual(
            parser.required_on_presence_map,
            config.DEFAULT_REQUIRED_ON_PRESENCE_MAP,
        )
        parser.required_roles.append("extra")
        parser.required_on_presence_map["extra"] = "value"
        self.assertNotIn("extra", config.DEFAULT_REQUIRED_ROLES)
        self.assertNotIn("extra", config.DEFAULT_REQUIRED_ON_PRESENCE_MAP)

    def test_parse_missing_role_strict(self) -> None:
        parser = DummyParserMissing(strict=True)
        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "tpl.docx"
            template_path.write_text("stub", encoding="utf-8")
            with self.assertRaises(ValueError):
                parser.parse(str(template_path))

    def test_export_json_default_path(self) -> None:
        parser = DummyParser()
        rules = {
            "title_L1": StyleRule(role="title_L1"),
            "body_L1": StyleRule(role="body_L1"),
        }
        output_path = config.DEFAULT_OUTPUT_PATH
        if output_path.exists():
            output_path.unlink()
        parser.export_json(rules, output_path=None)
        self.assertTrue(output_path.exists())
        with output_path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
        self.assertIn("roles", data)
        self.assertIn("title_L1", data["roles"])
        output_path.unlink()

    def test_export_json_custom_path(self) -> None:
        parser = DummyParser()
        rules = {
            "title_L1": StyleRule(role="title_L1"),
            "body_L1": StyleRule(role="body_L1"),
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / "nested" / "rules.json"
            parser.export_json(rules, output_path=str(out_path))
            self.assertTrue(out_path.exists())

    def test_export_json_roles_schema2(self) -> None:
        parser = DummyParser()
        rules = {
            "title_L1": StyleRule(role="title_L1"),
            "body_L1": StyleRule(role="body_L1"),
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / "rules.json"
            parser.export_json(rules, output_path=str(out_path))
            data = json.loads(out_path.read_text(encoding="utf-8"))
        self.assertEqual(data.get("schema_version"), "2.0")
        self.assertIn("roles", data)
        self.assertIn("title_L1", data["roles"])
        self.assertIn("body_L1", data["roles"])

    def test_export_json_strict_missing_role(self) -> None:
        parser = DummyParser(strict=True)
        rules = {"body_L1": StyleRule(role="body_L1")}
        with self.assertRaises(ValueError):
            parser.export_json(rules, output_path=None)

    def test_parse_theme_map(self) -> None:
        theme_xml = b"""<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="MajorLatin"/>
        <a:ea typeface="MajorEA"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="MinorLatin"/>
        <a:ea typeface="MinorEA"/>
      </a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>
"""
        theme_map = _parse_theme_map(theme_xml)
        self.assertEqual(theme_map["majorAscii"], "MajorLatin")
        self.assertEqual(theme_map["majorHAnsi"], "MajorLatin")
        self.assertEqual(theme_map["majorEastAsia"], "MajorEA")
        self.assertEqual(theme_map["minorAscii"], "MinorLatin")
        self.assertEqual(theme_map["minorHAnsi"], "MinorLatin")
        self.assertEqual(theme_map["minorEastAsia"], "MinorEA")

    def test_parse_theme_map_invalid_xml(self) -> None:
        theme_map = _parse_theme_map(b"<a:theme>")
        self.assertEqual(theme_map, {})


if __name__ == "__main__":
    unittest.main()
