import unittest

from src.style_rule import REQUIRED_FIELDS, StyleRule, serialize_style_rules


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


class StyleRuleTests(unittest.TestCase):
    def test_validate_enums(self) -> None:
        rule = StyleRule(
            role="body_L1",
            alignment="CENTER",
            line_spacing_rule="MULTIPLE",
            line_spacing_unit="MULTIPLE",
        )
        rule.validate()

    def test_validate_invalid_alignment(self) -> None:
        rule = StyleRule(role="body_L1", alignment="MIDDLE")
        with self.assertRaises(ValueError):
            rule.validate()

    def test_validate_invalid_line_spacing_rule(self) -> None:
        rule = StyleRule(role="body_L1", line_spacing_rule="AUTO")
        with self.assertRaises(ValueError):
            rule.validate()

    def test_validate_invalid_line_spacing_unit(self) -> None:
        rule = StyleRule(role="body_L1", line_spacing_unit="PX")
        with self.assertRaises(ValueError):
            rule.validate()

    def test_to_dict_includes_null_fields(self) -> None:
        rule = StyleRule(role="body_L1")
        data = rule.to_dict(include_missing_fields=False)
        self.assertIn("font_name", data)
        self.assertIsNone(data["font_name"])
        self.assertNotIn("missing_fields", data)

    def test_to_dict_missing_fields_gate(self) -> None:
        rule = StyleRule(role="body_L1", missing_fields=["font_size_pt"])
        data = rule.to_dict(include_missing_fields=False)
        self.assertNotIn("missing_fields", data)
        data = rule.to_dict(include_missing_fields=True)
        self.assertEqual(data["missing_fields"], ["font_size_pt"])

    def test_to_dict_missing_fields_none(self) -> None:
        rule = StyleRule(role="body_L1")
        data = rule.to_dict(include_missing_fields=True)
        self.assertIn("missing_fields", data)
        self.assertIsNone(data["missing_fields"])

    def test_to_dict_font_size_name(self) -> None:
        rule = StyleRule(role="body_L1", font_size_pt=12.0)
        data = rule.to_dict(include_missing_fields=False)
        self.assertEqual(data["font_size_name"], "小四")

    def test_serialize_style_rules_missing_fields(self) -> None:
        rules = {
            "title_L1": StyleRule(role="title_L1"),
            "body_L1": _complete_rule("body_L1"),
        }
        output = serialize_style_rules(rules, allow_fallback=True, strict=False)
        self.assertNotIn("missing_fields", output["title_L1"])

        output = serialize_style_rules(rules, allow_fallback=False, strict=False)
        self.assertEqual(
            set(output["title_L1"]["missing_fields"]), set(REQUIRED_FIELDS)
        )
        self.assertIn("missing_fields", output["body_L1"])
        self.assertIsNone(output["body_L1"]["missing_fields"])

    def test_serialize_style_rules_validates(self) -> None:
        rules = {"body_L1": StyleRule(role="body_L1", alignment="MIDDLE")}
        with self.assertRaises(ValueError):
            serialize_style_rules(rules, allow_fallback=True, strict=False)

    def test_serialize_style_rules_strict_missing_fields(self) -> None:
        rules = {"body_L1": StyleRule(role="body_L1")}
        with self.assertRaises(ValueError):
            serialize_style_rules(rules, allow_fallback=False, strict=True)

    def test_serialize_style_rules_schema_v2(self) -> None:
        rules = {
            "title_L1": StyleRule(role="title_L1"),
            "body_L1": _complete_rule("body_L1"),
        }
        role_links = [
            {
                "title_role": "title_L1",
                "body_role": "body_L1",
                "level": 1,
                "section": "abstract",
                "extra": "ignored",
            },
            {"title_role": "title_L1", "body_role": "missing", "level": 1},
            {"title_role": 1, "body_role": "body_L1"},
            {"title_role": "title_L1", "body_role": "body_L1"},
        ]
        meta = {"allow_fallback": True}
        payload = serialize_style_rules(
            rules,
            allow_fallback=True,
            strict=False,
            schema_version="2.0",
            role_links=role_links,
            meta=meta,
        )
        self.assertEqual(payload["schema_version"], "2.0")
        self.assertEqual(payload["meta"], meta)
        self.assertIn("roles", payload)
        self.assertIn("title_L1", payload["roles"])
        self.assertEqual(len(payload["role_links"]), 1)
        self.assertEqual(payload["role_links"][0]["level"], 1)
        self.assertNotIn("section", payload["role_links"][0])
        self.assertNotIn("extra", payload["role_links"][0])

    def test_serialize_style_rules_requires_schema_version(self) -> None:
        rules = {"body_L1": _complete_rule("body_L1")}
        with self.assertRaises(ValueError):
            serialize_style_rules(
                rules,
                allow_fallback=True,
                strict=False,
                meta={"allow_fallback": True},
            )

    def test_serialize_style_rules_invalid_schema_version(self) -> None:
        rules = {"body_L1": _complete_rule("body_L1")}
        with self.assertRaises(ValueError):
            serialize_style_rules(
                rules,
                allow_fallback=True,
                strict=False,
                schema_version="1.0",
            )


if __name__ == "__main__":
    unittest.main()
