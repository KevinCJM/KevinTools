import unittest

from src import template_types


class TemplateTypesTests(unittest.TestCase):
    def test_default_template_types_present(self) -> None:
        types_list = list(template_types.iter_template_types())
        keys = {item.key for item in types_list}
        self.assertIn("auto", keys)
        self.assertIn("generic", keys)
        self.assertIn("school_a", keys)
        self.assertIn("school_b", keys)

    def test_resolve_template_type_unknown(self) -> None:
        resolved = template_types.resolve_template_type("unknown")
        self.assertEqual(resolved.key, "generic")

    def test_resolve_template_type_auto(self) -> None:
        resolved = template_types.resolve_template_type("auto")
        self.assertEqual(resolved.key, "auto")

    def test_template_type_choices(self) -> None:
        choices = template_types.get_template_type_choices()
        self.assertTrue(choices)
        key, label = choices[0]
        self.assertIsInstance(key, str)
        self.assertIsInstance(label, str)
        self.assertEqual(key, "auto")


if __name__ == "__main__":
    unittest.main()
