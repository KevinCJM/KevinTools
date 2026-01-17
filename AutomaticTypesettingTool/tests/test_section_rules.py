import unittest

from src.section_rules import (
    BodyRangeRule,
    SectionPosition,
    validate_section_rules,
    iter_default_section_rules,
)


class SectionRulesTests(unittest.TestCase):
    def test_default_rules_contains_required_sections(self) -> None:
        rules = list(iter_default_section_rules())
        keys = {rule.key for rule in rules}
        self.assertIn("original_statement", keys)
        self.assertIn("acknowledgement", keys)

    def test_default_rules_have_keywords_and_names(self) -> None:
        rules = list(iter_default_section_rules())
        for rule in rules:
            self.assertTrue(rule.display_name)
            self.assertTrue(rule.title_keywords)

    def test_default_rules_positions_and_ranges(self) -> None:
        rules = list(iter_default_section_rules())
        for rule in rules:
            self.assertIsInstance(rule.position, SectionPosition)
            self.assertIsInstance(rule.body_range, BodyRangeRule)

    def test_default_rules_validation(self) -> None:
        rules = list(iter_default_section_rules())
        validate_section_rules(rules)


if __name__ == "__main__":
    unittest.main()
