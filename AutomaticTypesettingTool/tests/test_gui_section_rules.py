import unittest

from src import gui
from src.section_rules import BodyRangeRule, SectionPosition


class GuiSectionRuleParsingTests(unittest.TestCase):
    def test_split_text_tokens(self) -> None:
        text = "致谢, 感谢；鸣谢/谢谢|敬谢"
        tokens = gui._split_text_tokens(text)
        self.assertEqual(tokens, ["致谢", "感谢", "鸣谢", "谢谢", "敬谢"])

    def test_parse_position_text(self) -> None:
        self.assertEqual(gui._parse_position_text("首页"), SectionPosition.FIRST_PAGE)
        self.assertEqual(gui._parse_position_text("前置"), SectionPosition.FRONT)
        self.assertEqual(gui._parse_position_text("正文"), SectionPosition.BODY)
        self.assertEqual(gui._parse_position_text("后置"), SectionPosition.BACK)
        self.assertEqual(gui._parse_position_text("尾页"), SectionPosition.LAST_PAGE)
        self.assertEqual(gui._parse_position_text("front"), SectionPosition.FRONT)

    def test_parse_body_range_text(self) -> None:
        self.assertEqual(
            gui._parse_body_range_text("直到下一个标题"),
            BodyRangeRule.UNTIL_NEXT_TITLE,
        )
        self.assertEqual(
            gui._parse_body_range_text("到空行"),
            BodyRangeRule.UNTIL_BLANK,
        )
        self.assertEqual(
            gui._parse_body_range_text("固定段数"),
            BodyRangeRule.FIXED_PARAGRAPHS,
        )

    def test_parse_int(self) -> None:
        self.assertEqual(gui._parse_int("3"), 3)
        self.assertIsNone(gui._parse_int(""))
        self.assertIsNone(gui._parse_int("x"))


if __name__ == "__main__":
    unittest.main()
