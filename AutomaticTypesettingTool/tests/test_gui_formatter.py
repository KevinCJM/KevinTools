import unittest

from src.gui_formatter import format_result_payload


class GuiFormatterTests(unittest.TestCase):
    def test_format_result_payload_schema_v2(self) -> None:
        payload = {
            "schema_version": "2.0",
            "roles": {
                "title_L1": {"font_name": "宋体"},
                "body_L1": {"font_name": "宋体"},
                "abstract_title": {"font_name": "宋体"},
                "abstract_body": {"font_name": "宋体"},
            },
        }
        text = format_result_payload(payload)
        self.assertIn("【多级标题/正文】", text)
        self.assertIn("【1级标题】", text)
        self.assertIn("【1级正文】", text)
        self.assertIn("【摘要】", text)
        self.assertIn("【摘要标题】", text)
        self.assertIn("【摘要正文】", text)

    def test_format_result_payload_invalid(self) -> None:
        self.assertEqual(format_result_payload([]), "解析结果格式异常")

    def test_format_result_payload_meta_groups(self) -> None:
        payload = {
            "schema_version": "2.0",
            "roles": {
                "abstract_en_title": {"font_name": "Times New Roman"},
                "abstract_en_body": {"font_name": "Times New Roman"},
                "toc_body_L1": {"font_name": "宋体"},
                "footnote_reference": {"font_name": "宋体"},
                "footnote_text": {"font_name": "宋体"},
            },
            "meta": {
                "page_margins": {"sections": [], "summary": {}},
                "table_borders": {"summary": {}, "tables": []},
                "title_spacing": {},
                "footnote_numbering": {},
            },
        }
        text = format_result_payload(payload)
        self.assertIn("【摘要】", text)
        self.assertIn("【英文摘要标题】", text)
        self.assertIn("【目录】", text)
        self.assertIn("【目录正文 L1】", text)
        self.assertIn("【脚注】", text)
        self.assertIn("【脚注正文样式】", text)
        self.assertIn("【页面边距】", text)
        self.assertIn("【表格边框】", text)
        self.assertIn("【标题空行】", text)
        self.assertIn("【脚注编号】", text)


if __name__ == "__main__":
    unittest.main()
