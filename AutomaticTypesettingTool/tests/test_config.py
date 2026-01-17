import os
import unittest
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory

from src import config


class ConfigTests(unittest.TestCase):
    def test_default_output_path(self) -> None:
        self.assertEqual(config.DEFAULT_OUTPUT_PATH.name, "style_rules.json")
        self.assertEqual(config.DEFAULT_OUTPUT_PATH.parent.name, "output")

    def test_log_dir(self) -> None:
        self.assertEqual(config.LOG_DIR.name, "logs")

    def test_build_log_path(self) -> None:
        ts = datetime(2024, 1, 2, 3, 4, 5)
        log_path = config.build_log_path(ts)
        expected_name = "template_parser_20240102_030405.log"
        self.assertEqual(log_path.parent, config.LOG_DIR)
        self.assertEqual(log_path.name, expected_name)

    def test_build_log_path_default_timestamp(self) -> None:
        log_path = config.build_log_path()
        self.assertEqual(log_path.parent, config.LOG_DIR)

    def test_base_dirs_paths(self) -> None:
        self.assertEqual(config.OUTPUT_DIR, config.PROJECT_ROOT / "output")
        self.assertEqual(config.LOG_DIR, config.PROJECT_ROOT / "logs")

    def test_default_multilevel_settings(self) -> None:
        self.assertEqual(config.DEFAULT_MAX_HEADING_LEVEL, 6)
        self.assertEqual(config.DEFAULT_REQUIRED_ROLES, ["title_L1", "body_L1"])
        self.assertEqual(
            config.DEFAULT_REQUIRED_ON_PRESENCE_MAP,
            {
                "abstract_title": "abstract_body",
                "abstract_en_title": "abstract_en_body",
                "reference_title": "reference_body",
            },
        )

    def test_ensure_base_dirs(self) -> None:
        config.ensure_base_dirs()
        self.assertTrue(config.OUTPUT_DIR.exists())
        self.assertTrue(config.LOG_DIR.exists())
        self.assertTrue(config.OUTPUT_DIR.is_dir())
        self.assertTrue(config.LOG_DIR.is_dir())

    def test_cleanup_logs_removes_old_files(self) -> None:
        original_log_dir = config.LOG_DIR
        with TemporaryDirectory() as tmpdir:
            config.LOG_DIR = Path(tmpdir)
            try:
                old_log = config.LOG_DIR / f"{config.LOG_FILE_PREFIX}_old.log"
                new_log = config.LOG_DIR / f"{config.LOG_FILE_PREFIX}_new.log"
                old_log.write_text("old", encoding="utf-8")
                new_log.write_text("new", encoding="utf-8")
                base_time = datetime(2024, 1, 10, 12, 0, 0)
                old_time = base_time.timestamp() - 6 * 86400
                new_time = base_time.timestamp() - 2 * 86400
                os.utime(old_log, (old_time, old_time))
                os.utime(new_log, (new_time, new_time))

                removed = config.cleanup_logs(retention_days=5, now=base_time)

                self.assertEqual(removed, 1)
                self.assertFalse(old_log.exists())
                self.assertTrue(new_log.exists())
            finally:
                config.LOG_DIR = original_log_dir


if __name__ == "__main__":
    unittest.main()
