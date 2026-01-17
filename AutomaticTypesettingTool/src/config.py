from __future__ import annotations

from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "output"
LOG_DIR = PROJECT_ROOT / "logs"
DEFAULT_OUTPUT_PATH = OUTPUT_DIR / "style_rules.json"
TEMPLATE_TYPES_PATH = OUTPUT_DIR / "template_types.json"

LOG_FILE_PREFIX = "template_parser"
LOG_TIMESTAMP_FORMAT = "%Y%m%d_%H%M%S"

DEFAULT_MAX_HEADING_LEVEL = 6
DEFAULT_REQUIRED_ROLES = ["title_L1", "body_L1"]
DEFAULT_REQUIRED_ON_PRESENCE_MAP = {
    "abstract_title": "abstract_body",
    "abstract_en_title": "abstract_en_body",
    "reference_title": "reference_body",
}


def ensure_base_dirs() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)


def build_log_path(ts: datetime | None = None) -> Path:
    if ts is None:
        ts = datetime.now()
    name = f"{LOG_FILE_PREFIX}_{ts.strftime(LOG_TIMESTAMP_FORMAT)}.log"
    return LOG_DIR / name


def cleanup_logs(retention_days: int = 5, now: datetime | None = None) -> int:
    if retention_days <= 0:
        return 0
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
    except OSError:
        return 0
    base_time = now or datetime.now()
    cutoff = base_time.timestamp() - retention_days * 86400
    removed = 0
    for path in LOG_DIR.glob(f"{LOG_FILE_PREFIX}_*.log"):
        try:
            if path.stat().st_mtime < cutoff:
                path.unlink()
                removed += 1
        except OSError:
            continue
    return removed
