from __future__ import annotations

import sys
from pathlib import Path


def _ensure_project_root() -> None:
    root = Path(__file__).resolve().parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


def main() -> None:
    _ensure_project_root()
    from src.gui import main as gui_main

    gui_main()


if __name__ == "__main__":
    main()
