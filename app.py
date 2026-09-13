from __future__ import annotations

import sys
import os
from pathlib import Path

from PySide6.QtCore import QTimer

from sheet_hub.ui import create_app


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


if __name__ == "__main__":
    override = os.getenv("SHEET_DATA_HUB_DATA_DIR", "").strip()
    application, main_window = create_app(
        resource_path("assets/app.ico"),
        Path(override) if override else None,
    )
    main_window.show()
    if "--smoke-test" in sys.argv:
        QTimer.singleShot(1200, application.quit)
    raise SystemExit(application.exec())
