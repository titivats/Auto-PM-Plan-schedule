from __future__ import annotations

from pathlib import Path

from . import APP_NAME


DEFAULT_TEMPLATE_PATH = Path(r"C:\Users\sthit\Desktop\ALL BACKLINE PM PLAN - JAN - 2026.xls")
STATE_FILENAME = "app_state.json"
WINDOW_ICON_RELATIVE_PATH = Path("assets") / "app_icon.ico"
EXECUTABLE_NAME = "PMPlanAutoSchedule"

__all__ = [
    "APP_NAME",
    "DEFAULT_TEMPLATE_PATH",
    "EXECUTABLE_NAME",
    "STATE_FILENAME",
    "WINDOW_ICON_RELATIVE_PATH",
]
