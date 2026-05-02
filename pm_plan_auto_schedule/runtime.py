from __future__ import annotations

import os
import sys
from pathlib import Path

from .config import STATE_FILENAME


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_root() -> Path:
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def bundle_root() -> Path:
    return Path(getattr(sys, "_MEIPASS", app_root()))


def resource_path(*parts: str) -> Path:
    return bundle_root().joinpath(*parts)


def state_file_path() -> Path:
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "PMPlanAutoSchedule" / STATE_FILENAME
    return app_root() / STATE_FILENAME
