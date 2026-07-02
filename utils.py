"""General utilities for validation, display, and OS integration."""

from __future__ import annotations

import platform
import subprocess
from pathlib import Path


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}


def is_blank(value: object) -> bool:
    """Return True when a cell value should be treated as blank."""
    if value is None:
        return True

    text = str(value).strip()
    return text == "" or text.lower() == "nan"


def normalize_label(value: object) -> str:
    """Normalize labels such as yes, Yes, and YES."""
    if is_blank(value):
        return ""
    return str(value).strip().lower()


def format_duration(seconds: float | None) -> str:
    """Format seconds as HH:MM:SS."""
    if seconds is None:
        return "--:--:--"

    whole_seconds = max(0, int(seconds))
    hours, remainder = divmod(whole_seconds, 3600)
    minutes, secs = divmod(remainder, 60)

    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def validate_directory(path: Path, label: str) -> None:
    """Validate that a path exists and is a directory."""
    if not path.exists() or not path.is_dir():
        raise ValueError(f"{label} does not exist or is not a folder: {path}")


def validate_excel_file(path: Path) -> None:
    """Validate that a path exists and is a supported Excel workbook."""
    if not path.exists() or not path.is_file():
        raise ValueError(f"Excel workbook does not exist or is not a file: {path}")

    if path.suffix.lower() not in EXCEL_EXTENSIONS:
        allowed = ", ".join(sorted(EXCEL_EXTENSIONS))
        raise ValueError(f"Excel workbook must be one of: {allowed}")


def open_folder(path: Path) -> None:
    """Open a folder in the operating system's file explorer."""
    path = path.resolve()

    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)

    system = platform.system()

    if system == "Windows":
        subprocess.Popen(["explorer", str(path)])
    elif system == "Darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])
