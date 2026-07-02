"""Text logger for processing_log.txt."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from types import TracebackType


class ProcessingLogger:
    """Write processing events to a plain text log file."""

    def __init__(self, output_folder: Path) -> None:
        self.output_folder = output_folder
        self.log_path = output_folder / "processing_log.txt"
        self._handle = None

    def __enter__(self) -> "ProcessingLogger":
        self.output_folder.mkdir(parents=True, exist_ok=True)
        self._handle = self.log_path.open("w", encoding="utf-8", newline="")
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        if exc_value is not None:
            self.error(f"Unhandled error: {exc_value}")

        if self._handle is not None:
            self._handle.close()
            self._handle = None

    def info(self, message: str) -> None:
        """Write an informational message."""
        self._write("INFO", message)

    def warning(self, message: str) -> None:
        """Write a warning message."""
        self._write("WARNING", message)

    def error(self, message: str) -> None:
        """Write an error message."""
        self._write("ERROR", message)

    def _write(self, level: str, message: str) -> None:
        if self._handle is None:
            return

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._handle.write(f"[{timestamp}] {level}: {message}\n")
        self._handle.flush()