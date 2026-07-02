"""Excel workbook reader for dataset labels."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

import pandas as pd

from utils import is_blank, normalize_label


@dataclass(frozen=True)
class ImageRecord:
    """One image row from an Excel worksheet."""

    image_name: str
    label: str
    row_number: int


class ExcelReader:
    """Read worksheet names and image-label rows from columns B and C."""

    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path

    def get_sheet_names(self) -> list[str]:
        """Return every worksheet name in workbook order."""
        with pd.ExcelFile(self.workbook_path, engine="openpyxl") as workbook:
            return list(workbook.sheet_names)

    def iter_image_records(
        self, sheet_name: str, file_column: str, label_column: str
    ) -> Iterator[ImageRecord]:
        """Yield records from the specified file and label columns."""
        try:
            frame = pd.read_excel(
                self.workbook_path,
                sheet_name=sheet_name,
                header=0,
                dtype=str,
                engine="openpyxl",
            )
        except Exception:
            return

        frame.columns = [str(c).strip() for c in frame.columns]

        if file_column not in frame.columns or label_column not in frame.columns:
            return

        for index, row in frame.iterrows():
            raw_image_name = row[file_column]
            raw_label = row[label_column]

            image_name = "" if is_blank(raw_image_name) else str(raw_image_name).strip()
            label = normalize_label(raw_label)

            if not image_name and not label:
                continue

            yield ImageRecord(
                image_name=image_name,
                label=label,
                row_number=int(index) + 2,
            )

    def count_candidate_rows(self, sheet_names: list[str]) -> int:
        """Count nonblank rows used for progress estimation."""
        total = 0

        for sheet_name in sheet_names:
            try:
                frame = pd.read_excel(
                    self.workbook_path,
                    sheet_name=sheet_name,
                    header=0,
                    dtype=str,
                    engine="openpyxl",
                )
                total += int(frame.dropna(how="all").shape[0])
            except Exception:
                continue

        return total