"""Report writers for completed processing runs."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import TYPE_CHECKING

from utils import format_duration

if TYPE_CHECKING:
    from processor import ProcessingStats


def write_summary_report(output_folder: Path, stats: "ProcessingStats") -> Path:
    """Generate summary.txt."""
    output_folder.mkdir(parents=True, exist_ok=True)
    path = output_folder / "summary.txt"

    lines = [
        "Dataset Organization Summary",
        "============================",
        f"Total Sheets: {stats.total_sheets}",
        f"Total Images Processed: {stats.total_images_processed}",
    ]
    
    for label, count in sorted(stats.label_counts.items()):
        lines.append(f"  - {label.capitalize()}: {count}")

    lines.extend([
        f"Missing Images: {stats.missing_images}",
        f"Duplicate Filenames Renamed: {stats.duplicate_filenames_renamed}",
        f"Skipped Rows: {stats.skipped_rows}",
        f"Processing Time: {format_duration(stats.processing_time_seconds)}",
    ])

    # Multi-label specific stats (only present in multi-label mode)
    if stats.total_rows_processed > 0:
        lines.append("")
        lines.append("Multi-Label Details")
        lines.append("-------------------")
        lines.append(f"Total Rows Processed: {stats.total_rows_processed}")
        lines.append(f"Images Skipped (not found): {stats.images_skipped}")
        lines.append(f"Rows With Multiple Labels: {stats.multi_label_rows}")
        lines.append(f"Rows With No Selected Label: {stats.no_label_rows}")

        if stats.copies_per_folder:
            lines.append("")
            lines.append("Copies Per Folder:")
            for folder, count in sorted(stats.copies_per_folder.items()):
                lines.append(f"  - {folder}: {count}")

    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path



def write_missing_image_report(output_folder: Path, stats: "ProcessingStats") -> Path:
    """Generate missing_images.csv."""
    output_folder.mkdir(parents=True, exist_ok=True)
    path = output_folder / "missing_images.csv"

    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Sheet Name", "Image Name", "Reason"])

        for entry in stats.missing_entries:
            writer.writerow([entry.sheet_name, entry.image_name, entry.reason])

    return path