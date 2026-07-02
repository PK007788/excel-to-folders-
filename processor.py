"""Core processing workflow for organizing images by Excel labels."""

from __future__ import annotations

import threading
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from excel_reader import ExcelReader
from file_manager import FileManager
from logger import ProcessingLogger
from reports import write_missing_image_report, write_summary_report
from utils import validate_directory, validate_excel_file


@dataclass(frozen=True)
class ProcessingOptions:
    """Options selected in the GUI."""

    file_column: str = "Image name"
    label_column: str = "Retinopathy grade"
    prefix_base_folder_name: bool = True
    generate_summary_report: bool = True
    generate_missing_image_report: bool = True
    dry_run_mode: bool = False


@dataclass(frozen=True)
class ProcessingConfig:
    """Paths and options required for one processing run."""

    dataset_folder: Path
    workbook_path: Path
    output_folder: Path
    options: ProcessingOptions


@dataclass(frozen=True)
class MissingEntry:
    """One row for missing_images.csv."""

    sheet_name: str
    image_name: str
    reason: str


@dataclass
class ProcessingStats:
    """Counters collected during processing."""

    total_sheets: int = 0
    total_images_processed: int = 0
    label_counts: dict[str, int] = field(default_factory=dict)
    missing_images: int = 0
    duplicate_filenames_renamed: int = 0
    skipped_rows: int = 0
    processing_time_seconds: float = 0.0
    missing_entries: list[MissingEntry] = field(default_factory=list)


@dataclass(frozen=True)
class ProgressUpdate:
    """Progress payload sent from the worker thread to the GUI."""

    current_sheet: str
    current_image: str
    images_processed: int
    categories_found: int
    elapsed_seconds: float
    estimated_remaining_seconds: float | None
    progress_fraction: float


ProgressCallback = Callable[[ProgressUpdate], None]
LogCallback = Callable[[str], None]


class DatasetProcessor:
    """Organize every worksheet's images into Yes and No folders."""

    def __init__(
        self,
        config: ProcessingConfig,
        progress_callback: ProgressCallback | None = None,
        log_callback: LogCallback | None = None,
        cancel_event: threading.Event | None = None,
    ) -> None:
        self.config = config
        self.progress_callback = progress_callback
        self.log_callback = log_callback
        self.cancel_event = cancel_event or threading.Event()
        self.stats = ProcessingStats()

    def process(self) -> ProcessingStats:
        """Run the complete dataset organization workflow."""
        self._validate_inputs()

        start_time = time.monotonic()
        reader = ExcelReader(self.config.workbook_path)
        sheet_names = reader.get_sheet_names()
        self.stats.total_sheets = len(sheet_names)

        total_work_items = max(reader.count_candidate_rows(sheet_names), 1)

        file_manager = FileManager(
            output_folder=self.config.output_folder,
            dry_run=self.config.options.dry_run_mode,
        )
        file_manager.prepare_output_folders()

        with ProcessingLogger(self.config.output_folder) as log:
            mode = "DRY RUN" if self.config.options.dry_run_mode else "LIVE RUN"
            log.info(f"Started processing in {mode} mode.")
            log.info(f"Workbook: {self.config.workbook_path}")
            log.info(f"Dataset folder: {self.config.dataset_folder}")
            log.info(f"Output folder: {self.config.output_folder}")

            for sheet_name in sheet_names:
                if self.cancel_event.is_set():
                    log.warning("Processing cancelled by user.")
                    break

                self._send_status(f"Processing {sheet_name}")
                sanitized_sheet_name = sheet_name.replace(" ", "")
                base_folder = self.config.dataset_folder / sanitized_sheet_name

                if not base_folder.exists() or not base_folder.is_dir():
                    base_folder = self.config.dataset_folder / sheet_name

                if not base_folder.exists() or not base_folder.is_dir():
                    self._record_missing(
                        sheet_name=sheet_name,
                        image_name="",
                        reason=f"Matching folder not found: {self.config.dataset_folder / sheet_name}",
                        log=log,
                    )
                    self._emit_progress(sheet_name, "", start_time, total_work_items)
                    continue

                log.info(f"Processing worksheet '{sheet_name}' with folder '{base_folder}'.")

                for record in reader.iter_image_records(
                    sheet_name,
                    self.config.options.file_column,
                    self.config.options.label_column,
                ):
                    if self.cancel_event.is_set():
                        log.warning("Processing cancelled by user.")
                        break

                    image_name = record.image_name
                    label = record.label

                    if not image_name:
                        self.stats.skipped_rows += 1
                        log.warning(
                            f"Skipped blank image name in sheet '{sheet_name}', "
                            f"row {record.row_number}."
                        )
                        self._emit_progress(sheet_name, image_name, start_time, total_work_items)
                        continue

                    if not label:
                        self.stats.skipped_rows += 1
                        log.warning(
                            f"Skipped '{image_name}' in sheet '{sheet_name}', "
                            f"row {record.row_number}: blank label."
                        )
                        self._emit_progress(sheet_name, image_name, start_time, total_work_items)
                        continue

                    source_path = base_folder / image_name
                    if not source_path.exists() or not source_path.is_file():
                        self._record_missing(
                            sheet_name=sheet_name,
                            image_name=image_name,
                            reason=f"Image not found: {source_path}",
                            log=log,
                        )
                        self._emit_progress(sheet_name, image_name, start_time, total_work_items)
                        continue

                    destination_path, renamed = file_manager.copy_image(
                        source_path=source_path,
                        base_folder_name=sheet_name,
                        label=label,
                        prefix_base_folder_name=(
                            self.config.options.prefix_base_folder_name
                        ),
                    )

                    if renamed:
                        self.stats.duplicate_filenames_renamed += 1

                    self.stats.total_images_processed += 1
                    self.stats.label_counts[label] = self.stats.label_counts.get(label, 0) + 1

                    action = "Would copy" if self.config.options.dry_run_mode else "Copied"
                    log.info(f"{action}: {source_path} -> {destination_path}")
                    self._emit_progress(sheet_name, image_name, start_time, total_work_items)

        self.stats.processing_time_seconds = time.monotonic() - start_time

        self.config.output_folder.mkdir(parents=True, exist_ok=True)
        if self.config.options.generate_summary_report:
            write_summary_report(self.config.output_folder, self.stats)
        if self.config.options.generate_missing_image_report:
            write_missing_image_report(self.config.output_folder, self.stats)

        return self.stats

    def _validate_inputs(self) -> None:
        validate_directory(self.config.dataset_folder, "Dataset folder")
        validate_excel_file(self.config.workbook_path)
        self.config.output_folder.mkdir(parents=True, exist_ok=True)

    def _record_missing(
        self,
        sheet_name: str,
        image_name: str,
        reason: str,
        log: ProcessingLogger,
    ) -> None:
        self.stats.missing_images += 1
        self.stats.missing_entries.append(
            MissingEntry(
                sheet_name=sheet_name,
                image_name=image_name,
                reason=reason,
            )
        )
        log.warning(reason)

    def _emit_progress(
        self,
        current_sheet: str,
        current_image: str,
        start_time: float,
        total_work_items: int,
    ) -> None:
        if self.progress_callback is None:
            return

        elapsed = time.monotonic() - start_time
        completed_units = (
            self.stats.total_images_processed
            + self.stats.missing_images
            + self.stats.skipped_rows
        )
        fraction = min(1.0, completed_units / max(total_work_items, 1))

        remaining = None
        if completed_units > 0:
            seconds_per_item = elapsed / completed_units
            remaining = max(0.0, (total_work_items - completed_units) * seconds_per_item)

        self.progress_callback(
            ProgressUpdate(
                current_sheet=current_sheet,
                current_image=current_image,
                images_processed=self.stats.total_images_processed,
                categories_found=len(self.stats.label_counts),
                elapsed_seconds=elapsed,
                estimated_remaining_seconds=remaining,
                progress_fraction=fraction,
            )
        )

    def _send_status(self, message: str) -> None:
        if self.log_callback:
            self.log_callback(message)