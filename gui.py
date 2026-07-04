"""CustomTkinter user interface for the dataset organizer."""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

# pyrefly: ignore [missing-import]
import customtkinter as ctk

from excel_reader import ExcelReader
from processor import DatasetProcessor, ProcessingConfig, ProcessingOptions
from utils import format_duration, open_folder


class DataPreviewWindow(ctk.CTkToplevel):
    """Pop-up window that shows the first rows of an Excel sheet."""

    def __init__(self, parent, headers: list[str], rows: list[list[str]], sheet_name: str) -> None:
        super().__init__(parent)
        self.title(f"Data Preview — {sheet_name}")
        self.geometry("900x500")
        self.minsize(600, 300)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            self,
            text=f"First {len(rows)} rows of '{sheet_name}'",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="ew")

        # Scrollable frame for the table
        scroll_frame = ctk.CTkScrollableFrame(self, orientation="horizontal")
        scroll_frame.grid(row=1, column=0, padx=18, pady=(0, 16), sticky="nsew")

        # Column headers
        for col_idx, header in enumerate(headers):
            lbl = ctk.CTkLabel(
                scroll_frame,
                text=header,
                font=ctk.CTkFont(size=13, weight="bold"),
                anchor="w",
                padx=8,
                pady=4,
            )
            lbl.grid(row=0, column=col_idx, padx=2, pady=2, sticky="ew")

        # Data rows
        for row_idx, row_data in enumerate(rows):
            for col_idx, cell_value in enumerate(row_data):
                display = cell_value if len(cell_value) <= 40 else cell_value[:37] + "..."
                lbl = ctk.CTkLabel(
                    scroll_frame,
                    text=display,
                    anchor="w",
                    padx=8,
                    pady=2,
                )
                lbl.grid(row=row_idx + 1, column=col_idx, padx=2, pady=1, sticky="ew")

        # Close button
        ctk.CTkButton(self, text="Close", width=100, command=self.destroy).grid(
            row=2, column=0, pady=(0, 16)
        )

        self.lift()
        self.focus_force()


class DatasetOrganizerApp(ctk.CTk):
    """Main desktop application window."""

    def __init__(self) -> None:
        super().__init__()

        self.title("Dataset Organizer")
        self.geometry("960x780")
        self.minsize(860, 720)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.dataset_var = tk.StringVar()
        self.workbook_var = tk.StringVar()
        self.output_var = tk.StringVar()

        self.file_column_var = tk.StringVar(value="Image name")
        self.label_column_var = tk.StringVar(value="Retinopathy grade")

        self.prefix_var = tk.BooleanVar(value=True)
        self.summary_var = tk.BooleanVar(value=True)
        self.missing_var = tk.BooleanVar(value=True)
        self.dry_run_var = tk.BooleanVar(value=False)

        self.status_var = tk.StringVar(value="Ready")
        self.sheet_var = tk.StringVar(value="-")
        self.image_var = tk.StringVar(value="-")
        self.processed_var = tk.StringVar(value="0")
        self.categories_var = tk.StringVar(value="0")
        self.elapsed_var = tk.StringVar(value="00:00:00")
        self.remaining_var = tk.StringVar(value="--:--:--")

        self.progress_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None

        # Multi-label state
        self._mode_var = tk.StringVar(value="Single Label Dataset")
        self._multi_image_vars: list[tuple[str, tk.BooleanVar]] = []
        self._multi_label_vars: list[tuple[str, tk.BooleanVar]] = []

        self._build_ui()
        self.after(100, self._poll_queue)

    # ── UI Construction ─────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        container = ctk.CTkScrollableFrame(self, corner_radius=0)
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)

        # ── Header ──
        header = ctk.CTkFrame(container, fg_color="transparent")
        header.grid(row=0, column=0, padx=28, pady=(24, 12), sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            header,
            text="Dataset Organizer",
            font=ctk.CTkFont(size=26, weight="bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, sticky="ew")

        subtitle = ctk.CTkLabel(
            header,
            text="Organize image datasets into labelled folders using an Excel workbook.",
            text_color=("gray35", "gray75"),
            anchor="w",
        )
        subtitle.grid(row=1, column=0, pady=(4, 0), sticky="ew")

        # ── Inputs (paths) ──
        inputs = ctk.CTkFrame(container)
        inputs.grid(row=1, column=0, padx=28, pady=12, sticky="ew")
        inputs.grid_columnconfigure(1, weight=1)

        self._add_path_row(
            inputs,
            row=0,
            label="Dataset Folder",
            variable=self.dataset_var,
            command=self._browse_dataset,
        )
        self._add_path_row(
            inputs,
            row=1,
            label="Excel Workbook",
            variable=self.workbook_var,
            command=self._browse_workbook,
        )

        # Preview Data button (row 1, column 3)
        self._preview_button = ctk.CTkButton(
            inputs, text="Preview", width=90, command=self._preview_data
        )
        self._preview_button.grid(row=1, column=3, padx=(4, 18), pady=14)

        self._add_path_row(
            inputs,
            row=2,
            label="Output Folder",
            variable=self.output_var,
            command=self._browse_output,
        )

        # ── Mode Toggle ──
        mode_frame = ctk.CTkFrame(container)
        mode_frame.grid(row=2, column=0, padx=28, pady=12, sticky="ew")
        mode_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            mode_frame,
            text="Processing Mode",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="ew")

        self._mode_toggle = ctk.CTkSegmentedButton(
            mode_frame,
            values=["Single Label Dataset", "Multi-Label Dataset"],
            variable=self._mode_var,
            command=self._on_mode_changed,
        )
        self._mode_toggle.grid(row=1, column=0, padx=18, pady=(0, 16), sticky="ew")

        # ── Single-label column inputs ──
        self._single_label_frame = ctk.CTkFrame(container)
        self._single_label_frame.grid(row=3, column=0, padx=28, pady=12, sticky="ew")
        self._single_label_frame.grid_columnconfigure(1, weight=1)

        self._add_text_row(
            self._single_label_frame,
            row=0,
            label="File Name Column",
            variable=self.file_column_var,
        )
        self._add_text_row(
            self._single_label_frame,
            row=1,
            label="Label Column",
            variable=self.label_column_var,
        )

        # ── Multi-label column inputs ──
        self._multi_label_frame = ctk.CTkFrame(container)
        self._multi_label_frame.grid_columnconfigure(0, weight=1)
        self._multi_label_frame.grid_columnconfigure(1, weight=1)
        # Hidden by default — not gridded yet

        self._load_columns_button = ctk.CTkButton(
            self._multi_label_frame,
            text="Load Columns from Workbook",
            command=self._load_columns,
        )
        self._load_columns_button.grid(
            row=0, column=0, columnspan=2, padx=18, pady=(16, 8), sticky="ew"
        )

        ctk.CTkLabel(
            self._multi_label_frame,
            text="Image Columns",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        ).grid(row=1, column=0, padx=18, pady=(8, 4), sticky="ew")

        ctk.CTkLabel(
            self._multi_label_frame,
            text="Label Columns",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        ).grid(row=1, column=1, padx=18, pady=(8, 4), sticky="ew")

        self._image_col_scroll = ctk.CTkScrollableFrame(
            self._multi_label_frame, height=140
        )
        self._image_col_scroll.grid(
            row=2, column=0, padx=18, pady=(0, 16), sticky="nsew"
        )

        self._label_col_scroll = ctk.CTkScrollableFrame(
            self._multi_label_frame, height=140
        )
        self._label_col_scroll.grid(
            row=2, column=1, padx=18, pady=(0, 16), sticky="nsew"
        )

        # ── Options ──
        options = ctk.CTkFrame(container)
        options.grid(row=5, column=0, padx=28, pady=12, sticky="ew")
        options.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(
            options,
            text="Options",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, padx=18, pady=(16, 8), sticky="ew")

        ctk.CTkCheckBox(
            options,
            text="Prefix Base Folder Name",
            variable=self.prefix_var,
        ).grid(row=1, column=0, padx=18, pady=8, sticky="w")

        ctk.CTkCheckBox(
            options,
            text="Generate Summary Report",
            variable=self.summary_var,
        ).grid(row=1, column=1, padx=18, pady=8, sticky="w")

        ctk.CTkCheckBox(
            options,
            text="Generate Missing Image Report",
            variable=self.missing_var,
        ).grid(row=2, column=0, padx=18, pady=(8, 16), sticky="w")

        ctk.CTkCheckBox(
            options,
            text="Dry Run Mode",
            variable=self.dry_run_var,
        ).grid(row=2, column=1, padx=18, pady=(8, 16), sticky="w")

        # ── Progress ──
        progress = ctk.CTkFrame(container)
        progress.grid(row=6, column=0, padx=28, pady=12, sticky="nsew")
        progress.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkLabel(
            progress,
            text="Progress",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, columnspan=4, padx=18, pady=(16, 8), sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(progress)
        self.progress_bar.grid(row=1, column=0, columnspan=4, padx=18, pady=10, sticky="ew")
        self.progress_bar.set(0)

        self._add_stat(progress, "Current Sheet", self.sheet_var, 2, 0)
        self._add_stat(progress, "Current Image", self.image_var, 2, 1)
        self._add_stat(progress, "Images Processed", self.processed_var, 2, 2)
        self._add_stat(progress, "Categories Found", self.categories_var, 2, 3)
        self._add_stat(progress, "Elapsed Time", self.elapsed_var, 4, 0)
        self._add_stat(progress, "Estimated Time Remaining", self.remaining_var, 4, 1)
        self._add_stat(progress, "Status", self.status_var, 4, 2)

        # ── Action Buttons ──
        actions = ctk.CTkFrame(container, fg_color="transparent")
        actions.grid(row=7, column=0, padx=28, pady=(12, 24), sticky="ew")
        actions.grid_columnconfigure(0, weight=1)

        button_frame = ctk.CTkFrame(actions, fg_color="transparent")
        button_frame.grid(row=0, column=1, sticky="e")

        self.start_button = ctk.CTkButton(
            button_frame,
            text="Start",
            width=130,
            command=self._start_processing,
        )
        self.start_button.grid(row=0, column=0, padx=(0, 10))

        self.cancel_button = ctk.CTkButton(
            button_frame,
            text="Cancel",
            width=130,
            state="disabled",
            fg_color="#9B2C2C",
            hover_color="#7F1D1D",
            command=self._cancel_processing,
        )
        self.cancel_button.grid(row=0, column=1, padx=(0, 10))

        self.open_button = ctk.CTkButton(
            button_frame,
            text="Open Output Folder",
            width=170,
            command=self._open_output,
        )
        self.open_button.grid(row=0, column=2)

        # Store the container reference for mode switching
        self._container = container

    # ── Reusable row builders ───────────────────────────────────────

    def _add_path_row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
    ) -> None:
        ctk.CTkLabel(parent, text=label, anchor="w").grid(
            row=row,
            column=0,
            padx=(18, 12),
            pady=14,
            sticky="w",
        )
        entry = ctk.CTkEntry(parent, textvariable=variable)
        entry.grid(row=row, column=1, padx=(0, 12), pady=14, sticky="ew")
        ctk.CTkButton(parent, text="Browse", width=110, command=command).grid(
            row=row,
            column=2,
            padx=(0, 18),
            pady=14,
        )

    def _add_text_row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        label: str,
        variable: tk.StringVar,
    ) -> None:
        ctk.CTkLabel(parent, text=label, anchor="w").grid(
            row=row,
            column=0,
            padx=(18, 12),
            pady=14,
            sticky="w",
        )
        entry = ctk.CTkEntry(parent, textvariable=variable)
        entry.grid(row=row, column=1, columnspan=2, padx=(0, 18), pady=14, sticky="ew")

    def _add_stat(
        self,
        parent: ctk.CTkFrame,
        label: str,
        variable: tk.StringVar,
        row: int,
        column: int,
    ) -> None:
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=column, padx=18, pady=10, sticky="ew")

        ctk.CTkLabel(
            frame,
            text=label,
            text_color=("gray35", "gray75"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")

        ctk.CTkLabel(
            frame,
            textvariable=variable,
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w",
            wraplength=180,
        ).grid(row=1, column=0, sticky="ew")

    # ── Mode Switching ──────────────────────────────────────────────

    def _on_mode_changed(self, value: str) -> None:
        """Toggle between single-label and multi-label column panels."""
        if value == "Single Label Dataset":
            self._multi_label_frame.grid_forget()
            self._single_label_frame.grid(
                row=3, column=0, padx=28, pady=12, sticky="ew"
            )
        else:
            self._single_label_frame.grid_forget()
            self._multi_label_frame.grid(
                row=3, column=0, padx=28, pady=12, sticky="ew"
            )

    # ── Multi-label column loading ──────────────────────────────────

    def _load_columns(self) -> None:
        """Read column headers from the workbook and populate checkbox lists."""
        workbook = self.workbook_var.get().strip()
        if not workbook:
            messagebox.showwarning("No Workbook", "Please select an Excel workbook first.")
            return

        try:
            reader = ExcelReader(Path(workbook))
            sheets = reader.get_sheet_names()
            if not sheets:
                messagebox.showwarning("Empty Workbook", "The workbook has no sheets.")
                return
            columns = reader.get_column_names(sheets[0])
            if not columns:
                messagebox.showwarning("No Columns", "No columns found in the first sheet.")
                return
        except Exception as exc:
            messagebox.showerror("Load Error", str(exc))
            return

        # Clear previous checkboxes
        for widget in self._image_col_scroll.winfo_children():
            widget.destroy()
        for widget in self._label_col_scroll.winfo_children():
            widget.destroy()

        self._multi_image_vars.clear()
        self._multi_label_vars.clear()

        for col_name in columns:
            # Image column checkbox
            img_var = tk.BooleanVar(value=False)
            self._multi_image_vars.append((col_name, img_var))
            ctk.CTkCheckBox(
                self._image_col_scroll,
                text=col_name,
                variable=img_var,
            ).pack(anchor="w", padx=8, pady=3)

            # Label column checkbox
            lbl_var = tk.BooleanVar(value=False)
            self._multi_label_vars.append((col_name, lbl_var))
            ctk.CTkCheckBox(
                self._label_col_scroll,
                text=col_name,
                variable=lbl_var,
            ).pack(anchor="w", padx=8, pady=3)

        messagebox.showinfo(
            "Columns Loaded",
            f"Found {len(columns)} columns in sheet '{sheets[0]}'.\n"
            "Select your image and label columns.",
        )

    # ── Data Preview ────────────────────────────────────────────────

    def _preview_data(self) -> None:
        """Open a preview window showing the first rows of the Excel file."""
        workbook = self.workbook_var.get().strip()
        if not workbook:
            messagebox.showwarning("No Workbook", "Please select an Excel workbook first.")
            return

        try:
            reader = ExcelReader(Path(workbook))
            sheets = reader.get_sheet_names()
            if not sheets:
                messagebox.showwarning("Empty Workbook", "The workbook has no sheets.")
                return
            headers, rows = reader.get_preview_rows(sheets[0])
            if not headers:
                messagebox.showwarning("No Data", "Could not read data from the first sheet.")
                return
        except Exception as exc:
            messagebox.showerror("Preview Error", str(exc))
            return

        DataPreviewWindow(self, headers, rows, sheets[0])

    # ── File Browsing ───────────────────────────────────────────────

    def _browse_dataset(self) -> None:
        path = filedialog.askdirectory(title="Choose Dataset Folder")
        if path:
            self.dataset_var.set(path)

    def _browse_workbook(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose Data File",
            filetypes=[
                ("Supported files", "*.xlsx *.xlsm *.xltx *.xltm *.csv"),
                ("Excel workbooks", "*.xlsx *.xlsm *.xltx *.xltm"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.workbook_var.set(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Choose Output Folder")
        if path:
            self.output_var.set(path)

    # ── Processing ──────────────────────────────────────────────────

    def _start_processing(self) -> None:
        if self.worker and self.worker.is_alive():
            return

        try:
            config = self._build_config()
        except ValueError as exc:
            messagebox.showerror("Missing Input", str(exc))
            return

        self.cancel_event.clear()
        self._set_running_state(True)
        self._reset_progress()
        self.status_var.set("Running")

        self.worker = threading.Thread(
            target=self._run_worker,
            args=(config,),
            daemon=True,
        )
        self.worker.start()

    def _build_config(self) -> ProcessingConfig:
        dataset = self.dataset_var.get().strip()
        workbook = self.workbook_var.get().strip()
        output = self.output_var.get().strip()

        if not dataset:
            raise ValueError("Please choose a dataset folder.")
        if not workbook:
            raise ValueError("Please choose an Excel workbook.")
        if not output:
            raise ValueError("Please choose an output folder.")

        is_multi = self._mode_var.get() == "Multi-Label Dataset"

        if is_multi:
            file_columns = tuple(
                name for name, var in self._multi_image_vars if var.get()
            )
            label_columns = tuple(
                name for name, var in self._multi_label_vars if var.get()
            )

            if not file_columns:
                raise ValueError(
                    "Please select at least one Image Column.\n"
                    "Click 'Load Columns from Workbook' first."
                )
            if not label_columns:
                raise ValueError(
                    "Please select at least one Label Column.\n"
                    "Click 'Load Columns from Workbook' first."
                )

            return ProcessingConfig(
                dataset_folder=Path(dataset),
                workbook_path=Path(workbook),
                output_folder=Path(output),
                options=ProcessingOptions(
                    multi_label_mode=True,
                    file_columns=file_columns,
                    label_columns=label_columns,
                    prefix_base_folder_name=self.prefix_var.get(),
                    generate_summary_report=self.summary_var.get(),
                    generate_missing_image_report=self.missing_var.get(),
                    dry_run_mode=self.dry_run_var.get(),
                ),
            )
        else:
            file_column = self.file_column_var.get().strip()
            label_column = self.label_column_var.get().strip()

            if not file_column or not label_column:
                raise ValueError("Please specify both File Name Column and Label Column.")

            return ProcessingConfig(
                dataset_folder=Path(dataset),
                workbook_path=Path(workbook),
                output_folder=Path(output),
                options=ProcessingOptions(
                    file_column=file_column,
                    label_column=label_column,
                    prefix_base_folder_name=self.prefix_var.get(),
                    generate_summary_report=self.summary_var.get(),
                    generate_missing_image_report=self.missing_var.get(),
                    dry_run_mode=self.dry_run_var.get(),
                ),
            )

    def _run_worker(self, config: ProcessingConfig) -> None:
        processor = DatasetProcessor(
            config=config,
            progress_callback=lambda payload: self.progress_queue.put(("progress", payload)),
            log_callback=lambda message: self.progress_queue.put(("status", message)),
            cancel_event=self.cancel_event,
        )

        try:
            stats = processor.process()
        except Exception as exc:
            self.progress_queue.put(("error", str(exc)))
            return

        self.progress_queue.put(("complete", stats))

    def _poll_queue(self) -> None:
        while True:
            try:
                event, payload = self.progress_queue.get_nowait()
            except queue.Empty:
                break

            if event == "progress":
                self._apply_progress(payload)
            elif event == "status":
                self.status_var.set(str(payload))
            elif event == "error":
                self._set_running_state(False)
                self.status_var.set("Error")
                messagebox.showerror("Processing Error", str(payload))
            elif event == "complete":
                self._set_running_state(False)
                status = "Cancelled" if self.cancel_event.is_set() else "Complete"
                self.status_var.set(status)
                self.progress_bar.set(1 if not self.cancel_event.is_set() else self.progress_bar.get())
                messagebox.showinfo("Finished", self._completion_message(payload))

        self.after(100, self._poll_queue)

    def _apply_progress(self, progress) -> None:
        self.sheet_var.set(progress.current_sheet or "-")
        self.image_var.set(progress.current_image or "-")
        self.processed_var.set(str(progress.images_processed))
        self.categories_var.set(str(progress.categories_found))
        self.elapsed_var.set(format_duration(progress.elapsed_seconds))
        self.remaining_var.set(format_duration(progress.estimated_remaining_seconds))
        self.progress_bar.set(progress.progress_fraction)

    def _completion_message(self, stats) -> str:
        lines = [
            f"Images processed: {stats.total_images_processed}",
            f"Categories found: {len(stats.label_counts)}",
        ]
        if stats.label_counts:
            lines.append("Counts:")
            for label, count in sorted(stats.label_counts.items()):
                lines.append(f"  - {label}: {count}")

        lines.extend([
            f"Missing images/folders: {stats.missing_images}",
            f"Skipped rows: {stats.skipped_rows}",
            f"Time: {format_duration(stats.processing_time_seconds)}"
        ])

        # Multi-label specific info
        if stats.total_rows_processed > 0:
            lines.append("")
            lines.append(f"Rows processed: {stats.total_rows_processed}")
            lines.append(f"Multi-label rows: {stats.multi_label_rows}")
            lines.append(f"No-label rows: {stats.no_label_rows}")

        return "\n".join(lines)

    def _cancel_processing(self) -> None:
        self.cancel_event.set()
        self.status_var.set("Cancelling...")

    def _open_output(self) -> None:
        output = self.output_var.get().strip()
        if not output:
            messagebox.showinfo("Output Folder", "Please choose an output folder first.")
            return
        try:
            open_folder(Path(output))
        except Exception as exc:
            messagebox.showerror("Open Folder", str(exc))

    def _reset_progress(self) -> None:
        self.progress_bar.set(0)
        self.sheet_var.set("-")
        self.image_var.set("-")
        self.processed_var.set("0")
        self.categories_var.set("0")
        self.elapsed_var.set("00:00:00")
        self.remaining_var.set("--:--:--")

    def _set_running_state(self, running: bool) -> None:
        self.start_button.configure(state="disabled" if running else "normal")
        self.cancel_button.configure(state="normal" if running else "disabled")