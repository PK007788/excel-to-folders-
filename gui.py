"""CustomTkinter user interface for the dataset organizer."""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from processor import DatasetProcessor, ProcessingConfig, ProcessingOptions
from utils import format_duration, open_folder


class DatasetOrganizerApp(ctk.CTk):
    """Main desktop application window."""

    def __init__(self) -> None:
        super().__init__()

        self.title("Retinal Dataset Organizer")
        self.geometry("920x680")
        self.minsize(820, 620)

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

        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        container = ctk.CTkFrame(self, corner_radius=0)
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(container, fg_color="transparent")
        header.grid(row=0, column=0, padx=28, pady=(24, 12), sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            header,
            text="Retinal Dataset Organizer",
            font=ctk.CTkFont(size=26, weight="bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, sticky="ew")

        subtitle = ctk.CTkLabel(
            header,
            text="Copy images into Yes and No folders using labels from an Excel workbook.",
            text_color=("gray35", "gray75"),
            anchor="w",
        )
        subtitle.grid(row=1, column=0, pady=(4, 0), sticky="ew")

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
        self._add_path_row(
            inputs,
            row=2,
            label="Output Folder",
            variable=self.output_var,
            command=self._browse_output,
        )
        self._add_text_row(
            inputs,
            row=3,
            label="File Name Column",
            variable=self.file_column_var,
        )
        self._add_text_row(
            inputs,
            row=4,
            label="Label Column",
            variable=self.label_column_var,
        )

        options = ctk.CTkFrame(container)
        options.grid(row=2, column=0, padx=28, pady=12, sticky="ew")
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

        progress = ctk.CTkFrame(container)
        progress.grid(row=3, column=0, padx=28, pady=12, sticky="nsew")
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

        actions = ctk.CTkFrame(container, fg_color="transparent")
        actions.grid(row=4, column=0, padx=28, pady=(12, 24), sticky="ew")
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

    def _browse_dataset(self) -> None:
        path = filedialog.askdirectory(title="Choose Dataset Folder")
        if path:
            self.dataset_var.set(path)

    def _browse_workbook(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose Excel Workbook",
            filetypes=[
                ("Excel workbooks", "*.xlsx *.xlsm *.xltx *.xltm"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.workbook_var.set(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Choose Output Folder")
        if path:
            self.output_var.set(path)

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
                lines.append(f"  - {label.capitalize()}: {count}")
                
        lines.extend([
            f"Missing images/folders: {stats.missing_images}",
            f"Skipped rows: {stats.skipped_rows}",
            f"Time: {format_duration(stats.processing_time_seconds)}"
        ])
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