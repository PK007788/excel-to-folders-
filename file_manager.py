"""File manager for copying images to the output dataset."""

from __future__ import annotations

import shutil
from pathlib import Path


class FileManager:
    """Manages output folders and file copying."""

    def __init__(self, output_folder: Path, dry_run: bool = False) -> None:
        self.output_folder = output_folder
        self.dry_run = dry_run

    def prepare_output_folders(self) -> None:
        """Create the output folder."""
        if not self.dry_run:
            self.output_folder.mkdir(parents=True, exist_ok=True)

    def copy_image(
        self,
        source_path: Path,
        base_folder_name: str,
        label: str,
        prefix_base_folder_name: bool,
    ) -> tuple[Path, bool]:
        """Copy an image to the appropriate Yes/No folder.

        Returns:
            A tuple of (destination_path, renamed).
            renamed is True if the file name was altered to avoid a collision.
        """
        # Determine the target directory based on label
        target_dir = self.output_folder / label.capitalize()
        if not self.dry_run:
            target_dir.mkdir(exist_ok=True)
        
        # Determine the new file name
        original_name = source_path.name
        if prefix_base_folder_name:
            new_name = f"{base_folder_name}_{original_name}"
        else:
            new_name = original_name

        destination_path = target_dir / new_name
        
        # Handle collisions
        renamed = False
        counter = 1
        stem = destination_path.stem
        suffix = destination_path.suffix
        
        while destination_path.exists():
            renamed = True
            destination_path = target_dir / f"{stem}_{counter}{suffix}"
            counter += 1

        if not self.dry_run:
            shutil.copy2(source_path, destination_path)

        return destination_path, renamed

    def copy_image_to_label_folder(
        self,
        source_path: Path,
        label_folder_name: str,
        prefix_base_folder_name: bool,
        base_folder_name: str,
    ) -> tuple[Path, bool]:
        """Copy an image into a folder named after a label column header.

        Returns:
            A tuple of (destination_path, renamed).
            renamed is True if the file name was altered to avoid a collision.
        """
        target_dir = self.output_folder / label_folder_name
        if not self.dry_run:
            target_dir.mkdir(exist_ok=True)

        original_name = source_path.name
        if prefix_base_folder_name:
            new_name = f"{base_folder_name}_{original_name}"
        else:
            new_name = original_name

        destination_path = target_dir / new_name

        # Handle collisions
        renamed = False
        counter = 1
        stem = destination_path.stem
        suffix = destination_path.suffix

        while destination_path.exists():
            renamed = True
            destination_path = target_dir / f"{stem}_{counter}{suffix}"
            counter += 1

        if not self.dry_run:
            shutil.copy2(source_path, destination_path)

        return destination_path, renamed
