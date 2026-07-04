"""Entry point for the dataset organizer."""

from gui import DatasetOrganizerApp


def main() -> None:
    """Launch the desktop application."""
    app = DatasetOrganizerApp()
    app.mainloop()


if __name__ == "__main__":
    main()