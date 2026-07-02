# General Dataset Organizer

A desktop Python application for organizing image datasets for machine learning or general categorization tasks based on metadata from an Excel workbook.

## Features
- **Dynamic Columns:** Customize which Excel columns to use for "File Name" and "Category Label".
- **Dynamic Folders:** Automatically creates a unique folder for every category it discovers in your dataset.
- **Collision Handling:** Safely handles duplicate filenames.
- **Reporting:** Generates processing reports and missing file manifests automatically.

## How it works

The tool reads metadata labels from an Excel file, searches your source dataset for the corresponding image files, and copies them into an organized output directory structured by category:

```text
Output/
    Category_1/
        image1.jpg
        image2.png
    Category_2/
        image3.tif
    Category_3/
        ...
```

## Running the tool

```bash
pip install -r requirements.txt
python main.py
```