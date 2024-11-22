<p align="right">
  [English]
  [<a href="README-ko.md">한국어</a>]
</p>

# Excel2Image

`Excel2Image` is a Python script that converts Excel sheets into image files (JPG). It allows users to take data from an Excel file and generate an image representation of that data, which can be useful for reporting, presentations, or archiving. The script dynamically adjusts the size of the output image based on the content of each sheet.

## Features:
- Converts each sheet of an Excel file into a JPG image.
- Supports Korean fonts on macOS, Windows, and Linux.
- Dynamically adjusts the size of the output image based on sheet content.
- Handles merged cells and text alignment.

## Requirements:
- Python 3
- `openpyxl` library (for reading Excel files)
- `matplotlib` library (for generating images)
- `PIL` (Pillow) library (for saving the image)

## Installation:

1. Install the required libraries: `pip3 install -r requirements.txt`

## Usage:

Run the script from the command line:

- `<excel_file_path>`: The path to the Excel file.
- `<output_directory>`: The directory where the image files will be saved.

For example: excel2image sample.xlsx .

## License:
This project is licensed under the MIT License.