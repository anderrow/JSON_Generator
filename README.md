# Excel Generator

A Python utility that reads the source CSV files and generates a single Excel workbook.

## Overview

This tool reads every `.csv` file in `SOURCE FILES`, filters the rows, prefixes the `Keys` column with the project name, and writes everything into one Excel workbook with all tables stacked in a single sheet.

## Requirements

- Python 3.x
- Source `.csv` files inside `../SOURCE FILES/`

## Project Structure

Required file structure:
```text
_Proccesview/
|- SOURCE FILES/
|  `- *.csv
|- JSON_Generator/
|  |- ExcelGenerator.py
|  |- requirements.txt
|  `- run_env.py
`- EXCEL FILES/         (generated)
   `- AllProjects.xlsx
```

## Usage

1. Ensure the source `.csv` files are in the correct location.
2. Run the generator:
   ```bash
   python run_env.py
   ```

`run_env.py` creates its virtual environment outside the project folder
(for example under `%LOCALAPPDATA%` on Windows), so Nextcloud does not
sync it.

## Creating an Executable (Optional)

If you need an executable version:

1. Install pyinstaller:
   ```bash
   pip install pyinstaller
   ```

2. Create the executable:
   ```bash
   pyinstaller --onefile --windowed ExcelGenerator.py
   ```

## Dependencies

Main dependencies:
- polars
- openpyxl

See `requirements.txt` for the complete list.
