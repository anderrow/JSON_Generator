# JSON Generator

A Python utility that converts ProcessViewMessages Excel data into JSON files.

## Overview

This tool automatically processes the `ProcessViewMessages.xlsm` Excel file and generates corresponding JSON files for each sheet, organizing the output by category (AVA, UFA, VILOFOSS).

## Requirements

- Python 3.x
- `ProcessViewMessages.xlsm` file in the parent directory

## Project Structure

Required file structure:
```
PROCESSVIEW MESSAGES/
├── ProcessViewMessages.xlsm
├── JSON Generator/
│   ├── JsonGenerator.py
│   ├── requirements.txt
│   └── run_env.py
└── JSON FILES/         (generated)
    ├── AVA.json
    ├── UFA.json
    └── VILOFOSS.json
```

## Usage

1. Ensure `ProcessViewMessages.xlsm` is in the correct location
2. Run the generator:
   ```bash
   python JsonGenerator.py
   ```

## Creating an Executable (Optional)

If you need an executable version:

1. Install pyinstaller:
   ```bash
   pip install pyinstaller
   ```

2. Create the executable:
   ```bash
   pyinstaller --onefile --windowed JsonGenerator.py
   ```

   **Note**: The executable creation process may take 5-20 minutes due to pandas and JSON dependencies.

## Dependencies

Main dependencies:
- pandas
- openpyxl

See `requirements.txt` for the complete list.
