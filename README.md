# Excel Generator

Python utility that reads ProcessView CSV exports and generates a single normalized Excel workbook for ProtonView.

## What It Does

- Reads every `.csv` file from `../SOURCE FILES/`.
- Keeps the `Keys` column plus the language columns that come after it.
- Filters out rows whose base language text is empty or contains `Undefined` or `Unkown message`.
- Prefixes each key with the project name taken from the file name.
- Normalizes the language columns across all input files.
- Adds a `ProtonView` column required by the target format.
- Writes everything into one workbook: `../EXCEL FILES/AllProjects.xlsx`.

## Expected Input

Each CSV file must:

- live inside `../SOURCE FILES/`
- contain a `Keys` column
- contain at least one language column after `Keys`
- use language column names in the format `xx-YY`, for example `en-US` or `es-ES`

The project name is taken from the file name before the first underscore.

Example:

- `AVA_Messages.csv` -> project prefix `AVA_`
- `UFA_Alarms.csv` -> project prefix `UFA_`

## Processing Rules

- Files with missing or invalid headers are skipped.
- The first language column after `Keys` is treated as the base language.
- Rows are removed when the base language cell:
  - is empty
  - contains `Undefined`
  - contains `Unkown message`
- The language column `ta-GG` is excluded.
- If a key is not already prefixed with `<PROJECT>_`, the prefix is added automatically.
- Missing language columns are backfilled from `en-US` when available, otherwise from the first available language in that file.
- Legacy per-project workbooks such as `AVA.xlsx` are removed when the combined workbook is generated.

## Output

The script creates:

- workbook: `../EXCEL FILES/AllProjects.xlsx`
- sheet name: `ProtonView`

The generated sheet contains:

1. `ProtonView`
2. `Keys`
3. all normalized language columns sorted alphabetically

The `ProtonView` column is filled with:

```text
ProtonView_Resources_Controllers.IPlcResource
```

The workbook is written through a temporary file and then replaced atomically. If `AllProjects.xlsx` is open in Excel, the script stops with a permission error instead of partially overwriting the file.

## Project Structure

```text
_Proccesview/
|- SOURCE FILES/
|  `- *.csv
|- JSON_Generator/
|  |- ExcelGenerator.py
|  |- GenerateStrings.bat
|  |- README.md
|  |- requirements.txt
|  `- run_env.py
`- EXCEL FILES/
   `- AllProjects.xlsx
```

## Usage

Run from the repository folder:

```bash
python run_env.py
```

Windows shortcut:

```bat
GenerateStrings.bat
```

## Virtual Environment

`run_env.py` creates and uses a virtual environment outside the repository folder:

- Windows: `%LOCALAPPDATA%\PythonProjectVenvs\...`
- Linux/macOS: `$XDG_DATA_HOME/PythonProjectVenvs/...` or `~/.local/share/PythonProjectVenvs/...`

This avoids syncing a local `venv` through Nextcloud.

If an old in-repo `venv` still exists, `run_env.py` warns about it but does not use it.

## Dependencies

Main packages:

- `polars==0.18.14`
- `openpyxl`
- `xlsx2csv`

Install manually if needed:

```bash
pip install -r requirements.txt
```

## Optional Executable Build

```bash
pip install pyinstaller
pyinstaller --onefile --windowed ExcelGenerator.py
```
