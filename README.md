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

## Sync To ProcessView

This repository also includes a deployment helper that can:

- run `JsonGenerator.py`
- find the active CSV exports in `../SOURCE FILES/`
- map each project code to the correct `processview` translation folder
- run `git pull --ff-only` in the `processview` repo
- copy the generated JSON files into the destination folder
- optionally run `git add`, `git commit`, and `git push`

Current project mapping:

- `AVA` -> `hmi/translations/ava`
- `CAR` -> `hmi/translations/crevin`
- `DSM` -> `hmi/translations/dsm`
- `UFA` -> `hmi/translations/ufa`
- `VILO` -> `hmi/translations/vilo`

The script expects the source CSV names to follow one of these patterns:

- `PROJECTNAME_VERSION.csv`
- `PROJECTNAME.csv`

Examples:

- `AVA_V7.csv`
- `UFA.csv`

Run a preview without changing anything:

```bash
python run_processview_sync.py --dry-run
```

Run the full sync using the default detected `processview` clone:

```bash
python run_processview_sync.py
```

Run the full sync and push the translation update:

```bash
python run_processview_sync.py --push
```

If the `processview` repo is in a custom location, pass it explicitly or set an environment variable:

```bash
python run_processview_sync.py --processview-repo "C:\Users\Ander\Documents\GitHub\processview"
```

```powershell
$env:PROCESSVIEW_REPO="C:\Users\Ander\Documents\GitHub\processview"
python run_processview_sync.py --push
```

Useful flags:

- `--project AVA` sync only one project
- `--skip-generate` reuse the current JSON files
- `--skip-pull` skip `git pull --ff-only`
- `--allow-dirty` allow copying into a repo with local changes when `--skip-pull` is used
- `--prune` delete supported language files in `processview` that are missing from the generated folder
- `--commit-message "..."` override the automatic commit message

Safety checks:

- If more than one CSV exists for the same project, the sync stops to avoid ambiguity.
- If the `processview` repo has local changes, pull is blocked.
- Only the mapped translation folders are staged and committed.
- If a supported language file is missing in the generated folder, the sync uses `en.json` as a fallback for that language.
- By default the sync only overwrites supported ProcessView language files (`bg`, `da`, `de`, `en`, `fr`, `nl`, `uk`) and leaves missing target files untouched unless `--prune` is used.

## GitLab Automation

Yes, this can run without depending on anyone's PC, but the CSV inputs must be reachable by GitLab CI.

Recommended setup:

1. Put the source CSV files in a Git repo that the pipeline can read.
2. Run the generator and sync in GitLab CI.
3. Let the pipeline push the updated translation JSON into `processview`.

Important limitation:

- The current local folder `../SOURCE FILES/` is outside this git repository, so GitLab cannot see it unless you move those CSV files into a repo or clone a separate source repo in CI.
- A separate repo for generated JSON is usually not needed. The important repo is the one that stores the source CSV files.

This repository now includes a ready-to-adapt GitLab pipeline template in [`.gitlab-ci.yml`](C:\Users\Ander\Nextcloud\Amabox share folder\_Proccesview\JSON_Generator\.gitlab-ci.yml) and [ci/gitlab_sync_processview.sh](C:\Users\Ander\Nextcloud\Amabox share folder\_Proccesview\JSON_Generator\ci\gitlab_sync_processview.sh).

The pipeline supports two modes:

- If the CSV files are committed inside the same GitLab repo under `SOURCE FILES/`, it uses them directly.
- If the CSV files live in another GitLab repo, set `SOURCE_REPO_URL` and optionally `SOURCE_REPO_REF` and `SOURCE_REPO_SUBDIR`, and the pipeline clones that repo before generating JSON.

GitLab CI variables you need:

- `PROCESSVIEW_WRITE_TOKEN`: required, token with write access to `mes-tools/processview`
- `PROCESSVIEW_REPO_URL`: optional, defaults to `https://gitlab.com/mes-tools/processview.git`
- `PROCESSVIEW_TARGET_BRANCH`: optional, defaults to `develop_ibo`
- `GIT_BOT_NAME`: optional commit author name
- `GIT_BOT_EMAIL`: optional commit author email
- `SOURCE_REPO_URL`: optional if the CSV files are in another repo
- `SOURCE_REPO_TOKEN`: optional read token for the source repo when `CI_JOB_TOKEN` is not enough
- `SOURCE_REPO_REF`: optional branch or tag for the source repo
- `SOURCE_REPO_SUBDIR`: optional path inside the source repo, defaults to `SOURCE FILES`

The generator now also accepts these environment variables, which is what makes CI portable:

- `JSON_GENERATOR_SOURCE_DIR`
- `JSON_GENERATOR_OUTPUT_DIR`

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
