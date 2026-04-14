# Author: Anderrow
# JSON Generator using Polars
# https://github.com/anderrow

import glob
import json
import os
import re
import time

import polars as pl

start_time = time.perf_counter()

# Path to the CSV files
source_folder = "../SOURCE FILES/"
excel_files = glob.glob(os.path.join(source_folder, "*.csv"))
dfs = {}


def get_exportable_columns(df, sheet_name):
    export_columns = []
    skipped_columns = 0

    for column_name in df.columns[1:]:
        normalized_name = "" if column_name is None else str(column_name).strip()

        if not normalized_name or re.fullmatch(r"_duplicated_\d+", normalized_name):
            skipped_columns += 1
            continue
        export_columns.append(column_name)

    if skipped_columns:
        print(
            f"Warning: skipping {skipped_columns} empty header column(s) in {sheet_name}."
        )

    return export_columns


# -------------------------- READ ALL CSV FILES -------------------------- #
for path in excel_files:
    sheet_name = os.path.splitext(os.path.basename(path))[0]

    try:
        df = pl.read_csv(path)

        if df.height == 0 or df.columns is None or any(col is None for col in df.columns):
            print(f"Warning: file '{sheet_name}' has invalid or missing headers, skipping...")
            continue

        dfs[sheet_name] = df
        print(
            f"File '{sheet_name}' read successfully with {df.height} rows and {df.width} columns."
        )
    except Exception as e:
        print(f"Warning: error processing sheet '{sheet_name}': {e}")

print(f"current time {time.perf_counter() - start_time:.2f} seconds")

# ---------------------------- FILTER DATAFRAMES ---------------------------- #
df_filtered_dict = {}

for sheet_name, df in dfs.items():
    df = df.with_columns(pl.col(df.columns[2]).fill_null("").cast(pl.Utf8))

    df_filtered = df.filter(~pl.col(df.columns[2]).str.contains(r"(?i).*Undefined"))
    df_filtered2 = df_filtered.filter(
        ~pl.col(df.columns[2]).str.contains(r"(?i)Unkown message")
    )

    df_filtered_subset = df_filtered2.select(df_filtered2.columns[1:9])
    df_cleaned = df_filtered_subset.filter(pl.col(df_filtered_subset.columns[1]) != "")

    df_filtered_dict[sheet_name] = df_cleaned

    print("=" * 50)
    print(f"Sheet: {sheet_name}")
    print(f"Original length: {len(df)}")
    print(f"Filtered length: {len(df_cleaned)}")
    print(f"Data has been reduced by a {100 - len(df_cleaned) * 100 / len(df):.2f}%")
    print("=" * 50)

# -------------------------- VALIDATE COLUMN NAMES -------------------------- #
column_name_pattern = r"^[a-z]{2}-[A-Z]{2}$"
export_columns_by_sheet = {}

for sheet_name, df in df_filtered_dict.items():
    export_columns = get_exportable_columns(df, sheet_name)
    export_columns_by_sheet[sheet_name] = export_columns

    for column_name in export_columns:
        if not re.match(column_name_pattern, column_name):
            raise ValueError(
                f"\nColumn name '{column_name}' of {sheet_name} does not match format [a-z][a-z]-[A-Z][A-Z]"
            )

# --------------------------- GENERATE JSON FILES --------------------------- #
for sheet_name, df in df_filtered_dict.items():
    output_folder = f"../JSON FILES/{sheet_name}"
    os.makedirs(output_folder, exist_ok=True)

    export_columns = export_columns_by_sheet.get(sheet_name, [])
    if not export_columns:
        print(f"\nSkipping {sheet_name}: no exportable columns found.")
        continue

    print(f"\nFolder: {sheet_name}")

    keys = df.select(df.columns[0]).to_series().to_list()
    last_json_filename = None

    for column_name in export_columns:
        column_name_simple = re.sub(r"^([a-zA-Z]+)-.*", r"\1", column_name)
        json_filename = os.path.join(output_folder, f"{column_name_simple}.json")

        values = df.select(column_name).to_series().to_list()
        df_to_json = dict(zip(keys, values))

        with open(json_filename, "w", encoding="utf-8") as json_file:
            json.dump(df_to_json, json_file)

        last_json_filename = json_filename
        print(f"JSON file created: {json_filename.replace(os.sep, '/')}")

    if last_json_filename is not None:
        print(f"Completed: {sheet_name}")

end_time = time.perf_counter()
elapsed = end_time - start_time
print(f"\nTotal execution time: {elapsed:.2f} seconds")
