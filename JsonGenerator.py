# Author: Anderrow
# JSON Generator using Polars and openpyxl
# https://github.com/anderrow

import time
import polars as pl
import openpyxl
import json, os, re

start_time = time.perf_counter()

# Path to the Excel file
path = "../ProcessViewMessages.xlsm"


# Get the sheet names, filtering out the ones you don't want
wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
sheet_names = [s for s in wb.sheetnames if s not in ["ProtonView", "Pointer overview"]]
wb.close()

dfs = {}

for sheet in sheet_names:
    try:
        # Read each sheet individually using Polars
        df = pl.read_excel(path, sheet_name=sheet)
        
        # Check if the sheet is empty or headers are invalid (None)
        if df.height == 0 or df.columns is None or any(col is None for col in df.columns):
            print(f"⚠️ Sheet '{sheet}' has invalid or missing headers, skipping...")
            continue
        
        # Store the DataFrame in the dictionary using the sheet name as key
        dfs[sheet] = df
    except Exception as e:
        print(f"⚠️ Error processing sheet '{sheet}': {e}")


# ---------------------------- FILTER DATAFRAMES ---------------------------- #

# Initialize an empty dict for filtered DataFrames
df_filtered_dict = {}

# Each item in the dictionary has a sheet name and a DataFrame
for sheet_name, df in dfs.items():
    # Replace nulls in the 3rd column with empty string and cast to UTF8 (convert to a unicode string)
    df = df.with_columns(
        pl.col(df.columns[2]).fill_null("").cast(pl.Utf8)
    )

    # Filter out rows ending with "Undefined" (case-insensitive)
    df_filtered = df.filter(
        ~pl.col(df.columns[2]).str.contains(r"(?i).*Undefined$")
    )

    # Filter out rows containing 'Unkown message' (yes, typo included)
    df_filtered2 = df_filtered.filter(
        ~pl.col(df.columns[2]).str.contains(r"(?i)Unkown message")
    )

    # Select columns 2 to 8
    df_filtered_subset = df_filtered2.select(df_filtered2.columns[1:9])

    # Remove rows where the second column is empty (No message found)
    df_cleaned = df_filtered_subset.filter(
        pl.col(df_filtered_subset.columns[1]) != ""
    )

    df_filtered_dict[sheet_name] = df_cleaned

    print("=" * 50)
    print(f"Sheet: {sheet_name}")
    print(f"Original length: {len(df)}")
    print(f"Filtered length: {len(df_cleaned)}")
    print(f"Data has been reduced by a {100 - len(df_cleaned) * 100 / len(df):.2f}%")
    print("=" * 50)

# --------------------------------END OF FILTER DATAFRAMES--------------------------------#

# -------------------------- VALIDATE COLUMN NAMES -------------------------- #

column_name_pattern = r'^[a-z]{2}-[A-Z]{2}$'

for sheet_name, df in df_filtered_dict.items():
    for i in range(1, df.shape[1]):  # shape returns a tuple with dimensions: [0] = number of rows, [1] = number of columns
        column_name = df.columns[i]
        if not re.match(column_name_pattern, column_name):
            raise ValueError(
                f"\n❌ Column name '{column_name}' does not match format [a-z][a-z]-[A-Z][A-Z]"
            )

#---------------------------END OF VALIDATE COLUMN NAMES -----------------------#

# --------------------------- GENERATE JSON FILES --------------------------- #

for sheet_name, df in df_filtered_dict.items():
    output_folder = f"../JSON FILES/{sheet_name}"
    os.makedirs(output_folder, exist_ok=True)

    print(f"\n📁 {sheet_name}")
    print("**" * len((f"*✅ JSON file created:{sheet_name.replace(os.sep, '/')}*")))

    for i in range(1, df.shape[1]):
        column_name = df.columns[i]  # Keep the name of the column with index i
        column_name_simple = re.sub(r"^([a-zA-Z]+)-.*", r"\1", column_name)

        # Generate the name of the file based on the name of the Excel sheet
        json_filename = os.path.join(output_folder, f"{column_name_simple}.json")

        # Convert the DataFrame to a dictionary with Column1 as the key and Column2 as the value
        keys = df.select(df.columns[0]).to_series().to_list()
        values = df.select(column_name).to_series().to_list()
        df_to_json = dict(zip(keys, values))

        with open(json_filename, 'w', encoding="utf-8") as json_file:
            json.dump(df_to_json, json_file)

        print(f"*✅ JSON file created: {json_filename.replace(os.sep, '/')} *")

    print("*" * len((f"*✅ JSON file created: {json_filename.replace(os.sep, '/')}  *")))

# --------------------------------END OF GENERATE JSON FILES--------------------------------#


end_time = time.perf_counter()
elapsed = end_time - start_time
print(f"\n Total execution time: {elapsed:.2f} seconds")