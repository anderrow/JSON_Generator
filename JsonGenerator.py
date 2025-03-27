# Author: Anderrow
# https://github.com/anderrow

import pandas as pd
import json
import os

path = "../ProcessViewMessages.xlsm"
# Read all sheets into a dictionary of DataFrames
dfs = pd.read_excel(path, sheet_name=None)

del dfs["ProtonView"]  # ProtonView sheet is not needed
del dfs["Pointer overview"]  # Pointer Overview is not needed

# --------------------------------FILTER DATAFRAMES--------------------------------#
df_filtered_dict = {}  # Init

for sheet_name, df in dfs.items():
    # Save the reference for the df in the dict

    # Replace NaN with an empty string before converting the third column to string
    df.iloc[:, 2] = df.iloc[:, 2].fillna('')

    # Convert the third column to string to avoid issues with non-string values
    df.iloc[:, 2] = df.iloc[:, 2].astype(str)

    # Filter dataframe erasing rows that ends with Undefine (Case sensitive is set to false, Nan values are skipped)
    df_filtered = df[~df.iloc[:, 2].str.contains(r'.*Undefined$', case=False, na=False)]

    # Filter dataframe for 'Unknown message'
    df_filtered2 = df_filtered[~df_filtered.iloc[:, 2].str.contains('Unkown message', case=False, na=False)]

    # Select the second column (Keys) and the third column (Messages)
    df_filtered_subset = df_filtered2.iloc[:, [1, 2]]

    # Erase rows if there is a NaN value (Not message found)
    df_cleaned = df_filtered_subset.dropna()

    # Keep the dataframe filtered
    df_filtered_dict[sheet_name] = df_cleaned

    # Print Before and after
    print("=" * 50)
    print(f"Sheet: {sheet_name}")
    print(f"Original length: {len(df)}")
    print(f"Filtered length: {len(df_cleaned)}")
    print(f"Data has been reduced by a {100 - len(df_cleaned) * 100 / len(df):.2f}%")
    print("=" * 50)

# --------------------------------END OF FILTER DATAFRAMES--------------------------------#

# --------------------------------GENERATE JSON FILES--------------------------------#
# Create the folder in the parent directory if it doesn't exist
output_folder = "../JSON FILES"
os.makedirs(output_folder, exist_ok=True)

# Loop to save the JSON files
for sheet_name, df in df_filtered_dict.items():
    # Generate the name of the file based on the name of the excel sheet
    json_filename = os.path.join(output_folder, f"{sheet_name}.json")

    # Convert the df to a dict and save it as a JSON
    df_to_json = df.to_dict(orient='records')

    with open(json_filename, 'w') as json_file:
        json.dump(df_to_json, json_file)  # Removed indent parameter for no line breaks

    print(f"JSON file {json_filename} has been created.")
# --------------------------------END OF GENERATE JSON FILES--------------------------------#
