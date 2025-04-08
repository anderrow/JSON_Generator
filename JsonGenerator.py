# Author: Anderrow
# Read the README.md file for more info.
# https://github.com/anderrow

import pandas as pd
import json
import os
import re

path = "../ProcessViewMessages.xlsm"
# Read all sheets into a dictionary of DataFrames
dfs = pd.read_excel(path, sheet_name=None)

# --------------------------------FILTER DATAFRAMES--------------------------------#
del dfs["ProtonView"]  # ProtonView sheet is not needed
del dfs["Pointer overview"]  # Pointer Overview is not needed

df_filtered_dict = {}  # Init

for sheet_name, df in dfs.items():
    # Replace NaN with an empty string before converting the third column to string
    df.iloc[:, 2] = df.iloc[:, 2].fillna('')

    # Convert the third column to string to avoid issues with non-string values
    df.iloc[:, 2] = df.iloc[:, 2].astype(str)

    # Filter dataframe erasing rows that ends with Undefine (Case sensitive is set to false, Nan values are skipped)
    df_filtered = df[~df.iloc[:, 2].str.contains(r'.*Undefined$', case=False, na=False)]

    # Filter dataframe for 'Unknown message'
    df_filtered2 = df_filtered[~df_filtered.iloc[:, 2].str.contains('Unkown message', case=False, na=False)]

    # Select the second column (Keys) and the third column (Messages)
    df_filtered_subset = df_filtered2.iloc[:, 1:9]

    # Erase rows if there is a NaN value (Not message found)
    df_cleaned = df_filtered_subset[df_filtered_subset.iloc[:, 1] != ""]

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

#Define the column name pattern need it with regex for avoid extrange file names. 
column_name_pattern = r'^[a-z]{2}-[A-Z]{2}$'

#Verify that the columns have a correct name
for sheet_name, df in df_filtered_dict.items():
    for column in range(1, df.shape[1]):
        column_name = df.columns[column] #Keep name of the column with index column  

        #verify that the column complies with the defined regex pattern
        if not re.match(column_name_pattern, column_name):
            raise ValueError(f"\n The column name '{column_name}' doesn't have the requiered format [a-z][a-z]-[A-Z][A-Z]")

# Loop to save the JSON files
for sheet_name, df in df_filtered_dict.items():
    # Create the folder in the parent directory if it doesn't exist
    output_folder = f"../JSON FILES/{sheet_name}"
    os.makedirs(output_folder, exist_ok=True)
    print("*"*50)
    
    #Loop to save each column of the dataframe
    for column in range(1, df.shape[1]):
        column_name = df.columns[column] #Keep name of the column with index column  
           
        column_name_simple = re.sub(r"^([a-zA-Z]+)-.*", r"\1", column_name)

        # Generate the name of the file based on the name of the excel sheet
        json_filename = os.path.join(output_folder, f"{column_name_simple}.json")

        # Convert the DataFrame to a dictionary with Column1 as the key and Column2 as the value
        df_to_json = dict(zip(df.iloc[:, 0], df.iloc[:, column]))

        with open(json_filename, 'w') as json_file:
            json.dump(df_to_json, json_file)  # Saving without indent for a compact format

        print(f"JSON file {json_filename} has been created")
# --------------------------------END OF GENERATE JSON FILES--------------------------------#
