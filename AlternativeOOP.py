import pandas as pd
import json
import os

class ExcelProcessor:
    def __init__(self, path, output_folder):
        # Initialize the class with the Excel file path and output folder
        self.path = path
        self.output_folder = output_folder

    def process_excel(self):
        # Read all sheets from the Excel file
        dfs = pd.read_excel(self.path, sheet_name=None)
        
        # Remove unnecessary sheets
        del dfs["ProtonView"]  # ProtonView sheet is not needed
        del dfs["Pointer overview"]  # Pointer Overview is not needed

        df_filtered_dict = {}  # Dictionary to store filtered DataFrames

        # Process each sheet
        for sheet_name, df in dfs.items():
            # Clean and prepare the data
            df.iloc[:, 2] = df.iloc[:, 2].fillna('')
            df.iloc[:, 2] = df.iloc[:, 2].astype(str)
            df_filtered = df[~df.iloc[:, 2].str.contains(r'.*Undefined$', case=False, na=False)]
            df_filtered2 = df_filtered[~df_filtered.iloc[:, 2].str.contains('Unkown message', case=False, na=False)]
            df_filtered_subset = df_filtered2.iloc[:, [1, 2]]
            df_cleaned = df_filtered_subset.dropna()
            df_cleaned = df_filtered_subset[df_filtered_subset.iloc[:, 1] != ""]

            # Store the filtered DataFrame in the dictionary
            df_filtered_dict[sheet_name] = df_cleaned

            # Print stats for each sheet
            self.print_stats(sheet_name, len(df), len(df_cleaned))

        return df_filtered_dict

    def print_stats(self, sheet_name, original_len, filtered_len):
        # Print statistics about the processing of each sheet
        print("=" * 50)
        print(f"Sheet: {sheet_name}")
        print(f"Original length: {original_len}")
        print(f"Filtered length: {filtered_len}")
        print(f"Data has been reduced by a {100 - filtered_len * 100 / original_len:.2f}%")
        print("=" * 50)

    def generate_json_files(self, df_filtered_dict):
        # Create the output folder if it doesn't exist
        os.makedirs(self.output_folder, exist_ok=True)

        # Save the filtered DataFrames as JSON files
        for sheet_name, df in df_filtered_dict.items():
            json_filename = os.path.join(self.output_folder, f"{sheet_name}.json")
            df_to_json = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))

            # Save the JSON file
            with open(json_filename, 'w') as json_file:
                json.dump(df_to_json, json_file)

            print(f"JSON file {json_filename} has been created.")

# ===================================
# Example usage of the ExcelProcessor class

# Path to the Excel file and output folder for JSON files
path = "../ProcessViewMessages.xlsm"
output_folder = "../JSON FILES"

# Create an instance of the ExcelProcessor class
processor = ExcelProcessor(path, output_folder)

# Process the Excel file
df_filtered_dict = processor.process_excel()

# Generate JSON files from the filtered DataFrames
processor.generate_json_files(df_filtered_dict)
