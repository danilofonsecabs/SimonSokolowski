import os
import pandas as pd
from typing import List
import thefuzz

def extract_values_from_folder(folder_path: str, output_file: str):
    # List to store the extracted values
    extracted_values = []

    target_strings = ["22", "Service Listing Descripton", "service", "Description"]
    # Iterate through all Excel files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(('.xls', '.xlsx')):
            file_path = os.path.join(folder_path, file_name)
            # Load the workbook
            xls = pd.ExcelFile(file_path)

            # Iterate through the sheets and look for the ones containing "summary" and "22"
            for sheet_name in xls.sheet_names:
                if "summary" not in sheet_name.lower() and "download" not in sheet_name.lower() and "22" in sheet_name:
                    # Read the specific sheet
                    sheet_data = pd.read_excel(xls, sheet_name=sheet_name)

                    # Flags to check if "22" and "Total" are found
                    found_22 = False
                    found_total = False

                    # Iterate through columns 1 and 2, then 0 and 2
                    for col_idx in [(1, 2), (0, 2)]:
                        # Flag to start extracting values
                        start_extracting = False

                        # Iterate through the rows and look for the specific pattern
                        for index, row in sheet_data.iterrows():
                            # Use the appropriate column index
                            cell_value = str(row[sheet_data.columns[col_idx[0]]])
                            if any(target in cell_value for target in target_strings):
                                found_22 = True
                                start_extracting = True
                            if start_extracting and "Total" in str(cell_value):
                                found_total = True
                                break
                            if start_extracting and pd.notnull(cell_value) and pd.notnull(row[sheet_data.columns[col_idx[1]]]):
                                # Include the workbook name as one of the values
                                extracted_values.append((file_name, f"{cell_value}", f"{row[sheet_data.columns[col_idx[1]]]}"))
                        if found_22 and found_total:
                            break

                    # If "22" or "Total" not found, print the file name
                    if not found_22 or not found_total:
                        print(f"Skipped file: {file_name}")

    # Convert extracted values to DataFrame
    df_output = pd.DataFrame(extracted_values, columns=['Workbook Name', 'Description', 'Value'])

    # Write the extracted values to an Excel file
    df_output.to_excel(output_file, index=False)


# Specify the folder path where your Excel files are located
folder_path = r"C:\Users\dfonseca\Documents\Simon Sokolowski\Client files"
# Specify the output file path
output_file_path = r"C:\Users\dfonseca\Documents\Simon Sokolowski\extraction.xlsx"

# Extract the values and save them to the output file
extract_values_from_folder(folder_path, output_file_path)

os.startfile(output_file_path)