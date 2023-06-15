import pandas as pd
import glob

# Set the path to the directory containing the Excel files
excel_files_path = '/path/to/excel/files/*.xlsx'

# Initialize an empty list to store the dataframes
dataframes = []

# Iterate over each Excel file in the directory
for file in glob.glob(excel_files_path):
    # Read the Excel file into a pandas dataframe
    df = pd.read_excel(file)
    # Append the dataframe to the list
    dataframes.append(df)

# Concatenate all dataframes into a single dataframe
consolidated_df = pd.concat(dataframes)

# Set the path to save the CSV file
csv_file_path = '/path/to/save/result.csv'

# Save the consolidated dataframe to a CSV file
consolidated_df.to_csv(csv_file_path, index=False)

# Print a success message
print(f"Data extracted, consolidated, and saved to {csv_file_path}.")
