import pandas as pd
import os

# File paths
source_folder = 'C:/Users/<Source Folder from step one>'
archive_folder = 'C:/Users/<Folder Path>/archive'
consolidated_file = 'C:/Users/<Folder Path>/Consolidated.xlsx'

# Read consolidated data file or create an empty dataframe if file doesn't exist
try:
    consolidated_data = pd.read_excel(consolidated_file, index_col=0)
except FileNotFoundError:
    consolidated_data = pd.DataFrame()

# Process each file in the source folder
for filename in os.listdir(source_folder):
    if filename.endswith('.xlsx') and filename != 'Consolidated.xlsx' and filename != 'MasterData.xlsx' and filename != 'HistoricalData.xlsx':
        file_path = os.path.join(source_folder, filename)

        # Read data from file
        data = pd.read_excel(file_path)

        # Remove extra spaces in column names
        data.columns = data.columns.str.strip()

        # Align columns
        if not consolidated_data.empty:
            data = data[consolidated_data.columns.tolist()]

        # Concatenate data to consolidated data
        consolidated_data = pd.concat([consolidated_data, data], ignore_index=True)

        # Move file to archive folder
        os.rename(file_path, os.path.join(archive_folder, filename))

# Save consolidated data to file
consolidated_data.to_excel(consolidated_file)
