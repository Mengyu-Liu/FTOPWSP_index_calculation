import os
import pandas as pd

# There are three water resources prediction models (CwatM,H08, CR-GLOBWB) utilized, each motivated by three GCMs (GFDL-ESM2M, HadGEM2-ES, MIROC5)
# Take CwatM model motivated by GFDL-ESM2M GCM model as an example
# Note that the original water supple and demand data have been transformed from NC files to Excel files
# Specify your directory path, which stores the water supply and demand data from CwatM-GFDL-ESM2M coupled model

directory = 'E:/paper_editing/Cwatm_excel_gfdl'

# Iterate through all files in the directory
for filename in os.listdir(directory):
    if filename.endswith(".xlsx"):
        # Construct the complete file path
        excel_file_path = os.path.join(directory, filename)

        # Read Excel spreadsheet data
        df = pd.read_excel(excel_file_path, header=0)
        df = df.iloc[:, 1:]

        # Calculate the max and min of all data
        max_value = df.stack().max()
        min_value = df.stack().min()
        print(max_value)
        print(min_value)
        # Adjust the normalization coefficient based on keywords in the filename
        if 'adomww' in filename:
            normalization_factor = 0.069
        elif 'ainww' in filename:
            normalization_factor = 0.08
        elif 'airrww' in filename:
            normalization_factor = 0.221
        elif 'evap' in filename:
            normalization_factor = 0.072
        elif 'qg' in filename:
            normalization_factor = 0.057
        elif 'qr' in filename:
            normalization_factor = 0.071
        elif 'qs' in filename:
            normalization_factor = 0.167
        elif 'qsb' in filename:
            normalization_factor = 0.057
        elif 'soilmoist' in filename:
            normalization_factor = 0.056
        elif 'swe' in filename:
            normalization_factor = 0.150
        else:
            normalization_factor = 1.0  # If no keywords are matched, default to 1.0

        # Normalize data
        df_normalized = ((df - min_value) / (max_value-min_value)) * normalization_factor
        max_value_norm = df_normalized.stack().max()
        min_value_norm = df_normalized.stack().min()

        # Calculate the squared difference
        df_minus_max = df_normalized - max_value_norm
        df_minus_max_square = df_minus_max ** 2
        df_minus_min = df_normalized - min_value_norm
        df_minus_min_square = df_minus_min ** 2

        # Save the squared difference results as a new Excel file
        output_filename_max = f'{filename[:-5]}_minus_max_square.xlsx'
        output_filename_min = f'{filename[:-5]}_minus_min_square.xlsx'
        output_path_max = os.path.join('E:/paper_editing/Cwatm_calculation', output_filename_max)
        output_path_min = os.path.join('E:/paper_editing/Cwatm_calculation', output_filename_min)
        df_minus_max_square.to_excel(output_path_max, index=False)
        df_minus_min_square.to_excel(output_path_min, index=False)

        print(f"The calculation results for {filename} have been saved as {output_filename_max} and {output_filename_min}")

print("All files have been processed and saved")

# The computation of other coupled models are conducted under the similar coding framework



# For each water resources prediction model, the "squared results" from three GCMs need to be averaged into "squared roots" under fuzzy criteria
# Take CwatM model as an example
# Specify the directory path
directory = 'E:/paper_editing/Cwatm_calculation'
output_direct = 'E:/paper_editing/Cwatm_squaredroot'

# Store the parts of the processed filenames that do not include the GCM model name (e.g., 'gfdl'), along with the corresponding list of files
processed_files = {}

# Iterate through all files in the directory
for filename in os.listdir(directory):
    if filename.endswith(".xlsx"):
        # Extract key information from the filename
        parts = filename.split('_')
        if len(parts) >= 3:
            keyword = parts[1]
            other_parts = '_'.join(parts[:1] + parts[2:])  # Remove the 'gfdl' part (for example, from 'rcp26_gfdl_adomww_minus_max_square.xlsx' to 'rcp26_adomww_minus_max_square.xlsx')."

            # Check if files with the same remaining sections have already been processed
            if other_parts not in processed_files:
                processed_files[other_parts] = [filename]
            else:
                processed_files[other_parts].append(filename)

# For each group of files with the same remaining sections, perform a merge calculation operation
for other_parts, filenames in processed_files.items():
    # "Only perform calculation operations on groups containing at least three files
    if len(filenames) >= 3:
        print(f"Processing file group: {filenames}")
        dfs = []
        for filename in filenames:
            file_path = os.path.join(directory, filename)
            df = pd.read_excel(file_path, header=0)
            #df.replace(0, pd.NA, inplace=True)
            df.fillna(0, inplace=True)
            df.reset_index(inplace=True)
            dfs.append(df)

        # Merge files
        merged_df = pd.concat(dfs, axis=1, keys=[f'trial{i + 1}' for i in range(len(filenames))])
        print(f"Merging of {filenames} files completed")

        # Calculate average
        average_df = merged_df.groupby(axis=1, level=1).mean()
        print(f"Average calculation for {filenames} files completed")

        # Squared roots
        square_root_df = average_df ** 0.5
        print(f"Square root calculation for {filenames} files completed")

        # Construct output path
        output_filename = f"{other_parts}_root_merge.xlsx"
        output_path = os.path.join(output_direct, output_filename)

        # Save the computation results
        square_root_df.to_excel(output_path, index=False)
        print(f"The calculation results for {other_parts} and its related files have been saved as {output_filename}")
    else:
        print(f"The {other_parts} group does not have enough files to perform the calculation")

print("All files have been processed and saved")
#'''


# The final step is to merge the "square roots" to calculate the FTOPWSP index
# Still, take CwatM model under the rcp2.6 scenario as an example
# The final FTOPWSP index value is the average of the FTOPWSP indices values calculated by the three water resources prediction models


# Specify the directory path
file_paths_max = [
    f'E:/paper_editing/Cwatm_squaredroot/rcp26_{var}_minus_max_square.xlsx_root_merge.xlsx'
    for var in ['adomww', 'ainww', 'airrww', 'evap', 'qg', 'qr', 'qs', 'qsb', 'soilmoist', 'swe']
]

file_paths_min = [
    f'E:/paper_editing/Cwatm_squaredroot/rcp26_{var}_minus_min_square.xlsx_root_merge.xlsx'
    for var in ['adomww', 'ainww', 'airrww', 'evap', 'qg', 'qr', 'qs', 'qsb', 'soilmoist', 'swe']
]

# storing DataFrames
dfs_max = []
dfs_min = []

# Read all Excel files and replace missing values with 0
for file_path in file_paths_max:
    df = pd.read_excel(file_path, header=0).fillna(0)
    dfs_max.append(df)

for file_path in file_paths_min:
    df = pd.read_excel(file_path, header=0).fillna(0)
    dfs_min.append(df)

# Calculate FTOPWSP
df_FTOPWSP = (sum(dfs_max) / (sum(dfs_max) + sum(dfs_min) ))
print("The calcualtion of FTOPWSP26 is finished")

# Save calculation results
df_FTOPWSP.to_excel('E:/paper_editing/Cwatm_FTOPWSP/RCP26_FTOPWSP.xlsx', index=False)
print("The results from Cwatm model under RCP26 scenario have been saved")
