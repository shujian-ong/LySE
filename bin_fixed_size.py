import pandas as pd
import numpy as np
import os
from typing import List

# --- Configuration ---
# Update to the new input file path
INPUT_FILE_PATH = 'Input_file_path'

# Default bin size (can be adjusted by the user)
BIN_SIZE = 100

# The required sheet names from the input file
INPUT_SHEET_NUMBERS = ['63', '90', '121', '127', '130', '139', '143']
INPUT_SHEET_NAMES = [f'pSJ{num} Corrected' for num in INPUT_SHEET_NUMBERS]

# Define the 12 columns that represent the individual mutation counts.
# We assume they follow the standard naming convention used in your previous data.
MUTATION_COLS = [
    f'{ref}_to_{alt}_count'
    for ref in ['A', 'T', 'C', 'G']
    for alt in ['A', 'T', 'C', 'G'] if ref != alt
]

# The output file name will now include the bin size
OUTPUT_FILE_PATH = f'binned_mutation_averages_{BIN_SIZE}.xlsx'


# --- Core Logic Function ---

def calculate_binned_averages(df: pd.DataFrame, bin_size: int, mutation_cols: List[str]) -> pd.DataFrame:
    """
    Calculates the binned average of all 12 individual mutation types and
    uses the first position of the bin as the bin identifier.
    """
    # 1. Input validation for required columns
    missing_cols = [col for col in ['position'] + mutation_cols if col not in df.columns]
    if missing_cols:
        print(f"Skipping sheet due to missing mandatory columns: {missing_cols}")
        # Return an empty DataFrame or raise an error, depending on desired behavior
        return pd.DataFrame()

    # 2. Create the bin column
    # (position - 1) // bin_size gives a zero-indexed bin number
    df['bin_index'] = (df['position'] - 1) // bin_size

    # 3. Define aggregation rules: min for position, mean for mutation counts
    agg_rules = {
        'position': 'min',  # Get the first position in the bin
    }
    # Add all mutation columns to the aggregation rules, requesting the mean
    agg_rules.update({col: 'mean' for col in mutation_cols})

    # 4. Group by the bin index and aggregate
    binned_results = df.groupby('bin_index').agg(agg_rules).reset_index(drop=True)

    # 5. Rename the aggregated position column to reflect the output requirement
    binned_results.rename(columns={'position': 'first_position_in_bin'}, inplace=True)

    # 6. Select and order final columns: First position, then all averaged mutation counts
    final_cols = ['first_position_in_bin'] + mutation_cols
    final_output = binned_results[final_cols]

    # Rename the aggregated mutation columns for clarity (e.g., 'true_A_to_G_count' becomes 'true_A_to_G_count_avg')
    final_output.columns = ['first_position_in_bin'] + [f'{col}_avg' for col in mutation_cols]

    return final_output


# --- Main Execution Block ---

# 1. Update the output file name with the configured BIN_SIZE
OUTPUT_FILE_PATH = f'binned_mutation_averages_{BIN_SIZE}.xlsx'

# 2. Read all required sheets
print(f"\nReading data from {INPUT_FILE_PATH}...")
input_data_frames = {}

try:
    with pd.ExcelFile(INPUT_FILE_PATH) as xls:
        for sheet_name in INPUT_SHEET_NAMES:
            print(f"-> Reading sheet: {sheet_name}")
            # Ensure only the required columns are loaded to save memory
            required_cols = ['position'] + MUTATION_COLS
            df = pd.read_excel(xls, sheet_name=sheet_name)
            input_data_frames[sheet_name] = df

except Exception as e:
    print(f"Error reading input file. Check the path and sheet names. Error: {e}")
    # Print the absolute path for easier debugging if the file doesn't exist
    print(f"Attempted to access file at: {os.path.abspath(INPUT_FILE_PATH)}")
    exit()

# 3. Calculate binned averages for all sheets
binned_results = {}
print(f"\nCalculating binned averages for {BIN_SIZE} positions...")

for sheet_name, df in input_data_frames.items():
    print(f"-> Processing {sheet_name}...")
    binned_df = calculate_binned_averages(df, BIN_SIZE, MUTATION_COLS)
    if not binned_df.empty:
        # Use a cleaner sheet name for the output file
        output_sheet_name = f'{sheet_name.replace(" Corrected", "")} Binned Avg'
        binned_results[output_sheet_name] = binned_df
    else:
        print(f"Warning: No valid data generated for {sheet_name}. Skipping.")

# 4. Write the final DataFrames to the new Excel file in separate sheets
if binned_results:
    print(f"\nWriting final binned mutation averages to {OUTPUT_FILE_PATH}...")

    with pd.ExcelWriter(OUTPUT_FILE_PATH, engine='xlsxwriter') as writer:
        for sheet_name, df_out in binned_results.items():
            df_out.to_excel(writer, sheet_name=sheet_name, index=False)

    print("Binning complete. Final summary file created.")
    
    # Display a sample of the first generated result
    first_sheet_name = list(binned_results.keys())[0]
    print(f"\n--- Sample of '{first_sheet_name}' Output ---")
    print(binned_results[first_sheet_name].head())
else:
    print("\nNo sheets were successfully processed. Output file not created.")