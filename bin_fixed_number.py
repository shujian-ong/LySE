import pandas as pd
import numpy as np

# --- ⚙️ Configuration ---
INPUT_FILE_PATH = "Input_file_path"
OUTPUT_FILE_PATH = 'Output_file_path'

# --- 🎯 ADJUSTABLE PARAMETER ---
# The number of equally sized bins to divide the 'position' column into.
NUM_BINS = 20 # <-- Adjust this value as needed (e.g., 5, 10, 20, 50)
DECIMAL_PLACES = 10 # Desired decimal places for the resulting mean rates

# --- 🧬 Define Columns ---
# Mutation rate columns (H-S) that you want to calculate the mean for in each bin
MUTATION_RATE_COLS = [
    'A_to_T_count', 'A_to_C_count', 'A_to_G_count',
    'T_to_A_count', 'T_to_C_count', 'T_to_G_count',
    'C_to_A_count', 'C_to_T_count', 'C_to_G_count',
    'G_to_A_count', 'G_to_T_count', 'G_to_C_count'
]
# Base count/total count columns (for context)
COUNT_COLS = ['A_count', 'T_count', 'C_count', 'G_count', 'total_count']

# --- Core Logic Function ---

def calculate_binned_mean_rates(df: pd.DataFrame, num_bins: int, dp: int, sheet_name: str) -> pd.DataFrame:
    """
    Bins positions into equally sized intervals for a single DataFrame and calculates
    the mean mutation rate for each interval.
    """
    print(f"  -> Processing sheet: {sheet_name}")

    # 1. Data Cleaning and Preparation
    df['position'] = pd.to_numeric(df['position'], errors='coerce')
    df.dropna(subset=['position'], inplace=True)
    
    # Check if all required columns exist and the DataFrame is not empty
    required_cols = ['position'] + MUTATION_RATE_COLS
    if df.empty or not all(col in df.columns for col in required_cols):
        missing = [col for col in required_cols if col not in df.columns]
        if df.empty:
             print(f"  -> Skipping sheet '{sheet_name}': DataFrame is empty after cleaning.")
        else:
             print(f"  -> Skipping sheet '{sheet_name}': Missing required columns: {missing}")
        return None
    
    # 2. Binning the 'position' column (Equal Position Width)
    min_pos = df['position'].min()
    max_pos = df['position'].max()
    
    # Create bins for the position column
    bins = np.linspace(min_pos, max_pos, num_bins + 1)
    
    # Use pd.cut to assign each position to one of the bins
    df['position_bin'] = pd.cut(
        df['position'],
        bins=bins,
        include_lowest=True,
        right=True
    )
    
    # 3. Define aggregation dictionary
    agg_dict = {
        'position': ['min', 'max'],
        'ref_base': lambda x: '/'.join(x.astype(str).unique()),
        'position_bin': 'count'
    }
    
    # Sum counts, average rates
    for col in COUNT_COLS:
        agg_dict[col] = 'sum'
    for col in MUTATION_RATE_COLS:
        agg_dict[col] = 'mean'
    
    # 4. Group by bins and aggregate
    binned_results = df.groupby('position_bin').agg(agg_dict).reset_index(drop=True)
    
    # 5. Flatten multi-level column names
    binned_results.columns = [
        '_'.join(map(str, col)).strip('_') for col in binned_results.columns.values
    ]
    
    # 6. Final Formatting and Selection
    final_cols = {
        'position_min': 'Start_Position',
        'position_max': 'End_Position',
        'position_bin_count': 'Positions_in_Bin',
        'ref_base_<lambda>': 'Unique_Ref_Bases',
    }
    
    # Rename and round rate columns
    for col in MUTATION_RATE_COLS:
        final_cols[f'{col}_mean'] = f'Mean_{col.replace("_count", "_Rate")}'
    
    # Keep summed count columns for context
    for col in COUNT_COLS:
        final_cols[f'{col}_sum'] = f'Total_{col}'
        
    # Select and rename columns
    binned_results = binned_results.rename(columns=final_cols)[list(final_cols.values())]
    
    # Apply rounding to all mean rate columns
    mean_rate_cols = [col for col in binned_results.columns if col.startswith('Mean_')]
    binned_results[mean_rate_cols] = binned_results[mean_rate_cols].round(dp)
    
    return binned_results

# --- Main Execution Block ---

print(f"Reading all sheets from: **{INPUT_FILE_PATH}**")
try:
    # Read all sheets (sheet_name=None returns a dict of DataFrames)
    all_data = pd.read_excel(INPUT_FILE_PATH, sheet_name=None)
except FileNotFoundError:
    print(f"Error: File not found at {INPUT_FILE_PATH}. Please check the path.")
    exit()

print(f"Found {len(all_data)} sheets. Starting binning with **{NUM_BINS}** equal position bins.")

binned_results_dict = {}

# Process each sheet
for sheet_name, df in all_data.items():
    df_binned = calculate_binned_mean_rates(df.copy(), NUM_BINS, DECIMAL_PLACES, sheet_name)
    
    if df_binned is not None:
        binned_results_dict[sheet_name] = df_binned

# --- Write Results to a New Excel File ---

if not binned_results_dict:
    print("\nNo data was successfully binned. Check the input file and column names.")
else:
    print(f"\nWriting {len(binned_results_dict)} binned results to **{OUTPUT_FILE_PATH}**...")

    try:
        with pd.ExcelWriter(OUTPUT_FILE_PATH, engine='xlsxwriter') as writer:
            for sheet_name, df_out in binned_results_dict.items():
                # Ensure the sheet name is valid (max 31 chars, no invalid chars)
                output_sheet_name = f'Binned_{sheet_name[:24]}' # Truncate if too long
                df_out.to_excel(writer, sheet_name=output_sheet_name, index=False)
                
                # Auto-adjust column widths
                worksheet = writer.sheets[output_sheet_name]
                for idx, col in enumerate(df_out.columns):
                    max_len = max(df_out[col].astype(str).map(len).max(), len(col))
                    worksheet.set_column(idx, idx, max_len + 2)
                    
        print(f"\n✅ All sheets binned and saved successfully to **{OUTPUT_FILE_PATH}**.")
    except Exception as e:
        print(f"\nError writing to Excel: {e}")