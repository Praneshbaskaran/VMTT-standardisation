import pandas as pd
import os
import re
import glob

def read_file(file_path):
    """
    Reads a file based on its extension (csv or xlsx)
    
    Args:
        file_path (str): Path to the file to read
    Returns:
        dict: Dictionary of DataFrames for each sheet
    """
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.csv':
        return {'Sheet1': pd.read_csv(file_path)}
    elif file_extension == '.xlsx':
        return pd.read_excel(file_path, sheet_name=None)
    else:
        raise ValueError(f"Unsupported file format: {file_extension}. Please use .csv or .xlsx files.")

def save_file(df_dict, file_path):
    """
    Saves a dictionary of DataFrames to an Excel file
    
    Args:
        df_dict (dict): Dictionary of DataFrames to save
        file_path (str): Path where to save the file
    """
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.xlsx':
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in df_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        raise ValueError(f"Unsupported file format: {file_extension}. Please use .xlsx files.")

def remove_column_numbering(df):
    """
    Removes trailing decimal numbers from column names while preserving the data.
    Example: Converts columns like:
    'Custom field (Using Legal Entity (Application)).1', 
    'Custom field (Using Legal Entity (Application)).2'
    to:
    'Custom field (Using Legal Entity (Application))',
    'Custom field (Using Legal Entity (Application))'
    
    Args:
        df (pd.DataFrame): DataFrame with numbered columns
    Returns:
        pd.DataFrame: DataFrame with cleaned column names
    """
    # Create a mapping of old column names to new ones (without the trailing numbers)
    column_mapping = {}
    for col in df.columns:
        # Remove trailing .number pattern
        new_col = re.sub(r'\.\d+$', '', col)
        column_mapping[col] = new_col
    
    # Rename the columns while preserving the data
    df.columns = [column_mapping[col] for col in df.columns]
    return df

def standardize_sheet(base_df, new_df):
    """
    Standardizes a single sheet based on the base file structure.
    Uses the 11th row (index 9) as headers for standardization.
    Keeps rows 0-9 untouched. Adds missing headers in the 11th row and fills data columns with empty values.
    Rearranges data columns to match the standardized header order.
    """
    header_row = 9  # 0-based index for 11th row

    if len(base_df) <= header_row or len(new_df) <= header_row:
        print("WARNING: Not enough rows to standardize. Skipping sheet.")
        return new_df

    base_headers = list(base_df.iloc[header_row].fillna(''))
    new_headers = list(new_df.iloc[header_row].fillna(''))

    new_header_to_col = {h: i for i, h in enumerate(new_headers) if h}

    standardized_headers = []
    new_col_indices = []
    for h in base_headers:
        standardized_headers.append(h)
        new_col_indices.append(new_header_to_col[h] if h in new_header_to_col else None)

    num_cols = len(standardized_headers)

    # 1. Copy rows 0-9 as is, but pad or trim columns to match standardized_headers
    meta_rows = new_df.iloc[:header_row]
    if meta_rows.shape[1] < num_cols:
        # Pad with empty columns
        for i in range(meta_rows.shape[1], num_cols):
            meta_rows[i] = ''
    elif meta_rows.shape[1] > num_cols:
        # Trim extra columns
        meta_rows = meta_rows.iloc[:, :num_cols]
    meta_rows.columns = range(num_cols)

    # 2. Build the new header row (row 10, index 9)
    header_row_df = pd.DataFrame([standardized_headers], columns=range(num_cols))

    # 3. Rearrange/add columns for data rows (from row 11/index 10 onward)
    data_rows = new_df.iloc[header_row+1:]
    new_data = []
    for idx in range(len(data_rows)):
        row = data_rows.iloc[idx]
        new_row = []
        for col_idx in new_col_indices:
            if col_idx is not None and col_idx < len(row):
                new_row.append(row.iloc[col_idx])
            else:
                new_row.append('')
        new_data.append(new_row)
    if new_data:
        data_df = pd.DataFrame(new_data, columns=range(num_cols))
        standardized = pd.concat([meta_rows, header_row_df, data_df], ignore_index=True)
    else:
        standardized = pd.concat([meta_rows, header_row_df], ignore_index=True)

    standardized.columns = range(num_cols)
    return standardized

def set_column_headers(df, title):
    """
    Sets the first column header to the given title and the rest to empty strings.
    Args:
        df (pd.DataFrame): DataFrame to modify
        title (str): Title to set as the first column header
    Returns:
        pd.DataFrame: Modified DataFrame
    """
    if df.shape[1] == 0:
        return df
    new_columns = [title] + [''] * (df.shape[1] - 1)
    df = df.copy()
    df.columns = new_columns
    return df

def standardize_data(base_file, new_file):
    """
    Standardizes the structure of a new data file to match the base file.
    Handles multiple sheets and preserves Doc info and Summary sheets.
    Uses 11th row as headers for standardization in other sheets.
    """
    try:
        # Read both files with all sheets
        base_sheets = read_file(base_file)
        new_sheets = read_file(new_file)
        
        # Create dictionary for standardized sheets
        standardized_sheets = {}
        
        # Process each sheet
        for sheet_name in base_sheets.keys():
            if sheet_name in ['Doc info', 'Summary']:
                # Use the original sheet but update column headers
                sheet = new_sheets[sheet_name]
            else:
                # Standardize other sheets
                sheet = standardize_sheet(
                    base_sheets[sheet_name],
                    new_sheets[sheet_name]
                )
            # Set the column headers as requested
            sheet = set_column_headers(
                sheet,
                "Information Security Performance Indicator Reporting (ISPIRI)"
            )
            standardized_sheets[sheet_name] = sheet
            print(f"\nProcessing sheet: {sheet_name}")
            print(f"Base file columns: {len(base_sheets[sheet_name].columns)}")
            print(f"New file columns: {len(new_sheets[sheet_name].columns)}")
            print(f"Standardized columns: {len(standardized_sheets[sheet_name].columns)}")
        
        try:
            # Save the standardized data back to the original file
            save_file(standardized_sheets, new_file)
            print(f"\nFile {new_file} has been standardized successfully!")
            print("All sheets have their column headers replaced as requested.")
        except PermissionError:
            print(f"\nERROR: Could not save the file '{new_file}'.")
            print("Please make sure that:")
            print("1. The file is not currently open in Excel")
            print("2. You have write permissions for the file")
            print("3. The file is not set to read-only")
            print("\nPlease close the file if it's open and try again.")

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        if "Permission denied" in str(e):
            print("\nPlease make sure that:")
            print("1. The file is not currently open in Excel")
            print("2. You have write permissions for the file")
            print("3. The file is not set to read-only")
            print("\nPlease close the file if it's open and try again.")

def standardize_folder(base_file, folder_path):
    """
    Standardizes all .xlsx and .csv files in the given folder using the base file structure.
    Args:
        base_file (str): Path to the base file (xlsx or csv)
        folder_path (str): Path to the folder containing files to standardize
    """
    # Find all .xlsx and .csv files in the folder (not the base file itself)
    file_patterns = [os.path.join(folder_path, '*.xlsx'), os.path.join(folder_path, '*.csv')]
    files = []
    for pattern in file_patterns:
        files.extend(glob.glob(pattern))
    # Exclude the base file if it's in the same folder
    files = [f for f in files if os.path.abspath(f) != os.path.abspath(base_file)]
    if not files:
        print(f"No .xlsx or .csv files found in {folder_path}.")
        return
    for file_path in files:
        print(f"\n---\nProcessing file: {file_path}")
        try:
            standardize_data(base_file, file_path)
        except Exception as e:
            print(f"Failed to process {file_path}: {e}")

if __name__ == "__main__":
    import sys
    def get_input_path(prompt):
        path = input(prompt)
        return path.strip().strip('"')
    if len(sys.argv) == 3 and os.path.isdir(sys.argv[2]):
        # Usage: python standardize_data.py base_file folder_path
        base_file = sys.argv[1]
        folder_path = sys.argv[2]
        standardize_folder(base_file, folder_path)
    elif len(sys.argv) == 3:
        # Usage: python standardize_data.py base_file new_file
        base_file = sys.argv[1]
        new_file = sys.argv[2]
        standardize_data(base_file, new_file)
    else:
        print("No valid command-line arguments detected. Please provide the paths interactively.")
        base_file = get_input_path("Enter the path to the base file (.xlsx or .csv): ")
        target_path = get_input_path("Enter the path to the new file or folder: ")
        if os.path.isdir(target_path):
            standardize_folder(base_file, target_path)
        else:
            standardize_data(base_file, target_path)
