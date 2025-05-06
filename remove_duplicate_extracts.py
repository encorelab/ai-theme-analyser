# remove_duplicate_extracts.py

import pandas as pd
import argparse
from pathlib import Path
import sys
# import datetime # No longer needed

# --- Constants ---
DUPLICATES_SHEET_NAME = "Duplicate Extracts"
ALL_MATCHES_SHEET_NAME = "All Matches" # Needed for lookup
CODINGS_SHEET_NAME = "Merged Codings" # For --coded mode
OUTPUT_SUFFIX = "_cleaned"
# ---

# --- Helper functions ---
def get_file_path_from_arg(path_str, file_desc):
    # ... (implementation from previous version) ...
    file_path = Path(path_str).resolve();
    if not file_path.is_file(): print(f"Error: {file_desc} file not found: '{path_str}'"); sys.exit(1);
    return file_path

def parse_comma_sep_string(input_str, lower=False):
    # ... (implementation from previous version) ...
    if pd.isna(input_str) or not isinstance(input_str, str) or not input_str.strip(): return []
    items = [item.strip() for item in input_str.split(',') if item.strip()]
    if lower: items = [item.lower() for item in items]
    return items

# remove_codes function is NO LONGER NEEDED for --coded mode

# --- MODIFIED run_coded_removal function ---
def run_coded_removal(duplicates_file_path, target_file_path):
    """Removes entire rows from the Merged Codings sheet based on lookup."""
    # Changed description: Now removes ROWS, not just codes.
    print(f"\n--- Running Coded Extract ROW Removal (using original_excerpt lookup) ---")
    print(f"Reading duplicates info & All Matches from: {duplicates_file_path}")
    print(f"Reading target codings from: {target_file_path} (Sheet: '{CODINGS_SHEET_NAME}')")

    log_messages = []
    indices_to_drop = set() # Use set for unique indices

    try:
        df_duplicates = pd.read_excel(duplicates_file_path, sheet_name=DUPLICATES_SHEET_NAME)
        df_all_matches = pd.read_excel(duplicates_file_path, sheet_name=ALL_MATCHES_SHEET_NAME)
        df_target = pd.read_excel(target_file_path, sheet_name=CODINGS_SHEET_NAME)
        target_sheet_name = CODINGS_SHEET_NAME
    except FileNotFoundError: print(f"Error: One or more input files not found."); sys.exit(1)
    except ValueError as e: print(f"Error reading sheets. Check names ('{DUPLICATES_SHEET_NAME}', '{ALL_MATCHES_SHEET_NAME}', '{CODINGS_SHEET_NAME}'): {e}"); sys.exit(1)
    except Exception as e: print(f"An unexpected error occurred reading files: {e}"); sys.exit(1)

    # --- Validate columns ---
    required_dup_cols = ['Erroneous Filenames', 'matched_example', 'associated_codes']
    required_all_matches_cols = ['filename', 'matched_example', 'original_excerpt']
    # Target needs filename and excerpt for matching. Codings column no longer strictly needed for removal logic.
    required_target_cols = ['filename', 'excerpt']
    if not all(col in df_duplicates.columns for col in required_dup_cols): print(f"Error: Missing required columns in '{DUPLICATES_SHEET_NAME}'. Need: {required_dup_cols}"); sys.exit(1)
    if not all(col in df_all_matches.columns for col in required_all_matches_cols): print(f"Error: Missing required columns in '{ALL_MATCHES_SHEET_NAME}'. Need: {required_all_matches_cols}"); sys.exit(1)
    if not all(col in df_target.columns for col in required_target_cols): print(f"Error: Missing required columns in '{CODINGS_SHEET_NAME}'. Need: {required_target_cols}"); sys.exit(1)

    # --- Preprocess for matching ---
    for df in [df_duplicates, df_all_matches, df_target]:
        for col in ['filename', 'matched_example', 'original_excerpt', 'excerpt', 'codings', 'code']: # Common cols
            if col in df.columns:
                # Use fillna('') BEFORE astype(str) to handle non-string types robustly
                df[col] = df[col].fillna('').astype(str).str.strip()


    print("Identifying rows for removal...")
    processed_rows = 0
    # Iterate through the rows needing action based on the duplicates file
    for index, dup_row in df_duplicates.iterrows():
        processed_rows += 1
        print(f"  Processing Duplicate Extracts row {index + 1}/{len(df_duplicates)}...", end='\r')

        erroneous_files = parse_comma_sep_string(dup_row['Erroneous Filenames'])
        dup_matched_example = dup_row['matched_example'] # Already preprocessed
        # associated_codes = set(parse_comma_sep_string(dup_row['associated_codes'])) # No longer needed for row removal logic

        if not erroneous_files or not dup_matched_example: # Removed check for codes_to_remove
            continue # Skip row if essential info is missing

        for filename in erroneous_files:
            filename_lower = filename.lower() # Prepare for case-insensitive compare
            # --- Intermediate Lookup in All Matches ---
            all_matches_lookup = (df_all_matches['filename'].str.lower() == filename_lower) & \
                                 (df_all_matches['matched_example'] == dup_matched_example) # Assume matched_example is already preprocessed
            matching_orig_excerpts = df_all_matches.loc[all_matches_lookup, 'original_excerpt'].unique()
            # ---

            if len(matching_orig_excerpts) == 0:
                continue # Skip filename if no original excerpt found

            # --- Target Lookup using Original Excerpt(s) ---
            for original_excerpt in matching_orig_excerpts:
                if not original_excerpt: continue # Skip if original excerpt is empty

                original_excerpt_lower = original_excerpt.lower() # Prepare for case-insensitive compare

                # Find target rows matching filename and original_excerpt (case-insensitive)
                target_match_condition = (df_target['filename'].str.lower() == filename_lower) & \
                                         (df_target['excerpt'].str.lower() == original_excerpt_lower)
                target_indices = df_target[target_match_condition].index

                # If matches found, add their indices to the set for deletion
                if not target_indices.empty:
                     for target_idx in target_indices:
                          if target_idx not in indices_to_drop: # Log only once per identified row
                               log_messages.append(f"Row {target_idx+2}: Identified for deletion. File='{filename}', OriginalExcerpt='{original_excerpt[:60]}...' matched erroneous condition.")
                     indices_to_drop.update(target_indices) # Add all found indices


    print(f"\nProcessing complete. Identified {len(indices_to_drop)} unique rows for deletion.")

    # --- Perform Deletion ---
    if indices_to_drop:
        df_target_modified = df_target.drop(index=list(indices_to_drop)).reset_index(drop=True)
        print(f"Removed {len(indices_to_drop)} rows.")

        # --- Save the modified DataFrame ---
        output_path = target_file_path.parent / f"{target_file_path.stem}{OUTPUT_SUFFIX}{target_file_path.suffix}"
        print(f"\nSaving cleaned data to: {output_path}")
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                 df_target_modified.to_excel(writer, sheet_name=target_sheet_name, index=False)
            print("File saved successfully.")
        except Exception as e: print(f"Error saving cleaned file: {e}")
    else:
        print("\nNo rows were identified for deletion. Original target file remains unchanged.")
        # df_target_modified = df_target # No need to assign if no changes

    # --- Log Summary ---
    print("\n--- Coded ROW Removal Summary ---") # Updated summary title
    print(f"Total rows removed: {len(indices_to_drop)}")
    # for msg in log_messages: print(msg) # Optionally print detail logs
    print("---------------------------------")


# --- run_magnitude_removal function (Keep as is from previous version) ---
# It already performs row removal correctly based on its criteria.
def run_magnitude_removal(duplicates_file_path, target_file_path):
    """Removes entire rows from the Magnitude Codes sheet using original_excerpt lookup."""
    print(f"\n--- Running Magnitude Code Row Removal (using original_excerpt lookup) ---")
    print(f"Reading duplicates info & All Matches from: {duplicates_file_path}")
    print(f"Reading target magnitude codes from: {target_file_path} (First Sheet)")
    log_messages = []
    indices_to_drop = set()
    try:
        df_duplicates = pd.read_excel(duplicates_file_path, sheet_name=DUPLICATES_SHEET_NAME)
        df_all_matches = pd.read_excel(duplicates_file_path, sheet_name=ALL_MATCHES_SHEET_NAME)
        excel_file = pd.ExcelFile(target_file_path)
        if not excel_file.sheet_names: print(f"Error: Target file '{target_file_path}' has no sheets."); sys.exit(1)
        target_sheet_name = excel_file.sheet_names[0]; df_target = excel_file.parse(target_sheet_name)
        print(f"Reading from target sheet: '{target_sheet_name}'")
    except FileNotFoundError: print(f"Error: One or more input files not found."); sys.exit(1)
    except ValueError as e: print(f"Error reading sheets. Check names ('{DUPLICATES_SHEET_NAME}', '{ALL_MATCHES_SHEET_NAME}', target): {e}"); sys.exit(1)
    except Exception as e: print(f"An unexpected error occurred reading files: {e}"); sys.exit(1)

    required_dup_cols = ['Erroneous Filenames', 'matched_example', 'associated_codes']
    required_all_matches_cols = ['filename', 'matched_example', 'original_excerpt']
    required_target_cols = ['filename', 'excerpt', 'code']
    if not all(col in df_duplicates.columns for col in required_dup_cols): print(f"Error: Missing required columns in '{DUPLICATES_SHEET_NAME}'. Need: {required_dup_cols}"); sys.exit(1)
    if not all(col in df_all_matches.columns for col in required_all_matches_cols): print(f"Error: Missing required columns in '{ALL_MATCHES_SHEET_NAME}'. Need: {required_all_matches_cols}"); sys.exit(1)
    if not all(col in df_target.columns for col in required_target_cols): print(f"Error: Missing required columns in target sheet '{target_sheet_name}'. Need: {required_target_cols}"); sys.exit(1)

    for df in [df_duplicates, df_all_matches, df_target]:
        for col in ['filename', 'matched_example', 'original_excerpt', 'excerpt', 'codings', 'code']:
            if col in df.columns: df[col] = df[col].fillna('').astype(str).str.strip()

    print("Identifying rows for removal...")
    processed_rows = 0
    for index, dup_row in df_duplicates.iterrows():
        processed_rows += 1
        print(f"  Processing Duplicate Extracts row {index + 1}/{len(df_duplicates)}...", end='\r')
        erroneous_files = parse_comma_sep_string(dup_row['Erroneous Filenames'])
        dup_matched_example = dup_row['matched_example']
        codes_to_remove_check = set(parse_comma_sep_string(dup_row['associated_codes'], lower=True))
        if not erroneous_files or not dup_matched_example or not codes_to_remove_check: continue

        for filename in erroneous_files:
            filename_lower = filename.lower()
            all_matches_lookup = (df_all_matches['filename'].str.lower() == filename_lower) & \
                                 (df_all_matches['matched_example'] == dup_matched_example)
            matching_orig_excerpts = df_all_matches.loc[all_matches_lookup, 'original_excerpt'].unique()
            if len(matching_orig_excerpts) == 0: continue

            for original_excerpt in matching_orig_excerpts:
                if not original_excerpt: continue
                original_excerpt_lower = original_excerpt.lower()
                target_match_condition = (df_target['filename'].str.lower() == filename_lower) & \
                                         (df_target['excerpt'].str.lower() == original_excerpt_lower)
                target_indices = df_target[target_match_condition].index
                for target_idx in target_indices:
                    target_code = df_target.loc[target_idx, 'code'] # Already preprocessed
                    if target_code.lower() in codes_to_remove_check:
                        if target_idx not in indices_to_drop:
                             log_messages.append(f"Row {target_idx+2}: Identified for deletion. File='{filename}', OriginalExcerpt='{original_excerpt[:60]}...', Code='{target_code}' matches associated code.")
                        indices_to_drop.add(target_idx)

    print(f"\nProcessing complete. Identified {len(indices_to_drop)} unique rows for deletion.")
    if indices_to_drop:
        df_target_modified = df_target.drop(index=list(indices_to_drop)).reset_index(drop=True)
        print(f"Removed {len(indices_to_drop)} rows.")
        output_path = target_file_path.parent / f"{target_file_path.stem}{OUTPUT_SUFFIX}{target_file_path.suffix}"
        print(f"\nSaving cleaned data to: {output_path}")
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                 df_target_modified.to_excel(writer, sheet_name=target_sheet_name, index=False)
            print("File saved successfully.")
        except Exception as e: print(f"Error saving cleaned file: {e}")
    else:
        print("\nNo rows were identified for deletion. Original target file remains unchanged.")

    print("\n--- Magnitude Removal Summary ---")
    print(f"Total rows removed: {len(indices_to_drop)}")
    print("------------------------------")


# --- main function (Keep as is) ---
def main():
    parser = argparse.ArgumentParser(
        description="Removes specified codes or rows from target Excel files based on a 'Duplicate Extracts' sheet, using an intermediate lookup via 'All Matches' sheet.",
        formatter_class=argparse.RawTextHelpFormatter
        )
    mode_group = parser.add_mutually_exclusive_group(required=True)
    mode_group.add_argument('--coded', action='store_true', help="Mode for removing ROWS from Merged Codings file.") # Updated help text
    mode_group.add_argument('--magnitude', action='store_true', help="Mode for removing ROWS from Magnitude Codes file.")
    parser.add_argument('--duplicates-file', required=True, help="Path to Excel with 'Duplicate Extracts' AND 'All Matches'.")
    parser.add_argument('--target-file', required=True, help="Path to the target Excel file to be cleaned (--coded needs 'Merged Codings' sheet, --magnitude needs magnitude sheet).")
    args = parser.parse_args()
    duplicates_file = get_file_path_from_arg(args.duplicates_file, "Duplicates/All Matches")
    target_file = get_file_path_from_arg(args.target_file, "Target")
    if args.coded:
        run_coded_removal(duplicates_file, target_file) # Now removes rows
    elif args.magnitude:
        run_magnitude_removal(duplicates_file, target_file)
    else: print("Error: No mode selected."); sys.exit(1)

if __name__ == "__main__":
    main()