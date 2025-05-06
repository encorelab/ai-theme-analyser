import pandas as pd
from pathlib import Path
from thefuzz import fuzz
import openpyxl # Required by pandas for xlsx operations
import sys
import argparse # For command-line arguments
import docx # For reading docx files
import datetime
import ast

# --- Configuration ---
MATCH_THRESHOLD = 85 # Initial fuzzy match score threshold (0-100) for Stage 1.
DOCX_MATCH_THRESHOLD = 75 # Initial threshold for matching example text within DOCX content (Stage 2)
DOCX_LOWER_THRESHOLD = 55 # Lower threshold for DOCX check retry (Stage 2)

CODINGS_SHEET_NAME = 'Merged Codings'
CODES_SHEET_NAME = 'Updated Used Codes'
OUTPUT_FILENAME_BASE = "duplicate_extract_output" # Base filename
ALL_MATCHES_SHEET_NAME = 'All Matches'
DUPLICATE_SHEET_NAME = "Duplicate Extracts" # Changed sheet name for stage 1 output / stage 2 input+output
# --- End Configuration ---

# --- Helper Functions ---

def get_file_path_from_arg(path_str):
    """Validates a file path provided as an argument."""
    file_path = Path(path_str).resolve()
    if not file_path.is_file():
        print(f"Error: File not found at provided path: '{file_path}'")
        sys.exit(1)
    return file_path

def get_dir_path_from_arg(path_str):
    """Validates a directory path provided as an argument."""
    dir_path = Path(path_str).resolve()
    if not dir_path.is_dir():
        print(f"Error: Directory not found at provided path: '{dir_path}'")
        sys.exit(1)
    return dir_path

def read_docx_text(file_path: Path):
    """Reads text content from a DOCX file, handling paragraphs."""
    full_text_list = []
    try:
        if not file_path.is_file(): # Check existence here
             print(f"  Warning: DOCX file not found: {file_path}", end='\r')
             return None
        doc = docx.Document(file_path)
        for para in doc.paragraphs:
             # Append non-empty paragraph text
             if para.text.strip():
                  full_text_list.append(para.text)
        # Return paragraphs joined by newline, or None if empty
        return '\n'.join(full_text_list) if full_text_list else None
    except docx.opc.exceptions.PackageNotFoundError:
        print(f"  Warning: Not a valid DOCX file: {file_path}", end='\r')
        return None # Indicate file not found/readable
    except Exception as e:
        print(f"  Warning: Error reading DOCX file {file_path}: {e}", end='\r')
        return None # Indicate other reading error

def validate_example_in_docx(example_text, docx_full_text, initial_threshold, lower_threshold):
    """
    Checks if example_text fuzzy matches docx_full_text using initial,
    lower threshold, and paragraph-chunking strategies.
    """
    if not example_text or not docx_full_text:
        return False, "No Example or DOCX Text" # Return status and reason

    example_lower = str(example_text).lower() # Ensure string and lower
    docx_lower = docx_full_text.lower()

    # 1. Initial check (whole doc)
    score_whole = fuzz.partial_ratio(example_lower, docx_lower)
    if score_whole >= initial_threshold:
        return True, f"Whole Doc Match ({score_whole}>={initial_threshold})"

    # 2. Lower threshold check (whole doc)
    if score_whole >= lower_threshold:
        # Matched only because of the lower threshold
        return True, f"Whole Doc Match on Lower Threshold ({score_whole}>={lower_threshold})"

    # 3. Chunking check (Paragraphs)
    paragraphs = docx_full_text.split('\n') # Assumes read_docx_text joins with \n
    for i, para in enumerate(paragraphs):
        if not para.strip(): continue # Skip empty paragraphs
        para_lower = para.lower()
        score_chunk = fuzz.partial_ratio(example_lower, para_lower)
        if score_chunk >= initial_threshold:
            # Matched a specific paragraph at the initial threshold
            return True, f"Paragraph {i+1} Match ({score_chunk}>={initial_threshold})"

    # If all checks fail
    return False, f"No Match Found (Highest Score: {score_whole})"

# --- MODIFIED Helper Function for Stage 2 Validation ---
def find_max_paragraph_match_score(example_text, docx_full_text):
    """
    Finds the maximum fuzzy partial_ratio score between the example_text
    and any single paragraph in the docx_full_text.

    Returns:
        tuple: (max_score, best_match_info string)
               max_score is -1 if no valid paragraphs or text found.
    """
    if not example_text or not docx_full_text:
        return -1, "No Example or DOCX Text"

    example_lower = str(example_text).lower().strip() # Ensure string, lower, strip whitespace
    if not example_lower: # Skip if example text itself is empty after stripping
         return -1, "Empty Example Text"

    # Split docx_full_text into paragraphs (assuming read_docx_text joins with \n)
    paragraphs = docx_full_text.split('\n')

    max_score = -1
    best_match_info = "No matching paragraphs found"
    processed_paragraph = False

    for i, para in enumerate(paragraphs):
        para_strip = para.strip()
        if not para_strip: continue # Skip empty paragraphs

        processed_paragraph = True
        para_lower = para_strip.lower()
        score = fuzz.partial_ratio(example_lower, para_lower)

        if score > max_score:
            max_score = score
            # Provide slightly more context in the note
            excerpt_preview = para_strip[:60] + ('...' if len(para_strip) > 60 else '') # Increased preview length
            best_match_info = f"Paragraph #{i+1} (Score: {max_score}) <<{excerpt_preview}>>"

    if not processed_paragraph:
         # This means the DOCX had text but possibly no newline separators or only whitespace paragraphs
         # Fallback: check against the whole text if no paragraphs were processed
         score_whole = fuzz.partial_ratio(example_lower, docx_full_text.lower())
         if score_whole > max_score:
              max_score = score_whole
              best_match_info = f"Whole Doc Fallback (Score: {max_score})"
         elif max_score == -1: # Only update if max_score is still -1
              best_match_info = "No valid paragraphs found in DOCX"


    # Return highest score found across all paragraphs (or whole doc fallback)
    return max_score, best_match_info

# --- Stage Functions ---

def run_stage1(args, match_threshold_to_use): # Use passed threshold
    """Runs Stage 1: Find duplicates and save initial analysis with details."""
    print("--- Running Stage 1: Finding Potential Duplicates ---")
    input_excel_path = args.input_excel
    output_dir_path = args.output_dir

    # --- Generate Timestamped Output Path for Stage 1 ---
    timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"{OUTPUT_FILENAME_BASE}_{timestamp_str}.xlsx"
    output_path = output_dir_path / output_filename

    # 1. Load Data (same as before)
    print(f"\nLoading data from {input_excel_path}...")
    try:
        df_codings = pd.read_excel(input_excel_path, sheet_name=CODINGS_SHEET_NAME)
        df_codes = pd.read_excel(input_excel_path, sheet_name=CODES_SHEET_NAME)
        print(f"Successfully loaded '{CODINGS_SHEET_NAME}' ({len(df_codings)} rows) and '{CODES_SHEET_NAME}' ({len(df_codes)} rows).")
    except Exception as e:
        print(f"Error loading data: {e}")
        sys.exit(1)

    # 2. Preprocessing & Validation (same as before)
    required_coding_cols = ['filename', 'excerpt']
    required_code_cols = ['code', 'examples']
    if not all(col in df_codings.columns for col in required_coding_cols) or \
       not all(col in df_codes.columns for col in required_code_cols):
         print(f"\nError: Missing required columns in input sheets.")
         sys.exit(1)
    df_codings['excerpt'] = df_codings['excerpt'].astype(str).fillna('').str.strip().str.lower()
    df_codes['examples'] = df_codes['examples'].astype(str).fillna('').str.strip().str.lower()
    df_codings['filename'] = df_codings['filename'].astype(str).fillna('').str.strip()
    df_codes['code'] = df_codes['code'].astype(str).fillna('').str.strip()
    print("Data preprocessing complete.")

    # 3. Fuzzy Matching (same as before - uses match_threshold_to_use)
    print(f"\nPerforming fuzzy matching (Threshold = {match_threshold_to_use})...")
    matches_list = []
    total_excerpts = len(df_codings)
    processed_count = 0
    for idx_coding, coding_row in df_codings.iterrows():
        excerpt = coding_row['excerpt']
        filename = coding_row['filename']
        processed_count += 1
        if not excerpt: continue
        if processed_count % 100 == 0 or processed_count == total_excerpts:
             print(f"  Processing excerpt {processed_count}/{total_excerpts}...", end='\r')

        for idx_code, code_row in df_codes.iterrows():
            example = code_row['examples'] # Already lowercased
            code = code_row['code']
            if not example: continue
            score = fuzz.partial_ratio(excerpt, example)
            if score >= match_threshold_to_use: # Use the argument here
                  matches_list.append({
                      'filename': filename,
                      'matched_example': code_row['examples'], # Store original case example
                      'matched_code': code,
                      'original_excerpt': coding_row['excerpt'], # Store original case excerpt
                      'match_score': score
                  })
    print(f"\nMatching complete. Found {len(matches_list)} potential matches.")

    # --- Prepare output dataframes ---
    if not matches_list:
        print("No matches found. Stage 1 output will reflect this.")
        df_all_matches_output = pd.DataFrame({'Status':['No matches found.']})
        df_duplicate_extracts_output = pd.DataFrame({'Status': ['No matches found in Stage 1.']})
    else:
        df_all_matches = pd.DataFrame(matches_list)
        if not df_all_matches.empty:
            df_all_matches = df_all_matches.sort_values(
                by=['filename', 'matched_code', 'match_score'], ascending=[True, True, False]
            ).drop_duplicates()

        if df_all_matches.empty:
            print("No valid matches remained after processing. Stage 1 output will reflect this.")
            df_all_matches_output = pd.DataFrame({'Status':['No valid matches remaining after processing.']})
            df_duplicate_extracts_output = pd.DataFrame({'Status': ['No valid matches after processing in Stage 1.']})
        else:
            # Use the potentially filtered df_all_matches for the first output sheet
            df_all_matches_output = df_all_matches

            # 4. Analyze for Duplicates & Aggregate Details (MODIFIED)
            print("\nAnalyzing matches for examples linked to multiple filenames...")

            # 4a. Identify examples with >1 unique filename
            grouped_filenames = df_all_matches.groupby('matched_example')['filename'].agg(lambda fns: set(fn for fn in fns if fn))
            duplicate_examples_index = grouped_filenames[grouped_filenames.apply(len) > 1].index

            if duplicate_examples_index.empty:
                print("No examples associated with multiple filenames found.")
                df_duplicate_extracts_output = pd.DataFrame({'Status': ['No examples found associated with multiple filenames in Stage 1.']})
            else:
                print(f"Found {len(duplicate_examples_index)} examples associated with multiple filenames. Aggregating details...")

                # 4b. Filter the original matches to include only those related to duplicate examples
                # Use df_all_matches here, which contains all necessary columns ('original_excerpt', 'match_score' etc.)
                df_duplicates_details_full = df_all_matches[df_all_matches['matched_example'].isin(duplicate_examples_index)].copy()

                # 4c. Define helper function to format the contributing matches for a group
                def format_contributing_matches(group):
                    """Formats the filename, score, and excerpt for each row in the group."""
                    details_list = []
                    # Sort group for consistent output (by filename, then score descending)
                    group = group.sort_values(by=['filename', 'match_score'], ascending=[True, False])
                    for _, row in group.iterrows():
                        # Truncate excerpt for readability in Excel cell
                        excerpt_preview = str(row['original_excerpt']) # Ensure string
                        max_len = 150 # Max excerpt length to show in preview
                        if len(excerpt_preview) > max_len:
                             excerpt_preview = excerpt_preview[:max_len] + "..."

                        details_list.append(
                            f"File: '{row['filename']}' (Score: {row['match_score']:.0f}) -> Excerpt: \"{excerpt_preview}\""
                        )
                    # Join with a clear separator (double newline works well in Excel wrap text)
                    return "\n\n".join(details_list)

                # 4d. Aggregate the filtered details by matched_example
                df_duplicate_extracts_output = df_duplicates_details_full.groupby('matched_example').agg(
                     # Aggregate unique filenames, comma-separated
                     associated_filenames=('filename', lambda fns: ', '.join(sorted(list(set(fn for fn in fns if fn))))),
                     # Aggregate unique codes, comma-separated
                     associated_codes=('matched_code', lambda cds: ', '.join(sorted(list(set(cd for cd in cds if cd))))),
                     # Apply the formatting function to each group (pass the group's index to get the sub-dataframe)
                     contributing_matches=('filename', lambda x: format_contributing_matches(df_duplicates_details_full.loc[x.index])) # Pass group index
                ).reset_index()

                # Ensure columns are in desired order
                df_duplicate_extracts_output = df_duplicate_extracts_output[[
                    'matched_example',
                    'associated_codes',
                    'associated_filenames',
                    'contributing_matches' # New detailed column
                ]]

    # 5. Write Stage 1 Output
    print(f"\nWriting Stage 1 results to: {output_path.resolve()}")
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Use the prepared output DataFrames
            df_all_matches_output.to_excel(writer, sheet_name=ALL_MATCHES_SHEET_NAME, index=False)
            df_duplicate_extracts_output.to_excel(writer, sheet_name=DUPLICATE_SHEET_NAME, index=False)

        print("\nSuccessfully wrote Stage 1 output file.")
    except PermissionError:
         print(f"\nError: Permission denied writing to {output_path}.")
         print("Please ensure the file is not open in another application and you have write permissions.")
         sys.exit(1)
    except Exception as e:
        print(f"\nAn error occurred while writing the Stage 1 output file: {e}")
        sys.exit(1)

    print("\n--- Stage 1 Finished ---")


def run_stage2(args, docx_thresh, docx_lower_thresh):
    """Runs Stage 2: Validate duplicates against DOCX files."""
    print("--- Running Stage 2: Validating Duplicates against DOCX ---")
    input_excel_path = args.input_excel # This is the output file from Stage 1
    docx_dir_path = args.docx_dir

    # 1. Load Data from Stage 1 Output
    print(f"\nLoading data from {input_excel_path}...")
    try:
        # Read only the sheet needed for validation first
        df_duplicate_extracts = pd.read_excel(input_excel_path, sheet_name=DUPLICATE_SHEET_NAME)
        # Check if it's just the status message sheet
        if 'Status' in df_duplicate_extracts.columns and not 'matched_example' in df_duplicate_extracts.columns:
             print(f"'{DUPLICATE_SHEET_NAME}' sheet contains status message. No duplicates to validate.")
             print("\n--- Stage 2 Finished (No Action Taken) ---")
             return # Exit stage 2 cleanly
        print(f"Successfully loaded '{DUPLICATE_SHEET_NAME}' sheet ({len(df_duplicate_extracts)} rows).")
    except FileNotFoundError:
        print(f"Error: Input file for Stage 2 not found: {input_excel_path}")
        sys.exit(1)
    except ValueError as e:
         print(f"Error reading sheet '{DUPLICATE_SHEET_NAME}' from {input_excel_path}: {e}")
         print(f"Ensure the Stage 1 output file exists and contains this sheet.")
         sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred loading Stage 2 input: {e}")
        sys.exit(1)

    # 2. Check required columns
    required_cols_s2 = ['matched_example', 'associated_filenames']
    if not all(col in df_duplicate_extracts.columns for col in required_cols_s2):
        print(f"Error: Missing required columns in '{DUPLICATE_SHEET_NAME}' sheet. Expected: {required_cols_s2}")
        sys.exit(1)

    # 3. Initialize/Clear Validation Columns
    df_duplicate_extracts['Correct Filenames'] = ''
    df_duplicate_extracts['Erroneous Filenames'] = ''
    df_duplicate_extracts['Validation Notes'] = '' # To store reason for match/no match

    # 4. DOCX Validation Loop (REVISED LOGIC)
    print(f"\nValidating {len(df_duplicate_extracts)} duplicate examples against DOCX files...")
    print(f"Using Paragraph Matching - Initial Threshold: {docx_thresh}, Lower Threshold: {docx_lower_thresh}")
    print("Logic: Correct if >= High Threshold. If no file meets High, Correct if >= Low Threshold.")

    validation_results = [] # Store final results for the DataFrame update

    total_duplicates_to_check = len(df_duplicate_extracts)
    for index, row in df_duplicate_extracts.iterrows():
        print(f"  Validating example row {index + 1}/{total_duplicates_to_check}...", end='\r')

        # --- Handle potentially list-like matched_example --- (Keep as is)
        example_cell_value = row['matched_example']
        example_texts_to_check = []; is_list_like = False
        try:
            parsed_value = ast.literal_eval(str(example_cell_value))
            if isinstance(parsed_value, list):
                example_texts_to_check = [str(item).strip() for item in parsed_value if str(item).strip()]
                is_list_like = True if len(example_texts_to_check) > 1 else False
            else:
                single_text = str(parsed_value).strip();
                if single_text: example_texts_to_check = [single_text]
        except (ValueError, SyntaxError, TypeError, MemoryError):
            single_text = str(example_cell_value).strip();
            if single_text: example_texts_to_check = [single_text]
        # ---

        filenames_str = str(row['associated_filenames']) if pd.notna(row['associated_filenames']) else ''
        filenames_to_check = [fn.strip() for fn in filenames_str.split(',') if fn.strip()]

        # --- Intermediate storage for results per file for this row---
        file_results = [] # List of {'filename': str, 'score': int, 'reason': str, 'error': bool}

        # --- Pass 1: Check each file and store detailed results ---
        if filenames_to_check and example_texts_to_check:
            for filename in filenames_to_check:
                potential_file_path = docx_dir_path / filename
                docx_text = read_docx_text(potential_file_path)

                if docx_text is None: # File not found or error reading
                    file_results.append({'filename': filename, 'score': -1, 'reason': "Not found or read error.", 'error': True})
                    continue

                # Find highest score for this file across all example texts
                highest_score_for_file = -1
                best_match_reason_for_file = "No matching paragraph found for any example text."

                for i, example_text in enumerate(example_texts_to_check):
                    max_para_score, para_match_info = find_max_paragraph_match_score(
                        example_text, docx_text
                    )
                    if max_para_score > highest_score_for_file:
                        highest_score_for_file = max_para_score
                        example_prefix = f"Example Text #{i+1}: " if is_list_like else ""
                        best_match_reason_for_file = f"{example_prefix}{para_match_info}"

                # Store the result for this file
                file_results.append({
                    'filename': filename,
                    'score': highest_score_for_file,
                    'reason': best_match_reason_for_file, # Store the best reason found
                    'error': False
                })
        # --- End Pass 1 ---

        # --- Pass 2: Apply decision logic based on all file results for this row ---
        correct_files_set = set()
        erroneous_files_set = set()
        row_notes = []

        if not filenames_to_check:
            row_notes.append("No associated filenames listed.")
        elif not example_texts_to_check:
            row_notes.append("Matched example text was empty or invalid.")
            erroneous_files_set.update(filenames_to_check) # Mark all as erroneous
        else:
            # Check if any valid file met the high threshold
            high_threshold_met_somewhere = any(res['score'] >= docx_thresh for res in file_results if not res['error'])

            for res in file_results:
                filename = res['filename']
                score = res['score']
                reason = res['reason']
                is_error = res['error']

                if is_error:
                    erroneous_files_set.add(filename)
                    row_notes.append(f"{filename}: {reason}")
                    continue

                # Apply the refined logic
                if high_threshold_met_somewhere:
                    # Only files meeting the high threshold are correct
                    if score >= docx_thresh:
                        correct_files_set.add(filename)
                        row_notes.append(f"{filename}: CORRECT - High Threshold Match ({reason})")
                    else:
                        erroneous_files_set.add(filename)
                        score_note = f"(Score: {score})" if score >= 0 else ""
                        logic_note = f"(High match found elsewhere)"
                        row_notes.append(f"{filename}: ERRONEOUS {score_note} {logic_note}")
                else:
                    # No file met the high threshold, accept lower threshold matches
                    if score >= docx_lower_thresh:
                        correct_files_set.add(filename)
                        row_notes.append(f"{filename}: CORRECT - Lower Threshold Match ({reason}) (No high match found)")
                    else:
                        erroneous_files_set.add(filename)
                        score_note = f"(Score: {score})" if score >= 0 else ""
                        logic_note = f"(Below lower threshold)"
                        row_notes.append(f"{filename}: ERRONEOUS {score_note} {logic_note}")
            # --- End Pass 2 ---

        # Store final results for the row
        validation_results.append({
            'index': index,
            'Correct Filenames': ', '.join(sorted(list(correct_files_set))),
            'Erroneous Filenames': ', '.join(sorted(list(erroneous_files_set))),
            'Validation Notes': '; '.join(row_notes) # Join all notes for the row
        })
        # --- End Row Processing ---

    if total_duplicates_to_check > 0: print("\nValidation complete.")

    # 5. Update DataFrame with validation results
    update_df = pd.DataFrame(validation_results).set_index('index')
    df_duplicate_extracts.update(update_df)

    # 6. Prepare to Write Output (Overwrite Stage 1 file)
    print(f"\nPreparing to update {input_excel_path}...")
    try:
        # Read all existing sheets from the input file to preserve them
        print("  Reading existing sheets...")
        excel_data = pd.read_excel(input_excel_path, sheet_name=None) # Reads all into dict
    except Exception as e:
         print(f"Error reading existing sheets from {input_excel_path} for update: {e}")
         sys.exit(1)

    # Replace the old duplicate sheet data with the updated DataFrame
    excel_data[DUPLICATE_SHEET_NAME] = df_duplicate_extracts

    # Write all sheets back, overwriting the original file
    print(f"  Writing updated data back to {input_excel_path}...")
    try:
        with pd.ExcelWriter(input_excel_path, engine='openpyxl') as writer:
            for sheet_name, df_sheet in excel_data.items():
                # Handle potential NaT or other problematic types before writing if needed
                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"\nSuccessfully updated {input_excel_path} with DOCX validation results.")
    except PermissionError:
         print(f"\nError: Permission denied writing to {input_excel_path}.")
         print("Please ensure the file is not open in another application and you have write permissions.")
         sys.exit(1)
    except Exception as e:
        print(f"\nAn unexpected error occurred while updating the output file: {e}")
        sys.exit(1)

    print("\n--- Stage 2 Finished ---")


def main():
    parser = argparse.ArgumentParser(
        description="Analyze codings, find duplicates (Stage 1), and validate against DOCX files (Stage 2).",
        formatter_class=argparse.RawTextHelpFormatter # Preserve newline formatting in help
        )

    # --- Stage Arguments ---
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--stage1', action='store_true', help="Run Stage 1: Find potential duplicates and save initial analysis.")
    group.add_argument('--stage2', action='store_true', help="Run Stage 2: Validate duplicates from Stage 1 output against DOCX files.")

    # --- Path Arguments ---
    parser.add_argument(
        '--input-excel',
        required=True,
        help="Path to the input Excel file.\n"
             "- For Stage 1: Source file with 'Merged Codings' and 'Updated Used Codes' sheets.\n"
             "- For Stage 2: The output file generated by Stage 1 (e.g., duplicate_extract_output.xlsx)."
        )
    parser.add_argument(
        '--output-dir',
        help="Path to the directory for the output Excel file (REQUIRED for Stage 1)."
        )
    parser.add_argument(
        '--docx-dir',
        help="Path to the directory containing input DOCX files (REQUIRED for Stage 2)."
        )

    # --- Optional Threshold Arguments ---
    parser.add_argument('--match-threshold', type=int, default=MATCH_THRESHOLD, help=f"Fuzzy match threshold for Stage 1 (default: {MATCH_THRESHOLD})")
    parser.add_argument('--docx-threshold', type=int, default=DOCX_MATCH_THRESHOLD, help=f"Initial fuzzy match threshold for Stage 2 DOCX validation (default: {DOCX_MATCH_THRESHOLD})")
    parser.add_argument('--docx-lower-threshold', type=int, default=DOCX_LOWER_THRESHOLD, help=f"Lower fuzzy match threshold for Stage 2 DOCX retry (default: {DOCX_LOWER_THRESHOLD})")


    args = parser.parse_args()

    # --- Argument Validation ---
    args.input_excel = get_file_path_from_arg(args.input_excel)

    if args.stage1:
        if not args.output_dir:
            parser.error("--output-dir is required for --stage1")
        args.output_dir = get_dir_path_from_arg(args.output_dir)

        # Get the threshold value from args (which has the default or user override)
        current_match_threshold = args.match_threshold
        print(f"Using Match Threshold for Stage 1: {current_match_threshold}")
        # Pass the threshold as an argument - NO 'global' needed
        run_stage1(args, current_match_threshold)

    elif args.stage2:
        if not args.docx_dir:
             parser.error("--docx-dir is required for --stage2")
        args.docx_dir = get_dir_path_from_arg(args.docx_dir)

        # Get threshold values from args
        current_docx_threshold = args.docx_threshold
        current_docx_lower_threshold = args.docx_lower_threshold
        print(f"Using DOCX Thresholds for Stage 2: Initial={current_docx_threshold}, Lower={current_docx_lower_threshold}")
        # Pass thresholds as arguments - NO 'global' needed
        run_stage2(args, current_docx_threshold, current_docx_lower_threshold)



if __name__ == "__main__":
    # Levenshtein check
    try:
        import Levenshtein
    except ImportError:
        print("\nWarning: 'python-Levenshtein' library not found. Fuzzy matching may be slower.")
        print("Install with: pip install python-Levenshtein\n")
    main()