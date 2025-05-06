import json
import pandas as pd
import argparse
import os

# Added sheet_name parameter with default 0
def convert_excel_to_json(excel_filepath, json_filepath, sheet_name=0):
    """
    Converts an Excel spreadsheet (in the specific format) to a JSON file.
    Now also reads the 'frequency' column and allows specifying the sheet.

    Args:
        excel_filepath: The path to the input Excel file.
        json_filepath: The path to the output JSON file.
        sheet_name: The name or index of the sheet to convert (default is 0, the first sheet).
    """
    try:
        # Read the specified Excel sheet into a Pandas DataFrame
        # Updated to use the sheet_name argument
        df = pd.read_excel(excel_filepath, sheet_name=sheet_name)

        # Validate the column names
        expected_columns = ["code", "description", "examples", "construct", "frequency"]
        if not all(col in df.columns for col in expected_columns):
            missing_columns = [col for col in expected_columns if col not in df.columns] # Find missing
            # Raise error specifying missing columns
            raise ValueError(f"Excel sheet '{sheet_name}' must contain columns: {', '.join(missing_columns)}")

        # Convert DataFrame to a list of dictionaries
        data = []
        for _, row in df.iterrows():
            # Handle potential NaN or non-numeric values in frequency before converting to int
            try:
                # Fill NaN with 0, then convert to int
                frequency_val = int(pd.to_numeric(row["frequency"], errors='coerce').fillna(0))
            except ValueError:
                 # If coercion still fails (e.g., text), default to 0 or handle as needed
                print(f"Warning: Non-numeric frequency found for code '{row['code']}' in sheet '{sheet_name}'. Setting frequency to 0.")
                frequency_val = 0

            data.append({
                "code": row["code"],
                "description": row["description"],
                "examples": row["examples"],
                "construct": row["construct"],
                "frequency": frequency_val # Use the cleaned integer frequency
            })

        # Write the list of dictionaries to a JSON file
        with open(json_filepath, 'w') as f:
            json.dump(data, f, indent=4)

        # Updated success message to include sheet name
        print(f"Successfully converted sheet '{sheet_name}' from '{excel_filepath}' to '{json_filepath}'")

    except FileNotFoundError:
        print(f"Error: File not found - {excel_filepath}")
    except ValueError as e:
        # Updated error message to include sheet name
        print(f"Error: Invalid Excel format or missing columns in sheet '{sheet_name}' - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    parser = argparse.ArgumentParser(description="Convert an Excel codebook (including frequency) to a JSON file.")
    parser.add_argument("excel_file", help="Path to the input Excel file.")
    parser.add_argument("json_file", help="Path to the output JSON file (e.g., output.json).")
    # Added sheet_name argument, mirroring the first script
    parser.add_argument("-s", "--sheet_name", help="Name or index of the sheet to convert (default is 0, the first sheet).", default=0)

    args = parser.parse_args()

    # Create the output directory if it doesn't exist
    output_dir = os.path.dirname(args.json_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Convert sheet_name to int if it's a number, mirroring the first script
    try:
        sheet_name = int(args.sheet_name)
    except ValueError:
        sheet_name = args.sheet_name # Keep as string if conversion fails (it's a name)

    # Pass the sheet_name to the conversion function
    convert_excel_to_json(args.excel_file, args.json_file, sheet_name)

if __name__ == "__main__":
    main()