import json
import pandas as pd
import argparse
import os

def convert_excel_to_json_no_frequency(excel_filepath, json_filepath, sheet_name=0):
    """
    Converts an Excel spreadsheet (in the specific format) to a JSON file.
    This version *excludes* the 'frequency' column entirely from the output JSON.

    Args:
        excel_filepath: The path to the input Excel file.
        json_filepath: The path to the output JSON file.
        sheet_name: The name or index of the sheet to convert (default is 0, the first sheet).
    """
    try:
        # Read the Excel file into a Pandas DataFrame
        df = pd.read_excel(excel_filepath, sheet_name=sheet_name)

        # Validate the column names (frequency can be present, but will be ignored)
        expected_columns = ["code", "description", "examples", "construct"]  # No 'frequency'
        if not all(col in df.columns for col in expected_columns):
            missing_columns = [col for col in expected_columns if col not in df.columns]
            raise ValueError(f"Excel file must contain columns: {', '.join(missing_columns)}")

        # Convert DataFrame to a list of dictionaries
        data = []
        for _, row in df.iterrows():
            data.append({
                "code": row["code"],
                "description": row["description"],
                "examples": row["examples"],
                "construct": row["construct"],
                # NO FREQUENCY
            })

        # Write the list of dictionaries to a JSON file
        with open(json_filepath, 'w') as f:
            json.dump(data, f, indent=4)

        print(f"Successfully converted sheet '{sheet_name}' from '{excel_filepath}' to '{json_filepath}' (frequency column excluded)")

    except FileNotFoundError:
        print(f"Error: File not found - {excel_filepath}")
    except ValueError as e:
        print(f"Error: Invalid Excel format - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


def main():
    parser = argparse.ArgumentParser(description="Convert an Excel codebook to a JSON file.")
    parser.add_argument("excel_file", help="Path to the input Excel file.")
    parser.add_argument("json_file", help="Path to the output JSON file (e.g., output.json).")
    parser.add_argument("-s", "--sheet_name", help="Name or index of the sheet to convert (default is 0, the first sheet).", default=0) # Added sheet_name argument

    args = parser.parse_args()

    # Create the output directory if it doesn't exist
    output_dir = os.path.dirname(args.json_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Convert sheet_name to int if it's a number
    try:
        sheet_name = int(args.sheet_name)
    except ValueError:
        sheet_name = args.sheet_name
    
    convert_excel_to_json_no_frequency(args.excel_file, args.json_file, sheet_name)

if __name__ == "__main__":
    main()