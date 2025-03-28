import json
import pandas as pd
import argparse
import os

def convert_excel_to_json(excel_filepath, json_filepath):
    """
    Converts an Excel spreadsheet (in the specific format) to a JSON file.
    Now also reads the 'frequency' column.

    Args:
        excel_filepath: The path to the input Excel file.
        json_filepath: The path to the output JSON file.
    """
    try:
        # Read the Excel file into a Pandas DataFrame
        df = pd.read_excel(excel_filepath)

        # Validate the column names
        expected_columns = ["code", "description", "examples", "construct", "frequency"]
        if not all(col in df.columns for col in expected_columns):
            raise ValueError(f"Excel file must contain columns: {', '.join(expected_columns)}")

        # Convert DataFrame to a list of dictionaries
        data = []
        for _, row in df.iterrows():
            data.append({
                "code": row["code"],
                "description": row["description"],
                "examples": row["examples"],
                "construct": row["construct"],
                "frequency": int(row["frequency"])  # Convert frequency to integer
            })

        # Write the list of dictionaries to a JSON file
        with open(json_filepath, 'w') as f:
            json.dump(data, f, indent=4)

        print(f"Successfully converted '{excel_filepath}' to '{json_filepath}'")

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

    args = parser.parse_args()

    # Create the output directory if it doesn't exist
    output_dir = os.path.dirname(args.json_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    convert_excel_to_json(args.excel_file, args.json_file)

if __name__ == "__main__":
    main()