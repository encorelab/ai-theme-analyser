import os
import datetime
import argparse
import json
import pandas as pd

from config import *
from code_generation import CodeGenerationClient
from code_merger_client import CodeMergerClient
from theme_generator import ThemeGeneratorClient
from code_application import ThematicCodingClient
from within_case_analysis import IntraTextAnalyzerClient
from report_generation import CrossDocumentAnalyzerClient
from code_compressor_client import CodeCompressorClient
from utils import (extract_paragraphs_from_docx, 
                   write_coding_results_to_excel, 
                   generate_codes,
                   perform_analysis_and_reporting, 
                   perform_intra_text_analysis, 
                   perform_cross_document_analysis, 
                   load_analysis_results_from_file, 
                   load_themes_from_file, 
                   load_codes_from_file,
                   load_codes_from_file_as_dictionary,
                   visualize_individual_theme_subgraphs,
                   visualize_theme_overview,
                   visualize_single_file_graph,
                   visualize_network,
                   read_full_dataset_codes,
                   convert_codes_dict_to_dataframe,
                   generate_code_stats,
                   convert_df_to_codes_dict,
                   read_used_codes_with_def,
                   replace_and_update_codes,
                   split_data_by_class,
                   compress_code_examples) 


def perform_thematic_analysis(directory, batch_size, client_flag):

    output_file = os.path.join(OUTPUT_DIR, f"analyzed_results_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")
    
    # Stage 2 - Part 1
    if client_flag == "generate_initial_codes": 
        code_generator = CodeGenerationClient()
        themes = load_themes_from_file("themes.json") 
        all_codes, all_files_excerpt_codings, new_codes_by_file = generate_codes(
            directory, themes, code_generator, initial_codes={}, num_docs=NUM_DOCS_FOR_CODE_GENERATION) 
    
        
        output_file = os.path.join(
            OUTPUT_DIR,
            f"initial_code_generation_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        )
        write_coding_results_to_excel(all_files_excerpt_codings,
                                     new_codes_by_file, output_file)
        print(f"\nGenerated Codes:\n{all_codes}")

    # Stage 2 - Part 2
    elif client_flag == "verify_initial_codes": 
        code_generator = CodeGenerationClient()
        themes = load_themes_from_file("themes.json") 
        initial_codes_json = load_codes_from_file_as_dictionary("codes.json")
        all_codes, all_files_excerpt_codings, new_codes_by_file = generate_codes(
            directory, themes, code_generator, initial_codes=initial_codes_json, num_docs=NUM_DOCS_FOR_CODE_GENERATION) 
    
        output_file = os.path.join(
            OUTPUT_DIR,
            f"code_generation_verification_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        )
        write_coding_results_to_excel(all_files_excerpt_codings,
                                     new_codes_by_file, output_file)
        print(f"\nGenerated Codes:\n{all_codes}")

    # Stage 3 - Part 1
    elif client_flag == "generate_full_dataset_codes": 
        code_generator = CodeGenerationClient()
        themes = load_themes_from_file("themes.json") 
        initial_codes_json = load_codes_from_file_as_dictionary("codes.json")
        all_codes, all_files_excerpt_codings, new_codes_by_file = generate_codes(
            directory, themes, code_generator, initial_codes=initial_codes_json, num_docs=None) 
    
        output_file = os.path.join(
            OUTPUT_DIR,
            f"full_dataset_code_generation_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        )
        write_coding_results_to_excel(all_files_excerpt_codings,
                                     new_codes_by_file, output_file)
        print(f"\nGenerated Codes:\n{all_codes}")

    # Stage 3 - Part 2
    elif client_flag == "generate_code_stats": 
        
        full_dataset_file_path = input("Enter the path to the full_dataset_code_generation xlsx file: ")
        initial_codes_file_path = "codes.json"
    
        output_file_path = os.path.join(
            OUTPUT_DIR,
            f"code_stats_generation_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        )
        generate_code_stats(full_dataset_file_path, initial_codes_file_path, output_file_path)

    # Stage 3 - Part 3
    elif client_flag == "merge_codes": 
        
        file_path = input("Enter the path to the full_dataset_code_generation xlsx file: ")

        # 1. Read the 'used_codes_with_def' sheet
        used_codes_df = read_used_codes_with_def(file_path)
        if used_codes_df is None:
            return  # Exit if there was an error reading the file

        # 2. Convert DataFrame to dictionary
        full_dataset_codes = convert_df_to_codes_dict(used_codes_df)

        # 3. Load themes
        themes = load_themes_from_file("themes.json") 
        if not themes:
            print("Error: No themes loaded. Exiting.")
            return

        # 4. Instantiate CodeMergerClient
        code_merger = CodeMergerClient()

        # 5. Merge themes (call the merge_themes method)
        merged_codes_result = code_merger.merge_themes(full_dataset_codes, themes, MERGE_CODES_GREATER_THAN)
        # Check if merged_codes_result is valid JSON
        if isinstance(merged_codes_result, str):
            try:
                merged_codes_result = json.loads(merged_codes_result)
            except json.JSONDecodeError:
                print("Error: The response from merge_themes is not valid JSON.")
                return
        elif not isinstance(merged_codes_result, dict):
            print("Error: The response from merge_themes has unexpected return type")
            return

        # 6. Prepare data for Excel output
        merged_codes_data = []
        
        for merged_code, details in merged_codes_result.items():
            merged_codes_data.append({
                "code": merged_code,
                "description": details["new_description"],
                "examples": details["examples"],
                "merged_codes": details["merged_codes"],
            })

        # Create DataFrame
        merged_codes_df = pd.DataFrame(merged_codes_data)

        # 7. Write to Excel
        output_filepath = os.path.join(
            OUTPUT_DIR,
            f"merged_codes_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
        )
        try:
            with pd.ExcelWriter(output_filepath, engine="openpyxl") as writer:
                merged_codes_df.to_excel(writer, sheet_name="Merged Codes", index=False)
            print(f"Successfully merged codes and saved to '{output_filepath}'")
        except Exception as e:
            print(f"An error occurred while writing to Excel: {e}")

    # Stage 3 - Part 4
    elif client_flag == "replace_merged_codes":
        full_dataset_file_path = input(
            "Enter the path to the code_stats_generation xlsx file: "
        )
        merged_codes_file_path = input(
            "Enter the path to the merged_codes xlsx file: "
        )
        output_filepath = os.path.join(
            OUTPUT_DIR,
            f"merged_codings_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
        )

        replace_and_update_codes(
            full_dataset_file_path, merged_codes_file_path, output_filepath
        )

    # Stage 3 - Part 5
    elif client_flag == "split_by_class":
        merged_codings_file_path = input(
            "Enter the path to the merged_codings xlsx file (with class column): "
        )
        split_data_by_class(merged_codings_file_path)

    # Stage 3 - Part 6
    elif client_flag == "compress_code_examples":
        codes_file_path = input("Enter the path to the codes JSON file: ")
        compression_type = input("Enter 1 to compress only examples or 2 to compress examples and descriptions: ")

        compressor = CodeCompressorClient()

        compress_code_examples(codes_file_path, compression_type, compressor)

    # Stage 4 - Part 1
    elif client_flag == "generate_themes":
        theme_generator = ThemeGeneratorClient()

        codes_filepath = input("Enter the path to the codes JSON file: ")
        themes_filepath = "themes.json"

        themes = load_themes_from_file(themes_filepath)
        all_code_data = load_codes_from_file(codes_filepath)

        if themes is None or all_code_data is None:
            return  # Exit if loading failed

        theme_generator = ThemeGeneratorClient() #make sure this class is accessible.
        themes_hierarchy = theme_generator.generate_themes(all_code_data, themes)

        # Extract filename without extension
        codes_filename = os.path.splitext(os.path.basename(codes_filepath))[0]

        # Save themes_hierarchy as a JSON file
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        themes_output_file = os.path.join(OUTPUT_DIR, f"{codes_filename}_themes_hierarchy_{timestamp}.json")

        with open(themes_output_file, 'w') as f:
            json.dump(themes_hierarchy, f, indent=4)

        print(f"\nGenerated Themes and saved to:\n{themes_output_file}")
        print(f"\nGenerated Themes:\n{themes_hierarchy}")


    # Stage 4 - Part 2
    elif client_flag == "visualize_themes":
        full_dataset_file = input("Enter the path to the themes_hierarchy JSON file: ")
        try:
            themes_hierarchy = load_themes_from_file(full_dataset_file)
            
            # Extract filename without extension
            full_dataset_filename = os.path.splitext(os.path.basename(full_dataset_file))[0]

            # Save themes_hierarchy as a JSON file
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            themes_hierarchy_output_file = os.path.join(OUTPUT_DIR, f"{full_dataset_filename}_themes_overview_{timestamp}.png")
            visualize_theme_overview(themes_hierarchy, themes_hierarchy_output_file)

        except FileNotFoundError:
            print(f"Error: File not found at {full_dataset_file}")
        except json.JSONDecodeError:
            print(f"Error: Invalid JSON format in {full_dataset_file}")
        except Exception as e:
            print(f"An error occurred: {e}")

    # Stage 4 - Part 3
    elif client_flag == "visualize_codes":
        full_dataset_file = input("Enter the path to the themes_hierarchy JSON file: ")
        output_dir_name = input("Enter the name of the output directory: ")
        try:
            themes_hierarchy = load_themes_from_file(full_dataset_file)
            visualize_individual_theme_subgraphs(themes_hierarchy, output_dir=output_dir_name)
        except FileNotFoundError:
            print(f"Error: File not found at {full_dataset_file}")
        except json.JSONDecodeError:
            print(f"Error: Invalid JSON format in {full_dataset_file}")
        except Exception as e:
            print(f"An error occurred: {e}")

    # Stage 5
    elif client_flag == "visualize_individual_file":
        # 1. Prompt user for the themes_hierarchy JSON file
        full_dataset_file = input(
            "Enter the path to the themes_hierarchy JSON file: "
        )
        try:
            themes_hierarchy = load_themes_from_file(full_dataset_file)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Error loading themes hierarchy: {e}")
            themes_hierarchy = None  # Set to None to indicate failure

        # Only proceed if themes_hierarchy was loaded successfully
        if themes_hierarchy:
            # 2. Prompt user for the XLSX file with coding data
            xlsx_file = input(
                "Enter the path to the XLSX file containing coding data: "
            )
            try:
                coding_df = pd.read_excel(
                    xlsx_file, sheet_name=0
                )  # Assuming data is on the first sheet
            except FileNotFoundError:
                print(f"Error: XLSX file not found at {xlsx_file}")
                coding_df = None  # Set to None to indicate failure
            except Exception as e:
                print(f"Error reading XLSX file: {e}")
                coding_df = None

            # Only proceed if coding_df was loaded successfully
            if coding_df is not None:
                # 3. Prompt user for the filename to analyze
                filename_to_analyze = input(
                    "Enter the filename to analyze (e.g., 104.docx): "
                )

                # 4. Filter the DataFrame for the specified filename
                filtered_df = coding_df[
                    coding_df["filename"] == filename_to_analyze
                ]

                if filtered_df.empty:
                    print(f"No data found for filename: {filename_to_analyze}")
                else:
                    # 5. Extract and count the codes, tracking frequencies
                    all_codes = []
                    code_frequencies = {}  # Dictionary to track code frequencies
                    for coding_str in filtered_df["codings"]:
                        codes = [
                            code.strip().split("-", 1)[1] if "-" in code else code.strip() 
                            for code in coding_str.split(",")
                        ]  # Split into list of codes
                        for code in codes:
                            all_codes.append(code)
                            code_frequencies[code] = (
                                code_frequencies.get(code, 0) + 1
                            )

                    # 6. Filter the theme_hierarchy to include only relevant codes and update frequencies
                    def filter_and_update_hierarchy(
                        hierarchy, relevant_codes, code_frequencies
                    ):
                        """
                        Filters the theme hierarchy to include only relevant codes, updates frequencies, and modifies code names.
                        """
                        filtered_hierarchy = {}

                        for meta_theme, meta_theme_data in hierarchy.items():
                            filtered_themes = {}
                            for theme, theme_data in meta_theme_data.get(
                                "themes", {}
                            ).items():
                                filtered_sub_themes = {}
                                for sub_theme, sub_theme_data in theme_data.get(
                                    "sub-themes", {}
                                ).items():
                                    filtered_codes = []
                                    updated_code_frequencies = {}  # Store updated code frequencies
                                    for code in sub_theme_data.get("codes", []):
                                        # Modify code name here
                                        shortened_code = code.split("-", 1)[1] if "-" in code else code
                                        if shortened_code in relevant_codes:
                                            filtered_codes.append(shortened_code)
                                            # Use the shortened code to look up the frequency
                                            updated_code_frequencies[shortened_code] = code_frequencies.get(shortened_code, 0)

                                    if filtered_codes:
                                        # Update code frequencies in the sub-theme
                                        sub_theme_data["codes"] = filtered_codes
                                        sub_theme_data["code_frequencies"] = updated_code_frequencies
                                        # Calculate sub-theme frequency
                                        sub_theme_data["frequency"] = sum(
                                            sub_theme_data["code_frequencies"].values()
                                        )
                                        filtered_sub_themes[sub_theme] = sub_theme_data

                                if filtered_sub_themes:
                                    # Calculate theme frequency
                                    theme_data["sub-themes"] = filtered_sub_themes
                                    theme_data["frequency"] = sum(
                                        sub_theme_data["frequency"]
                                        for sub_theme_data in filtered_sub_themes.values()
                                    )
                                    filtered_themes[theme] = theme_data

                            if filtered_themes:
                                # Calculate meta-theme frequency
                                meta_theme_data["themes"] = filtered_themes
                                meta_theme_data["frequency"] = sum(
                                    theme_data["frequency"]
                                    for theme_data in filtered_themes.values()
                                )
                                filtered_hierarchy[meta_theme] = meta_theme_data

                        return filtered_hierarchy

                    filtered_themes_hierarchy = filter_and_update_hierarchy(
                        themes_hierarchy, all_codes, code_frequencies
                    )

                    # 7. Save the filtered themes_hierarchy to a file
                    output_dir = os.path.join("within_case_network_graphs", filename_to_analyze)
                    os.makedirs(output_dir, exist_ok=True)  # Create the directory if it doesn't exist

                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    filtered_themes_hierarchy_filename = f"{filename_to_analyze}_filtered_themes_{timestamp}.json"
                    filtered_themes_hierarchy_filepath = os.path.join(output_dir, filtered_themes_hierarchy_filename)

                    with open(filtered_themes_hierarchy_filepath, "w") as f:
                        json.dump(filtered_themes_hierarchy, f, indent=4)
                    print(f"Saved filtered themes hierarchy to '{filtered_themes_hierarchy_filepath}'")

                    # 8. Visualize the filtered hierarchy for the specific filename
                    visualize_single_file_graph(
                        filtered_themes_hierarchy,
                        filename_to_analyze,
                        output_dir=output_dir,
                    )

    elif client_flag == "intra_text_analyzer":
        # intra_text_analyzer = IntraTextAnalyzerClient()
        # analysis_results = load_analysis_results_from_file("analysis_results.json") 
        # perform_intra_text_analysis(analysis_results, intra_text_analyzer)
      
        # data = json.loads(data122)
        # visualize_network(data)
        pass

    elif client_flag == "cross_document_analyzer":
        cross_document_analyzer = CrossDocumentAnalyzerClient()
        intra_text_output_file = os.path.join(OUTPUT_DIR, "intra_text_analysis.xlsx")  
        perform_cross_document_analysis(intra_text_output_file, cross_document_analyzer)

    else:
        print("Invalid client flag.")


def main():
    parser = argparse.ArgumentParser(description="Perform different thematic analysis steps.")
    parser.add_argument("--client", required=True, choices=[
        "generate_initial_codes", 
        "verify_initial_codes",
        "generate_full_dataset_codes",
        "generate_code_stats",
        "merge_codes",
        "replace_merged_codes",
        "split_by_class",
        "compress_code_examples",
        "generate_themes", 
        "visualize_themes", 
        "visualize_codes",
        "visualize_individual_file", 
        "intra_text_analyzer", 
        "cross_document_analyzer"
    ], help="Specify the client to run.")
    args = parser.parse_args()

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    perform_thematic_analysis(INPUT_DIR, BATCH_SIZE, args.client)

if __name__ == "__main__":
    main()