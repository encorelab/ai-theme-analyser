# src/fix_code_generation.py

import json
import time
import logging
import pandas as pd # Added for type hinting if needed
import random # *** Import random for jitter ***

import vertexai
from vertexai.generative_models import GenerativeModel, Part, SafetySetting, FinishReason

from config import PROJECT_ID, LOCATION, GEMINI_MODEL, LARGE_GENERATION_CONFIG, SAFETY_SETTINGS
from src.utils import remove_json_markdown

# Configure logging (consider sharing a logger instance if desired)
LOG_FILE = "log.txt"
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")

class FixCodeGeneratorClient:
    def __init__(self):
        vertexai.init(project=PROJECT_ID, location=LOCATION)
        self.model = GenerativeModel(
            GEMINI_MODEL,
            system_instruction=
            """You are an expert qualitative researcher specializing in thematic analysis. Your task is to analyze text excerpts where specific codes have been applied and generate missing definitions for those codes based on their name, the construct they belong to, and the example excerpt provided. Ensure the definitions accurately reflect the potential meaning within the given context."""
        )

    def generate_missing_definitions(self,
                                     construct: dict,
                                     missing_codes_data: dict) -> list[dict]:
        """
        Generates definitions for codes that are missing from the definitions list.
        Uses exponential backoff with jitter for retries on specific errors.

        Args:
            construct: A dictionary containing information about the current theme/construct
                       (e.g., {'theme': 'Theme Name', 'definition': '...'}).
            missing_codes_data: A dictionary where keys are missing code names and values
                                are dictionaries containing {'excerpt': '...', 'filename': '...'}.

        Returns:
            A list of dictionaries, where each dictionary represents a row for the
            code_justifications dataframe. Returns an empty list if generation fails.
        """
        if not missing_codes_data:
            return []

        # ... (prompt setup remains the same) ...
        construct_name = construct.get('theme', 'Unknown Construct')
        construct_definition = construct.get('definition', 'No definition provided.')
        missing_codes_prompt_list = []
        for code, data in missing_codes_data.items():
            missing_codes_prompt_list.append(
                f"- Code Name: {code}\n  - Example Excerpt: \"{data['excerpt']}\"\n  - Original Filename: {data['filename']}"
            )
        missing_codes_prompt_str = "\n".join(missing_codes_prompt_list)
        prompt = (
            f"You are analyzing codes applied within the construct/theme: '{construct_name}'.\n"
            f"Construct Definition: {construct_definition}\n\n"
            "The following codes related to this construct were applied to text excerpts in a specific document, but they currently lack definitions in our codebook. Your task is to generate a definition for EACH of these missing codes based on the code name, the construct definition, and the example excerpt where it was first applied.\n\n"
            "Missing Codes and Examples:\n"
            f"{missing_codes_prompt_str}\n\n"
            "Task:\n"
            "Generate a JSON list containing definitions for EACH of the missing codes listed above. For each code, provide the following information in a JSON object:\n"
            "  - 'code': The full name of the missing code (string).\n"
            "  - 'filename': The original filename where the code was applied (string, use the filename provided with the example).\n"
            "  - 'examples': The example excerpt provided above (string).\n"
            "  - 'construct': The name of the construct ('{construct_name}') (string).\n"
            "  - 'description': A concise definition of what the code likely represents, based on its name, the construct, and the example excerpt (string).\n"
            "  - 'justification': A brief explanation for why this code is relevant or applicable to the example excerpt, justifying its potential creation (string).\n"
            "  - 'probability': Assign a confidence score (float, e.g., 0.9) indicating the likelihood that your generated definition accurately captures the intended meaning.\n\n"
            "Output Format: Provide ONLY the JSON list, starting with '[' and ending with ']'. Each object in the list should represent one code definition.\n\n"
            "Example of a single JSON object within the list:\n"
            """
{
  "code": "Theme-CodeName",
  "filename": "doc1.docx",
  "examples": "Example text excerpt...",
  "construct": "Theme Name",
  "description": "A generated description of the code.",
  "justification": "A generated justification for the code's relevance.",
  "probability": 0.9
}
"""
            "\nEnsure the final output is a valid JSON list containing definitions for ALL the requested missing codes, and use single quotes when quoting the analyzed text.\n"
        )
        logging.info(f"Generating definitions for {len(missing_codes_data)} missing codes in construct '{construct_name}'.")
        logging.info(f"Missing codes: {list(missing_codes_data.keys())}")

        # --- Retry Logic Setup ---
        max_retries = 10
        base_delay = 2  # Initial delay in seconds (increased slightly)
        max_delay = 60 # Maximum delay in seconds
        # --- End Retry Logic Setup ---

        generated_definitions = []
        current_delay = base_delay # Initialize delay for the first retry

        for attempt in range(max_retries):
            try:
                # Call the model
                response = self.model.generate_content(
                    [prompt],
                    generation_config=LARGE_GENERATION_CONFIG,
                    safety_settings=SAFETY_SETTINGS,
                ).text

                clean_response = remove_json_markdown(response)

                # --- *** Replace non-standard escapes *** ---
                potentially_fixed_response = clean_response.replace("\\\"", "'")


                # Strip leading/trailing whitespace before checking start/end
                stripped_response = potentially_fixed_response.strip() # Use the fixed version

                if not stripped_response.startswith('[') or not stripped_response.endswith(']'):
                    raise json.JSONDecodeError(
                        "Response, after stripping whitespace and fixing escapes, does not start with '[' or end with ']'.",
                        stripped_response, 0
                        )

                # Parse the potentially fixed response string
                json_response = json.loads(potentially_fixed_response)

                if not isinstance(json_response, list):
                    raise ValueError("Response is not a list.")

                required_keys = {"code", "filename", "examples", "construct", "description", "justification", "probability"}
                validated_definitions = []
                for item in json_response:
                    if not isinstance(item, dict):
                        logging.warning(f"Item in response is not a dictionary: {item}")
                        continue
                    if not required_keys.issubset(item.keys()):
                        logging.warning(f"Dictionary item missing required keys: {item}")
                        continue
                    validated_definitions.append(item)

                if len(validated_definitions) != len(missing_codes_data):
                     logging.warning(f"LLM did not return definitions for all requested codes. Requested: {len(missing_codes_data)}, Returned: {len(validated_definitions)}")

                generated_definitions = validated_definitions
                logging.info(f"Successfully generated definitions for {len(generated_definitions)} codes.")
                break # Exit loop on success

            # --- Exception Handling ---
            except json.JSONDecodeError as e:
                logging.error(f"JSON decode error (attempt {attempt+1}/{max_retries}): {e}. Response: {clean_response[:500]}...")
                print(f"Error decoding JSON (attempt {attempt+1}/{max_retries}). Check logs.")
                if attempt < max_retries - 1:
                    # Calculate sleep time with backoff and jitter
                    sleep_time = current_delay + random.uniform(0, 1)
                    logging.info(f"Retrying in {sleep_time:.2f} seconds...")
                    print(f"Retrying in {sleep_time:.2f} seconds...")
                    time.sleep(sleep_time)
                    # Increase delay for next potential retry, capped at max_delay
                    current_delay = min(current_delay * 2, max_delay)
                else:
                    logging.error("Max JSON decode retries reached for fix_code_generation.")
                    print("Max JSON decode retries reached. Giving up on this batch.")
                    return [] # Return empty list on persistent failure

            except ValueError as e:
                 logging.error(f"Data validation error (attempt {attempt+1}/{max_retries}): {e}. Response: {clean_response[:500]}...")
                 print(f"Data validation error (attempt {attempt+1}/{max_retries}). Check logs.")

                 if attempt < max_retries - 1:
                    # --- Backoff Logic (Example if you want to retry ValueErrors) ---
                    sleep_time = current_delay + random.uniform(0, 1)
                    logging.info(f"Retrying after validation error in {sleep_time:.2f} seconds...")
                    print(f"Retrying after validation error in {sleep_time:.2f} seconds...")
                    time.sleep(sleep_time)
                    current_delay = min(current_delay * 2, max_delay)
                    # --- End Backoff Logic ---
                 else:
                    logging.error("Max validation retries reached (or validation error not retryable).")
                    return []

            except Exception as e:
                # Handle API errors (like rate limits) or other unexpected errors
                logging.exception(f"An unexpected error occurred during definition generation (attempt {attempt+1}/{max_retries}): {e}")
                print(f"An unexpected error occurred (attempt {attempt+1}/{max_retries}). Check logs.")
                # Check if error is likely retryable (e.g., rate limit, temporary server error)
                if "429" in str(e) or "Quota exceeded" in str(e) or "Resource has been exhausted" in str(e) or "503" in str(e):
                     if attempt < max_retries - 1:
                        # --- Backoff Logic for API errors ---
                        sleep_time = current_delay + random.uniform(0, 1)
                        print(f"Retryable API error detected. Retrying in {sleep_time:.2f} seconds...")
                        logging.info(f"Retryable API error detected. Retrying in {sleep_time:.2f} seconds...")
                        time.sleep(sleep_time)
                        current_delay = min(current_delay * 2, max_delay) # Apply backoff
                        # --- End Backoff Logic ---
                     else:
                        logging.error("Max retries reached after retryable API error.")
                        print("Max retries reached after retryable API error. Giving up on this batch.")
                        return []
                else:
                    # For non-retryable errors, fail immediately
                    logging.error("Non-retryable error encountered.")
                    print("Non-retryable error encountered. Giving up on this batch.")
                    return [] # Return empty list immediately

        return generated_definitions