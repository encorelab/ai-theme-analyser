import json
import time
import logging

import vertexai
from vertexai.generative_models import GenerativeModel
from src.utils import remove_json_markdown


from config import PROJECT_ID, LOCATION, GEMINI_MODEL, LARGE_GENERATION_CONFIG, SAFETY_SETTINGS

# Configure logging
LOG_FILE = "log.txt"
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")


class IntensityGenerationClient:
    def __init__(self):
        vertexai.init(project=PROJECT_ID, location=LOCATION)
        self.model = GenerativeModel(
            GEMINI_MODEL,
            system_instruction=(
                "You are a thematic analysis researcher tasked with applying magnitude coding to descriptive codes. "
                "For a given text excerpt, you will be provided with a list of codes applied to that excerpt, along with the definitions of those codes and their associated themes. "
                "Assign a Likert scale value between 1 and 7 (where 1 is low and 7 is high) to EACH code, representing the magnitude to which the excerpt's expression of the code aligns with BOTH the code's definition AND the theme's definition. "
                "Provide a VERY BRIEF (1-2 sentence) justification for EACH magnitude rating. "
                "Return a JSON object where the keys are the code names, and the values are nested objects containing 'magnitude' (the rating) and 'justification' (the brief explanation). "
                "Ensure the JSON is valid and well-formatted."
            )
        )

    def generate_intensity(self, excerpt, codes_applied, code_definitions, themes):
        """
        Generates intensity ratings and justifications for ALL codes in an excerpt.
        """
        prompt = "Analyze the following text excerpt:\n\n"
        prompt += f"Excerpt: {excerpt}\n\n"
        prompt += "Codes Applied:\n"

        theme_info = {}
        for code in codes_applied:
            prompt += f"- {code}\n"  # List the codes
            if code in code_definitions:  # Get construct (theme) only if code definition exists
                construct = code_definitions[code]['construct']
                if construct not in theme_info:
                    for theme_data in themes:
                        if theme_data['theme'] == construct:
                            theme_info[construct] = theme_data
                            break

        prompt += "\nCode and Theme Definitions:\n"
        for code in codes_applied:  # Iterate through CODES, not definitions (important change)
            if code in code_definitions:
                construct = code_definitions[code]['construct']
                theme_data = theme_info.get(construct)

                prompt += f"- Code: {code}\n"
                prompt += f"  - Definition: {code_definitions[code]['description']} (Examples: {code_definitions[code]['examples']})\n"
                if theme_data:
                    prompt += f"  - Theme: {theme_data['theme']}\n"
                    prompt += f"    - Definition: {theme_data['definition']}\n"
                else:
                    prompt += f"  - Theme: (Theme info not found for '{construct}')\n"
            else:
                prompt += f"- {code}: Definition not found.\n" # Important: Handle missing definition.


        prompt += "\nProvide magnitude ratings (1-7) AND BRIEF justifications for EACH code in the following JSON format:\n\n"
        prompt += """
{
  "Code 1": {
    "magnitude": 3,
    "justification": "Brief reason."
  },
  "Code 2": {
    "magnitude": 6,
    "justification": "Brief reason."
  }
}
"""
        prompt += "\nEnsure the JSON is valid and does not contain any extra characters or formatting.\n"

        print(f"\nIntensity coding prompt:\n\n{prompt}")
        logging.info(f"Intensity coding prompt: {prompt}")

        max_retries = 10
        delay = 5

        for attempt in range(max_retries):
            try:
                response = self.model.generate_content(
                    [prompt],
                    generation_config=LARGE_GENERATION_CONFIG,
                    safety_settings=SAFETY_SETTINGS
                ).text
                print(f"\nIntensity coding response:\n\n{response}")

                #Remove the markdown
                clean_response = remove_json_markdown(response)
                json_response = json.loads(clean_response)

                # Validate the structure of the response
                for code, data in json_response.items():
                    if not isinstance(data, dict) or "magnitude" not in data or "justification" not in data:
                        raise ValueError(f"Invalid JSON structure for code '{code}'.  Expected keys 'magnitude' and 'justification'.")
                    if not isinstance(data["magnitude"], int) or not 1 <= data["magnitude"] <= 7:
                         raise ValueError(f"Invalid magnitude value for code '{code}'. Expected an integer between 1 and 7.")
                    if not isinstance(data["justification"], str) or len(data["justification"]) == 0:
                        raise ValueError(f"Invalid justification for code '{code}'. Expected a non-empty string.")
                return json_response
            
            except json.JSONDecodeError as e:
                # Add prompt suggestions for JSON errors
                if attempt == 0:
                    prompt += "\n\n Ensure that if an excerpt ends with a single quote, that a double quotation mark is used to close the excerpt text."
                elif attempt == 1:
                    prompt += "\n\n If a double quote is used in an excerpt's quoted text, ensure each double quotation mark is escaped with a backslash."
                elif attempt == 2:
                    prompt += "\n\n Double-check that all key values start and end with a double quotation mark, a closing quotation mark appears to be missing."
                elif attempt ==3:
                    prompt += "\n\n Ensure the JSON is valid and well-formatted."

                logging.error(f"JSON decode error (attempt {attempt + 1}/{max_retries}): {e}")
                print(f"Error decoding JSON (attempt {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    logging.info(f"Retrying in {delay} seconds...")
                    print(f"Retrying in {delay} seconds...")
                    time.sleep(delay)

            except ValueError as e:
                logging.error(f"Data validation error (attempt {attempt + 1}/{max_retries}): {e}")
                print(f"Data validation error (attempt {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries -1:
                    logging.info(f"Retrying in {delay} seconds...  Adjusting prompt for next attempt.")
                    print(f"Retrying in {delay} seconds... Adjusting prompt for next attempt.")
                    prompt += f"\nError: {e}. Please correct the JSON output."
                    time.sleep(delay)
            
            except Exception as e:
                if "429" in str(e) or "Quota exceeded" in str(e):
                    print(f"Rate limit error: {e}. Retrying in {delay} seconds...")
                    time.sleep(delay)
                    delay *= 2
                else:
                    logging.exception(f"An unexpected error occurred: {e}")
                    print(f"An unexpected error occurred: {e}")
                    return None

        logging.error("Max retries reached. Giving up.")
        print("Max retries reached. Giving up.")
        return None