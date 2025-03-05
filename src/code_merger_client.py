# code_merger_client.py
import json
import time
import logging

import vertexai
from vertexai.generative_models import GenerativeModel, Part, SafetySetting, FinishReason

from config import PROJECT_ID, LOCATION, GEMINI_MODEL, LARGE_GENERATION_CONFIG, SAFETY_SETTINGS, RESEARCH_QUESTION_FILE
from src.utils import remove_json_markdown


# Configure logging
LOG_FILE = "log.txt"
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")


class CodeMergerClient:
    def __init__(self):
        vertexai.init(project=PROJECT_ID, location=LOCATION)
        self.model = GenerativeModel(GEMINI_MODEL)

    def merge_themes(self, codes, themes, merge_threshold):
        """
        Merges semantically similar codes within themes based on a threshold.
        """

        merged_codes_result = {}  # Initialize the dictionary

        for theme in themes:
            theme_codes = {
                code: details
                for code, details in codes.items()
                if details["theme"] == theme["theme"]
            }

            if len(theme_codes) > merge_threshold:
                # Prepare data for LLM
                codes_for_llm = []
                for code, details in theme_codes.items():
                    codes_for_llm.append({
                        "code": code,
                        "description": details["description"],
                        "examples": details["examples"],
                    })
                codes_payload = json.dumps(codes_for_llm, indent=2)

                # Construct the prompt
                prompt = (
                    "You are an expert qualitative researcher tasked with merging semantically similar codes "
                    "within a thematic coding framework.  Below is a JSON object representing codes belonging "
                    f"to the theme '{theme}'.\n\n"
                    f"{codes_payload}\n\n"
                    "Merge ONLY semantically very similar codes.  The merged codes should accurately represent the "
                    "original codes and their associated excerpts.  Do NOT merge codes that are conceptually distinct, "
                    "even if they share some keywords.  The goal is to consolidate very similar concepts, NOT to reduce "
                    "the number of codes drastically.\n\n"
                    "Provide a JSON response in the following format:\n\n"
                    "```json\n"
                    "{\n"
                    '  "New Code Name 1": {\n'
                    '    "new_description": "Concise description of the merged code.",\n'
                    '    "examples": ["Example excerpt 1", "Example excerpt 2", ...],\n'
                    '    "merged_codes": ["Original Code 1", "Original Code 2", ...]\n'
                    "  },\n"
                    '  "New Code Name 2": {\n'
                    '    "new_description": "Concise description of the merged code.",\n'
                    '    "examples": ["Example excerpt 3", "Example excerpt 4", ...],\n'
                    '    "merged_codes": ["Original Code 3", "Original Code 4", ...]\n'
                    "  }\n"
                    "  ...\n"
                    "}\n"
                    "```\n\n"
                    "If no codes within the theme should be merged, return an empty JSON object: `{}`. "
                     "Ensure the new description accurately reflects the combined meaning of the merged codes, "
                    "and the examples are representative excerpts from the original codes. Ensure that the JSON object does not contain any markdown formatting, comments, or additional text besides the JSON."
                )
                print(f"\nMerge codes prompt for theme '{theme}':\n\n{prompt}")
                # Call the model
                response = self.model.generate_content(
                    [prompt], generation_config=LARGE_GENERATION_CONFIG, safety_settings=SAFETY_SETTINGS
                )
                
                try:
                    response_text = response.text
                    print(f"\nMerge codes response for theme '{theme}':\n\n{response_text}")
                    clean_response = remove_json_markdown(response_text)
                    merged_codes_result.update(json.loads(clean_response))

                except (json.JSONDecodeError, IndexError) as e:
                    print(f"Error processing response for theme '{theme}': {e}")
                    print(f"Raw response: {response_text}") # Print raw text for inspection
                    continue # Continue to the next theme
            else:
                print(f"Skipping merging for theme '{theme}' (number of codes: {len(theme_codes)} is not greater than the threshold: {merge_threshold})")

        return merged_codes_result