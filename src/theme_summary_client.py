import json
import time
import logging

import vertexai
from vertexai.generative_models import GenerativeModel
from src.utils import remove_json_markdown
from config import RESEARCH_QUESTION_FILE


from config import PROJECT_ID, LOCATION, GEMINI_MODEL, LARGE_GENERATION_CONFIG, SAFETY_SETTINGS

# Configure logging
LOG_FILE = "log.txt"
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")


class ThemeSummaryClient:
    def __init__(self):
        vertexai.init(project=PROJECT_ID, location=LOCATION)
        self.model = GenerativeModel(
            GEMINI_MODEL,
            system_instruction=(
                "You are a qualitative researcher specializing in thematic analysis. "
                "You will be provided with a theme (or construct) name and definition, "
                "a sub-theme name and description, a list of codes relevant to this "
                "sub-theme with code definitions, and all extracts coded with codes "
                "that belong to this sub-theme. You will also receive the research "
                "question that guides this study."
            )
        )
    def generate_theme_summary(self, theme, sub_theme, excerpts, code_definitions, theme_definitions):
        """
        Generates a comprehensive summary/report for a sub-theme.
        """

        # Load research question
        try:
            with open(RESEARCH_QUESTION_FILE, "r") as f:
                research_question = f.read().strip()
        except FileNotFoundError:
            print(f"Error: Research question file not found - {RESEARCH_QUESTION_FILE}")
            research_question = ""  # Set an empty string if the file is not found

        # Construct the prompt using a JSON-like structure within f-strings
        prompt = f"""
Analyze the following information and provide a comprehensive summary/report for the sub-theme:

{{"research_question": "{research_question}",
"theme": {json.dumps(theme_definitions[theme], indent=2)},
"sub_theme": "{sub_theme}",
"codes": {json.dumps(code_definitions, indent=2)},
"excerpts": {json.dumps(excerpts, indent=2)}
}}

Your task is to generate a comprehensive summary/report for this sub-theme.
The summary should:
1.  Include a concise introduction to the sub-theme.
2.  Use the frequency of related codes within the sub-theme to inform the common patterns to be identified.
3.  Identify common patterns and insights within the data, supported by quotes from excerpts.
4.  Use an individual or two to exemplify these patterns more personally, using filename IDs to identify extracts associated with particular individuals
5.  Relate the findings to the research question: "{research_question}".
"""

        print(f"\nTheme summary prompt:\n\n{prompt}")
        logging.info(f"Theme summary prompt: {prompt}")

        max_retries = 10
        delay = 5

        for attempt in range(max_retries):
            try:
                response = self.model.generate_content(
                    [prompt],
                    generation_config=LARGE_GENERATION_CONFIG,
                    safety_settings=SAFETY_SETTINGS
                ).text
                print(f"\nTheme summary response:\n\n{response}")

                #No need to remove markdown in this class
                #clean_response = remove_json_markdown(response)
                #json_response = json.loads(clean_response)
                # Return the raw text response
                return response


            except json.JSONDecodeError as e:
                logging.error(f"JSON decode error (attempt {attempt + 1}/{max_retries}): {e}")
                print(f"Error decoding JSON (attempt {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    logging.info(f"Retrying in {delay} seconds...")
                    print(f"Retrying in {delay} seconds...")
                    time.sleep(delay)
            except Exception as e:
                if "429" in str(e) or "Quota exceeded" in str(e):
                    print(f"Rate limit error: {e}. Retrying in {delay} seconds...")
                    time.sleep(delay)
                    delay *= 2  # Exponential backoff
                else:
                    logging.exception(f"An unexpected error occurred: {e}")
                    print(f"An unexpected error occurred: {e}")
                    return None  # Or raise the exception if you want to stop execution

        logging.error("Max retries reached. Giving up.")
        print("Max retries reached. Giving up.")
        return None