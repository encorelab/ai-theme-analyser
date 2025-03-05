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


class CodeGenerationClient:
    def __init__(self):
        vertexai.init(project=PROJECT_ID, location=LOCATION)
        self.model = GenerativeModel(
            GEMINI_MODEL,
            system_instruction=
            """You are a research assistant specializing in thematic analysis of qualitative data. Your task is to code excerpts from documents and generate a comprehensive list of codes for the specified theme. Focus on capturing the key concepts, emotions, and magnitudes expressed in the text."""
        )

    def generate_codes(self,
                               text_chunk,
                               codes,
                               construct):
        """
        Generates initial codes for a single code construct (or theme) by analyzing 
        an extract of text and considering existing codes.
        """

        construct_name = construct['theme']
        construct_definition = construct['definition']
        construct_examples = construct['examples']
        construct_exclusion_criteria = construct['exclude']

        # Combine codes, descriptions, and constructs into a single JSON payload
        codes_with_data = []
        for code, code_data in codes.items():
            if code_data['theme'] == construct_name: # Only include codes for current construct
                codes_with_data.append({
                    "code": code, 
                    "description": code_data['description'],
                    "theme": code_data['theme'],
                })
        codes_payload = json.dumps({
            "codes": codes_with_data
        },
                                   indent=2)

        # Read the research question from the file
        try:
            with open(RESEARCH_QUESTION_FILE, "r") as f:
                research_question = f.read().strip()
        except FileNotFoundError:
            print(f"Error: Research question file not found - {RESEARCH_QUESTION_FILE}")
            research_question = ""  # Set an empty string if the file is not found

        if construct_exclusion_criteria:
            theme_prompt = f"Theme:\n\n{construct_name}: {construct_definition} (e.g., {construct_examples}); Exclusion criteria: {construct_exclusion_criteria}\n\n"
        else:
            theme_prompt = f"Theme:\n\n{construct_name}: {construct_definition} (e.g., {construct_examples})\n\n"

        # Prompt template
        prompt = (
            f"{theme_prompt}\n\n"
            "Existing codes (if any) with descriptions related to the specified theme:\n\n"
            f"{codes_payload}\n\n"
            "Text you will analyze:\n\n"
            f"{text_chunk}"
            "Task:\n\n"
            "1. Analyze the provided text to find extracts related to both the specified theme and pre-university (and pre-Praxis project) experiences. You may break down paragraphs into smaller extracts as needed to ensure accurate and granular coding.\n"
            "2. If extracts in the provided text are found that include direct, compelling, and explicit reference to the specified theme and relate to pre-university experiences, then assign appropriate codes to the sentences or multi-sentence excerpts.\n"
            "    - If a relevant code already exists in the JSON object above, then use it by adding a coded_excerpts item.\n"
            "    - If an excerpt includes direct, compelling, and explicit relevance to the theme but does not have an existing relevant code, then you MUST also add and define the new code.\n"
            "3. If you are adding and defining a new code:\n" 
            "    - New codes are added as short labels to the new_codes key in the JSON response. Provide the proposed name, excerpt, construct it belongs to, definition, as well as justification and confidence for adding the new code.\n"
            "    - Note: new code names should not include any commas within its name"
            "4. If some or all excerpts are not directly relevant to the parent theme, exclude them from your response entirely, and do not feel obligated to create new codes.\n\n"
            "Provide the response in a JSON object in the following format:\n\n"
            """
{
  "coded_excerpts": {
    "Excerpt 1": ["Code 1", "Code 3"],
    "Excerpt 2": ["Code 2"],
    "Excerpt 3": ["Code 1"]
  },
  "new_codes": {
    "[Theme name]-[New Code Name]": {
      "excerpt": "[Text quote receiving this new code]",
      "theme": "[Name of specified theme]",
      "description": "[Proposed description of new code]", 
      "justification": "[Justification for proposing new code based on excerpt]"
      "probability": "[Classification confidence score from 0-1, where 1 is full confidence]
    }
  }
}
"""
            "\nEnsure the JSON is valid and does not contain any extra characters or formatting.\n"
        )
        print(f"\nThematic coding prompt:\n\n{prompt}")

        logging.info("Starting coding process.")
        logging.info(f"Codes sent to model: {codes} ({len(codes)})")
        logging.info(f"Number of words in excerpt: {len(text_chunk.split())}")

        max_retries = 10
        delay = 1  # seconds

        for attempt in range(max_retries):
            try:
                # Call the model to predict and get results in string format
                response = self.model.generate_content(
                    [prompt],
                    generation_config=LARGE_GENERATION_CONFIG,
                    safety_settings=SAFETY_SETTINGS).text
                print(f"\nThematic coding response:\n\n{response}")
                clean_response = remove_json_markdown(response)
                json_response = json.loads(clean_response)

                break  # Exit the loop if successful
            except json.JSONDecodeError as e:
                if attempt == 0:
                    prompt += "\n\n Ensure that if an excerpt ends with a single quote, that a double quotation mark is used to close the excerpt text."
                elif attempt == 1:
                    prompt += "\n\n If a double quote is used in an excerpt's quoted text, ensure each double quotation mark is escaped with a backslash."
                elif attempt == 2:
                    prompt += "\n\n Double-check that all key values start and end with a double quotation mark, a closing quotation mark appears to be missing."
            
                # Log JSON decode error
                logging.error(
                    f"JSON decode error (attempt {attempt+1}/{max_retries}): {e}"
                )
                print(
                    f"Error decoding JSON (attempt {attempt+1}/{max_retries}): {e}"
                )
                if attempt < max_retries - 1:
                    logging.info(f"Retrying in {delay} seconds...")
                    print(f"Retrying in {delay} seconds...")
                    time.sleep(delay)
                else:
                    logging.error("Max retries reached. Giving up.")
                    print("Max retries reached. Giving up.")
                    # Handle the error (e.g., skip this excerpt, return an empty result)
                    return {}, {}
            except Exception as e:
                if "429" in str(e) or "Quota exceeded" in str(e):  # Check for rate limit error
                    print(f"Rate limit error: {e}. Retrying in {delay} seconds...")
                    time.sleep(delay)
                    delay *= 2  # Double the delay
                else:
                    # Log any other errors
                    logging.exception(f"An unexpected error occurred: {e}")
                    print(f"An unexpected error occurred: {e}")
                    return {}, {}  # Or handle the error appropriately

        excerpt_codings = json_response.get('coded_excerpts', {})
        new_codes = json_response.get('new_codes', {})

        # Log successful response
        logging.info(
            f"AI response received (attempt {attempt+1}/{max_retries}).")
        logging.info(f"New codes proposed: {list(new_codes.keys())}")

        # Add new code excerpts and codes to excerpt_codings, if not already added
        for new_code_name, new_code_data in new_codes.items():
            excerpt = new_code_data['excerpt']
            if excerpt in excerpt_codings:
                if new_code_name not in excerpt_codings[excerpt]:  # Check if code already exists
                    excerpt_codings[excerpt].append(new_code_name)
            else:
                excerpt_codings[excerpt] = [new_code_name]

        # Remove excerpts with empty code lists
        excerpt_codings = {
            excerpt: codes
            for excerpt, codes in excerpt_codings.items() if codes
        }

        return excerpt_codings, new_codes