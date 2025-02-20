# code_compressor_client.py
import json
import vertexai
from vertexai.generative_models import GenerativeModel

from config import PROJECT_ID, LOCATION, GEMINI_MODEL, LARGE_GENERATION_CONFIG, SAFETY_SETTINGS
from utils import remove_json_markdown


class CodeCompressorClient:
    def __init__(self):
        vertexai.init(project=PROJECT_ID, location=LOCATION)
        self.model = GenerativeModel(GEMINI_MODEL)

    def compress_examples(self, codes_list, compression_type): 
        """
        Compresses the 'examples' in a LIST of code dictionaries using an LLM.
        Expects and returns a LIST of dictionaries.
        """

        codes_payload = json.dumps(codes_list, indent=2)  

        if compression_type == "1":
            prompt = (
                "You are an expert qualitative researcher summarizing text extracts.\n"
                "Below is a JSON object representing codes, their descriptions, themes (constructs), and examples (extracts), along with frequency counts.\n\n"
                f"{codes_payload}\n\n"
                "Your task is to return a JSON object with the EXACT SAME structure, keys, and values, "
                "EXCEPT that the value of the 'examples' key might be summarized if they are long. "
                "The goal is to reduce the overall size of the JSON object while retaining the original "
                "meaning, *especially* the relationship between the examples, the code, and the construct.\n\n"
                "Important Instructions:\n"
                "- Return ONLY a valid JSON object, with NO additional text or markdown formatting.\n"
                "- DO NOT change the 'code', 'description', 'construct', or 'frequency' values.\n"
                "- If an 'examples' value is already concise, keep it unchanged.\n"
                "- If an 'examples' value is long, summarize it concisely, preserving the core meaning.\n"
                "- Ensure that the summarized examples still clearly relate to the corresponding 'code' and 'construct'.\n"
                "- Maintain the original JSON structure (a LIST of dictionaries).  Do NOT add or remove any keys.\n"
                "- Return an empty LIST if there are any issues\n"
                "Example Input:\n"
                "```json\n"
                "[\n"
                " {\n"
                '    "code": "Code1",\n'
                '    "description": "A description of Code1",\n'
                '    "examples": "A very long and detailed example extract... (long text)",\n'
                '    "construct": "ThemeA",\n'
                '    "frequency": 38\n'
                " },\n"
                " {\n"
                '    "code": "Code2",\n'
                '    "description": "Description of Code2",\n'
                '    "construct": "ThemeB",\n'
                '    "examples": "Short example.",\n'
                '    "frequency": 5\n'
                " }\n"
                "]\n"
                "```\n\n"
                "Example Output (Illustrative):\n"
                "```json\n"
                "[\n"
                " {\n"
                '   "code": "Code1",\n'
                '   "description": "A description of Code1",\n'
                '   "construct": "ThemeA",\n'
                '   "examples": "Summarized example of Code1...",\n'  # Summarized
                '   "frequency": 38\n'
                "  },\n"
                " {\n"
                '   "code": "Code2",\n'
                '   "description": "Description of Code2",\n'
                '   "construct": "ThemeB",\n'
                '   "examples": "Short example.",\n'
                '   "frequency": 5\n'
                "  }\n"
                "]\n"
                "```\n"
            )

        if compression_type == "2":
            prompt = (
                "You are an expert qualitative researcher summarizing text extracts.\n"
                "Below is a JSON object representing codes, their descriptions, themes (constructs), and examples (extracts), along with frequency counts.\n\n"
                f"{codes_payload}\n\n"
                "Your task is to return a JSON object with the EXACT SAME structure, keys, and values, "
                "EXCEPT that the value of the 'description' and 'examples' keys might be summarized if they are long. "
                "The goal is to reduce the overall size of the JSON object while retaining the original "
                "meaning, *especially* the relationship between the examples, the code, and the construct.\n\n"
                "Important Instructions:\n"
                "- Return ONLY a valid JSON object, with NO additional text or markdown formatting.\n"
                "- DO NOT change the 'code', 'construct', or 'frequency' values.\n"
                "- If a 'description' or 'examples' value is already concise, keep it unchanged.\n"
                "- If a 'description' or 'examples' value is long, summarize it concisely, preserving the core meaning.\n"
                "- Ensure that the summarized description and examples still clearly relate to the corresponding 'code' and 'construct'.\n"
                "- Maintain the original JSON structure (a LIST of dictionaries).  Do NOT add or remove any keys.\n"
                "- Return an empty LIST if there are any issues\n"
                "Example Input:\n"
                "```json\n"
                "[\n"
                " {\n"
                '    "code": "Code1",\n'
                '    "description": "A description of Code1",\n'
                '    "examples": "A very long and detailed example extract... (long text)",\n'
                '    "construct": "ThemeA",\n'
                '    "frequency": 38\n'
                " },\n"
                " {\n"
                '    "code": "Code2",\n'
                '    "description": "A very long and detailed description... (long text)",\n'
                '    "construct": "ThemeB",\n'
                '    "examples": "Short example.",\n'
                '    "frequency": 5\n'
                " }\n"
                "]\n"
                "```\n\n"
                "Example Output (Illustrative):\n"
                "```json\n"
                "[\n"
                " {\n"
                '   "code": "Code1",\n'
                '   "description": "A description of Code1",\n'
                '   "construct": "ThemeA",\n'
                '   "examples": "Summarized example of Code1...",\n'  # Summarized
                '   "frequency": 38\n'
                "  },\n"
                " {\n"
                '   "code": "Code2",\n'
                '   "description": "Summarized description of Code2",\n'
                '   "construct": "ThemeB",\n'
                '   "examples": "Short example.",\n'
                '   "frequency": 5\n'
                "  }\n"
                "]\n"
                "```\n"
            )

        print(f"\nCompress examples prompt:\n\n{prompt}")

        try:
            response = self.model.generate_content(
                [prompt], generation_config=LARGE_GENERATION_CONFIG, safety_settings=SAFETY_SETTINGS
            )
            response_text = response.text
            print(f"\nCompress examples response:\n\n{response_text}")
            clean_response = remove_json_markdown(response_text)
            compressed_codes = json.loads(clean_response)

            # Validate the output format. 
            if not isinstance(compressed_codes, list):
                print("Error: LLM did not return a list. Returning empty list.")
                return [] 

            for item in compressed_codes:
                if not isinstance(item, dict) or not all(key in item for key in ["code", "description", "construct", "examples", "frequency"]):
                    print("Error:  LLM returned a list, but an item is malformed. Returning empty list.")
                    return [] 

            return compressed_codes 

        except Exception as e:
            print(f"Error during compression: {e}")
            return [] 