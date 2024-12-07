import json

import vertexai
from vertexai.generative_models import GenerativeModel, Part, SafetySetting, FinishReason

from config import PROJECT_ID, LOCATION, GEMINI_MODEL, LARGE_GENERATION_CONFIG, SAFETY_SETTINGS
from utils import remove_json_markdown


class ThemeGeneratorClient:
    def __init__(self):
        vertexai.init(project=PROJECT_ID, location=LOCATION)
        self.model = GenerativeModel(GEMINI_MODEL,
            system_instruction="""You are a research assistant specializing in thematic analysis of qualitative data. Your task is to generate a hierarchical list of potential meta-themes, themes, sub-themes, and codes based on the provided codes and themes, along with a brief description of each. Ensure the hierarchy is clear, concise, and captures the overarching patterns and meanings represented by the codes and themes."""
        )

    def generate_themes(self, codes, themes):
        # Combine codes and themes into a single JSON payload
        codes_payload = json.dumps(codes, indent=2)
        themes_payload = json.dumps(themes, indent=2)

        # Prompt template
        prompt = (
            "The following is a JSON object containing the full list of themes with descriptions, the full list of codes with descriptions, and the theme each code is affiliated with, used for thematic coding of engineering student positionality statements throughout a design course.\n\n"
            f"{{'themes':{themes_payload}, 'codes':{codes_payload}}}\n\n"
            "Provide a comprehensive hierarchical list of meta-themes, themes, sub-themes, and codes (whereby themes and codes are exact matches to the JSON data above), along with descriptions for each meta-theme, theme, and sub-theme, as a JSON object in the following format:\n\n"
"""
{
  "Meta-theme 1": {
    "description": "Description of meta-theme 1",
    "themes": {
      "Theme 1": {
        "description": "Description of theme 1",
        "sub-themes": {
          "Sub-theme 1": {
            "description": "Description of sub-theme 1",
            "codes": ["Code 1", "Code 2"]
          },
          "Sub-theme 2": {
            "description": "Description of sub-theme 2",
            "codes": ["Code 3", "Code 4"]
          }
        }
      },
      "Theme 2": {
        "description": "Description of theme 2",
        "sub-themes": {
          "Sub-theme 3": {
            "description": "Description of sub-theme 3",
            "codes": ["Code 5", "Code 6"]
          }
        }
      }
    }
  },
  "Meta-theme 2": {
    // ...
  }
}
"""
            "\nEnsure (1) there are multiple sub-themes for each theme and (2) the JSON is valid and does not contain any extra characters or formatting.\n"
        )
        print(f"\nGenerate themes prompt:\n\n{prompt}")

        # Call the model to predict and get results in string format
        response = self.model.generate_content([prompt], generation_config=LARGE_GENERATION_CONFIG, safety_settings=SAFETY_SETTINGS).text
        print(f"\nGenerate themes response:\n\n{response}")

        clean_response = remove_json_markdown(response)

        # Convert the results to a python dictionary
        themes_hierarchy = json.loads(clean_response)

        return themes_hierarchy
