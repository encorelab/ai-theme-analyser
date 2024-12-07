import json

import vertexai
from vertexai.generative_models import GenerativeModel, Part, SafetySetting, FinishReason

from config import PROJECT_ID, LOCATION, GEMINI_MODEL, LARGE_GENERATION_CONFIG, SAFETY_SETTINGS


class CrossDocumentAnalyzerClient:
    def __init__(self):
        vertexai.init(project=PROJECT_ID, location=LOCATION)
        self.model = GenerativeModel(GEMINI_MODEL,
            system_instruction="""You are a research assistant specializing in thematic analysis of qualitative data. Your task is to analyze intra-text analysis results across multiple documents and identify overarching patterns, clusters of themes, common intersections/contradictions/connections, or other syntheses that would be worth discussing in a thematic analysis report."""
        )

    def analyze_cross_document(self, analyze_cross_document):

        json_payload = json.dumps(analyze_cross_document, indent=2)

        # Prompt template with emotion codebook and classification request
        prompt = (
            "The following are engineering student positionality statements excerpts to be analyzed, organised in a JSON object by paragraph indices and paragraph contents.\n\n"
            f"{json_payload}\n\n"
            "Your output should be a structured summary highlighting the key findings, organized by relevant categories or patterns. This should be written in the tone of a results section of a thematic analysis in a social sciences academic paper."
        )

        print(f"\nCross document analysis prompt:\n\n{prompt}")

        # Call the model to predict and get results in string format
        response = self.model.generate_content([prompt], generation_config=LARGE_GENERATION_CONFIG, safety_settings=SAFETY_SETTINGS).text
        print(f"\nCross document analysis response:\n\n{response}")

        return response