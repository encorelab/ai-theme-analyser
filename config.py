import os
from dotenv import load_dotenv
import vertexai.preview.generative_models as generative_models
from vertexai.generative_models import GenerativeModel, Part, SafetySetting, FinishReason


load_dotenv() # Load environment variables from .env file

# Project Configuration
PROJECT_ID = os.getenv("PROJECT_ID")
LOCATION = os.getenv("LOCATION")
GEMINI_MODEL = os.getenv("GEMINI_MODEL")

BATCH_SIZE = 1

# Directory Setup
INPUT_DIR = "input_files"
OUTPUT_DIR = "output_files"
SPREADSHEET_FILENAME = "ParentSatisfaction_24Jun24.xlsx" #Modify as needed
RESEARCH_QUESTION_FILE = "research_question.txt"
CONSTRUCT_FILE = "construct.txt"

# Stage 2: Initial Code Generation
NUM_DOCS_FOR_CODE_GENERATION = 40

# Data Extraction Targets (modify as needed)
SAFETY_SETTINGS = [
    SafetySetting(
        category=SafetySetting.HarmCategory.HARM_CATEGORY_HATE_SPEECH,
        threshold=SafetySetting.HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
    ),
    SafetySetting(
        category=SafetySetting.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
        threshold=SafetySetting.HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
    ),
    SafetySetting(
        category=SafetySetting.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
        threshold=SafetySetting.HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
    ),
    SafetySetting(
        category=SafetySetting.HarmCategory.HARM_CATEGORY_HARASSMENT,
        threshold=SafetySetting.HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
    ),
]

DEFAULT_GENERATION_CONFIG = {
    "max_output_tokens": 1024,
    "temperature": 0,
    "top_p": 0.95,
}

LARGE_GENERATION_CONFIG = {
    "max_output_tokens": 8192,
    "temperature": 0,
    "top_p": 0.95,
}