# AI-Theme-Analyzer

This Python script uses Google's Gemini language model on Vertex AI to perform thematic analysis on textual data across multiple DOCX files. Given a set of a priori themes, this tool inductively generates initial codes for human review, codes the full dataset, inductively generates sub- and meta-themes, then creates network graphs and reports to support within-case analyses and cross-case analyses of documents. 

## Features

- **DOCX File Parsing:** Reads data from DOCX files, extracting and formatting text for analysis.
- **Gemini-Powered Theme and Code Generation:** Employs the Gemini model to generate themes and codes based on initial data and predefined constructs.
- **Thematic Analysis:**  Identifies themes within individual documents and across multiple documents.
- **Iterative Human-in-the-loop Approach:**  This process is designed to promote transparency and human-LLM negotiation of meaning and control, through (1) human modification of system instructions, tasks, and themes/codebooks to guide the LLM, (2) LLM generated justification for the decisions made (i.e., proposed codes and theme mappings), and (3) human control to review and modify the output of each stage of analysis, used as input for later stages. 
- **Within-case Analysis:** Analyzes relationships between themes within the same document, identifying intersections, contradictions, and connections.
- **Cross-case Analysis:** Summarizes patterns and trends across multiple documents.
- **Structured Output:** Generates Excel spreadsheets to store coding results, thematic analysis, and cross-case analysis.
- **Visualization:** Creates network graphs to visualize relationships between themes and codes.
- **Timestamped Results:** Appends a timestamp to output file names for easy tracking.

## Prerequisites

- **Google Cloud Project:** A Google Cloud project with Vertex AI enabled and the Gemini model available.
- **Python Environment:** Python 3.7 or higher.
- **API Key:** A valid API key for the Gemini model, obtained through your Google Cloud project.
- **themes.json:** A JSON file defining the themes and their definitions. Create this file based on the `themes_template.json` file.

## Installation

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/encorelab/ai-theme-analyser.git
   cd ai-theme-analyzer

2. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Log in to Google Cloud:** 
   ```bash
   gcloud auth application-default login
   ```

4. **Set Environment Variables:**
Create a file named .env in the project directory and add the following:
```
PROJECT_ID=your-project-id
LOCATION=your-project-location
GEMINI_MODEL=gemini-1.5-pro-002 # or your preferred model
```

## Usage
1. **Prepare Input Documents:**
- Place your DOCX files in the input_files directory

2. **Create themes.json:**
- Create a themes.json file in the project directory, defining your themes and their definitions based on the structure in themes_template.json

3. **Run the Script:**
- Use the --client argument to specify the analysis stage to run:
   ```bash
   python main.py --client code_generator   # Generate initial codes
   python main.py --client theme_generator  # Generate themes and hierarchy
   python main.py --client coding_client    # Apply codes to documents
   python main.py --client intra_text_analyzer  # Analyze relationships within documents 
   python main.py --client cross_document_analyzer # Analyze patterns across documents
   ```

4. **View Results:**
- Analyzed data and visualizations will be saved in the output_files directory.

## Running Tests
To run the unit tests, use the following command:
   ```bash
   python -m unittest -v -f test_ai_emotion_analyzer.py  
   ```
# Important Notes
- **Iterative Process:** Thematic analysis is iterative. You may need to refine codes and themes as you analyze more data.
- **Customisation:** Adjust parameters like NUM_DOCS_FOR_CODE_GENERATION, BATCH_SIZE, and the content of themes.json to suit your specific needs and data.
- **Logging:** The script logs events to log.txt
- **Visualization:** Network graph visualizations provide insights into theme relationships.
