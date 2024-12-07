import os
import re
import json
from collections import defaultdict
import matplotlib.pyplot as plt
import networkx as nx
import docx
import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
import time
import logging

from config import RESEARCH_QUESTION_FILE

# Configure logging
LOG_FILE = "log.txt"
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, 
                    format="%(asctime)s - %(levelname)s - %(message)s")


def remove_json_markdown(text):
    """Removes JSON markdown from a string."""
    pattern = re.compile(r'```json\s*(.*?)\s*```', re.DOTALL)
    return pattern.sub(r'\1', text)


def extract_paragraphs_from_docx(filepath):
    """
    Extracts paragraphs from a docx file, formats them with markdown bolding for headings,
    and returns them as a list of strings.
    """
    doc = docx.Document(filepath)
    formatted_paragraphs = []
    for p in doc.paragraphs:
        if p.text.strip():
            if p.style.name.startswith('Heading 1') or p.style.name.startswith('Heading 2'):
                formatted_paragraphs.append(f"**{p.text}**")
            else:
                formatted_paragraphs.append(p.text)
    return formatted_paragraphs


def write_coding_results_to_excel(all_files_excerpt_codings: dict[str, dict[str, list[str]]], 
                                 new_codes_by_file: dict[str, dict[str, dict]], 
                                 output_file: str):
    """
    Writes the coding results and justifications to an Excel file.
    Ensures the Excel file is created even if the data is blank.
    """

    print(f"Received all_files_excerpt_codings: {all_files_excerpt_codings}")
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Prepare data for Excel output (codings) for ALL files
            all_codings_data = []
            for filename, excerpt_codings in all_files_excerpt_codings.items():
                for excerpt, codes in excerpt_codings.items():
                    # Remove illegal characters from excerpt
                    cleaned_excerpt = ILLEGAL_CHARACTERS_RE.sub(r'', excerpt)  
                    all_codings_data.append({
                        'filename': filename,
                        'excerpt': cleaned_excerpt,  # Use the cleaned excerpt
                        'codings': ', '.join(codes)
                    })

            # Write codings to Excel for ALL files (even if empty)
            codings_df = pd.DataFrame(all_codings_data)  # Create DataFrame even if empty
            codings_df.to_excel(writer, sheet_name='codings', index=False)
            writer.sheets['codings'].sheet_state = 'visible' 

            # Prepare data for Excel output (new_codes)
            all_justifications_data = []
            print("Exporting new codes by file to code_justifications sheet:\n" + json.dumps(new_codes_by_file, indent=4)) 
            for filename, new_codes in new_codes_by_file.items():
                for code, data in new_codes.items():

                    all_justifications_data.append({
                        'code': code,
                        'filename': filename,
                        'excerpt': data['excerpt'],
                        'theme': data.get('theme', ''),
                        'description': data.get('description', ''),
                        'justification': data['justification'],
                        'probability': data['probability']
                    })

            # Write code justifications to Excel (even if empty)
            justifications_df = pd.DataFrame(all_justifications_data)  # Create DataFrame even if empty
            justifications_df.to_excel(writer, sheet_name='code_justifications', index=False)
            writer.sheets['code_justifications'].sheet_state = 'visible' 

    except Exception as e:
        print(f"An error occurred while writing to Excel: {e}")


def load_themes_from_file(filepath):
    """
    Loads themes and their definitions from a JSON file.
    """
    try:
        with open(filepath, 'r') as f:
            themes_data = json.load(f)
        
        themes = []
        theme_definitions = {}
        for item in themes_data:
            themes.append(item["theme"])
            theme_definitions[item["theme"]] = item["definition"]

        return themes, theme_definitions

    except FileNotFoundError:
        print(f"Error: File not found - {filepath}")
        return [], {}
    
    
def generate_initial_codes(directory,
                           themes,
                           coding_client,
                           words_per_chunk=1200,
                           num_docs=None,
                           time_between_calls=6):
    """
    Generates initial codes with dynamic delay to avoid rate limiting.

    Args:
        directory: Directory with docx files.
        themes: Dictionary of code constructs and definitions.
        coding_client: CodeGeneratorClient instance.
        words_per_chunk: Approximate words per chunk.
        num_docs: Number of documents to process (all if None).
        time_between_calls: Target time between calls in seconds.

    Returns:
        Tuple of coding results.
    """
    start_time = time.time()
    all_codes = {}
    code_descriptions = {}
    code_constructs = {}
    all_files_excerpt_codings = {}
    new_codes_by_file = {}

    docx_files = [
        f for f in os.listdir(directory)
        if f.endswith('.docx') and not f.startswith('~$')
    ]
    if num_docs is not None:
        docx_files = docx_files[:num_docs]
    total_files = len(docx_files)
    processed_files = 0

    for filename in docx_files:
        filepath = os.path.join(directory, filename)
        if filename.endswith('.docx') and not filename.startswith('~$'):
            paragraphs = extract_paragraphs_from_docx(filepath)
            file_excerpt_codings = {}
            paragraph_chunks = chunk_paragraphs(paragraphs, words_per_chunk)

            for chunk in paragraph_chunks:
                for construct in themes:
                    call_start_time = time.time()  # Record call start time

                    excerpt_codings, all_codes, new_codes = generate_codes_for_chunk(
                        chunk, construct, coding_client, all_codes,
                        code_descriptions, code_constructs
                    )

                    # Merge excerpt_codings into file_excerpt_codings
                    for excerpt, codes in excerpt_codings.items():
                        if excerpt in file_excerpt_codings:
                            for code in codes:
                                if code not in file_excerpt_codings[excerpt]:
                                    file_excerpt_codings[excerpt].append(code)
                        else:
                            file_excerpt_codings[excerpt] = codes

                    # Append new_codes to the existing codes for the filename
                    if filename not in new_codes_by_file:
                        new_codes_by_file[filename] = {}
                    new_codes_by_file[filename].update(new_codes)

                    call_end_time = time.time()  # Record call end time
                    call_duration = call_end_time - call_start_time

                    # Calculate dynamic delay
                    delay = max(0, time_between_calls - call_duration) 
                    time.sleep(delay)

            all_files_excerpt_codings[filename] = file_excerpt_codings

        # --- Print timestamp and progress ---
        processed_files += 1
        elapsed_time = time.time() - start_time
        remaining_files = total_files - processed_files
        print(
            f"Processed {filename} ({processed_files}/{total_files} files). Time elapsed: {elapsed_time:.2f} seconds. Remaining: {remaining_files} files."
        )
    return all_codes, code_descriptions, code_constructs, all_files_excerpt_codings, new_codes_by_file

def chunk_paragraphs(paragraphs, words_per_chunk=1200):
    """Chunks a list of paragraphs into smaller lists based on word count."""
    paragraph_chunks = []
    current_chunk = []
    current_word_count = 0
    for paragraph in paragraphs:
        current_chunk.append(paragraph)
        current_word_count += len(paragraph.split())
        if current_word_count >= words_per_chunk:
            paragraph_chunks.append(current_chunk)
            current_chunk = []
            current_word_count = 0
    if current_chunk:  # Add the last chunk if not empty
        paragraph_chunks.append(current_chunk)
    return paragraph_chunks


def generate_codes_for_chunk(chunk, construct, code_generation_client, all_codes, code_descriptions, code_constructs):
    """Generates codes for a single chunk of text and accumulates the results."""
    # Create data_for_chunk as a string with markdown formatting
    data_for_chunk = ""
    for paragraph in chunk:
        data_for_chunk += f"{paragraph}\n\n"

    # Call generate_initial_codes with construct definition and formatted string
    excerpt_codings, new_codes = code_generation_client.generate_codes(
                        data_for_chunk, all_codes, construct 
                    )

    # Add new codes and descriptions to the ongoing lists
    for code, data in new_codes.items():
        if code not in all_codes:
            all_codes[code] = data.copy()

    return excerpt_codings, all_codes, new_codes


def perform_coding(directory, all_codes, code_descriptions, code_constructs, coding_client, paragraphs_per_batch=1000, codes_per_batch=11):
    """
    Performs coding on each docx file in batches, potentially updating the codes list.
    Writes coding results and justifications to Excel.
    Skips files that start with ~$.
    Groups codes by construct for batching.
    """

    start_time = time.time()  # Record the start time of the script
    all_files_excerpt_codings = {}  # Store codings for all excerpts in all files
    new_codes_by_file = {}

    docx_files = [f for f in os.listdir(directory) if f.endswith('.docx') and not f.startswith('~$')]
    total_files = len(docx_files)
    processed_files = 0

    for filename in docx_files:
        filepath = os.path.join(directory, filename)
        if filename.endswith('.docx') and not filename.startswith('~$'):
            filepath = os.path.join(directory, filename)
            paragraphs = extract_paragraphs_from_docx(filepath)

            # Batch paragraphs by word count (1000-1500 words per batch)
            paragraph_batches = []
            current_batch = []
            current_word_count = 0
            for paragraph in paragraphs:
                current_batch.append(paragraph)
                current_word_count += len(paragraph.split())
                if current_word_count >= paragraphs_per_batch:
                    paragraph_batches.append(current_batch)
                    current_batch = []
                    current_word_count = 0
            if current_batch:  # Add the last batch if not empty
                paragraph_batches.append(current_batch)

            all_excerpt_codings = {}  # Store codings for all excerpts in the file

            # Group codes by construct
            codes_by_construct = {}
            for code in all_codes:
                construct = code_constructs[code]
                if construct not in codes_by_construct:
                    codes_by_construct[construct] = []
                codes_by_construct[construct].append(code)

            code_batch = []
            current_construct_codes = []
            for construct, codes in codes_by_construct.items():
                current_construct_codes.extend(codes)
                if len(current_construct_codes) >= codes_per_batch:
                    code_batch = current_construct_codes[:codes_per_batch]
                    current_construct_codes = current_construct_codes[codes_per_batch:]
                    # Create batches for code_descriptions and code_constructs
                    code_descriptions_batch = {code: code_descriptions[code] for code in code_batch}
                    code_constructs_batch = {code: code_constructs[code] for code in code_batch}

                    for paragraph_batch in paragraph_batches:
                        # Create data_for_file as a string with markdown formatting
                        data_for_file = ""
                        for paragraph in paragraph_batch:
                            data_for_file += f"{paragraph}\n\n"

                        # Call apply_codes with code descriptions, constructs, and the formatted string
                        excerpt_codings, new_codes = coding_client.apply_codes(
                            data_for_file, code_batch, code_descriptions_batch, code_constructs_batch
                        )

                        # Add new codes and descriptions to the ongoing lists
                        for code, data in new_codes.items():
                            if code not in all_codes: 
                                all_codes.append(code)
                                if 'description' in data:
                                    code_descriptions[code] = data['description']
                                else:
                                    print(f"Warning: Description missing for new code '{code}'.")
                                if 'theme' in data:
                                    code_constructs[code] = data['theme']
                                else:
                                    print(f"Warning: Construct missing for new code '{code}'.")

                            # Add the new code to the excerpt_codings for the current excerpt
                            excerpt = data['excerpt']
                            if excerpt in excerpt_codings:
                                excerpt_codings[excerpt].append(code)
                            else:
                                excerpt_codings[excerpt] = [code]

                        all_excerpt_codings.update(excerpt_codings)  # Update the overall codings
                        new_codes_by_file[filename] = new_codes
            
            # Handle any remaining codes
            if current_construct_codes:
                code_batch = current_construct_codes
                # Create batches for code_descriptions and code_constructs
                code_descriptions_batch = {code: code_descriptions[code] for code in code_batch}
                code_constructs_batch = {code: code_constructs[code] for code in code_batch}

                for paragraph_batch in paragraph_batches:
                    # Create data_for_file as a string with markdown formatting
                    data_for_file = ""
                    for paragraph in paragraph_batch:
                        data_for_file += f"{paragraph}\n\n"

                    # Call apply_codes with code descriptions, constructs, and the formatted string
                    excerpt_codings, new_codes = coding_client.apply_codes(
                        data_for_file, code_batch, code_descriptions_batch, code_constructs_batch
                    )

                    # Add new codes and descriptions to the ongoing lists
                    for code, data in new_codes.items():
                        if code not in all_codes:  # Avoid duplicates
                            all_codes.append(code)
                            if 'description' in data:
                                code_descriptions[code] = data['description']
                            else:
                                print(f"Warning: Description missing for new code '{code}'.")
                            if 'theme' in data:
                                code_constructs[code] = data['theme']
                            else:
                                print(f"Warning: Construct missing for new code '{code}'.")

                        # Add the new code to the excerpt_codings for the current excerpt
                        excerpt = data['excerpt']  # Get the excerpt associated with the new code
                        if excerpt in excerpt_codings:
                            excerpt_codings[excerpt].append(code)  # Add the new code to the existing list
                        else:
                            excerpt_codings[excerpt] = [code]  # Create a new list with the new code

                    all_excerpt_codings.update(excerpt_codings)  # Update the overall codings

            all_files_excerpt_codings[filename] = all_excerpt_codings

        # --- Print timestamp and progress ---
        processed_files += 1
        elapsed_time = time.time() - start_time
        remaining_files = total_files - processed_files
        print(f"Processed {filename} ({processed_files}/{total_files} files). Time elapsed: {elapsed_time:.2f} seconds. Remaining: {remaining_files} files.")

    return all_codes, code_descriptions, code_constructs, all_files_excerpt_codings, new_codes_by_file


def perform_analysis_and_reporting(output_file, themes, analyzer_client, intra_text_analyzer):
    """
    Performs thematic and intra-text analysis on coded data and writes results to Excel.
    """
    codings_df = pd.read_excel(output_file, sheet_name='codings')

    analysis_output_file = os.path.join(OUTPUT_DIR, f"thematic_analysis_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")
    intra_text_output_file = os.path.join(OUTPUT_DIR, f"intra_text_analysis_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")

    with pd.ExcelWriter(analysis_output_file) as analysis_writer, pd.ExcelWriter(intra_text_output_file) as intra_text_writer:
        for filename in codings_df['filename'].unique():
            file_data = codings_df[codings_df['filename'] == filename]
            excerpts = {
                filename: {
                    row['paragraph_index']: row['original_text']
                    for _, row in file_data.iterrows()
                }
            }

            analysis_results = analyzer_client.analyze_themes(json.dumps(excerpts), themes)
            themes = analysis_results['themes_to_codes']

            # Write thematic analysis results to Excel
            analysis_df = pd.DataFrame([
                {
                    'filename': filename,
                    'paragraph_index': paragraph_index,
                    'themes': ', '.join(analysis_data['themes']),
                    'quote': analysis_data['quote'],
                    'justification': analysis_data['justification']
                }
                for paragraph_index, analysis_data in analysis_results['analysis'].items()
            ])
            analysis_df.to_excel(analysis_writer, sheet_name='analysis', index=False)

            # Perform intra-text analysis
            intra_text_results = intra_text_analyzer.analyze_intra_text(json.dumps(analysis_results))

            # Write intra-text analysis results to Excel
            for analysis_type in ['intersections', 'contradictions', 'connections']:
                intra_text_df = pd.DataFrame(intra_text_results[analysis_type])
                if not intra_text_df.empty:
                    intra_text_df.insert(0, 'filename', filename)
                    intra_text_df.insert(1, 'analysis_type', analysis_type)
                    intra_text_df.to_excel(intra_text_writer, sheet_name='intra_text', index=False, header=False, startrow=intra_text_writer.sheets['intra_text'].max_row)

    return intra_text_output_file


def perform_cross_document_analysis(intra_text_output_file, cross_document_analyzer):
    """
    Performs cross-document analysis on intra-text results and prints the summary.
    """
    intra_text_df = pd.read_excel(intra_text_output_file, sheet_name='intra_text')
    cross_document_summary = cross_document_analyzer.analyze_cross_document(intra_text_df.to_json(orient='records'))

    print("\nCross-Document Analysis Summary:\n")
    print(cross_document_summary)


def perform_intra_text_analysis(analysis_results, intra_text_analyzer):
    """
    Performs intra-text analysis on the provided analysis results and writes 
    the output to an Excel file.
    """
    intra_text_output_file = os.path.join(OUTPUT_DIR, f"intra_text_analysis_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")

    with pd.ExcelWriter(intra_text_output_file) as intra_text_writer:
        for filename, analysis_data in analysis_results.items():
            intra_text_results = intra_text_analyzer.analyze_intra_text(json.dumps(analysis_data))

            # Write intra-text analysis results to Excel
            for analysis_type in ['intersections', 'contradictions', 'connections']:
                intra_text_df = pd.DataFrame(intra_text_results[analysis_type])
                if not intra_text_df.empty:
                    intra_text_df.insert(0, 'filename', filename)
                    intra_text_df.insert(1, 'analysis_type', analysis_type)
                    intra_text_df.to_excel(intra_text_writer, sheet_name='intra_text', index=False, header=False, startrow=intra_text_writer.sheets['intra_text'].max_row)

    return intra_text_output_file


def load_analysis_results_from_file(filepath):
    """
    Loads thematic analysis results from a JSON file.
    """
    try:
        with open(filepath, 'r') as f:
            analysis_results = json.load(f)
        return analysis_results
    except FileNotFoundError:
        print(f"Error: File not found - {filepath}")
        return {}


def load_themes_from_file(filepath):
    """
    Loads themes from a JSON file.
    """
    try:
        with open(filepath, 'r') as f:
            themes = json.load(f)
        return themes
    except FileNotFoundError:
        print(f"Error: File not found - {filepath}")
        return {}


def load_codes_from_file(filepath):
    """
    Loads codes and their descriptions and constructs from a JSON file.
    """
    try:
        with open(filepath, 'r') as f:
            code_data = json.load(f)

        return code_data

    except FileNotFoundError:
        print(f"Error: File not found - {filepath}")
        return {}
    

def visualize_theme_overview(themes_hierarchy, filename="theme_overview.png"):
    """
    Visualizes the overview of meta-themes, themes, and sub-themes and saves it as an image.

    Args:
        themes_hierarchy: The hierarchical theme structure.
        filename: The name of the file to save the visualization.
    """

    graph = nx.DiGraph()
    node_labels = {}
    node_colors = {}
    color_map = plt.get_cmap("tab20")

    def add_nodes_and_edges(hierarchy, level=0, parent=None):
        """Recursively adds nodes and edges to the graph."""
        for name, data in hierarchy.items():
            node_id = str(name)
            graph.add_node(node_id)
            node_labels[node_id] = name
            node_colors[node_id] = color_map(level)
            if parent:
                graph.add_edge(parent, node_id)
            if "themes" in data:
                add_nodes_and_edges(data["themes"], level + 1, node_id)
            if "sub-themes" in data:
                add_nodes_and_edges(data["sub-themes"], level + 1, node_id)

    add_nodes_and_edges(themes_hierarchy)  # Only add meta-themes, themes, and sub-themes

    plt.figure(figsize=(12, 12))
    pos = nx.spring_layout(graph, k=0.3)
    nx.draw(graph,
            pos,
            labels=node_labels,
            with_labels=True,
            node_size=3000,
            node_color=[node_colors[node] for node in graph.nodes()],
            font_size=10,
            font_weight="bold",
            arrowsize=20)
    plt.savefig(filename)
    plt.show()

def visualize_family_members_subgraph(themes_hierarchy,
                                     filename="family_members_subgraph.png"):
    """
    Visualizes the subgraph of "Family members" with its sub-themes and codes, 
    skipping the meta-theme level and other theme-level nodes.
    """

    graph = nx.DiGraph()
    node_labels = {}
    node_colors = {}
    color_map = plt.get_cmap("tab20")

    def add_nodes_and_edges(hierarchy, level=0, parent=None):
        """Recursively adds nodes and edges to the graph, skipping unwanted levels."""
        for name, data in hierarchy.items():
            if level == 0:  # Skip meta-themes (level 0)
                add_nodes_and_edges(data["themes"], level + 1, None)
            elif level == 1 and name != "Family members":  # Skip other theme-level nodes
                continue
            else:
                node_id = str(name)
                graph.add_node(node_id)
                node_labels[node_id] = name
                node_colors[node_id] = color_map(level)
                if parent:
                    graph.add_edge(parent, node_id)
                if "sub-themes" in data:
                    add_nodes_and_edges(data["sub-themes"], level + 1,
                                        node_id)
                if "codes" in data:
                    for code in data["codes"]:
                        code_id = str(code) 
                        
                        # Remove "Family members-" prefix from code name
                        if code_id.startswith("Family members-"):
                            code_id = code_id.replace("Family members-", "") 

                        graph.add_node(code_id)
                        node_labels[code_id] = code
                        node_colors[code_id] = color_map(level + 1)
                        graph.add_edge(node_id, code_id)

    add_nodes_and_edges(themes_hierarchy)

    # Find the "Family members" node
    family_members_node = None
    for node in graph.nodes:
        if node == "Family members":
            family_members_node = node
            break

    if family_members_node:
        # Create subgraph of "Family members"
        family_members_nodes = [
            node for node in nx.descendants(graph, family_members_node)
        ] + [family_members_node]
        subgraph = graph.subgraph(family_members_nodes)

        plt.figure(figsize=(12, 12))
        pos = nx.spring_layout(subgraph, k=0.3)
        nx.draw(subgraph,
                pos,
                labels=node_labels,
                with_labels=True,
                node_size=3000,
                node_color=[node_colors[node] for node in subgraph.nodes()],
                font_size=10,
                font_weight="bold",
                arrowsize=20)
        plt.savefig(filename)
        plt.show()
    else:
        print("Error: 'Family members' node not found in the hierarchy.")

import json
import matplotlib.pyplot as plt
import networkx as nx

def visualize_network(data, filename="network_visualization.png"):
    """
    Visualizes a network graph from a JSON object containing nodes and edges.

    Args:
        data: A JSON object with "nodes" and "edges" lists.
        filename: The name of the file to save the visualization.
    """

    graph = nx.DiGraph()
    node_labels = {}
    node_colors = {}
    edge_labels = {}
    color_map = plt.get_cmap("tab20")

    # Add nodes to the graph
    for i, node in enumerate(data["nodes"]):
        node_id = node["id"]
        graph.add_node(node_id)
        node_labels[node_id] = node["label"]
        node_colors[node_id] = color_map(i)  # Assign color based on index

    # Add edges to the graph
    for edge in data["edges"]:
        source = edge["source"]
        target = edge["target"]
        relation = edge["relation"]
        graph.add_edge(source, target)
        edge_labels[(source, target)] = relation

    # Set figure size and layout
    plt.figure(figsize=(12, 12))
    pos = nx.spring_layout(graph, k=0.3)

    # Draw nodes with labels and colors
    nx.draw(graph,
            pos,
            labels=node_labels,
            with_labels=True,
            node_size=3000,
            node_color=[node_colors[node] for node in graph.nodes()],
            font_size=10,
            font_weight="bold",
            arrowsize=20)

    # Draw edge labels
    nx.draw_networkx_edge_labels(graph, pos, edge_labels=edge_labels, font_size=8)

    plt.savefig(filename)
    plt.show()