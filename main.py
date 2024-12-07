import os
import datetime
import argparse
import json

from config import *
from code_generation import CodeGenerationClient
from theme_generator import ThemeGeneratorClient
from code_application import ThematicCodingClient
#from theme_generator import ThematicAnalyzerClient
from within_case_analysis import IntraTextAnalyzerClient
from report_generation import CrossDocumentAnalyzerClient
from utils import (extract_paragraphs_from_docx, 
                   write_coding_results_to_excel, 
                   generate_initial_codes,
                   perform_coding, 
                   perform_analysis_and_reporting, 
                   perform_intra_text_analysis, 
                   perform_cross_document_analysis, 
                   load_analysis_results_from_file, 
                   load_themes_from_file, 
                   load_codes_from_file,
                   visualize_family_members_subgraph,
                   visualize_theme_overview,
                   visualize_network) 


def process_docx_files(directory, batch_size, client_flag):
    """
    Processes docx files based on the specified client flag.
    """

    output_file = os.path.join(OUTPUT_DIR, f"analyzed_results_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")
    
    if client_flag == "code_generator":
        code_generator = CodeGenerationClient()
        themes = load_themes_from_file("themes.json") 
        all_codes, code_descriptions, code_constructs, all_files_excerpt_codings, new_codes_by_file = generate_initial_codes(
            directory, themes, code_generator, num_docs=NUM_DOCS_FOR_CODE_GENERATION) 
    
        
        output_file = os.path.join(
            OUTPUT_DIR,
            f"initial_code_generation_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        )
        write_coding_results_to_excel(all_files_excerpt_codings,
                                     new_codes_by_file, output_file)
        print(f"\nGenerated Codes:\n{all_codes}")

    elif client_flag == "theme_generator":
        theme_generator = ThemeGeneratorClient()
        themes = load_themes_from_file("themes.json") 
        all_code_data = load_codes_from_file("codes.json")
        themes_hierarchy = theme_generator.generate_themes(all_code_data, themes)
        print(f"\nGenerated Themes:\n{themes_hierarchy}")
        visualize_theme_overview(themes_hierarchy)
        visualize_family_members_subgraph(themes_hierarchy)

    elif client_flag == "coding_client":
        coding_client = ThematicCodingClient()
        all_codes, code_descriptions, code_constructs = load_codes_from_file("codes.json")
        all_codes, code_descriptions, code_constructs, all_files_excerpt_codings, new_codes_by_file = perform_coding(directory, all_codes, code_descriptions, code_constructs, coding_client)
        write_coding_results_to_excel(all_files_excerpt_codings, new_codes_by_file, output_file)

    elif client_flag == "analyzer_client":
        analyzer_client = ThematicAnalyzerClient()
        themes = load_themes_from_file("themes.json") 
        perform_analysis_and_reporting(output_file, themes, analyzer_client, intra_text_analyzer)

    elif client_flag == "intra_text_analyzer":
        # intra_text_analyzer = IntraTextAnalyzerClient()
        # analysis_results = load_analysis_results_from_file("analysis_results.json") 
        # perform_intra_text_analysis(analysis_results, intra_text_analyzer)
        data129 = """
{
  "nodes": [
    {
      "id": "Social Influences",
      "label": "Social Influences",
      "description": "Influence from peers and teamwork experiences shaped engineering aspirations.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "Peer and Collaborative Learning",
      "label": "Peer and Collaborative Learning",
      "description": "Interactions with peers and collaborative projects significantly influenced engineering aspirations.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "Self Efficacy in creating or building things",
      "label": "Self-Efficacy in Creating/Building",
      "description": "Student expressed high self-efficacy in creating and building, evident in design projects and problem-solving.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "Design and Construction Skills",
      "label": "Design & Construction Skills",
      "description": "Student's self-assessment of design and construction skills were highly influential.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "Career Outlook",
      "label": "Career Outlook",
      "description": "Student explicitly stated a desire to pursue an engineering career.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "Intrinsic Motivation",
      "label": "Intrinsic Motivation",
      "description": "Personal interests and aptitudes drove the pursuit of engineering.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "Help society",
      "label": "Help Society",
      "description": "Desire to positively impact society through engineering, emphasizing community engagement and empathetic design.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "Community Engagement and Development",
      "label": "Community Engagement",
      "description": "Desire to contribute to community development and social well-being through engineering.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "Subject Affinity & Integration",
      "label": "Subject Affinity",
      "description": "Explicit interest in applying skills (communication, problem-solving) to engineering.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "Problem-Solving Orientation",
      "label": "Problem-Solving Orientation",
      "description": "Student's focus on problem-solving and its application to real-world challenges.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "Prior Experiences",
      "label": "Prior Experiences",
      "description": "Participation in formal and informal STEM activities significantly influenced interest in engineering.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "Informal Learning and Exploration",
      "label": "Informal Learning",
      "description": "Informal learning and self-directed exploration contributed to engineering interests.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    }

  ],
  "edges": [
    {"source": "Social Influences", "target": "Peer and Collaborative Learning", "relation": "contains"},
    {"source": "Self Efficacy in creating or building things", "target": "Design and Construction Skills", "relation": "contains"},
    {"source": "Career Outlook", "target": "Intrinsic Motivation", "relation": "contains"},
    {"source": "Career Outlook", "target": "Help society", "relation": "contains"},
    {"source": "Help society", "target": "Community Engagement and Development", "relation": "contains"},
    {"source": "Self Efficacy in creating or building things", "target": "Problem-Solving Orientation", "relation": "related_to"},
    {"source": "Peer and Collaborative Learning", "target": "Subject Affinity & Integration", "relation": "influences"},
    {"source": "Prior Experiences", "target": "Informal Learning and Exploration", "relation": "contains"}

  ]
}
"""
        data120 = """
{
  "nodes": [
    {
      "id": "1",
      "label": "Self-Efficacy in creating or building things",
      "description": "Student expressed enjoyment, confidence, or excellence in creating or building physical objects. Focus is on expressed self-efficacy illustrated by specific examples of engineering-related building projects.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "2",
      "label": "Design and Construction Skills",
      "description": "Students' self-assessment of their design and construction skills.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "3",
      "label": "Problem-Solving and Innovation",
      "description": "Students' confidence in their ability to solve problems creatively and innovatively.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "4",
      "label": "Subject Affinity & Integration",
      "description": "Student expressed explicit interest in the application of subjects or skills such as math, science, biology, chemistry, computers, designing, or problem-solving. They relate to the application being a motivating and or lucrative part of the engineering profession.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "5",
      "label": "SpecificInterest-Software",
      "description": "Student's specific interest in software design and development.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "6",
      "label": "Intrinsic Motivation",
      "description": "Personal interests and aptitudes drive the pursuit of engineering.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "7",
      "label": "Prior Experiences",
      "description": "Student mentioned joining engineering clubs, working on projects, or having other influential formal or informal learning experiences (i.e., in clubs, programmes, or non-organized activities) in elementary, high school, or out-of-school contexts that supported their interest in engineering-related work or tasks.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "8",
      "label": "Social Influences",
      "description": "The student described the influence to become an engineer came from more than themselves and people outside of the family (such as friends, peers, working in teams, or socio-geographic environments).",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    }
  ],
  "edges": [
    {
      "source": "1",
      "target": "2",
      "relation": "contains"
    },
    {
      "source": "1",
      "target": "3",
      "relation": "contains"
    },
    {
      "source": "4",
      "target": "5",
      "relation": "contains"
    },
    {
      "source": "4",
      "target": "6",
      "relation": "motivates"
    },
    {
      "source": "1",
      "target": "7",
      "relation": "influenced by"
    },
    {
      "source": "8",
      "target": "7",
      "relation": "influenced by"
    }
  ]
}
"""
        data139 = """
{
  "nodes": [
    {
      "id": "1",
      "label": "Early Interest in STEM",
      "description": "Long-standing interest in science and technology, evident from childhood.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "2",
      "label": "Design-focused Visual Arts",
      "description": "High school visual arts courses provided design experience.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "3",
      "label": "Interdisciplinary Projects & Summer Camps",
      "description": "Participation in collaborative design projects strengthened design skills and teamwork abilities.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "4",
      "label": "U of T DEEP Program",
      "description": "Exposure to engineering through the University of Toronto's DEEP program.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "5",
      "label": "Strong Design Self-Efficacy",
      "description": "Confidence and enjoyment in design and problem-solving, demonstrated through projects and experiences.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "6",
      "label": "Aspirational Engineering Career",
      "description": "Clear intention to pursue an engineering career.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "7",
      "label": "Desire for Societal Impact",
      "description": "Focus on creating positive change through engineering, particularly for people with disabilities.",
      "meta-theme": "Career Goals and Societal Impact"
    }
  ],
  "edges": [
    {
      "source": "1",
      "target": "6",
      "relation": "Motivating Factor"
    },
    {
      "source": "2",
      "target": "5",
      "relation": "Skill Development"
    },
    {
      "source": "2",
      "target": "6",
      "relation": "Motivating Factor"
    },
    {
      "source": "3",
      "target": "5",
      "relation": "Skill Development"
    },
    {
      "source": "3",
      "target": "6",
      "relation": "Motivating Factor"
    },
    {
      "source": "4",
      "target": "5",
      "relation": "Skill Development"
    },
    {
      "source": "4",
      "target": "6",
      "relation": "Motivating Factor"
    },
    {
      "source": "5",
      "target": "6",
      "relation": "Underlying Capability"
    },
    {
      "source": "6",
      "target": "7",
      "relation": "Career Goal"
    }
  ]
}
"""
        data119 = """
{
  "nodes": [
    {
      "id": "1",
      "label": "Developmental Influences on Engineering Aspirations",
      "description": "Factors shaping student's interest in engineering.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "2",
      "label": "Social Influences",
      "description": "Influence from peers, mentors, and collaborative projects.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "3",
      "label": "Prior experiences",
      "description": "Formal and informal learning experiences influencing engineering interest.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "4",
      "label": "Self-Efficacy and Subject Matter",
      "description": "Student's self-perceived abilities and subject interests.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "5",
      "label": "Self Efficacy in creating or building things",
      "description": "Confidence and enjoyment in creating physical objects.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "6",
      "label": "Subject Affinity & Integration",
      "description": "Explicit interest in applying STEM subjects to engineering.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
    {
      "id": "7",
      "label": "Career Goals and Societal Impact",
      "description": "Career aspirations and desire to contribute to society.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "8",
      "label": "Career Outlook",
      "description": "Explicit desire to pursue engineering as a career.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "9",
      "label": "Nature of field",
      "description": "Positive perceptions of the engineering field.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "10",
      "label": "Intrinsic Motivation",
      "description": "Personal interests driving the pursuit of engineering.",
      "meta-theme": "Career Goals and Societal Impact"
    }

  ],
  "edges": [
    {
      "source": "1",
      "target": "2",
      "relation": "contains"
    },
    {
      "source": "1",
      "target": "3",
      "relation": "contains"
    },
    {
      "source": "1",
      "target": "8",
      "relation": "influences"
    },
    {
      "source": "4",
      "target": "5",
      "relation": "contains"
    },
    {
      "source": "4",
      "target": "6",
      "relation": "contains"
    },
    {
      "source": "7",
      "target": "8",
      "relation": "contains"
    },
    {
      "source": "7",
      "target": "9",
      "relation": "contains"
    },
    {
      "source": "8",
      "target": "10",
      "relation": "contains"
    },
    {
      "source": "3",
      "target": "5",
      "relation": "influences"
    },
    {
      "source": "2",
      "target": "5",
      "relation": "influences"
    },
    {
      "source": "6",
      "target": "5",
      "relation": "supports"
    }
  ]
}
""" 
        data122 = """
{
  "nodes": [
    {
      "id": "1",
      "label": "Prior Experiences",
      "description": "Participation in Air Cadets, including glider flying and establishing ground school.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "2",
      "label": "Leadership & Mentorship Roles",
      "description": "Leading ground school and piloting familiarization flights in Air Cadets.",
      "meta-theme": "Developmental Influences on Engineering Aspirations"
    },
    {
      "id": "3",
      "label": "Nature of Field",
      "description": "Positive perceptions of engineering's problem-solving, creativity, positive impact, and efficiency aspects.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "4",
      "label": "Help Society",
      "description": "Desire to enable others to have similar or better experiences; focus on positive societal impact and environmental preservation in design.",
      "meta-theme": "Career Goals and Societal Impact"
    },
    {
      "id": "5",
      "label": "Self-Confidence Development",
      "description": "Willingness to try new things and accept failure; confidence in problem-solving and design.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    },
        {
      "id": "6",
      "label": "Problem-Solving and Innovation",
      "description": "Seeking creative and new solutions; holistic analysis of options to minimize negative impacts.",
      "meta-theme": "Self-Efficacy and Subject Matter"
    }

  ],
  "edges": [
    {
      "source": "1",
      "target": "2",
      "relation": "Includes"
    },
    {
      "source": "1",
      "target": "3",
      "relation": "Influenced by"
    },
    {
      "source": "1",
      "target": "4",
      "relation": "Motivates"
    },
    {
      "source": "2",
      "target": "4",
      "relation": "Motivates"
    },
    {
      "source": "3",
      "target": "5",
      "relation": "Developed through"
    },
    {
      "source": "3",
      "target": "6",
      "relation": "Highlights"
    },
    {
      "source": "4",
      "target": "3",
      "relation": "Informs"
    },
    {
      "source": "5",
      "target": "4",
      "relation": "Supports"
    },
    {
      "source": "6",
      "target": "4",
      "relation": "Supports"
    }
  ]
}"""
        data = json.loads(data122)
        visualize_network(data)

    elif client_flag == "cross_document_analyzer":
        cross_document_analyzer = CrossDocumentAnalyzerClient()
        intra_text_output_file = os.path.join(OUTPUT_DIR, "intra_text_analysis.xlsx")  
        perform_cross_document_analysis(intra_text_output_file, cross_document_analyzer)

    else:
        print("Invalid client flag.")


def main():
    parser = argparse.ArgumentParser(description="Process docx files using different clients.")
    parser.add_argument("--client", required=True, choices=[
        "code_generator", 
        "theme_generator", 
        "coding_client", 
        "analyzer_client", 
        "intra_text_analyzer", 
        "cross_document_analyzer"
    ], help="Specify the client to run.")
    args = parser.parse_args()

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    process_docx_files(INPUT_DIR, BATCH_SIZE, args.client)

if __name__ == "__main__":
    main()