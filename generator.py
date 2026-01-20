from google import genai
import os
from dotenv import load_dotenv
import sys
import argparse
import json
import markdown2
import subprocess

def export_as_pptx(json_data, target_directory="generated_pptx"):
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)

    for report in json_data:
        report_fn = report['file_name'].replace('.json', '.pptx')
        output_path = os.path.join(target_directory, report_fn)
        
        # Clean each string to ensure '#' is at the very start
        cleaned_content = [item.strip() for item in report['Report Content']]
        
        # Use a very large gap to force section breaks
        full_markdown = "\n\n\n\n".join(cleaned_content)
        
        temp_md_path = f"temp_{report['file_name']}.md"
        with open(temp_md_path, "w", encoding="utf-8") as f:
            f.write(full_markdown)

        try:
            subprocess.run([
                "pandoc", 
                temp_md_path, 
                "--standalone",
                "--slide-level=1", 
                "-o", output_path
            ], check=True)
            print(f"Exported: {output_path}")
        finally:
            if os.path.exists(temp_md_path):
                os.remove(temp_md_path)
                
def main():
    # Load and import API Keys
    load_dotenv()
    GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')

    # Parse input arguments
    parser = argparse.ArgumentParser(description="Generates a JSON of internal audit risk reports in Markdown format.")
    parser.add_argument("-r","--reports", type=int, help="number of reports to generate.")
    parser.add_argument("-v", '--verbose', action='store_true')
    parser.add_argument("--seed",
        type=str,
        help="Optional: Provide a specific industry or focus (e.g., 'Healthcare', 'Fintech')",
        default="General Corporate Operations"
    )
    args = parser.parse_args()

    # Verbose logging of command line input arguments
    if args.verbose:
        print(f"[DEBUG] Received reports={args.reports}")

    # Get response from Gemini API
    google_client = genai.Client(api_key=GEMINI_API_KEY)

    # Generator Prompt
    generator_prompt = f"""
    Act as a Senior Internal Audit Manager with 15+ years of experience in corporate governance, risk management, and compliance (GRC). Generate {args.reports} distinct Internal Audit Reports for different corporate business units.

    **Report Requirements:**
    One report must be an **Assurance Review** (stringent focus on risks/controls of specific teams) and the other must be an **Advisory Engagement** (focus on process improvement and consultative insights).

    Each report must be structured as a sequence of slides. Each of the following sections must correspond to its own slide (one slide per item):
    1. **Title Slide:** Title of the report, Business Unit name, and Date.
    2. **Executive Summary:** Objectives, Background, and Scope.
    3. **Details (Observation Slides):** Provide one slide for **each** observation/issue. Each slide must contain:
        * The issue description.
        * The corresponding Risk.
        * The Risk Rating (choose from: **ADEQUATE**, **FOR IMPROVEMENT**, or **INADEQUATE**).
        * A specific Recommendation.
        * The overall Status.
    4. **Recommendations Summary:** A consolidated list of recommendations for the business unit.
    5. **Management Action Plan:** Specific action items, owners, and deadlines.

    **Output Format:**
    Deliver the final response as a JSON file. 
    * Use a key titled "file_name" for the report's filename.
    * Use a key titled "Report Content" which houses an array of strings. 
    * **Crucial:** Each element in the array must represent exactly one slide. 
    * Use proper Markdown formatting (bold, headers, lists) within the strings.

    **Example structure for the JSON:**
    {{"file_name": "Report_Name.json", "Report Content": ["# Slide 1 Content", "## Slide 2 Content"]}}

    Please provide only the JSON output.
"""

    # response = google_client.models.generate_content(
    #     model='gemini-3-flash-preview', 
    #     contents=generator_prompt,
    #     config={
    #         'response_mime_type': 'application/json',
    #     }
    # )

    # # Load JSON
    # data = json.loads(response.text)
    # # print(data)
    # with open("output.json", "w", encoding="utf-8") as json_file:
    #     json.dump(data, json_file, ensure_ascii=False, indent=4)
    # # Convert data to PDFs
    # DEBUG: Load dummy response from JSON file
    with open('output.json', 'r') as file:
        data = json.load(file)
    export_as_pptx(data)
    
if __name__ == '__main__':
    main()
