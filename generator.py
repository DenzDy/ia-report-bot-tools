from google import genai
import os
from dotenv import load_dotenv
import sys
import argparse
import json
import markdown2
from fpdf import FPDF
from weasyprint import HTML

def export_as_pdf(json_data, target_directory="generated_reports"):
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)

    # Basic CSS to ensure the table has borders and proper padding
    style = """
    <style>
        body { font-family: 'DejaVu Sans', sans-serif; font-size: 12px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid black; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        h2 { color: #800000; }
    </style>
    """

    for report in json_data:
        # Clean the markdown: Replace double pipes with newlines
        md_content = report['Report Content']
        fixed_md = md_content.replace("||", "|\n|")

        # Convert to HTML with 'tables' extra enabled
        html_content = markdown2.markdown(fixed_md, extras=["tables"])

        # Combine CSS and HTML
        full_html = f"{style}{html_content}"

        # Generate the PDF using WeasyPrint (much better for tables)
        report_fn = report['file_name']
        output_path = os.path.join(target_directory, report_fn)
        
        HTML(string=full_html).write_pdf(output_path)
        print(f"Exported: {output_path}")


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

    response = google_client.models.generate_content(
        model='gemini-3-flash-preview', 
        contents=generator_prompt,
        config={
            'response_mime_type': 'application/json',
        }
    )

    # Load JSON
    data = json.loads(response.text)
    # print(data)
    with open("output.json", "w", encoding="utf-8") as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)
    # Convert data to PDFs
    # DEBUG: Load dummy response from JSON file
    # with open('output.json', 'r') as file:
    #     data = json.load(file)
    # export_as_pdf(data)
    
if __name__ == '__main__':
    main()
