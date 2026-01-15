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
    parser.add_argument("number_of_reports", type=int, help="number of reports to generate.")
    parser.add_argument("-v", '--verbose', action='store_true')
    parser.add_argument("--seed",
        type=str,
        help="Optional: Provide a specific industry or focus (e.g., 'Healthcare', 'Fintech')",
        default="General Corporate Operations"
    )
    args = parser.parse_args()

    # Verbose logging of command line input arguments
    if args.verbose:
        print(f"[DEBUG] Received number_of_reports={args.number_of_reports}")

    # Get response from Gemini API
    google_client = genai.Client(api_key=GEMINI_API_KEY)

    # Generator Prompt
    generator_prompt = f"""
    You are a Senior Internal Audit Manager with 15+ years of experience in corporate governance, risk management, and compliance (GRC).
    
    Generate 2 distinct Internal Audit Reports for different corporate business units.

    **Report Requirements**
    Each report should contain the following sections:
    1. Title
    2. Executive Summary
        - contains objectives, backgrounds, and scope
    3. Details
        - contains a list of observations/issues about the business unit.
        - each observation has a corresponding risk and risk rating defined at a later section
        - each observation will have a recommendation and overall status
    4. Recommendations
        - contains some recommendations for the business unit regarding the observation details
    5. Management Action Plan
        - contains an action plan for the business unit management

    **Risk Ratings**
    The risk ratings are shown below as follows:
        1. ADEQUATE
        2. FOR IMPROVEMENT
        3. INADEQUATE

    These risks are ranked from best to worst, or least concerning to most concerning.

    **Report Types**
    There are two report types, and each generated report will only be able to pick one of the two:
    1. Assurance Review
        - a more stringent report that focuses on the risks of specific teams under the business unit  (e.g. Finance, Accounting, etc.) 

    2. Advisory Engagement
        - focuses on understanding processes and providing recommendations to improve on processes.

    **Expected Output**:
    The expected output should be a JSON file which houses all the report content data, with a specific key titled "Report Content" which houses the whole report content in markdown format.
    Each JSON object should also have a file name key, which defines a file name for said report.
    Only give the JSON file as the output, and use JSON formatting. And use proper markdown formatting for bold, italicized, and other formattings.
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
    # Convert data to PDFs
    # DEBUG: Load dummy response from JSON file
    with open('output.json', 'r') as file:
        data = json.load(file)
    export_as_pdf(data)
    
if __name__ == '__main__':
    main()
