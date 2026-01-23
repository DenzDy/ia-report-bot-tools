from google import genai
import os
from dotenv import load_dotenv
import sys
import argparse
import json
import markdown2
import subprocess

def export_as_pptx(json_data, target_directory="generated_pptx", char_limit=800):
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)

    for report in json_data:
        report_fn = report['file_name'].replace('.json', '.pptx')
        output_path = os.path.join(target_directory, report_fn)
        
        final_slides = []
        for item in report['Report Content']:
            parts = item.strip().split('\n', 1)
            header = parts[0]
            content = parts[1] if len(parts) > 1 else ""
            
            # Split if content is excessively long
            if len(content) > char_limit:
                split_point = content.rfind(' ', 0, char_limit)
                final_slides.append(f"{header} (1/2)\n\n{content[:split_point]}")
                final_slides.append(f"{header} (2/2)\n\n{content[split_point:].strip()}")
            else:
                final_slides.append(item.strip())
        
        full_markdown = "\n\n---\n\n".join(final_slides)
        temp_md = "temp_report.md"
        
        with open(temp_md, "w", encoding="utf-8") as f:
            f.write(full_markdown)

        try:
            subprocess.run([
                "pandoc", 
                temp_md, 
                "--reference-doc=template.pptx", # Path to the file you just made
                "--slide-level=1", 
                "-o", output_path
            ], check=True)
            print(f"Success! Created {output_path}")
        finally:
            if os.path.exists(temp_md):
                os.remove(temp_md)

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
    parser.add_argument("--type",
        type=str,
        help="Optional: Provide a specific report type",
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

    **Report Requirements (IMPORTANT):**
    1. Focus on {args.seed}

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
    {{"file_name": "Report_Name.json", "Report Content": ["# Slide 1 Content", "# Slide 2 Content"]}}

    Please provide only the JSON output.
    
    Shown below is a toy example of a report's content:
    
    Title:  AC Analytics lacks project governance 

    ES: WE have noted 10 observations , 5 are high…. These are the common themes…. Scope… action plans… finding… per review conclusion (A, NI, INAQ) 

    Details: 

    Observation 1: ... + Risk rating (L/M/H) + Recommendation, Risk, Action plan, Overall status, 

    O2: ...

    O3: ... 

    O4: ... 

    O5: ...
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
    # # DEBUG: Load dummy response from JSON file
    with open('output.json', 'r') as file:
        data = json.load(file)
    export_as_pptx(data)
    
if __name__ == '__main__':
    main()
