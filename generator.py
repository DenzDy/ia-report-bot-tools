from openai import OpenAI
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
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')

    # Parse input arguments
    parser = argparse.ArgumentParser(description="Generates a JSON of internal audit risk reports in Markdown format.")
    parser.add_argument("-r","--reports", required=True, type=int, help="number of reports to generate.")
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
    openai_client = OpenAI(
        api_key=OPENAI_API_KEY
    )

    # System Prompt
    system_prompt = f"""
    ### ROLE
    Senior Internal Audit Manager (15+ years experience). Tone: Professional and authoritative.

    ### TASK
    Generate a structured Internal Audit Report JSON for PowerPoint.

    ### STRUCTURAL RULES
    - Return a JSON object: {{"file_name": "...", "Report Content": ["...", "..."]}}
    - Each array element = ONE slide.
    - Every string MUST start with "# " for the title.
    - Use Markdown tables for "Management Action Plan" slides.
    - Keep slide text under 600 characters.
    
    ### STRUCTURAL REQUIREMENTS (CRITICAL)
    - Return ONLY a JSON object: {{"file_name": "...", "Report Content": ["...", "..."]}}
    - Each array element in "Report Content" represents EXACTLY one PowerPoint slide.
    - EVERY slide string MUST start with "# " as the first character.
    - **Spacing Rule:** Use exactly two newline characters (\n\n) to separate every header, subheading, and paragraph. This prevents text clumping on the slide.

    ### SLIDE SEQUENCE
    1. # [Title Slide]: Include Business Unit and Date.
    2. # Executive Summary: Include ### Objectives, ### Background, and ### Scope.
    3. # Observation [N]: [Title]: Detail the finding. Use bold labels (**Issue:**, **Risk:**, **Risk Rating:**, **Recommendation:**) each on their own line separated by \n\n.
    4. # Recommendations Summary: Consolidated bulleted list.
    5. # Management Action Plan: Use a Markdown table: | Action Item | Owner | Deadline |.

    ### CONSTRAINTS
    - Minimum of 2 detailed observations per report.
    - Maximum length: Keep body text under 1000 characters per slide.
    - Risk Ratings MUST be: ADEQUATE, FOR IMPROVEMENT, or INADEQUATE.

    ### TITLE SLIDE RULES
    The first element of the array MUST follow this exact format:
    "# [Audit Report Title]
    ## [Business Unit Name]
    ### [Date]"

    ### EXAMPLES OF IDEAL OUTPUT
    {{
        "file_name": "IA_Report_HR_Payroll_Processing.json",
        "Report Content": [
            "# Internal Audit of Payroll Processing & Employee Benefits\n\n**Business Unit:** Human Resources (Global Operations)\n**Date:** January 25, 2026",
            "# Executive Summary\n\n### Objectives\nTo evaluate the accuracy of payroll disbursements and compliance with tax regulations.\n\n### Background\nThe HR unit manages payroll for 5,000 employees across three jurisdictions.\n\n### Scope\nReview of payroll cycles from Q3 2025 to Q4 2025, including manual adjustments and bonus calculations.",
            "# Observation 1: Lack of Segregation of Duties\n**Issue:** The same individual responsible for updating employee master data also executes the final payroll run.\n**Risk:** Potential for unauthorized salary adjustments or creation of 'ghost employees'.\n**Risk Rating:** **INADEQUATE**\n**Recommendation:** Segregate master data entry from payroll execution; implement a secondary reviewer for all payroll batches.,
            "# Observation 2: Delayed Deactivation of Terminated Employees\n**Issue:** Access to corporate systems remained active for 48 hours post-termination for 15% of sampled cases.\n**Risk:** Unauthorized data access or intellectual property theft.\n**Risk Rating:** **FOR IMPROVEMENT**\n**Recommendation:** Automate the link between HR termination logs and IT access management systems.,
            "# Recommendations Summary\n1. Enforce strict Segregation of Duties (SoD) in payroll software.\n2. Implement automated 'Leaver' protocols for system access.\n3. Conduct quarterly payroll audits against physical headcount records.",
            "# Management Action Plan\n* **Action Item:** Configure ERP permissions for SoD.\n  * **Owner:** HR Director / IT Manager\n  * **Deadline:** March 31, 2026\n* **Action Item:** Establish automated termination alerts.\n  * **Owner:** HR Operations Lead\n  * **Deadline:** February 28, 2026"
        ]
    }},
    {{
        "file_name": "IA_Report_Procurement_Vendor_Mgmt.json",
        "Report Content": [
        "# Strategic Sourcing and Vendor Risk Management Audit\n\n**Business Unit:** Corporate Procurement\n**Date:** January 20, 2026",
        "# Executive Summary\n\n### Objectives\nTo assess the vendor selection process and contract lifecycle management.\n\n### Background\nProcurement managed $50M in spend across 200+ vendors in the last fiscal year.\n\n### Scope\nFocus on vendors with annual spend exceeding $500k and the competitive bidding process.",
        "# Observation 1: Inconsistent Competitive Bidding\n**Issue:** 3 out of 10 large contracts were awarded without the required three-quote minimum without documented justification.\n**Risk:** Failure to achieve best value for money and potential vendor favoritism.\n**Risk Rating:** **FOR IMPROVEMENT**\n**Recommendation:** Mandate a 'Bid Exception Form' signed by the CFO for any non-competitive awards.,
        "# Observation 2: Missing Vendor Performance Evaluations\n**Issue:** Annual performance reviews for 'Tier 1' vendors were not conducted in 2025.\n**Risk:** Service level degradation and missed opportunities for contract renegotiation.\n**Risk Rating:** **INADEQUATE**\n**Recommendation:** Implement a standardized vendor scorecard and schedule quarterly review meetings.,
        "# Recommendations Summary\n1. Standardize the competitive bidding workflow.\n2. Launch a Vendor Performance Management (VPM) framework.\n3. Audit vendor insurance certificates for expiration.",
        "# Management Action Plan\n* **Action Item:** Update Procurement Policy to include Bid Exception requirements.\n  * **Owner:** Head of Procurement\n  * **Deadline:** April 15, 2026\n* **Action Item:** Conduct catch-up reviews for top 10 vendors.\n  * **Owner:** Procurement Category Manager\n  * **Deadline:** May 30, 2026"
        ]
    }}

    ### INPUT SEED
    Generate a new report based on the user's provided seed.
    """

    # User Prompt
    user_prompt = f"""
    Generate {args.reports} reports with focus on {args.seed} for the reports.
    """

    response = openai_client.responses.create(
        model='gpt-5-mini', 
        instructions=system_prompt,
        input=user_prompt
    )

    # Load JSON
    data = json.loads(response.text)
    # print(data)
    with open("output.json", "w", encoding="utf-8") as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)
    # Convert data to PDFs
    # DEBUG: Load dummy response from JSON file
    with open('output.json', 'r') as file:
        data = json.load(file)
    export_as_pptx(data)
    
if __name__ == '__main__':
    main()
