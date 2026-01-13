from google import genai
import os
from dotenv import load_dotenv
import sys
import argparse
import json
import markdown2
from fpdf import FPDF

def export_as_pdf(json_data, target_directory="generated_reports"):
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)

    font_path = "/usr/share/fonts/truetype/dejavu/"

    for report in json_data:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()

        # Load Unicode font
        pdf.add_font("DejaVu", style="", fname="./fonts/DejaVuSans.ttf")
        pdf.add_font("DejaVu", style="B", fname="./fonts/DejaVuSans-Bold.ttf")
        pdf.set_font("DejaVu", size=10)

        # 1. Get the content directly from the JSON
        content = report["Report Content"]

        # 2. Add extra spacing between sections to prevent "clutter"
        # This replaces standard newlines with HTML breaks that FPDF understands
        formatted_content = content.replace("\n", "<br>")

        try:
            # 3. Write directly to PDF using the 1.5x line spacing (8)
            pdf.write_html(formatted_content, 8) 
            
            output_path = os.path.join(target_directory, report["Document File Name"])
            pdf.output(output_path)
            print(f"✅ Clean PDF generated: {output_path}")
        except Exception as e:
            print(f"❌ Error: {e}")


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
**Role:** You are a Senior Internal Audit Manager with 15+ years of experience in corporate governance, risk management, and compliance (GRC).

**Task:** Generate {args.number_of_reports} distinct Internal Audit Reports for different corporate business units.

**Seed Instruction:** Focus specifically on {args.seed}. 

**Report Format Instructions (Strict PDF Rendering):**
Use ONLY these HTML tags for formatting: <h1>, <h3>, <b>, <center>, and <br>. 
DO NOT use Markdown symbols like #, *, -, or >.

For each of the {args.number_of_reports} reports, follow this exact structure:

1. <center><h1>[Centered Title of the Report]</h1></center><br>

2. <h3>Executive Summary</h3>
   [Provide a high-level overview of the audit scope and overall opinion.]<br><br>

3. <h3>Risk Assessment & In-Depth Analysis</h3>
   For each risk, use this exact vertical structure to ensure the analysis starts on its own line:
   <b>Risk [Number]: [Name] — Rating: [Value]</b>
   <br><b>In-Depth Analysis:</b> [Provide 3-4 sentences. The <br> tag above ensures this starts on a new line. For 'Severe' or 'Very severe' risks, include USD quantification.]<br><br>

4. <h3>Audit Findings</h3>
   <b>Finding 1: [Title]</b><br>[Detailed observation.]<br><br>
   <b>Finding 2: [Title]</b><br>[Detailed observation.]<br><br>

5. <h3>Recommendations</h3>
   <b>Recommendation 1: [Title]</b><br>[Remediation instructions.]<br><br>
   <b>Recommendation 2: [Title]</b><br>[Remediation instructions.]<br><br>

6. <h3>Management Response</h3>
   [Simulate "audit friction" in 2 out of the {args.number_of_reports} reports.]<br><br>

7. <h3>Conclusion</h3>
   [Final summary statement.]

**Output Format:**
Return ONLY a valid JSON array containing exactly {args.number_of_reports} objects.
* `Document File Name`: String (e.g., "IA_Report_Name_2026.pdf").
* `Report Content`: Full report string using the HTML tags above."""

    response = google_client.models.generate_content(
        model='gemini-3-flash-preview', 
        contents=generator_prompt,
        config={
            'response_mime_type': 'application/json',
        }
    )

    # Load JSON
    data = json.loads(response.text)

    # Convert data to PDFs
    export_as_pdf(data)

if __name__ == '__main__':
    main()
