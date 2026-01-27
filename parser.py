from pydantic import BaseModel
from typing import List
from pptx import Presentation
from google import genai
import os
from dotenv import load_dotenv
from typing import Dict
import json 
from google.genai import types

# JSON Object Definitions
class DetailsTable(BaseModel):
    observation: str
    risk: str
    risk_rating: str
    recommendation: str
    status: str

class ReportData(BaseModel):
    report_title : str
    executive_summary : str
    details: List[DetailsTable]
    recommendations: List[str]
    management_action_plan: List[str]

class BatchResponse(BaseModel):
    reports: List[ReportData]
    
def extract_slide_content(file_path):
    prs = Presentation(file_path)
    slides_content = []

    for i, slide in enumerate(prs.slides):
        slide_text_blocks = []
        
        for shape in slide.shapes:
            # Extract from standard text shapes
            if hasattr(shape, "text") and shape.text.strip():
                slide_text_blocks.append(shape.text.strip())
            
            # Extract from tables
            elif shape.has_table:
                table_data = []
                for row in shape.table.rows:
                    # Join cell text with a pipe (|) to maintain table structure
                    row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if row_text:
                        table_data.append(row_text)
                
                if table_data:
                    slide_text_blocks.append("\n".join(table_data))

        # Join all text found on this specific slide into one string
        full_slide_text = "\n".join(slide_text_blocks)
        slides_content.append(f"{full_slide_text}")
    
    return "\n".join(slides_content)

def generate_json(extracted_text):
    extraction_prompt = f"""
    Act as an Internal Audit Data Analyst. Your goal is to parse raw text extracted from MULTIPLE PowerPoint reports into a structured JSON format. 

    ### Input Format
    The source text contains multiple reports. Each report begins with a marker like [[START_FILE: filename]] and ends with [[END_FILE]].

    ### Data Mapping Instructions (Apply to EACH report):
    1. **report_title**: Locate the main title of the audit.
    2. **executive_summary**: Summarize overarching conclusions and background.
    3. **details**: A list of audit findings. For each finding:
    - **observation**: Factual description of findings.
    - **risk**: Potential negative impact.
    - **risk_rating**: Severity (e.g., INADEQUATE, FOR IMPROVEMENT, ADEQUATE).
    - **recommendation**: Suggestion for improvement for THAT finding.
    4. **recommendations**: A list of all corrective actions in the deck.
    5. **management_action_plan**: Specific management commitments.

    ### Constraints:
    - Return a SINGLE JSON object where each key is the filename provided in the [[START_FILE]] marker.
    - Output ONLY valid JSON.
    - If a field is not found, use "" or []. Do not omit keys.

    ### Target JSON Schema:
    {{
    "filename_1.pptx": {{
        "report_title": "string",
        "executive_summary": "string",
        "details": [{{"observation": "string", "risk": "string", "risk_rating": "string", "recommendation": "string"}}],
        "recommendations": ["string"],
        "management_action_plan": ["string"]
    }},
    "filename_2.pptx": {{ ... }}
    }}

    ### Source Text:
    {extracted_text}
    """
    # Load and import API Keys
    load_dotenv()
    GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')
    
    google_client = genai.Client(api_key=GEMINI_API_KEY)
    response = google_client.models.generate_content(
        model='gemini-3-flash-preview', 
        contents=extraction_prompt,
        config=types.GenerateContentConfig(
            responseMimeType='application/json',
            responseSchema=BatchResponse
        )
    )
    return json.loads(response.text)
def main():
    # TODO: Batch processing
    extracted_data : dict[str, ReportData] = {}
    all_pptx = [f for f in os.listdir('generated_pptx') if f.endswith('.pptx')]
    batch_files = all_pptx[0:5]
    combined_text = ""
    for filename in batch_files:
        # Local extraction is fast and consumes no quota
        text = extract_slide_content(f"generated_pptx/{filename}")
        combined_text += f"\n[[START_FILE: {filename}]]\n{text}\n[[END_FILE]]\n"
    batch_json = generate_json(combined_text)
    extracted_data.update(batch_json)

    with open("extracted_output.json", "w") as f:
        json.dump(extracted_data, f, indent=4)
if __name__ == '__main__':
    main()