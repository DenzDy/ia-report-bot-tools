from pydantic import BaseModel
from typing import List
from pptx import Presentation
from google import genai
import os
from dotenv import load_dotenv
import json 
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
    Act as an Internal Audit Data Analyst. Your goal is to parse raw text extracted from a PowerPoint report into a structured JSON format following a specific schema.

    ### Data Mapping Instructions:
    1. **report_title**: Locate the main title of the audit (usually on Slide 1).
    2. **executive_summary**: Summarize the overarching conclusion and background found in the introductory slides.
    3. **details**: This is a list of audit findings. For each finding/observation found in a table or slide:
    - **observation**: The factual description of what was found.
    - **risk**: The potential negative impact or consequence.
    - **risk_rating**: The severity (e.g., High, Medium, Low).
    - **recommendation**: The specific suggestion for improvement for THAT finding.
    - **status**: Current state (e.g., Open, Closed, In Progress).
    4. **recommendations**: A general list of all corrective actions suggested throughout the deck.
    5. **management_action_plan**: Specific responses or commitments made by management to address the findings.

    ### Constraints:
    - Output ONLY valid JSON.
    - Do not hallucinate data; if a field is not found, use an empty string "" or an empty list [].
    - Combine data if a single observation spans multiple slides.

    ### Target JSON Schema:
    {{
    "report_title": "string",
    "executive_summary": "string",
    "details": [
        {{
        "observation": "string",
        "risk": "string",
        "risk_rating": "string",
        "recommendation": "string",
        "status": "string"
        }}
    ],
    "recommendations": ["string"],
    "management_action_plan": ["string"]
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
        config={
            'response_mime_type': 'application/json',
        }
    )
    
    # Load JSON
    clean_text = response.text.replace("```json", "").replace("```", "").strip()
    return ReportData.model_validate_json(clean_text)
def main():
    extracted_data : dict[str, ReportData] = {}
    for filename in os.listdir('generated_pptx'):
        text = extract_slide_content(f"generated_pptx/{filename}")
        data = generate_json(text)
        extracted_data[filename] = data.model_dump()
    with open("extracted_output.json", "w") as f:
        json.dump(extracted_data, f, indent=4)
if __name__ == '__main__':
    main()