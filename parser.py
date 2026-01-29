from pydantic import BaseModel
from typing import List
from pptx import Presentation
import os
from dotenv import load_dotenv
from typing import Dict
import json 
from openai import OpenAI
import ast

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
    overall_audit_rating : str 
    overall_risk_description : str
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
        slides_content.append(f"[SLIDE START]\n{full_slide_text}\n[SLIDE END]\n")
    
    return "\n".join(slides_content)

def generate_json(extracted_text):
    extraction_prompt = f"""
    Act as an Internal Audit Data Analyst. Your goal is to parse raw text extracted from MULTIPLE PowerPoint reports into a structured JSON format. 

    ### Input Format
    The source text contains multiple reports. Each report begins with a marker like [[START_FILE: filename]] and ends with [[END_FILE]].
    Each slide's content is enclosed in [SLIDE START] amd [SLIDE END] markers.

    ### Data Mapping Instructions (Apply to EACH report):
    1. **report_title**: Locate the main title of the audit.
    2. **executive_summary**: Summarize overarching conclusions and background.
    3. **overall_audit_rating**: The overall audit rating, which can be INADEQUATE, FOR IMPROVEMENT, or ADEQUATE
    4. **overall_audit_conclusion**: The overall conclusions and findings of the report.   
    4. **details**: A list of audit findings. For each finding:
    - **observation**: Factual description of findings.
    - **risk**: Potential negative impact.
    - **risk_rating**: Severity (e.g., INADEQUATE, FOR IMPROVEMENT, ADEQUATE).
    - **recommendation**: Suggestion for improvement for THAT finding.
    5. **recommendations**: A list of all corrective actions in the deck.
    6. **management_action_plan**: Specific management commitments.

    ### Constraints:
    - Return a SINGLE JSON OBJECT where each key is the filename provided in the [[START_FILE]] marker.
    - Output ONLY valid JSON.
    - If a field is not found, use "" or []. Do not omit keys.

    ### OUTPUT RULES:
    - Use double quotes for all keys and string values
    - Do not include explanations or formatting

    ### Target JSON Schema:
    {{
    "filename_1.pptx": {{
        "report_title": "string",
        "executive_summary": "string",
        "overall_audit_rating": "string",
        "overall_audit_conclusion": "string",
        "details": [{{"observation": "string", "risk": "string", "risk_rating": "string", "recommendation": "string"}}],
        "recommendations": ["string"],
        "management_action_plan": ["string"]
    }},
    "filename_2.pptx": {{ ... }}
    }}
    """
    # Load and import API Keys
    load_dotenv()
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
    
    openai_client = OpenAI(
        api_key=OPENAI_API_KEY
    )
    response = openai_client.responses.create(
        model='gpt-5-mini', 
        instructions=extraction_prompt,
        input=extracted_text,
    )
    raw = response.output_text.strip()
    data = ast.literal_eval(raw)
    return data

def main():
    # TODO: Batch processing
    extracted_data : dict[str, ReportData] = {}
    all_pptx = [f for f in os.listdir('templates') if f.endswith('.pptx')]
    for i in range(0, len(all_pptx), 5):
        combined_text = ""
        batch_files = all_pptx[i:i+5]
        for filename in batch_files:
            # Local extraction is fast and consumes no quota
            text = extract_slide_content(f"templates/{filename}")
            print(f"[START OF FILE]\n{text}\n[END OF FILE]\n")
            combined_text += f"\n[[START_FILE: {filename}]]\n{text}\n[[END_FILE]]\n"

        # Create Batch JSONs and add it to final JSON list
        batch_json = generate_json(combined_text)
        for filename, data in batch_json.items():
            extracted_data[filename] = data

    # Create JSON file
    with open("ac_test_output.json", "w") as f:
        json.dump(extracted_data, f, indent=4)
if __name__ == '__main__':
    main()