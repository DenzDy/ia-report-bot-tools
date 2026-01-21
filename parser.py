from pydantic import BaseModel
from typing import List
from pptx import Presentation

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

def extract_text_from_pptx(file_path):
    prs = Presentation(file_path)
    full_text = []

    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                full_text.append(f"Slide {i+1}: {shape.text}")
    
    return "\n".join(full_text)

def main():
    text = extract_text_from_pptx()    