from pydantic import BaseModel
from typing import List

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
     