"""Structured output the LLM must return for one job posting.

Enums mirror the workbook dropdowns exactly. The LLM only codes INPUTS;
breadth score, DR strength, span count, and maturity are derived in code/Excel.
"""
from typing import Literal, Optional
from pydantic import BaseModel, Field

Score = Literal[0, 1, 2, 3]


class DimScore(BaseModel):
    score: Score
    evidence: str = Field(description="Verbatim quote (max ~25 words) from the posting supporting the score. Empty string if score is 0.")


class Dimensions(BaseModel):
    strategic: DimScore
    technical: DimScore
    data: DimScore
    cybersecurity: DimScore
    risk_compliance: DimScore
    stakeholder: DimScore
    governance: DimScore
    workforce: DimScore
    vendor_ecosystem: DimScore
    value: DimScore


class DecisionRights(BaseModel):
    low: bool = Field(description="advise/support/recommend language appears")
    moderate: bool = Field(description="coordinate/partner/facilitate language appears")
    stronger: bool = Field(description="lead/oversee/chair/govern language appears")
    high: bool = Field(description="approve/prioritize/set policy/enforce/halt/waive language appears")
    observed_language: str = Field(description="The actual verbs/phrases from the posting, quoted")


class Interacts(BaseModel):
    cio: bool
    ciso: bool
    legal: bool
    business_units: bool
    hr: bool
    procurement: bool
    privacy: bool
    board: bool


class PostingCoding(BaseModel):
    # Organization (null if the posting does not say)
    organization_name: str
    sector: Optional[Literal["Public", "Private", "Government", "Nonprofit"]] = None
    industry: Optional[str] = None
    org_size: Optional[str] = None
    headquarters: Optional[str] = None
    # Role
    exact_title: str
    title_bucket: Literal["CAIO", "VP of AI", "CDAO", "AI Leader", "Other"]
    employment_type: Literal["Full-time", "Interim", "Fractional", "Unknown"]
    reporting_line: Literal["CEO", "CIO", "CTO", "CDO", "CISO", "COO", "Legal", "Risk", "Board", "Unknown"]
    scope: Literal["Enterprise-wide", "Business-unit", "Product", "Data/Analytics", "Governance", "Federal Agency", "Regional"]
    team_structure: Literal["Owns AI team", "Chairs AI council", "Embedded in IT", "Works through matrix structure", "Unknown"]
    # Coding
    dimensions: Dimensions
    decision_rights: DecisionRights
    cyber_score: DimScore
    interacts: Interacts
    notes: str = Field(description="Ambiguities, conflicting signals, or anything a second reader should check")
