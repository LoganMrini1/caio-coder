"""Structured output schema for CAIO job-posting coding."""

from typing import Literal, Optional
from pydantic import BaseModel, Field


class DimScore(BaseModel):
    score: int = Field(
        description="Integer 0-3: 0 not mentioned, 1 indirect, 2 explicit, 3 explicit authority or ownership."
    )
    evidence: str = Field(
        description="Verbatim evidence quote from the posting, max about 25 words. Empty if score is 0."
    )


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
    low: bool = Field(
        description="True if advise, support, or recommend language appears in connection with the role."
    )
    moderate: bool = Field(
        description="True if coordinate, partner, or facilitate language appears in connection with the role."
    )
    stronger: bool = Field(
        description="True if lead, oversee, chair, or govern language appears in connection with the role."
    )
    high: bool = Field(
        description="True if approve, prioritize, set policy, enforce, halt, or waive language appears in connection with the role."
    )
    observed_language: str = Field(
        description="Actual authority verbs or phrases quoted from the posting."
    )


class Interacts(BaseModel):
    cio: bool
    ciso: bool
    legal: bool
    business_units: bool
    hr: bool
    procurement: bool
    privacy: bool
    board: bool
    other: Optional[str] = None


class PostingCoding(BaseModel):
    organization_name: str
    sector: Optional[
        Literal["Public", "Private", "Government", "Nonprofit"]
    ] = None
    industry: Optional[str] = None
    org_size: Optional[str] = None
    headquarters: Optional[str] = None

    exact_title: str

    title_bucket: Literal[
        "CAIO",
        "VP of AI",
        "CDAO",
        "AI Leader",
        "Other",
    ]

    appointment_date: Optional[str] = None

    employment_type: Literal[
        "Full-time",
        "Interim",
        "Fractional",
        "Unknown",
    ]

    person_name: Optional[str] = None
    prior_role: Optional[str] = None
    prior_function: Optional[str] = None

    background_type: Literal[
        "Technical",
        "Business",
        "Legal",
        "Cyber",
        "Mixed",
        "Unknown",
    ]

    reporting_line: Literal[
        "CEO",
        "CIO",
        "CTO",
        "CDO",
        "CISO",
        "COO",
        "Legal",
        "Risk",
        "Board",
        "Unknown",
    ]

    scope: Literal[
        "Enterprise-wide",
        "Business-unit",
        "Product",
        "Data/Analytics",
        "Governance",
        "Federal Agency",
        "Regional",
    ]

    team_structure: Literal[
        "Owns AI team",
        "Chairs AI council",
        "Embedded in IT",
        "Works through matrix structure",
        "Unknown",
    ]

    dimensions: Dimensions

    decision_rights: DecisionRights

    decision_rights_override: Optional[
        Literal[
            "Low direct authority",
            "Orchestration role, moderate authority",
            "Stronger authority",
            "High decision-rights authority",
        ]
    ] = None

    cyber_score: DimScore

    interacts: Interacts

    orchestration_maturity_level: Literal[
        "Level 1: Symbolic CAIO",
        "Level 2: Advisory CAIO",
        "Level 3: Coordinating CAIO",
        "Level 4: Governing CAIO",
        "Level 5: Enterprise Orchestrator",
    ]

    caio_archetype: Literal[
        "AI Strategy Orchestrator",
        "AI Governance Orchestrator",
        "AI Technology Orchestrator",
        "Cyber/IS Orchestrator",
        "Workforce Transformation Orchestrator",
        "Hybrid Enterprise Orchestrator",
    ]

    source_quote_notes: Optional[str] = None
    general_notes: str = ""