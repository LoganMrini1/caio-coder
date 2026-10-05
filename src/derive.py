"""Values computed in code (not by the LLM).

Maturity mirrors the workbook's own formulas (columns BG, BH, BI) so the value we write
into 'Orchestration Maturity Level' matches the sheet's 'Suggested Maturity Level (auto)'.

The archetype rule is a v1 PLACEHOLDER built from the plan's archetype table.
"""

from dataclasses import dataclass

from src.schema import Dimensions, PostingCoding


DIMENSION_KEYS = list(Dimensions.model_fields)


LEVEL_NAMES = {
    1: "Level 1: Symbolic CAIO",
    2: "Level 2: Advisory CAIO",
    3: "Level 3: Coordinating CAIO",
    4: "Level 4: Governing CAIO",
    5: "Level 5: Enterprise Orchestrator",
}


# Archetype groups: which dimensions feed each one
# (averaged, so bigger groups aren't favored)
ARCHETYPE_GROUPS = [
    ("AI Strategy Orchestrator", ["strategic", "value"]),
    ("AI Governance Orchestrator", ["governance", "risk_compliance"]),
    ("AI Technology Orchestrator", ["technical", "data"]),
    ("Cyber/IS Orchestrator", ["cybersecurity", "CYBER_ASPECT"]),
    ("Workforce Transformation Orchestrator", ["workforce"]),
]


HYBRID_REQUIRES = [
    "strategic",
    "governance",
    "technical",
    "data",
    "risk_compliance",
    "workforce",
]


@dataclass
class Derived:
    breadth: int
    dr_rank: int
    span: int
    base_level: int
    level: int
    level_name: str
    archetype: str


def derive(coding: PostingCoding) -> Derived:
    scores = {
        k: getattr(coding.dimensions, k).score
        for k in DIMENSION_KEYS
    }

    breadth = sum(scores.values())

    dr = coding.decision_rights

    rank = (
        4
        if dr.high
        else 3
        if dr.stronger
        else 2
        if dr.moderate
        else 1
        if dr.low
        else 0
    )

    span = sum(
        bool(v)
        for v in coding.interacts.model_dump().values()
        if isinstance(v, bool)
    )

    base = (
        1
        if breadth < 6
        else 2
        if breadth < 12
        else 3
        if breadth < 18
        else 4
        if breadth < 24
        else 5
    )

    level = base

    if base >= 4 and rank < 3:
        level = 3
    elif base == 5 and (span < 6 or rank < 4):
        level = 4

    return Derived(
        breadth=breadth,
        dr_rank=rank,
        span=span,
        base_level=base,
        level=level,
        level_name=LEVEL_NAMES[level],
        archetype=_archetype(
            scores,
            coding.cyber_score.score,
        ),
    )


def _archetype(scores, cyber_aspect):
    if all(
        scores[k] >= 2
        for k in HYBRID_REQUIRES
    ):
        return "Hybrid Enterprise Orchestrator"

    best_name = ""
    best_avg = 0.0

    for name, keys in ARCHETYPE_GROUPS:
        vals = [
            cyber_aspect
            if k == "CYBER_ASPECT"
            else scores[k]
            for k in keys
        ]

        avg = sum(vals) / len(vals)

        if avg > best_avg:
            best_name = name
            best_avg = avg

    return best_name