"""Post-LLM checks. Returns a list of flag strings (empty list = clean).

Schema/enum validity is already enforced by PostingCoding.model_validate in run.py.
Here we check the things a schema can't: that every quote really appears in the posting.
"""
import re

from derive import DIMENSION_KEYS
from schema import PostingCoding


def _norm(s: str) -> str:
    """Lowercase and drop punctuation/bullets/quote styles so formatting differences don't cause false alarms."""
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def quote_in_text(quote: str, text: str) -> bool:
    haystack = _norm(text)
    # allow the model to use "..." to skip words; each fragment must appear
    parts = [p for p in re.split(r"\.\.\.|…", quote) if _norm(p)]
    return bool(parts) and all(_norm(p) in haystack for p in parts)


def validate(coding: PostingCoding, text: str) -> list[str]:
    flags = []
    items = [(k, getattr(coding.dimensions, k)) for k in DIMENSION_KEYS]
    items.append(("cyber_score", coding.cyber_score))
    for name, ds in items:
        if ds.score > 0 and not ds.evidence.strip():
            flags.append(f"{name}: score {ds.score} but no evidence quote")
        elif ds.score > 0 and not quote_in_text(ds.evidence, text):
            flags.append(f"{name}: quote not found in posting")

    dr = coding.decision_rights
    any_tier = dr.low or dr.moderate or dr.stronger or dr.high
    if any_tier and not dr.observed_language.strip():
        flags.append("decision_rights: tier marked but no language quoted")
    if not any_tier and dr.observed_language.strip():
        flags.append("decision_rights: language quoted but no tier marked")
    return flags