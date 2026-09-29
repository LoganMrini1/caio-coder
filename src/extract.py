"""Get the structured coding for one posting.

fake mode: returns made-up but schema-valid data so the rest of the pipeline can be tested
without an API key. Fake output is watermarked in the notes and written to separate files.

real mode: NOT WRITTEN YET. It needs the provider and API key. It should:
  1. send SYSTEM PROMPT (prompts/<version>) + build_user_message(posting)
  2. force JSON output that follows PostingCoding.model_json_schema()
  3. return the parsed dict
"""
import hashlib

from schema import Dimensions, PostingCoding


def build_user_message(posting: dict) -> str:
    return (
        f"Company: {posting['company']}\n"
        f"Job title: {posting['title']}\n"
        f"Location: {posting['location']}\n\n"
        f"POSTING TEXT:\n{posting['description']}"
    )


def get_coding(posting: dict, system_prompt: str, cfg: dict, fake: bool = False) -> dict:
    if fake:
        return _fake(posting)
    raise NotImplementedError(
        "Real extraction is not written yet. Run with --fake to test the pipeline, "
        "or add the provider call in src/extract.py once the API key/provider is known."
    )


def _fake(posting: dict) -> dict:
    text = posting["description"]
    quote = " ".join(text.split()[:8])
    seed = int(hashlib.sha1(posting["url"].encode()).hexdigest(), 16)

    def bits(i, n):
        return (seed >> (i * 5)) % n

    dims = {}
    for i, key in enumerate(Dimensions.model_fields):
        s = bits(i, 4)
        dims[key] = {"score": s, "evidence": quote if s else ""}
    cyber = bits(12, 4)
    return {
        "organization_name": posting["company"],
        "exact_title": posting["title"],
        "title_bucket": "Other",
        "employment_type": "Unknown",
        "reporting_line": "Unknown",
        "scope": "Enterprise-wide",
        "team_structure": "Unknown",
        "dimensions": dims,
        "decision_rights": {
            "low": bool(bits(13, 2)), "moderate": bool(bits(14, 2)),
            "stronger": bool(bits(15, 2)), "high": bool(bits(16, 2)),
            "observed_language": quote,
        },
        "cyber_score": {"score": cyber, "evidence": quote if cyber else ""},
        "interacts": {k: bool(bits(17 + j, 2)) for j, k in enumerate(
            ["cio", "ciso", "legal", "business_units", "hr", "procurement", "privacy", "board"])},
        "notes": "FAKE DATA - pipeline test only, not real coding.",
    }


# quick self-check that the fake output really matches the schema
if __name__ == "__main__":
    PostingCoding.model_validate(_fake({"url": "x", "company": "C", "title": "T", "location": "L", "description": "a b c d e f g h i"}))
    print("fake output is schema-valid")
    