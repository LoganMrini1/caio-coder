"""Get structured coding for one posting using Gemini."""

import hashlib
import os

from google import genai
from google.genai import types
from src.schema import Dimensions, PostingCoding
from dotenv import load_dotenv

load_dotenv()


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

    # Get API key from environment
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY is not set. Add it to your .env file."
        )

    # Create Gemini client
    client = genai.Client(api_key=api_key)

    # Get model name from config, with a default
    model = cfg.get("model", "gemini-3.8-flash")

    # Combine the system prompt and the actual job posting
    user_message = build_user_message(posting)

    # Ask Gemini for structured JSON matching our Pydantic schema
    response = client.models.generate_content(
        model=model,
        contents=user_message,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            response_mime_type="application/json",
            response_schema=PostingCoding,
        ),
    )

    # Parse and validate Gemini's response
    if not response.text:
        raise RuntimeError("Gemini returned an empty response.")

    result = PostingCoding.model_validate_json(response.text)

    return result.model_dump()


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
            "low": bool(bits(13, 2)),
            "moderate": bool(bits(14, 2)),
            "stronger": bool(bits(15, 2)),
            "high": bool(bits(16, 2)),
            "observed_language": quote,
        },
        "cyber_score": {
            "score": cyber,
            "evidence": quote if cyber else "",
        },
        "interacts": {
            k: bool(bits(17 + j, 2))
            for j, k in enumerate(
                [
                    "cio",
                    "ciso",
                    "legal",
                    "business_units",
                    "hr",
                    "procurement",
                    "privacy",
                    "board",
                ]
            )
        },
        "notes": "FAKE DATA - pipeline test only, not real coding.",
    }


# Quick self-check for fake mode
if __name__ == "__main__":
    PostingCoding.model_validate(
        _fake(
            {
                "url": "x",
                "company": "C",
                "title": "T",
                "location": "L",
                "description": "a b c d e f g h i",
            }
        )
    )
    print("fake output is schema-valid")