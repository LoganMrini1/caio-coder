
"""Get structured coding for one posting using Gemini or Claude."""

import hashlib
import os

from dotenv import load_dotenv
from google import genai
from google.genai import types
from anthropic import Anthropic

from src.schema import Dimensions, PostingCoding

load_dotenv()


def build_user_message(posting: dict) -> str:
    return (
        f"Company: {posting['company']}\n"
        f"Job title: {posting['title']}\n"
        f"Location: {posting['location']}\n\n"
        f"POSTING TEXT:\n{posting['description']}"
    )


def get_coding(
    posting: dict,
    system_prompt: str,
    cfg: dict,
    fake: bool = False,
) -> dict:
    if fake:
        return _fake(posting)

    provider = cfg.get("provider", "gemini").strip().lower()
    model = cfg.get("model")
    user_message = build_user_message(posting)

    if not model:
        raise RuntimeError("No model is configured in config.yaml.")

    if provider == "anthropic":
        api_key = os.getenv("ANTHROPIC_API_KEY")
        if not api_key:
            raise RuntimeError(
                "ANTHROPIC_API_KEY is not set. Add it to your .env file."
            )

        client = Anthropic(api_key=api_key)

        # Use Claude's structured-output support to return the existing
        # Pydantic schema. Temperature is omitted for compatibility with
        # current Claude models.
        response = client.messages.parse(
            model=model,
            max_tokens=8192,
            system=system_prompt,
            messages=[
                {
                    "role": "user",
                    "content": user_message,
                }
            ],
            output_format=PostingCoding,
        )

        result = response.parsed_output
        if result is None:
            raise RuntimeError("Claude returned no parsed coding result.")

        return result.model_dump()

    if provider == "gemini":
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise RuntimeError(
                "GEMINI_API_KEY is not set. Add it to your .env file."
            )

        client = genai.Client(api_key=api_key)

        response = client.models.generate_content(
            model=model,
            contents=user_message,
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                response_mime_type="application/json",
                response_schema=PostingCoding,
                temperature=cfg.get("temperature", 0),
            ),
        )

        if not response.text:
            raise RuntimeError("Gemini returned an empty response.")

        result = PostingCoding.model_validate_json(response.text)
        return result.model_dump()

    raise ValueError(
        f"Unsupported provider '{provider}'. "
        "Set provider to 'anthropic' or 'gemini' in config.yaml."
    )


def _fake(posting: dict) -> dict:
    text = posting["description"]
    quote = " ".join(text.split()[:8])
    seed = int(hashlib.sha1(posting["url"].encode()).hexdigest(), 16)

    def bits(i, n):
        return (seed >> (i * 5)) % n

    dims = {}
    for i, key in enumerate(Dimensions.model_fields):
        score = bits(i, 4)
        dims[key] = {
            "score": score,
            "evidence": quote if score else "",
        }

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
            key: bool(bits(17 + j, 2))
            for j, key in enumerate(
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
