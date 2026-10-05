"""Validation checks for Gemini CAIO coding output."""

from src.derive import DIMENSION_KEYS
from src.schema import PostingCoding


def validate_coding(coding: PostingCoding) -> list[str]:
    errors = []

    # Check all orchestration dimension scores.
    for key in DIMENSION_KEYS:
        dimension = getattr(coding.dimensions, key)

        if not 0 <= dimension.score <= 3:
            errors.append(
                f"{key} score must be between 0 and 3."
            )

        if dimension.score == 0 and dimension.evidence:
            errors.append(
                f"{key} has evidence even though its score is 0."
            )

        if dimension.score > 0 and not dimension.evidence.strip():
            errors.append(
                f"{key} has a non-zero score but no evidence quote."
            )

    # Check cybersecurity score.
    if not 0 <= coding.cyber_score.score <= 3:
        errors.append(
            "cyber_score must be between 0 and 3."
        )

    if (
        coding.cyber_score.score == 0
        and coding.cyber_score.evidence
    ):
        errors.append(
            "cyber_score has evidence even though its score is 0."
        )

    if (
        coding.cyber_score.score > 0
        and not coding.cyber_score.evidence.strip()
    ):
        errors.append(
            "cyber_score has a non-zero score but no evidence quote."
        )

    # Decision-rights language should have evidence when any tier is true.
    dr = coding.decision_rights

    any_dr = (
        dr.low
        or dr.moderate
        or dr.stronger
        or dr.high
    )

    if any_dr and not dr.observed_language.strip():
        errors.append(
            "Decision-rights language is marked true but no observed language was provided."
        )

    return errors
    