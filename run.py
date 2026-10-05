from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

import pandas as pd
import yaml
from dotenv import load_dotenv

from src.extract import get_coding
from src.make_slim import main as make_slim
from src.validate import validate_coding
from src.write_excel import prepare_output, write_posting
from src.schema import PostingCoding


ROOT = Path(__file__).resolve().parent


def load_config():
    with open(ROOT / "config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_prompt(version: str) -> str:
    prompt_path = ROOT / "prompts" / f"{version}.txt"

    if not prompt_path.exists():
        prompt_path = ROOT / "prompts" / version / "prompt.txt"

    if not prompt_path.exists():
        raise FileNotFoundError(
            f"Could not find prompt for version '{version}'."
        )

    return prompt_path.read_text(encoding="utf-8")


def cache_key(prompt_version: str, prompt_text: str, posting: dict) -> str:
    payload = {
        "prompt_version": prompt_version,
        "prompt": prompt_text,
        "company": posting.get("company", ""),
        "title": posting.get("title", ""),
        "location": posting.get("location", ""),
        "url": posting.get("url", ""),
        "description": posting.get("description", ""),
    }

    raw = json.dumps(
        payload,
        sort_keys=True,
        ensure_ascii=False,
    ).encode("utf-8")

    return hashlib.sha256(raw).hexdigest()


def load_cache(cache_path: Path) -> dict:
    if not cache_path.exists():
        return {}

    try:
        return json.loads(
            cache_path.read_text(encoding="utf-8")
        )
    except Exception:
        return {}


def save_cache(cache_path: Path, cache: dict):
    cache_path.parent.mkdir(parents=True, exist_ok=True)

    cache_path.write_text(
        json.dumps(
            cache,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Only process the first N postings.",
    )

    parser.add_argument(
        "--input",
        required=True,
        help="Input Excel workbook containing the job descriptions.",
    )

    parser.add_argument(
        "--template",
        default="data/CAIO_URA_Data_Collection copy.xlsx",
        help="CAIO URA Data Collection workbook to use as the output template.",
    )

    args = parser.parse_args()

    load_dotenv(ROOT / ".env")

    config = load_config()

    prompt_version = config.get(
        "prompt_version",
        "v1",
    )

    collected_by = config.get(
        "collected_by",
        "Logan",
    )

    prompt_text = load_prompt(prompt_version)

    print("=" * 60)
    print("CAIO CODING PIPELINE")
    print("=" * 60)
    print(f"Prompt version: {prompt_version}")
    print(f"Collected by:   {collected_by}")
    print(f"Input:          {args.input}")
    print(f"Template:       {args.template}")
    print()

    print("Reading job descriptions from input Excel...")

    make_slim(args.input)

    postings_path = ROOT / "data" / "postings.csv"

    if not postings_path.exists():
        raise FileNotFoundError(
            f"Expected {postings_path} after make_slim()."
        )

    postings = pd.read_csv(
        postings_path
    )

    if args.limit is not None:
        postings = postings.head(args.limit)

    if len(postings) == 0:
        raise ValueError(
            "No postings were found in the input workbook."
        )

    output_dir = ROOT / "output"
    output_dir.mkdir(exist_ok=True)

    output_path = (
        output_dir
        / f"CAIO_URA_results_{prompt_version}.xlsx"
    )

    print()
    print("Creating fresh output workbook...")

    prepare_output(
        args.template,
        output_path,
    )

    cache_path = (
        ROOT
        / "cache"
        / "extractions.json"
    )

    cache = load_cache(cache_path)

    processed = 0

    for _, row in postings.iterrows():

        posting = row.to_dict()

        company = posting.get(
            "company",
            "Unknown company",
        )

        title = posting.get(
            "title",
            "Unknown title",
        )

        print(
            f"[{processed + 1}/{len(postings)}] "
            f"{company} — {title}"
        )

        key = cache_key(
            prompt_version,
            prompt_text,
            posting,
        )

        if key in cache:
            print("  Using cached Gemini result.")
            coding_dict = cache[key]

        else:
            print("  Sending posting text to Gemini...")

            coding_dict = get_coding(
                posting=posting,
                system_prompt=prompt_text,
                cfg=config,
                fake=False,
            )

            cache[key] = coding_dict

            save_cache(
                cache_path,
                cache,
            )

        coding = PostingCoding.model_validate(
            coding_dict
        )

        errors = validate_coding(coding)

        if errors:
            print("  WARNING: validation issues:")

            for error in errors:
                print(f"    - {error}")

        excel_row = 3 + processed

        write_posting(
            output_path=output_path,
            row_num=excel_row,
            posting=posting,
            coding=coding,
            collected_by=collected_by,
        )

        print(
            f"  Wrote to CAIO Data row {excel_row}."
        )

        processed += 1

    print()
    print("=" * 60)
    print(f"Finished. Processed {processed} posting(s).")
    print(f"Output: {output_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()