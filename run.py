"""Run the CAIO coding pipeline from an Excel workbook.

The Excel workbook is automatically converted to data/postings.csv
before processing, so the entire pipeline can be run with one command.

Examples:
    python run.py --limit 1 --template data/CAIO_DATA_TEST.xlsx
    python run.py --limit 5 --template data/CAIO_DATA_TEST.xlsx
    python run.py --limit 60 --template data/CAIO_DATA.xlsx

Safe to stop and rerun: finished postings are skipped, and raw model
answers are cached.
"""

import argparse
import csv
import json
import sys
import hashlib
from pathlib import Path

import yaml
from pydantic import ValidationError

# Make src/ importable
sys.path.insert(0, str(Path(__file__).parent / "src"))

from make_slim import main as make_slim  # noqa: E402
from derive import derive  # noqa: E402
from extract import get_coding  # noqa: E402
from schema import PostingCoding  # noqa: E402
from validate import validate  # noqa: E402
from write_excel import existing_urls, open_results, write_row  # noqa: E402


def load_prompt(version: str) -> str:
    """Find prompts/<version>.* regardless of capitalization or extension."""
    for p in Path("prompts").iterdir():
        if p.stem.lower() == version.lower():
            return p.read_text(encoding="utf-8")
    return ""


def main():
    ap = argparse.ArgumentParser()

    ap.add_argument(
        "--template",
        help="path to the CAIO Excel workbook",
    )

    ap.add_argument(
        "--limit",
        type=int,
        help="only process the first N postings",
    )

    ap.add_argument(
        "--fake",
        action="store_true",
        help="use made-up AI answers to test the pipeline",
    )

    args = ap.parse_args()

    # Load configuration
    with open("config.yaml", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}

    version = str(cfg.get("prompt_version") or "v1")
    collected_by = cfg.get("collected_by") or "Script"

    # Determine which workbook to use
    template = args.template or cfg.get("workbook_path")

    if not template or not Path(template).exists():
        sys.exit(
            "Need the workbook: pass "
            "--template data/<your CAIO workbook>.xlsx"
        )

    # ---------------------------------------------------------
    # STEP 1: Convert Excel -> data/postings.csv automatically
    # ---------------------------------------------------------

    print(f"Reading Excel workbook: {template}")

    try:
        make_slim(template)
    except Exception as e:
        sys.exit(f"Could not create postings.csv from Excel file: {e}")

    csv_path = Path("data/postings.csv")

    if not csv_path.exists():
        sys.exit("data/postings.csv was not created.")

    # ---------------------------------------------------------
    # STEP 2: Load the system prompt
    # ---------------------------------------------------------

    system_prompt = load_prompt(version)

    if not system_prompt and not args.fake:
        sys.exit(
            f"No prompt file found for version '{version}' in prompts/."
        )

    # ---------------------------------------------------------
    # STEP 3: Set up output and cache
    # ---------------------------------------------------------

    tag = version + ("_fake" if args.fake else "")

    out_path = f"output/results_{tag}.xlsx"

    cache_dir = Path("cache") / tag
    cache_dir.mkdir(parents=True, exist_ok=True)

    Path("output").mkdir(exist_ok=True)

    # ---------------------------------------------------------
    # STEP 4: Read the slim CSV
    # ---------------------------------------------------------

    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        postings = list(csv.DictReader(f))

    if args.limit:
        postings = postings[: args.limit]

    print(f"Processing {len(postings)} posting(s)...")
    print()

    # ---------------------------------------------------------
    # STEP 5: Open output workbook
    # ---------------------------------------------------------

    wb, ws = open_results(template, out_path)

    done = existing_urls(ws)

    written = 0
    skipped = 0
    failed = 0
    flagged = 0

    err_log = Path("output") / f"errors_{tag}.log"

    # ---------------------------------------------------------
    # STEP 6: Process each posting
    # ---------------------------------------------------------

    for i, posting in enumerate(postings, 1):

        url = posting["url"]

        label = (
            f"[{i}/{len(postings)}] "
            f"{posting['company'][:30]} | "
            f"{posting['title'][:40]}"
        )

        # Skip postings already written to the output workbook
        if url in done:
            skipped += 1
            print(f"{label} -> SKIPPED (already completed)")
            continue

        try:
            # Cache file is based on posting URL
            cache_file = (
                cache_dir
                / (hashlib.sha1(url.encode()).hexdigest() + ".json")
            )

            # Use cached AI result if available
            if cache_file.exists():

                raw = json.loads(
                    cache_file.read_text(encoding="utf-8")
                )

                print(f"{label} -> using cached result")

            else:

                # Real Gemini call happens here
                raw = get_coding(
                    posting,
                    system_prompt,
                    cfg,
                    fake=args.fake,
                )

                # Save raw result so we don't have to pay for it again
                cache_file.write_text(
                    json.dumps(raw, indent=1),
                    encoding="utf-8",
                )

            # Validate against Pydantic schema
            coding = PostingCoding.model_validate(raw)

            # Validate evidence against the actual posting text
            flags = validate(
                coding,
                posting["description"],
            )

            # Calculate derived fields
            derived = derive(coding)

            # Write result to Excel
            write_row(
                ws,
                posting,
                coding,
                derived,
                flags,
                collected_by,
            )

            written += 1
            flagged += bool(flags)

            print(
                f"{label} -> {derived.level_name}"
                + (
                    f"  ({len(flags)} flag(s))"
                    if flags
                    else ""
                )
            )

            # Save periodically
            if written % 10 == 0:
                wb.save(out_path)

        except NotImplementedError as e:

            wb.save(out_path)
            sys.exit(str(e))

        except (ValidationError, Exception) as e:

            failed += 1

            print(
                f"{label} -> FAILED: {str(e)[:200]}"
            )

            with open(err_log, "a", encoding="utf-8") as f:
                f.write(
                    f"{url}\n"
                    f"{e}\n\n"
                )

    # ---------------------------------------------------------
    # STEP 7: Save final workbook
    # ---------------------------------------------------------

    try:
        wb.save(out_path)

    except PermissionError:
        sys.exit(
            f"Could not save {out_path}. "
            "Close it in Excel and rerun."
        )

    print()
    print(
        f"Done. written={written} "
        f"skipped={skipped} "
        f"failed={failed} "
        f"with_flags={flagged}"
    )

    print(f"Results: {out_path}")


if __name__ == "__main__":
    main()