"""Run the pipeline over data/postings.csv.

Examples (from the project root):
    python run.py --fake --limit 5 --template data/CAIO_Data_Copyable.xlsx     # test without an API key
    python run.py --limit 60 --template data/CAIO_Data_Copyable.xlsx           # real run (needs extract.py finished)

Safe to stop and rerun: finished postings are skipped, and raw model answers are cached.
"""
import argparse
import csv
import json
import sys
import hashlib
from pathlib import Path

import yaml
from pydantic import ValidationError

sys.path.insert(0, str(Path(__file__).parent / "src"))
from derive import derive              # noqa: E402
from extract import get_coding         # noqa: E402
from schema import PostingCoding       # noqa: E402
from validate import validate          # noqa: E402
from write_excel import existing_urls, open_results, write_row  # noqa: E402


def load_prompt(version: str) -> str:
    """Find prompts/<version>.* regardless of capitalization or extension (.md/.txt)."""
    for p in Path("prompts").iterdir():
        if p.stem.lower() == version.lower():
            return p.read_text(encoding="utf-8")
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", default="data/postings.csv")
    ap.add_argument("--template", help="path to the CAIO Data workbook to start from")
    ap.add_argument("--limit", type=int, help="only process the first N postings")
    ap.add_argument("--fake", action="store_true", help="use made-up AI answers to test the pipeline")
    args = ap.parse_args()

    cfg = yaml.safe_load(open("config.yaml", encoding="utf-8")) or {}
    version = str(cfg.get("prompt_version") or "v1")
    collected_by = cfg.get("collected_by") or "Script"
    template = args.template or cfg.get("workbook_path")
    if not template or not Path(template).exists():
        sys.exit("Need the workbook: pass --template data/<your CAIO workbook>.xlsx (file must exist).")

    system_prompt = load_prompt(version)
    if not system_prompt and not args.fake:
        sys.exit(f"No prompt file found for version '{version}' in prompts/.")

    tag = version + ("_fake" if args.fake else "")
    out_path = f"output/results_{tag}.xlsx"
    cache_dir = Path("cache") / tag
    cache_dir.mkdir(parents=True, exist_ok=True)
    Path("output").mkdir(exist_ok=True)

    with open(args.csv, encoding="utf-8-sig", newline="") as f:
        postings = list(csv.DictReader(f))
    if args.limit:
        postings = postings[: args.limit]

    wb, ws = open_results(template, out_path)
    done = existing_urls(ws)
    written = skipped = failed = flagged = 0
    err_log = Path("output") / f"errors_{tag}.log"

    for i, posting in enumerate(postings, 1):
        url = posting["url"]
        label = f"[{i}/{len(postings)}] {posting['company'][:30]} | {posting['title'][:40]}"
        if url in done:
            skipped += 1
            continue
        try:
            cache_file = cache_dir / (hashlib.sha1(url.encode()).hexdigest() + ".json")
            if cache_file.exists():
                raw = json.loads(cache_file.read_text(encoding="utf-8"))
            else:
                raw = get_coding(posting, system_prompt, cfg, fake=args.fake)
                cache_file.write_text(json.dumps(raw, indent=1), encoding="utf-8")
            coding = PostingCoding.model_validate(raw)
            flags = validate(coding, posting["description"])
            derived = derive(coding)
            write_row(ws, posting, coding, derived, flags, collected_by)
            written += 1
            flagged += bool(flags)
            print(f"{label} -> {derived.level_name}" + (f"  ({len(flags)} flag(s))" if flags else ""))
            if written % 10 == 0:
                wb.save(out_path)
        except NotImplementedError as e:
            wb.save(out_path)
            sys.exit(str(e))
        except (ValidationError, Exception) as e:  # keep going; log and move on
            failed += 1
            print(f"{label} -> FAILED: {str(e)[:120]}")
            with open(err_log, "a", encoding="utf-8") as f:
                f.write(f"{url}\n{e}\n\n")

    try:
        wb.save(out_path)
    except PermissionError:
        sys.exit(f"Could not save {out_path}. Close it in Excel and rerun (finished rows are remembered).")
    print(f"\nDone. written={written} skipped={skipped} failed={failed} with_flags={flagged}")
    print(f"Results: {out_path}")


if __name__ == "__main__":
    main()