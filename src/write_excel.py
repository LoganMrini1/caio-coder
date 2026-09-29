"""Append coded postings to the CAIO Data workbook."""
import re
import shutil
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from derive import DIMENSION_KEYS, Derived
from excel_map import (ARCHETYPE_COL, CONST_COLS, CYBER_COL, DIM_COLS, DR_COLS, DR_QUOTE_COL,
                       FIELD_COLS, FIRST_DATA_ROW, INTERACT_COLS, MATURITY_COL,
                       QUOTE_NOTES_COL, SHEET, copy_formulas, next_empty_row)
from schema import PostingCoding

_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _clean(v):
    return _ILLEGAL.sub("", v) if isinstance(v, str) else v


def open_results(template: str, out_path: str):
    """Start from a copy of the template; if the output file already exists, keep adding to it."""
    out = Path(out_path)
    if not out.exists():
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(template, out)
    wb = load_workbook(out)
    return wb, wb[SHEET]


def existing_urls(ws) -> set:
    col = CONST_COLS["source_url"]
    urls, r = set(), FIRST_DATA_ROW
    while ws[f"{col}{r}"].value not in (None, ""):
        urls.add(str(ws[f"{col}{r}"].value))
        r += 1
    return urls


def _yn(b: bool) -> str:
    return "Yes" if b else "No"


def write_row(ws, posting: dict, coding: PostingCoding, derived: Derived, flags: list[str], collected_by: str) -> int:
    r = next_empty_row(ws)
    copy_formulas(ws, r)

    def put(col, val):
        ws[f"{col}{r}"] = _clean(val)

    # constants set by code
    put(CONST_COLS["date"], date.today().isoformat())
    put(CONST_COLS["collected_by"], collected_by)
    put(CONST_COLS["source_category"], "Job Posting")
    put(CONST_COLS["evidence_type"], "Job posting")
    put(CONST_COLS["source_url"], posting["url"])

    # simple fields from the LLM (skip nulls; notes handled below)
    for field, col in FIELD_COLS.items():
        if field == "notes":
            continue
        val = getattr(coding, field)
        if val is not None:
            put(col, val)

    # orchestration scores + cyber aspect
    for key in DIMENSION_KEYS:
        put(DIM_COLS[key], getattr(coding.dimensions, key).score)
    put(CYBER_COL, coding.cyber_score.score)

    # decision rights
    dr = coding.decision_rights
    for tier, col in DR_COLS.items():
        put(col, _yn(getattr(dr, tier)))
    put(DR_QUOTE_COL, dr.observed_language)

    # coordination flags
    for who, col in INTERACT_COLS.items():
        put(col, _yn(getattr(coding.interacts, who)))

    # evidence quotes
    quotes = [f'{k}: "{getattr(coding.dimensions, k).evidence}"' for k in DIMENSION_KEYS
              if getattr(coding.dimensions, k).evidence]
    if coding.cyber_score.evidence:
        quotes.append(f'cyber: "{coding.cyber_score.evidence}"')
    put(QUOTE_NOTES_COL, " | ".join(quotes))

    # derived in code
    put(MATURITY_COL, derived.level_name)
    if derived.archetype:
        put(ARCHETYPE_COL, derived.archetype)

    # notes + flags
    notes = coding.notes.strip()
    if flags:
        notes = (notes + " " if notes else "") + "FLAGS: " + "; ".join(flags)
    put(FIELD_COLS["notes"], notes)
    return r