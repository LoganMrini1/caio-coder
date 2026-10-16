"""Score the validation postings with Anthropic and GPT and compare them to the manual coding.

Usage (from the project root):
    python compare_models.py --workbook "data/CAIO Dataset.xlsx"
    python compare_models.py --workbook "data/CAIO Dataset.xlsx" --layout-only

What it does to the workbook (in place):
- 'Validation Analysis V1': each "Script — X" column becomes "Script (Gemini) — X",
  followed by new "Script (Anthropic) — X" and "Script (GPT) — X" columns.
- 'Validation Analysis V1 results': match rates and Pearson r for Anthropic and GPT
  next to the existing Gemini numbers.

Each validation row is matched to its posting through 'Script Output V1'
(row r there is posting r-3 in data/postings.csv).
"""
from __future__ import annotations

import argparse
import json
import statistics
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import copy
from pathlib import Path

import pandas as pd
import yaml
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from run import cache_key, load_cache, save_cache
from src.extract import get_coding
from src.schema import PostingCoding


ROOT = Path(__file__).resolve().parent

VAL_SHEET = "Validation Analysis V1"
RESULTS_SHEET = "Validation Analysis V1 results"
SCRIPT_SHEET = "Script Output V1"

PROVIDERS = {"anthropic": "Anthropic", "openai": "GPT"}
FIXED_COLS = 5  # A-E: Record ID, Organization Name, Exact Title, (blank), (blank)

DIM_LABELS = {
    "Strategic": "strategic", "Technical": "technical", "Data": "data",
    "Cybersecurity": "cybersecurity", "Risk/Compliance": "risk_compliance",
    "Stakeholder": "stakeholder", "Governance": "governance", "Workforce": "workforce",
    "Vendor/Ecosystem": "vendor_ecosystem", "Value": "value",
}
DR_TIERS = {"Low": "low", "Moderate": "moderate", "Stronger": "stronger", "High": "high"}
LEVEL_NAMES = {
    1: "Level 1: Symbolic CAIO", 2: "Level 2: Advisory CAIO", 3: "Level 3: Coordinating CAIO",
    4: "Level 4: Governing CAIO", 5: "Level 5: Enterprise Orchestrator",
}


# ---------------------------------------------------------------- derived values

def _dr_auto(c: PostingCoding) -> str:
    dr = c.decision_rights
    if dr.high:
        return "High decision-rights authority"
    if dr.stronger:
        return "Stronger authority"
    if dr.moderate:
        return "Orchestration role, moderate authority"
    if dr.low:
        return "Low direct authority"
    return ""


def _suggested_level(c: PostingCoding) -> str:
    """Mirrors the 'Script Output V1' formulas in columns BG, BH, BI and BC."""
    breadth = sum(getattr(c.dimensions, k).score for k in DIM_LABELS.values())
    dr = c.decision_rights
    rank = 4 if dr.high else 3 if dr.stronger else 2 if dr.moderate else 1 if dr.low else 0
    span = sum(v for v in c.interacts.model_dump().values() if isinstance(v, bool))
    base = 1 if breadth <= 5 else 2 if breadth <= 11 else 3 if breadth <= 17 else 4 if breadth <= 23 else 5
    if base >= 4:
        if rank < 3:
            level = 3
        elif base == 5:
            level = 5 if span >= 6 and dr.high else 4
        else:
            level = 4
    else:
        level = base
    return LEVEL_NAMES[level]


def value_for(variable: str, c: PostingCoding):
    """The value a model's coding gives for one validation variable."""
    simple = {
        "Employment Type": c.employment_type,
        "Background Type": c.background_type,
        "Reporting Line": c.reporting_line,
        "Scope": c.scope,
        "Team Structure": c.team_structure,
        "Cybersecurity Aspect Score": c.cyber_score.score,
        "Orchestration Maturity Level": c.orchestration_maturity_level,
        "CAIO Archetype": c.caio_archetype,
        "Decision-Rights Strength (auto)": _dr_auto(c),
        "Decision-Rights Strength (Override — optional)": c.decision_rights_override or "",
        "Decision-Rights Strength (Final)": c.decision_rights_override or _dr_auto(c),
        "Suggested Maturity Level (auto)": _suggested_level(c),
    }
    if variable in simple:
        return simple[variable]
    if variable.endswith("Orchestration Score"):
        return getattr(c.dimensions, DIM_LABELS[variable.removesuffix(" Orchestration Score")]).score
    if variable.startswith("DR Language: "):
        tier = DR_TIERS[variable.split(": ")[1].split(" ")[0]]
        return "Yes" if getattr(c.decision_rights, tier) else "No"
    if variable == "Decision-Rights Language Breadth (0-4)":
        dr = c.decision_rights
        return sum([dr.low, dr.moderate, dr.stronger, dr.high])
    raise KeyError(f"No rule for validation variable '{variable}'.")


# ---------------------------------------------------------------- sheet layout

def _variables(ws) -> list[str]:
    """Variable names in column order, read from the 'Manual — X' headers."""
    return [
        str(cell.value).split(" — ", 1)[1]
        for cell in ws[1]
        if isinstance(cell.value, str) and cell.value.startswith("Manual — ")
    ]


def _group_width(ws) -> int:
    header = [c.value for c in ws[1]]
    return 5 if any(isinstance(h, str) and h.startswith("Script (Gemini) — ") for h in header) else 3


def _copy_cell(src, dst):
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format


def restructure_validation_sheet(wb):
    """Old groups: Manual | Script | match.  New: Manual | Gemini | Anthropic | GPT | match."""
    old = wb[VAL_SHEET]
    if _group_width(old) == 5:
        return  # already done

    variables = _variables(old)
    new = wb.create_sheet(VAL_SHEET + " new", index=wb.sheetnames.index(VAL_SHEET))
    last_row = old.max_row

    # fixed columns A-E
    for col in range(1, FIXED_COLS + 1):
        for r in range(1, last_row + 1):
            _copy_cell(old.cell(r, col), new.cell(r, col))
        letter = get_column_letter(col)
        new.column_dimensions[letter].width = old.column_dimensions[letter].width

    for i, var in enumerate(variables):
        old_manual = FIXED_COLS + 1 + 3 * i
        new_manual = FIXED_COLS + 1 + 5 * i
        width = old.column_dimensions[get_column_letter(old_manual)].width

        for r in range(1, last_row + 1):
            _copy_cell(old.cell(r, old_manual), new.cell(r, new_manual))
            _copy_cell(old.cell(r, old_manual + 1), new.cell(r, new_manual + 1))
            for offset in (2, 3):  # empty Anthropic / GPT cells, same look as the Gemini column
                _copy_cell(old.cell(r, old_manual + 1), new.cell(r, new_manual + offset))
                new.cell(r, new_manual + offset).value = None

        new.cell(1, new_manual + 1).value = f"Script (Gemini) — {var}"
        new.cell(1, new_manual + 2).value = f"Script (Anthropic) — {var}"
        new.cell(1, new_manual + 3).value = f"Script (GPT) — {var}"

        # match column: Manual vs Gemini (the old formulas pointed at the wrong cells)
        m, g = get_column_letter(new_manual), get_column_letter(new_manual + 1)
        for r in range(2, last_row + 1):
            if old.cell(r, 1).value:
                new.cell(r, new_manual + 4).value = f'=IF({m}{r}={g}{r},"Y","N")'

        for offset in range(4):
            new.column_dimensions[get_column_letter(new_manual + offset)].width = width

    # trailing columns (Match Method)
    old_tail = FIXED_COLS + 1 + 3 * len(variables)
    new_tail = FIXED_COLS + 1 + 5 * len(variables)
    for k in range(old.max_column - old_tail + 1):
        for r in range(1, last_row + 1):
            _copy_cell(old.cell(r, old_tail + k), new.cell(r, new_tail + k))
        new.column_dimensions[get_column_letter(new_tail + k)].width = (
            old.column_dimensions[get_column_letter(old_tail + k)].width
        )

    new.freeze_panes = old.freeze_panes
    wb.remove(old)
    new.title = VAL_SHEET


def model_col(variables: list[str], variable: str, which: str) -> int:
    offset = {"manual": 0, "gemini": 1, "anthropic": 2, "openai": 3}[which]
    return FIXED_COLS + 1 + 5 * variables.index(variable) + offset


# ---------------------------------------------------------------- posting lookup

def posting_rows(wb, postings: pd.DataFrame) -> dict[int, int]:
    """Validation row -> index into postings.csv, via 'Script Output V1'."""
    val, so = wb[VAL_SHEET], wb[SCRIPT_SHEET]
    variables = _variables(val)
    score_vars = [v for v in variables if v.endswith("Orchestration Score")]

    candidates: dict[tuple, list[int]] = {}
    for r in range(3, min(so.max_row, 2 + len(postings)) + 1):
        key = (str(so[f"E{r}"].value).strip(), str(so[f"J{r}"].value).strip())
        candidates.setdefault(key, []).append(r)

    lookup = {}
    for r in range(2, val.max_row + 1):
        if not val.cell(r, 1).value:
            continue
        key = (str(val.cell(r, 2).value).strip(), str(val.cell(r, 3).value).strip())
        rows = candidates.get(key)
        if not rows:
            print(f"  WARNING: no posting found for validation row {r} {key}")
            continue

        # duplicates: pick the one whose Gemini scores match this validation row
        def agreement(so_row):
            gem = [val.cell(r, model_col(variables, v, "gemini")).value for v in score_vars]
            out = [so.cell(so_row, 25 + i).value for i in range(10)]  # Y..AH
            return sum(str(a) == str(b) for a, b in zip(gem, out))

        lookup[r] = max(rows, key=agreement) - 3
    return lookup


# ---------------------------------------------------------------- results sheet

def _norm(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _pearson(xs, ys):
    pairs = [(float(x), float(y)) for x, y in zip(xs, ys)
             if _norm(x) != "" and _norm(y) != ""]
    if len(pairs) < 3:
        return None
    a, b = zip(*pairs)
    if statistics.pstdev(a) == 0 or statistics.pstdev(b) == 0:
        return None
    return statistics.correlation(a, b)


def compare(wb) -> dict:
    """{variable: {which: (matches, mismatches, rate, r)}}"""
    val = wb[VAL_SHEET]
    variables = _variables(val)
    rows = [r for r in range(2, val.max_row + 1) if val.cell(r, 1).value]
    out = {}
    for var in variables:
        out[var] = {}
        manual = [val.cell(r, model_col(variables, var, "manual")).value for r in rows]
        numeric = var.endswith("Score") or var.endswith("(0-4)")
        for which in ("gemini", "anthropic", "openai"):
            # a row counts once the model has coded it (Employment Type is never blank then)
            ran = [which == "gemini" or _norm(val.cell(r, model_col(variables, variables[0], which)).value) != ""
                   for r in rows]
            script = [val.cell(r, model_col(variables, var, which)).value for r in rows]
            scored = [(m, s) for m, s, ok in zip(manual, script, ran) if ok]
            if not scored:
                out[var][which] = None
                continue
            hits = sum(_norm(m) == _norm(s) for m, s in scored)
            r = _pearson(*zip(*scored)) if numeric else None
            out[var][which] = (hits, len(scored) - hits, hits / len(scored), r)
    return out


def write_results_sheet(wb, stats: dict):
    ws = wb[RESULTS_SHEET]
    header_style = ws["B1"]

    # Gemini numbers already in B-E stay as they are; just label them.
    for col, name in zip("BCDE", ["Exact Matches", "Mismatches", "Exact Match Rate", "Pearson r"]):
        ws[f"{col}1"] = f"{name} (Gemini)"

    # clear the old summary block in H:I
    for r in (2, 3):
        ws[f"H{r}"] = None
        ws[f"I{r}"] = None

    blocks = {"anthropic": "F", "openai": "J"}
    for which, start in blocks.items():
        label = PROVIDERS[which]
        first = ord(start) - ord("A") + 1
        for k, name in enumerate(["Exact Matches", "Mismatches", "Exact Match Rate", "Pearson r"]):
            cell = ws.cell(1, first + k, f"{name} ({label})")
            _copy_style(header_style, cell)
            ws.column_dimensions[get_column_letter(first + k)].width = 20

        for r in range(2, ws.max_row + 1):
            var = ws.cell(r, 1).value
            s = stats.get(var, {}).get(which) if var else None
            for k in range(4):
                ws.cell(r, first + k).value = None if s is None else s[k]
            ws.cell(r, first + 2).number_format = ws.cell(r, 4).number_format

    # summary: averages for all three models
    last = max(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value)
    ws["O1"], ws["P1"], ws["Q1"], ws["R1"] = "Summary", "Gemini", "Anthropic", "GPT"
    for col in "OPQR":
        _copy_style(header_style, ws[f"{col}1"])
    ws["O2"], ws["O3"] = "Average Exact Match Rate", "Average Pearson r correlation"
    for col, rate_col, r_col in (("P", "D", "E"), ("Q", "H", "I"), ("R", "L", "M")):
        ws[f"{col}2"] = f"=AVERAGE({rate_col}2:{rate_col}{last})"
        ws[f"{col}3"] = f"=AVERAGE({r_col}2:{r_col}{last})"
    ws.column_dimensions["O"].width = 30


def _copy_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)


# ---------------------------------------------------------------- main

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", default="data/CAIO Dataset.xlsx")
    parser.add_argument("--postings", default="data/postings.csv")
    parser.add_argument("--prompt", default="prompts/V1.txt",
                        help="Prompt file; V1 is the one the Gemini results were produced with.")
    parser.add_argument("--providers", nargs="+", default=list(PROVIDERS), choices=list(PROVIDERS))
    parser.add_argument("--workers", type=int, default=6,
                        help="How many postings to score at the same time.")
    parser.add_argument("--layout-only", action="store_true",
                        help="Only rename/add the columns; make no API calls.")
    args = parser.parse_args()

    load_dotenv(ROOT / ".env")
    cfg = yaml.safe_load((ROOT / "config.yaml").read_text(encoding="utf-8"))
    prompt_text = (ROOT / args.prompt).read_text(encoding="utf-8")
    postings = pd.read_csv(ROOT / args.postings).fillna("")

    wb_path = ROOT / args.workbook
    wb = load_workbook(wb_path)
    restructure_validation_sheet(wb)

    if args.layout_only:
        wb.save(wb_path)
        print(f"Columns updated in {wb_path}")
        return

    val = wb[VAL_SHEET]
    variables = _variables(val)
    lookup = posting_rows(wb, postings)
    failures = []

    for provider in args.providers:
        label = PROVIDERS[provider]
        model = cfg.get(f"{provider}_model")
        cache_path = ROOT / "cache" / f"extractions_{provider}.json"
        cache = load_cache(cache_path)
        print("=" * 60)
        print(f"{label} ({model}) — {len(lookup)} postings")
        print("=" * 60)

        jobs = {}
        for val_row, idx in lookup.items():
            posting = postings.iloc[idx].to_dict()
            key = cache_key(f"{provider}:{model}:{args.prompt}", prompt_text, posting)
            jobs[val_row] = (posting, key)

        todo = {r: j for r, j in jobs.items() if j[1] not in cache}
        print(f"{len(jobs) - len(todo)} cached, {len(todo)} to score")

        # several postings at once; cache is saved after each one finishes
        with ThreadPoolExecutor(max_workers=args.workers) as pool:
            futures = {
                pool.submit(get_coding, posting, prompt_text, cfg, provider=provider): (posting, key)
                for posting, key in todo.values()
            }
            for n, fut in enumerate(as_completed(futures), 1):
                posting, key = futures[fut]
                try:
                    cache[key] = fut.result()
                    save_cache(cache_path, cache)
                    print(f"[{n}/{len(todo)}] {posting['company']} — {posting['title']}")
                except Exception as e:  # keep going; report at the end
                    print(f"[{n}/{len(todo)}] ERROR {posting['company']}: {e}")
                    failures.append((label, posting["company"], str(e)))

        for val_row, (posting, key) in jobs.items():
            if key not in cache:
                continue
            coding = PostingCoding.model_validate(cache[key])
            for var in variables:
                val.cell(val_row, model_col(variables, var, provider)).value = value_for(var, coding)

    stats = compare(wb)
    write_results_sheet(wb, stats)
    wb.save(wb_path)

    print()
    print(f"{'Variable':<58} {'Gemini':>8} {'Anthropic':>10} {'GPT':>8}")
    for var, s in stats.items():
        cells = [f"{s[w][2]:.0%}" if s.get(w) else "-" for w in ("gemini", "anthropic", "openai")]
        print(f"{var[:58]:<58} {cells[0]:>8} {cells[1]:>10} {cells[2]:>8}")
    for w, name in (("gemini", "Gemini"), ("anthropic", "Anthropic"), ("openai", "GPT")):
        rates = [s[w][2] for s in stats.values() if s.get(w)]
        if rates:
            print(f"Average exact match rate — {name}: {sum(rates) / len(rates):.1%}")
    if failures:
        print(f"\n{len(failures)} posting(s) failed:")
        for f in failures:
            print("  ", *f)
    print(f"\nSaved: {wb_path}")


if __name__ == "__main__":
    main()
