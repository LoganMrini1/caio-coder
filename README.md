# caio-coder
Automated coding of CAIO job postings into the CAIO Data workbook.

Pipeline: links -> posting text -> LLM (structured JSON) -> validate -> derive -> Excel row.

Build order:
1. [x] Repo skeleton, schema (`src/schema.py`), workbook column map (`src/excel_map.py`)
2. [ ] Codebook + prompt v1 (`codebook/`, `prompts/v1.md`)
3. [ ] `ingest.py` / `fetch.py` (raw file descriptions first, web fetch fallback)
4. [ ] `extract.py` (Claude API, schema-constrained)
5. [ ] `validate.py` (schema + quote-in-text check)
6. [ ] `derive.py` (maturity level + archetype, mirroring workbook formulas)
7. [ ] `write_excel.py`

Setup: `python -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt`, copy `.env.example` to `.env`.
