# caio-coder

Automated coding of Chief AI Officer (CAIO) job postings for the *CAIO as AI Capability Orchestrator* research project.
It reads each posting with an AI model, scores it against the research codebook (orchestration dimensions, decision-rights language, cybersecurity aspect, stakeholder span), and writes one row per posting into the CAIO Data Excel workbook.

## How it works

```mermaid
flowchart LR
    RAW["Raw scraped file<br/>.xlsx, many columns"] -->|"src/make_slim.py<br/>(run once)"| CSV["data/postings.csv<br/>url, company, title,<br/>location, date, description"]

    subgraph PIPE["python run.py  (loops over every posting)"]
        direction TB
        CACHE{"Already<br/>in cache/?"}
        EXT["src/extract.py<br/>AI call<br/>prompts/v1 + schema"]
        RESP["Raw JSON answer<br/>saved to cache/"]
        SCHEMA["Schema check<br/>src/schema.py<br/>values must match<br/>the Excel dropdowns"]
        VAL["Evidence check<br/>src/validate.py<br/>every quote must appear<br/>in the posting text"]
        DER["src/derive.py<br/>maturity level + archetype<br/>computed in code"]
        XL["src/write_excel.py<br/>append row +<br/>copy formulas down"]

        CACHE -- no --> EXT --> RESP
        CACHE -- yes --> RESP
        RESP --> SCHEMA --> VAL --> DER --> XL
    end

    CSV --> CACHE
    TEMPLATE["CAIO Data workbook<br/>(template)"] --> XL
    XL --> OUT["output/results_v1.xlsx"]
    SCHEMA -. "invalid answer" .-> ERR["output/errors_v1.log"]
```

**Design decisions**

| Decision | Why |
|---|---|
| The AI only *codes inputs* (scores, Yes/No flags, categories). | Excel's own formulas compute the breadth score, decision-rights strength, span count and suggested maturity, so the sheet stays the single source of truth. |
| Answers are forced into a schema whose allowed values are copied from the workbook dropdowns. | The model cannot return a category Excel would reject. |
| Every non-zero score needs a verbatim quote, and the quote is checked against the posting. | Cheap guard against invented evidence; failures are flagged in *General Notes / Flags*. |
| Maturity level and archetype are calculated in code. | Reproducible and consistent with the sheet's formulas; the archetype rule is a v1 draft in `src/derive.py`. |
| Raw model answers are cached per prompt version; finished URLs are skipped. | Safe to stop and rerun, and re-running after a code change costs no new AI calls. |
| Prompts are versioned files (`prompts/v1`, `v2`, ...). | Each run records which instructions produced which results. |

## Status

- [x] Repo, schema, workbook column map
- [x] Prompt v1 (from the research plan rubric)
- [x] Slim postings file builder
- [x] Validation, derived fields, Excel writer, resumable runner
- [x] End-to-end test with fake AI answers
- [x] Real AI call in `src/extract.py` (waiting on API key and provider)
- [x] Full-dataset run
- [ ] Iterate on the prompt (v2, v3, ...) using real results
