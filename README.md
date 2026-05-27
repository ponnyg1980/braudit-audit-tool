# Braudit Audit Tool

Internal Streamlit tool for The Trademark Helpline / Braudit. Runs Steps 2–5
of the audit pipeline (de-duplicate, exclude, score, generate report) over a
scraped-results spreadsheet and produces a Word-document monitoring report.

> **v1 scope.** No Step 6 (forensic appendix) in this version — that comes in v2
> once it's wired to the Anthropic API. Step 6 is the only stage that needs an
> LLM, so v1 has zero API costs and is safe to run as much as you like.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open http://localhost:8501. The default development password is `braudit-dev`
(set in `.streamlit/secrets.toml`).

## Project layout

```
braudit_audit_tool/
├── app.py                       # Streamlit UI
├── requirements.txt
├── README.md
├── DEPLOY.md                    # Step-by-step deployment guide
├── .streamlit/
│   ├── config.toml              # Braudit-blue theme
│   └── secrets.toml.example     # Template for password / future API keys
└── pipeline/
    ├── filters.py               # Steps 2–4: dedupe, exclude, score
    └── report_builder.py        # Step 5: build the docx
```

## How it works

1. **Upload** — user uploads the `.xlsx` produced by the Braudit scrape job.
2. **Form** — user supplies client name, contact, account manager, preparer,
   and the search criteria (exact / similar match, classes, SIC).
3. **Run Audit** — the pipeline:
   - Reads the Trademarks / Companies / Google / Domains / Social sheets
   - Deduplicates by natural key (App #, Company #, URL)
   - Drops records whose mark text isn't `ROOT` or `ROOT ` + descriptor
   - Drops trademark records that don't touch any of the client's classes
   - Drops Companies House rows that don't include the client's SIC
   - Scores each surviving trademark on a 0–13 rubric (status × mark × type × class overlap)
   - Buckets into Negligible / Low / Medium / High Risk
   - Renders the Braudit-style Word document
4. **Download** — user downloads the `.docx`.

## What the tool does NOT do

- It does not call any LLM. Nothing leaves your environment.
- It does not store the uploaded file once the session ends.
- It does not (yet) do Step 6 (forensic verification of each record against
  the source USPTO / UKIPO / EUIPO register). Step 6 is the v2 lift.

## Maintaining the tool

The two files you'd most likely want to tweak are:

- `pipeline/filters.py` — the scoring rubric and the exclusion rules
- `pipeline/report_builder.py` — the docx layout, colours, table structure

Anything UI-related lives in `app.py`.

If you want to change the methodology (e.g. drop dead marks instead of tagging
them Negligible, or widen the mark-scope rule), edit `filters.py` and redeploy.
Streamlit Cloud auto-deploys on every git push.

---

*v1.0 · 21 May 2026*
