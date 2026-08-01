# RSS Plan Builder · Streamlit

An interactive Streamlit application that guides Qualified Persons through a
BCA-aligned Remote Site Supervision Plan and generates a polished native
Microsoft Word report.

## Features

- Eight-step workflow from project definition to QP(S) sign-off
- Flat supervision-activity selector: individual structural-checklist
  activities plus bored tunnelling, material-test, fabrication-yard and
  project-specific activities from the older Site Supervision Plan guide
- Activity suitability matrix with alternative-approach justification
- Multiple implementation phases within the same activity, including progress
  ranges, RSS extent, parallel supervision, phase gates and QP(S) review points
- Organisation-chart image upload and structured, evidence-backed competency
  verification
- Reusable people, technology, control and record profiles; activities refer to
  a selected profile and record only activity-specific requirements or
  variations
- Project-wide phasing, people, technology, controls and records remain in the
  BCA plan sections without repeating the same prose in every activity
- Readiness checks before export
- Native `.docx` report generated from the supplied BCA template, plus editable
  `.json` working-file downloads
- Optional overall site-plan image

## Run locally

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

Node.js is not required for the Streamlit version.

## Deploy on Streamlit Community Cloud

1. Push this folder to a GitHub repository.
2. Sign in to Streamlit Community Cloud.
3. Create an app and select `app.py` as the entry point.
4. Deploy. No secrets or external services are required.

## Important

The builder follows the structure and decision logic in the Guidebook for
Remote Site Supervision Version 2.0 (June 2026) and references the Guide Book
for Site Supervision Plan Version 1.1 (October 2023). It supports preparation
but does not replace the QP(S)'s professional judgement, statutory duties,
review of approved plans, or verification of the latest BCA requirements.
