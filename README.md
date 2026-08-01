# RSS Plan Builder · Streamlit

An interactive Streamlit application that guides Qualified Persons through a
BCA-aligned Remote Site Supervision Plan and generates a polished native
Microsoft Word report.

## Features

- Eight-step workflow from project definition to QP(S) sign-off
- Activity suitability matrix with alternative-approach justification
- Phased implementation guidance (15% / 30% / 50%)
- People, technology, contingency, quality and records sections
- Readiness checks before export
- Native `.docx` report and editable `.json` working-file downloads
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

