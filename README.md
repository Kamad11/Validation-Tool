# Bill Validator Local POC

This project runs a local chatbot + validation engine for electricity invoices.

## What It Does
- Option to load existing bundled contract files from the project root.
- Upload contract Excel files and upsert by `MPAN` (overwrite existing, insert new).
- Upload invoice PDFs, extract readable fields and line items.
- Hybrid invoice parsing:
  - Azure Document Intelligence (`prebuilt-invoice`) when available
  - `pypdf` extraction as fallback/alternative
  - automatic quality-based selection of parsed result
- Persist full extracted invoice text for grounded Q&A.
- Validate invoice vs contract with strict exact checks.
- Validate using invoice date range (supports variable periods like 24, 31, 92 days).
- Validate energy line costs (invoice GBP vs expected GBP from contract rate x invoice kWh).
- Invoice rate parsing is UOM-aware (`£/kWh` vs `p/kWh`) and normalized before comparison.
- Validation uses explicit tolerances:
  - Rate tolerance: `0.0001 GBP` (`0.01p`)
  - Money tolerance: `0.02 GBP`
  - Meter tolerance (percent): default `2.0%`, user-editable in UI
- Cross-check meter consumption using:
  - `Meter Data/meters.data` (MPAN-last4 mapping)
  - `Meter Data/half-hour.data` (day/night tariffs)
  - `Meter Data/day.data` (single-rate tariffs)
- For day/night validations, half-hour records are aggregated across the full 24-hour period from the invoice date range.
- Return `PASS/FAIL`, reason codes, evidence, and weighted score:
  - Green: >=95
  - Amber: 80-94
  - Red: <80
- Validation now also returns MPAN-level cost summaries:
  - Invoice energy total
  - Expected from contract x invoice usage
  - Expected from contract x meter usage
- Grounded chat answers only from uploaded/parsed evidence.
- Chat can answer from the full stored invoice content, with citations to invoice text chunks.
- UI shows separate score/status cards for:
  - Contract vs Invoice
  - Invoice vs Meter
- Detailed comparison values open on demand in a full-screen table modal.
- Chat opens from a floating toggle button and supports collapsible references.
- Validation can be run with meter comparison enabled/disabled from the UI toggle.
- Validation output includes a meter-data note indicating whether Wh->kWh normalization was applied.

## Run Locally
### Option A: Standard local Python (recommended for any machine)

```powershell
cd "C:\Users\kamad\OneDrive\Desktop\validation tool"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app\server.py
```

Then open:
- http://127.0.0.1:8000

### Option B: Bundled runtime used in this environment

```powershell
& 'C:\Users\kamad\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' app\server.py
```

Then open:
- http://127.0.0.1:8000

### Run smoke test

```powershell
python tests\smoke_test.py
```

## Optional Environment Variables
- `HOST` (default `127.0.0.1`)
- `PORT` (default `8000`)
- `AZURE_OPENAI_ENDPOINT` (optional override; hardcoded default is present in `app/server.py`)
- `AZURE_OPENAI_API_KEY` (for grounded chat completion)
- `AZURE_OPENAI_DEPLOYMENT` (optional override; hardcoded default is `gpt-5-mini`)
- `AZURE_OPENAI_API_VERSION` (optional override; hardcoded default is `2024-10-21`)
- `DOCUMENT_INTELLIGENCE_ENDPOINT` (optional override; hardcoded default in `app/server.py`)
- `DOCUMENT_INTELLIGENCE_API_KEY` (optional override; hardcoded default in `app/server.py`)
- `DOCUMENT_INTELLIGENCE_MODEL` (optional override; hardcoded default is `prebuilt-invoice`)
- `DOCUMENT_INTELLIGENCE_API_VERSION` (optional override; hardcoded default is `2024-11-30`)

Minimum required (PowerShell, current session only):
```powershell
$env:AZURE_OPENAI_API_KEY = "<your-api-key>"
```

## API Endpoints
- `POST /api/contracts/load-defaults` (loads known local contract files in project root)
- `POST /api/contracts/upsert` (multipart, `file`)
- `POST /api/invoices/parse` (multipart, `file`)
- `POST /api/invoices/validate` (multipart, `file` or json with `invoice_number`)
- `POST /api/chat` (`question`, optional `invoice_number`)
- `GET /api/state`

## Notes
- Chat is evidence-grounded using stored invoice/contract/validation context.
- If Azure chat fails, the response includes a short `Diagnostic` line to help identify config/connectivity issues.
- App is Python 3.13+ compatible (no runtime `cgi` dependency).

## Push To GitHub

```powershell
cd "C:\Users\kamad\OneDrive\Desktop\validation tool"
git init
git add .
git commit -m "Add requirements and local run instructions"
git branch -M main
git remote add origin https://github.com/Kamad11/Validation-Tool.git
git push -u origin main
```
