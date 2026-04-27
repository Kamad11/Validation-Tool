## Title
Invoice Validation Tool - Current Plan (Local POC)

## Current State
- Local-first single-process app (`app/server.py`) with static frontend (`static/index.html`).
- Invoice parsing via `pypdf` (not Azure Document Intelligence in current implementation).
- Contract import from Excel (`openpyxl`), validation by MPAN, optional meter comparison.
- Grounded chat over parsed invoice/contract/validation evidence using Azure OpenAI chat completions.
- Azure defaults are hardcoded for endpoint/deployment/api-version; API key remains runtime-configurable.

## Implemented
- Separate comparison flows:
  - Contract vs Invoice
  - Invoice vs Meter
- Separate comparison status/score cards in UI.
- Full-screen comparison modal opened on demand from score cards.
- Meter tolerance input (percent) in UI, passed to backend validation.
- Meter fallback behavior:
  - direct comparison when data exists
  - prediction marker when direct data is insufficient
  - explicit unavailable messaging when prediction is not possible
- Floating chat toggle button with compact chat panel.
- Chat message tiles (user/bot), collapsible references, formatted bot response rendering.
- Invoice address extraction fields for Q&A:
  - `invoice_supply_address`
  - `invoice_billing_address`

## Immediate Next Steps
- Add optional `.env` loading for local API key convenience (avoid repeated shell export).
- Add a lightweight `/api/health/azure` endpoint to validate Azure connectivity/deployment quickly.
- Add regression checks for:
  - comparison modal rendering
  - meter tolerance input propagation
  - chat formatting and references collapse behavior
- Add basic guardrails to avoid accidental API-key commits.

## Operational Notes
- Only required runtime secret for chat is `AZURE_OPENAI_API_KEY` (unless user overrides other Azure env vars).
- Keep data artifacts (`data_store/*`, `temp_uploads/*`, `__pycache__`) out of source commits.
