## Title
POC Grounded Invoice-Contract Validation Chatbot (MPAN-Based)

## Summary
Build an Azure-hosted POC chatbot where users can upload invoice PDFs and contract Excel files, ask free-form validation questions, and get:
- `Pass/Fail + Reasons`
- a `Validation Score (0-100)` with bands: Green `>=95`, Amber `80-94`, Red `<80`
- strict, source-grounded answers only (no hallucinated/general answers)

The bot validates by MPAN, supports both single-rate and day/night-rate contracts, and uses only:
- `meters.data` (mapping)
- `half-hour.data` (day/night tariff cases)
- `day.data` (single-rate cases)

## Key Implementation Changes
- Ingestion and normalization
  - Parse contract Excel into normalized MPAN records (overwrite existing MPAN, insert if new).
  - Parse invoice PDF via Azure Document Intelligence into structured fields and line items.
  - Parse meter files and build mapping from `meters.data` bracket value (`last4`) to MPAN last 4 digits.
- Validation engine
  - Match invoice line MPAN -> contract MPAN.
  - If contract has day/night rates, compute consumption split from `half-hour.data` using contract rate labels/time definitions.
  - If contract has single rate, validate against `day.data` totals.
  - Strict exact match on monetary/rate checks; return deterministic reason codes for every mismatch.
- Scoring model (weighted 100-point)
  - Start at 100 and subtract penalties by mismatch class (rate, standing charge, usage, period, missing mapping, missing source data).
  - Return score + band + reasoned deductions so customer can audit how score was produced.
- Chatbot behavior
  - Grounded Q&A only over uploaded/parsed artifacts (contract/invoice/meter).
  - Every answer includes evidence references (source file + extracted section/field).
  - If answer cannot be grounded, bot responds with explicit “insufficient evidence”.

## Azure Credentials/Config Needed From Customer
- Azure Document Intelligence
  - `DOCUMENT_INTELLIGENCE_ENDPOINT`
  - `DOCUMENT_INTELLIGENCE_API_KEY`
- Azure OpenAI (for chat + optional extraction post-processing)
  - `AZURE_OPENAI_ENDPOINT`
  - `AZURE_OPENAI_API_KEY`
  - `AZURE_OPENAI_CHAT_DEPLOYMENT`
  - `AZURE_OPENAI_API_VERSION`
- Azure Blob Storage
  - `AZURE_STORAGE_ACCOUNT` + `AZURE_STORAGE_KEY` (or SAS/connection string)
  - Container names for `raw-uploads`, `normalized-data`, `validation-results`
- App hosting/runtime
  - Target Azure App Service/Container endpoint
  - Allowed CORS origins for frontend
  - Environment name (`dev/poc`) and logging destination

## Test Plan
- Unit tests
  - MPAN normalization/matching (full MPAN + last4 mapping).
  - Contract parser for both provided Excel formats.
  - Meter parsers for `meters.data`, `half-hour.data`, `day.data`.
  - Score deduction logic and score-band classification.
- Integration tests
  - Upload invoice -> extract -> validate -> score output end-to-end.
  - Contract upsert overwrite behavior for same MPAN changes.
  - Day/night tariff invoice validated using half-hour aggregation.
  - Single-rate tariff invoice validated using day totals.
- Acceptance scenarios
  - Customer asks natural-language validation questions and receives grounded, evidence-linked responses.
  - Missing meter data for an MPAN is surfaced clearly without crashing pipeline.

## Assumptions
- This is a POC, so overwrite-in-place contract updates are acceptable (no version history yet).
- Some MPANs will not have meter data; validation still runs with explicit missing-data reasons.
- Day/night window definitions come from contract labels/config where present.
