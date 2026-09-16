# InfraPilot AI

InfraPilot AI is an on-premises AI infrastructure planning and GPU sizing tool. It converts workload requirements into a model recommendation, hardware mapping, three infrastructure options, a bill of materials, and budgetary financial estimates.

This repository contains the current management-demo release. It is designed to make the sizing logic understandable to infrastructure, AI platform, finance, and management stakeholders while keeping the catalog data editable in Excel.

## What It Does

InfraPilot AI follows this planning chain:

```text
Workload inputs
    -> workload rule and context demand
    -> eligible model shortlist
    -> selected model and GPU/server mapping
    -> minimum, moderate, and heavy infrastructure options
    -> BOM and financial planning summary
```

The current edition is focused on on-premises infrastructure. Cloud-specific pricing and cloud deployment options are planned as a future extension.

## Supported Features

### Workload and technical inputs

- Workload type: RAG/inference, embedding generation, fine-tuning, training, batch inference, forecasting, predictive ML, and multimodal workloads.
- Query complexity, response detail, conversation length, retrieval depth, and agentic complexity.
- Tool calls per request and language coverage.
- Knowledge-base size in TB.
- Peak concurrent sessions and peak requests per second.
- Required P95 TTFT, meaning P95 time to first token.
- Required generation speed in tokens per second per session.
- Precision preference.
- Availability, HA, DR, planning period, workload growth, and data growth.

### Model selection

- Derives a workload rule from the workload profile.
- Calculates context and token demand.
- Filters eligible models using context-window and language requirements.
- Shows the automatic candidate model.
- Allows the user to manually select another eligible model.
- Recalculates the complete hardware recommendation after a model change.

### Hardware and BOM planning

- Uses VRAM, throughput, prefill, decode, concurrency, and tensor-parallel capacity drivers.
- Rounds GPU quantities to complete supported server topologies.
- Shows three options:
  - Minimum Memory-Fit
  - Moderate Supported
  - Heavy Supported
- Allows installed GPU quantities to be tuned before final financial review.
- Recalculates VRAM, servers, HA, DR, and hardware totals from edited GPU quantities.

### Financial decision controls

Users can tune the following planning assumptions before reviewing the final financials:

- Server unit price
- Storage planning cost per TB
- Hot-tier NVMe cost per TB
- Network switch cost
- Software and support allowance
- Contingency percentage

The app clearly labels whether the scenario is the workbook baseline or a user-tuned scenario.

### Projects and catalogs

- Save projects locally in the browser.
- Load saved projects.
- Export and import project JSON.
- View catalog version used by the calculation.
- Add manual catalog records for review.
- Import and export manual catalog records.
- Maintain the official catalog in a separate Excel workbook.

## Repository Structure

```text
index.html                         Main application page
styles.css                         Application styling and responsive layout
app.js                             Calculator, model selection, BOM, and financial logic
Forge_Catalog_Reference_v1.xlsx    Human-maintained catalog workbook
AI_Sizing_infraSizing_*.xlsx       Original baseline workbook and historical reference
catalogs/                          Generated browser-friendly catalog snapshots
scripts/export-catalogs.py         Excel-to-JSON catalog exporter
scripts/create-catalog-workbook.py Creates the standalone catalog workbook
sync-catalogs.ps1                  One-command catalog synchronization
README.md                          This guide
```

## Requirements

- Windows, macOS, or Linux
- A modern browser: Edge, Chrome, Firefox, or Safari
- Python 3.10 or newer for catalog synchronization
- Python package: `openpyxl`
- Git, if pushing the project to GitHub

The app itself has no JavaScript package installation and no build step.

## Run Locally

### Fastest option

Open `index.html` in a browser.

### Recommended option

Run a local HTTP server from the project directory:

```powershell
py -m http.server 8080
```

Open:

```text
http://localhost:8080
```

Using HTTP is recommended because the app can then read `catalogs/manifest.json` and display the active catalog version reliably.

## Catalog and Excel Workflow

The Excel workbook is the human-maintained source of truth. The browser uses generated JSON snapshots because browsers should not execute `.xlsx` formulas directly at runtime.

### Update the official catalogs

1. Open `Forge_Catalog_Reference_v1.xlsx`.
2. Update the four catalog sheets:
   - `Model_Mapping_Catalog`
   - `Model_Selection`
   - `Hardware_Mapping_Catalog`
   - `GPU_Server_Selection`
3. Replace yellow `USER-DEFINED` rows only with source-backed data.
4. Include validation notes, source references, and stable IDs.
5. Save and close the workbook.
6. Synchronize the browser catalog files:

```powershell
./sync-catalogs.ps1
```

7. Review the generated files in `catalogs/`.
8. Test the app locally before publishing.

The generated catalog files are:

- `catalogs/Model_Mapping_Catalog.json`
- `catalogs/Model_Selection.json`
- `catalogs/Hardware_Mapping_Catalog.json`
- `catalogs/GPU_Server_Selection.json`
- `catalogs/manifest.json`

## Excel Baseline Relationships

The calculator implements the main workbook chain:

```text
Input
  -> Model_Selection
  -> GPU_Server_Selection
  -> BOM
  -> Financials
```

Implemented relationships include:

- Query, output, history, retrieval, and tool token calculations
- Required context-window calculation
- Model-call and token-per-second demand
- Workload rule selection
- Candidate model selection
- Model context eligibility
- Model-to-GPU mapping
- Model weight, runtime, and KV-cache planning memory
- TTFT and generation-speed validation
- Throughput, prefill, decode, concurrency, and tensor-parallel minimums
- Maximum capacity-driver selection
- Complete-server topology rounding
- Workload growth
- HA server additions
- DR server quantities
- Installed GPU VRAM
- Hardware totals
- Future storage and RAG hot-tier storage
- Network switch quantities and cost
- Software/support allowance
- CAPEX, contingency, and total planning budget

The default reference scenario has been checked against the workbook cached values, including 40/52/60 installed GPUs, 10/13/15 base servers, 18/23/26 growth-adjusted servers, 19/24/27 production plus HA servers, 10/12/14 DR servers, and 29/36/41 total servers.

## Management Demo Guidance

Present the output as a **budgetary planning BOM**, not a final procurement quote.

Before procurement, review:

- `Data Required` values
- Proxy benchmark rows
- `Review Required` catalog records
- Benchmark evidence for TTFT and throughput
- Current supplier and Dell configuration quotes
- Server topology and networking compatibility
- Model licensing and production approval

The tool is intended to support decisions and scenario comparison. It does not replace benchmark testing, vendor configuration validation, or architecture review.

## Publishing

This is a static web app and can be deployed to:

- GitHub Pages
- Netlify
- Cloudflare Pages
- Vercel
- Any static web server

Publish the complete project root, including:

- `index.html`
- `styles.css`
- `app.js`
- `catalogs/`
- `Forge_Catalog_Reference_v1.xlsx`

The Excel workbook is not required by the browser at runtime, but keeping it in the repository preserves the editable source catalog for the team.

## GitHub Setup

From this project directory, run:

```powershell
git init
git branch -M main
git remote add origin https://github.com/MasyDeloitte/GPU_AI_Infra.git
git add -A
git status
git commit -m "Initial InfraPilot AI on-prem GPU calculator"
git push -u origin main
```

If the remote already exists, use:

```powershell
git remote set-url origin https://github.com/MasyDeloitte/GPU_AI_Infra.git
git add -A
git commit -m "Update InfraPilot AI calculator"
git push origin main
```

The `.gitignore` excludes Python environments and temporary Excel lock files, but it does not exclude the real `.xlsx` workbooks. Therefore `git add -A` will include the Excel source files.

## Future Roadmap

### Catalog reliability

- Make the calculator consume the complete generated catalogs directly.
- Add catalog schema validation.
- Add duplicate-ID and missing-reference checks.
- Add catalog approval and promotion workflow.
- Add audit history for every catalog change.

### Calculation quality

- Add automated Excel-versus-browser fixture tests.
- Expand benchmark catalog coverage.
- Add benchmark evidence attachments and confidence tracking.
- Add model override reason and approval fields.
- Separate planning proxy results from production-certified results.

### Product and collaboration

- Add user accounts and project ownership.
- Add a backend database, preferably PostgreSQL for multi-user production use.
- Add organization sharing and permissions.
- Add project version history and scenario comparison.
- Add PDF and Excel BOM export.
- Add management approval workflow.

### Cloud edition

- Add AWS, Azure, and Google Cloud catalogs.
- Add cloud GPU instance mapping.
- Add region and availability-zone selection.
- Add cloud pricing and reserved-instance options.
- Compare on-prem CAPEX with cloud OPEX.
- Add hybrid deployment scenarios.

## Current Status

- Internal management demo: ready
- On-prem budgetary planning: usable
- Public production SaaS: future phase
- Procurement-certified sizing: requires benchmark and vendor validation

## License and Data Notice

Confirm the licensing and redistribution rights for model names, benchmark data, vendor hardware information, and pricing before publishing the repository publicly. Replace planning prices with approved supplier quotes for production decisions.
