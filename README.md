<p align="center">
  <img src="docs/icon.png" alt="SAP SF Permission Comparer" width="128">
</p>

<h1 align="center">SAP SF Permission Comparer</h1>

<p align="center">A web tool for comparing SAP SuccessFactors role-based permissions across environments and formats.</p>

**Live:** [sap-comparer.openclaw.vvlasy.cz](https://sap-comparer.openclaw.vvlasy.cz)

## Features

### Compare T3 vs PROD
Upload two PDF permission exports (T3/pre-prod and PROD) and get a structured diff showing:
- Permission differences by section
- Values changed between environments
- Raw line-level diff

### Compare PDF vs Excel Workbook
Cross-reference a PDF role export against an Excel permission matrix:
- Auto-detects role name from PDF filename
- Shows entries only in PDF, only in Excel, and mismatches
- Expandable matched values view
- One-click "Set to None" for unmatched Excel entries with modified file download

### Generate Excel from PDF
Extract all permissions from a PDF role export into a clean Excel spreadsheet.

## Tech Stack

- **Backend:** Python 3.12, FastAPI, pdfplumber, openpyxl
- **Frontend:** Vanilla HTML/CSS/JS (single page, no build step)
- **Deployment:** Docker, GitHub Actions CI/CD, K8s (Traefik ingress)

## Running Locally

```bash
pip install -r requirements.txt
uvicorn server:app --host 0.0.0.0 --port 8000
```

Then open [http://localhost:8000](http://localhost:8000).

## Docker

```bash
docker build -t sap-comparer .
docker run -p 8000:8000 sap-comparer
```

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `/` | Frontend SPA |
| `GET` | `/api/info` | Build metadata |
| `POST` | `/api/compare` | Compare two PDF exports (T3 vs PROD) |
| `POST` | `/api/compare-pdf-excel` | Compare PDF against Excel workbook |
| `POST` | `/api/modify-excel` | Set specified Excel cells to "None" and download |
| `POST` | `/api/pdf-to-excel` | Generate Excel from PDF permission export |
