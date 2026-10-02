# HSE Performance Dashboard — Live OneDrive Rev9

This package uses the approved Rev40 dashboard as the visual baseline.

Architecture:
OneDrive Excel (Anyone / Can view) -> GitHub Actions server-side download -> data.json -> GitHub Pages.

The browser does NOT fetch the OneDrive sharing URL directly, avoiding the CORS problem confirmed by the connection test.

## Required GitHub secret

Create repository secret:

`ONEDRIVE_EXCEL_URL`

Value:

the current OneDrive `Anyone -> Can view` Excel sharing URL.

## Workflow

`refresh-dashboard.yml` runs every 15 minutes and can also be started manually from GitHub Actions.

It downloads the workbook server-side, parses the existing workbook structure, creates `data.json`, and deploys the dashboard to GitHub Pages.

## Important

GitHub Pages is public. Do not place confidential raw personnel/incident data in the repository or generated `data.json` unless that exposure is acceptable. The dashboard itself only needs the structured data required for its displayed KPIs/charts.
