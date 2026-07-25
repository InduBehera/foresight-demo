# Foresight V24 — Deployable Forecasting Demo

A static demonstration web app for the Foresight AI-enabled pharma forecasting platform.

## What is fixed in V24

- **Generate Model works in a hosted static site**. Excel generation runs in the browser and downloads a formula-wired `.xlsx` workbook. No localhost service or Python backend is required.
- Base, Strong, and Weak EPI lines are visually separated for the demo and show all three values on hover.
- Epidemiology trend charts show Prevalence, Diagnosed, Treated, and Eligible values on hover.
- **Patient Flow / Patient Funnel** is renamed **Patient Potential Tower** throughout the visible interface.
- The existing EPI and Market workspaces, Scenario Simulators, Home, Model Library, Upload flow, Ask AI, Sources, and Validation are retained.

## Project structure

```text
index.html
vercel.json
.nojekyll
assets/
  js/v24-fixes.js
  vendor/jszip.min.js
```

## Run locally

From the project root:

```bash
python -m http.server 8080
```

Open `http://localhost:8080`.

Use a local web server rather than double-clicking `index.html`; browsers restrict some file-download and asset behaviours under `file://`.

## Deploy to Vercel

### Git-connected deployment

1. Push the contents of this folder to the root of a GitHub repository.
2. Import the repository in Vercel.
3. Use **Other** / static project settings. No build command is required and the project root is the output.
4. Deploy. Future pushes to the connected repository create new deployments.

### Vercel CLI

From this folder:

```bash
vercel
vercel --prod
```

Official documentation:
- https://vercel.com/docs/deployments/overview
- https://vercel.com/docs/cli/deploy

## Deploy to GitHub Pages

1. Push this folder to a branch.
2. In the repository, open **Settings → Pages**.
3. Select **Deploy from a branch**, then choose the branch and `/(root)`.
4. Save.

Official documentation:
- https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site

## Excel generation

The browser creates a valid `.xlsx` workbook using the packaged Excel ZIP writer. The EPI model contains:

- Model Information
- Model Architecture
- Input Register
- Assumptions
- Base / Strong / Weak forecast outputs
- Patient Potential Tower
- Validation

The Market model contains:

- Model Setup
- Market Inputs
- Baseline and Event-Adjusted outputs
- SKU Structure
- Validation

Inputs are visually distinguished from formulas in the workbook. The generated model is a governed draft and should be reviewed before business use.

## Public deployment note

This public-safe package excludes the uploaded source workbooks and disables source-workbook download. Excel model generation and scenario exports remain functional.
