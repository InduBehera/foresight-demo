# Foresight AD EPI Cycle Workspace v5

A self-contained static demo of the Foresight AI-enabled pharma forecasting platform, focused on **Atopic Dermatitis · Product X**, forecast cycles and governed scenario working.

## Deployable file

`index.html` contains the complete application, styling, data and JavaScript. It does not require a build step or backend.

## Included experience

- Forecasting Home / Cockpit
- Create, Import and Continue entry points
- Model Library with Original Approved Base and Forecast Cycles 1–4
- Cycle Workspace entry and readiness review
- Overview
- Patient Potential Tower
- Market Landscape
- Product Performance
- Scenario Simulator with 17 levers, visible Base values, saved scenarios and comparison
- Explore evidence workspace
- Contextual Ask AI questions and answers
- Reports and Data Dictionary
- Restored profile dropdown and account settings

## Profile menu

The top-right avatar opens:

- Profile
- Preferences
- Notification settings
- Help & data dictionary
- Sign out

## Update an existing GitHub repository

1. Unzip this package.
2. Copy the contents of `foresight-ad-epi-cycle-workspace-v5/` into the **repository root**.
3. Replace the existing `index.html`.
4. Keep or replace `vercel.json` with the included version.
5. Commit and push to the branch connected to Vercel.
6. Vercel should redeploy automatically after the push.

Example Git commands:

```bash
git add .
git commit -m "Update Foresight AD EPI cycle workspace to v5"
git push
```

## Deploy directly to Vercel

Import the GitHub repository as a static project.

- Framework preset: **Other**
- Build command: leave blank
- Output directory: leave blank
- Root directory: repository root

## Run locally

```bash
python -m http.server 8080
```

Open `http://localhost:8080`.

## Public deployment note

The package does not contain the confidential source forecasting workbook. It contains demo data and workbook-derived interface values for the prototype only.
