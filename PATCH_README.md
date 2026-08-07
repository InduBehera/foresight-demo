# Foresight V6.2 — Contextual Ask AI & Scenario Copilot

## Scope

This update changes the Ask AI experience for returning users. It does not modify the V24 Model Builder or the existing leadership summaries.

## What changed

- Replaced visible guidance-workbook questions with questions a country forecaster is more likely to ask while reviewing the active forecast.
- Added page-specific prompt groups for Home, Cycle Readiness, Forecast Cycle Comparison, Overview, Patient Potential Tower, Market Landscape, Product Performance, Scenario Builder, Saved Scenarios, Scenario Comparison, Explore and Reports.
- Added active context to Ask AI: forecast cycle, page, output metric, severity, focus year and scenario mode.
- Added structured answers using Finding, Evidence, Driver Interpretation, Forecasting Caution and Recommended Next Action.
- Added staged AI simulation actions for:
  - Diagnosis +2 percentage points
  - Product X access +5 percentage points
  - Three-point Product Y-to-Product X share transfer
  - Competitor entry / erosion one year earlier
  - An illustrative balanced $6.5B goal-seek option
- AI stages assumptions in the Scenario Simulator but does not run or apply them automatically.
- Added 100% market-share reconciliation guidance and explicit separation of patient-pool expansion, product share and access.
- Added submission-readiness Ask AI action to the Validation drawer.

## Governance

- Original Approved Base remains locked.
- Saved scenarios remain separate from the Current Cycle Draft.
- AI-proposed changes require user review and an explicit Run Simulation action.
- Scenario outputs are calculated by the model, not typed by the AI.

## Deployment

Upload the contents of this folder to the GitHub repository root. `index.html` already loads `model-builder.html?embedded=1` from the same root.
