# V24 Test Report

Validated in headless Chromium and checked with openpyxl.

## Passed

- Application loads with the V24 title.
- Browser Excel generator reports ready without a localhost service.
- EPI Generate Model downloads a valid `.xlsx` file.
- Generated EPI workbook opens and contains:
  - Model Information
  - Model Architecture
  - Input Register
  - Assumptions
  - Forecast Outputs
  - Patient Potential Tower
  - Validation
- Market Generate Model downloads a valid `.xlsx` file.
- Generated Market workbook opens and contains:
  - Model Setup
  - Market Inputs
  - Forecast Outputs
  - SKU Structure
  - Validation
- EPI scenario hover displays Base, Strong, and Weak values.
- Epidemiology hover displays Prevalence, Diagnosed, Treated, and Eligible values.
- Patient Potential Tower terminology is visible in the EPI workspace.
- Model Builder controls use the main page scroll; the controls container has no nested scrolling.

## Example hover checks

- Scenario 2031: Base, illustrative Strong, and illustrative Weak all displayed.
- Epidemiology 2031: Prevalence, Diagnosed, Treated, and Eligible all displayed.

## Note

The icon font is loaded from the Tabler CDN. Core navigation, charts, Excel generation, and downloads do not depend on that icon font.
