# aniket-kanani-3107.github.io

## Business Tax Planner Workbook

This repository now includes an Excel planner:

- `Business_Tax_Planner.xlsx`

### Sheets included

1. **Main (All in One)**
   - Assumptions and editable tax settings
   - YTD and predicted KPI summary
   - Tax payment planning cells and alert flags
2. **Last Year Data**
   - Monthly manual input table (Apr to Mar) for the immediately previous financial year
3. **Current Year Data**
   - Monthly manual input table (Apr to Mar)
   - `Status` column (`Actual` / `Pending`) to control whether a month uses entered values or model prediction
4. **All Predictions**
   - Weighted forecasting model
   - Scenario selector (`Base`, `Best`, `Worst`)
   - Predicted monthly and full-year outputs
   - Tax payable, net payable, quarterly estimates, and chart

### How to use

1. Enter prior-year monthly values in **Last Year Data**.
2. Enter current-year monthly values in **Current Year Data**.
3. Set month `Status` in **Current Year Data**:
   - `Actual`: month is completed and final values are available
   - `Pending`: month is not completed yet and should be forecasted
4. Update tax assumptions in **Main (All in One)**.
5. Optionally adjust forecasting weights and scenario in **All Predictions**.
6. Review predicted outcomes, payable tax, and monthly payment suggestions.
