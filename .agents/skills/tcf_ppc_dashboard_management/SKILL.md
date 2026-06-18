---
name: tcf-ppc-dashboard-management
description: Manage, run, debug, and test the TCF PPC Streamlit dashboard.
---

# TCF PPC Dashboard Management Skill

Use this skill when you need to run, configure, test, or debug the Streamlit dashboard application for Tata Motors TCF line tracking.

## Running the Streamlit App
To run the dashboard locally in the workspace:
```bash
python -m streamlit run app.py
```
- By default, it runs on port `8503` (or the next available port).
- The application automatically reloads when `app.py` is edited.

## Verifying Integrity & Tests
Always run the integration tests after making changes to the data parsing or calculations:
```bash
python test_app_integration.py
```
This script:
1. Reads sample files from the project folder and the consolidated Excel workbook.
2. Runs all parsers (`parse_paint_float`, `parse_wiring_tcf1`, etc.).
3. Verifies that generated shortage models match correct reference sheet figures.
4. Generates a report in `test_report.txt`.

## Project Directory Structure
- `app.py`: Main Streamlit script containing UI and business logic.
- `data/`: Local storage directory (created automatically).
  - `engine_stock.json`: Contains manually entered engine stocks (Clearance, Today VIN, TA Code).
  - `bom_master.json`: Cached version of the Bill of Materials.
  - `last_reset.json`: Stores timestamp of the last 6:30 AM reset.
- `Bom details.xlsx`: Reference BOM details sheet.
- `TCF VIN  & Paint Float mapping data-1.xlsx`: Consolidated workbook.

## Reruns and Session State
- Streamlit reruns the entire script on user interaction.
- Component state like input tables (such as engine manual stock inputs) are cached in `st.session_state` to prevent losing cursor focus or resets during typing.
