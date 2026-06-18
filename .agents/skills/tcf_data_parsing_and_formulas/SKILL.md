---
name: tcf-data-parsing-and-formulas
description: Reference guide for the data channels, excel columns, parsing rules, and shortage calculations in the dashboard.
---

# TCF Data Parsing and Formulas Reference Skill

Use this skill when modifying calculations, headers, or columns in `app.py` for TCF dashboard reporting.

## 1. Input Excel Mapping Tables
The dashboard parses inputs from several sheets. Ensure data structure matches the following formats:

### Paint Float Report
- **Source**: PPC Float Report.
- **Fields parsed**: `SHORT_VC`, `MODEL`, `TOTAL_FLOAT`, `PBS_FLOAT`, `PBS_TO_POLISHING`, `POLISHING_TO_TOPCOAT`, `TOPCOAT_TO_WETSANDING_ROOFBLACK`, `TOPCOAT_TO_WETSANDING_FRESH`, `WETSANDING_TO_SEALANT`, `TOTAL_UPTO_SEALANT`, `PT_ENTRY_TO_SEALENT`, `BIW_LIFTING_TO_PT`, `PT_BYPASS`.
- **Note**: Merges wiring, cockpit, and engine part numbers using the Short VC from the BOM master dataset.

### Wiring Harness Stock
- **Source**: TCF1 and TCF2 Wiring Files.
- **Header matching**: Searches for headers containing `"COVERAGE"` or `"FRESH VIN"` (usually Column J / index 9) to extract clearances.
- **Formula discrepancy fix**: The reference Excel workbook sums Column J for clearances. Column J must be used as the source of clearances for TCF1 and TCF2 shortage reports.

### DPT Production Plans
- **Source**: TCF1 and TCF2 DPT Plan sheets.
- **Fields parsed**: Maps vehicle sequencing and sums clearances by engine type or part.

### Cockpit Assemblies
- **Source**: TCF1 Cockpit, TCF2 Cockpit, and Nova Cockpit files.

## 2. Business Shortage Calculations
Calculations are executed in the summary generators:

### Wiring / Cockpit Shortage Calculations
For each part number:
- **Shortage vs PBS**: `Clearance - Today_VIN - PBS_Float`
- **Shortage vs Upto Sealant**: `Clearance - Today_VIN - Total_Upto_Sealant`
- **Shortage vs TOTAL Float**: `Clearance - Today_VIN - Total_Float`

### Engine Stock Calculations
- **Engine Balance**: `Clearance - Today_VIN`
- **With respect to Total Float**: `Clearance - Today_VIN - Total_Float`

## 3. Shift Reset Logic
- Reset occurs daily at **6:30 AM IST**.
- Clears all cached manual stock JSONs (`engine_stock.json`), BOM cache (`bom_master.json`), and deletes uploaded files in `data/` to start a fresh production shift.
