import pandas as pd
import numpy as np
import os
import json
import sys
from pathlib import Path
import app

ref_file = 'TCF VIN  & Paint Float mapping data-1.xlsx'
test_dir = Path("test_runs")
test_dir.mkdir(exist_ok=True)

# Extract sheets to separate files
xls = pd.ExcelFile(ref_file)
sheet_to_file = {
    'Paint Float': 'paint_float.xls',
    'TCF1 DPT Plan': 'tcf1_dpt_plan.xls',
    'TCF2 DPT Plan': 'tcf2_dpt_plan.xls',
    'TCF1 Wiring File': 'tcf1_wiring.xls',
    'TCF2 Wiring File': 'tcf2_wiring.xls',
    'TCF1 Cockpit': 'tcf1_cockpit.xls',
    'TCF2 Cockpit': 'tcf2_cockpit.xls',
    'Nova Cockpit': 'nova_cockpit.xls'
}

for sheet, fname in sheet_to_file.items():
    dest = test_dir / fname
    df_sheet = pd.read_excel(xls, sheet_name=sheet, header=None)
    df_sheet.to_excel(dest, index=False, header=False)

report_path = Path("test_report.txt")
with open(report_path, "w", encoding="utf-8") as out_file:
    def log(msg=""):
        out_file.write(str(msg) + "\n")
        print(str(msg).encode("ascii", errors="replace").decode("ascii"))

    log(f"Reading reference file: {ref_file}")
    log(f"Extracted sheets saved to: {test_dir}")

    log("\n--- Testing Part Master / BOM details loading ---")
    bom_df = app.load_bom_master()
    log(f"BOM master shape: {bom_df.shape}")
    log(f"BOM columns: {list(bom_df.columns)}")
    pm = bom_df.drop_duplicates(subset="SHORT_VC").reset_index(drop=True)

    log("\n--- Testing Parsers ---")
    pf_df = app.parse_paint_float("uploaded", test_dir / "paint_float.xls", pm)
    log(f"Paint Float shape: {pf_df.shape}")
    log(f"Sample Paint Float columns: {list(pf_df.columns[:5])} ... and float cols")

    dpt_frames = []
    for ft, fname, label in [
        ("TCF1 DPT Plan", "tcf1_dpt_plan.xls", "TCF1"),
        ("TCF2 DPT Plan", "tcf2_dpt_plan.xls", "TCF2")
    ]:
        df_parsed = app.parse_dpt_plan("uploaded", test_dir / fname, ft, label)
        log(f"{label} DPT parsed shape: {df_parsed.shape}")
        dpt_frames.append(df_parsed)

    dpt_all = pd.concat(dpt_frames, ignore_index=True) if dpt_frames else None
    if dpt_all is not None:
        dpt_all.drop(columns=["WIRING", "COCKPIT", "ENGINE"], inplace=True, errors="ignore")
        pm_map = pm[["SHORT_VC", "FRONT_WIRING", "COCKPIT", "ENGINE"]].rename(columns={"FRONT_WIRING": "WIRING"})
        dpt_all = dpt_all.merge(pm_map, on="SHORT_VC", how="left")
        log(f"Combined DPT All shape after mapping: {dpt_all.shape}")
        log(f"Mapped non-empty Wiring count: {dpt_all[dpt_all['WIRING'].notna() & (dpt_all['WIRING'] != '')].shape[0]}")
        log(f"Mapped non-empty Engine count: {dpt_all[dpt_all['ENGINE'].notna() & (dpt_all['ENGINE'] != '')].shape[0]}")

    wiring_tcf1 = app.parse_wiring_tcf1("uploaded", test_dir / "tcf1_wiring.xls")
    wiring_tcf2 = app.parse_wiring_tcf2("uploaded", test_dir / "tcf2_wiring.xls")
    log(f"Wiring TCF1 shape: {wiring_tcf1.shape if wiring_tcf1 is not None else None}")
    log(f"Wiring TCF2 shape: {wiring_tcf2.shape if wiring_tcf2 is not None else None}")

    cockpit_dfs = []
    for ft, fname, label in [
        ("TCF1 Cockpit", "tcf1_cockpit.xls", "TCF1"),
        ("TCF2 Cockpit", "tcf2_cockpit.xls", "TCF2"),
        ("Nova Cockpit", "nova_cockpit.xls", "TCF1")
    ]:
        df_parsed = app.parse_cockpit_file("uploaded", test_dir / fname, ft, label)
        log(f"Cockpit {ft} shape: {df_parsed.shape if df_parsed is not None else None}")
        cockpit_dfs.append(df_parsed)

    log("\n--- Testing Compute Summaries ---")

    mw_float = app.compute_model_wise_float(pf_df, dpt_all, "All Lines")
    log(f"Model Wise Float summary shape: {mw_float.shape}")
    log(mw_float[['Paint Float', 'MODEL', 'TOTAL FLOAT', 'PBS FLOAT', 'Today VIN']].to_string())

    wiring_sum = app.compute_wiring_summary(pf_df, wiring_tcf1, wiring_tcf2, pm, dpt_all, "All Lines")
    log(f"\nWiring Summary shape: {wiring_sum.shape}")
    log("Wiring Summary sample (worst 5 shortages):")
    log(wiring_sum.head(5).to_string())

    cockpit_sum = app.compute_cockpit_summary(pf_df, cockpit_dfs, pm, dpt_all, "All Lines")
    log(f"\nCockpit Summary shape: {cockpit_sum.shape}")
    log("Cockpit Summary sample (worst 5 shortages):")
    log(cockpit_sum.head(5).to_string())

    engine_req = app.build_engine_table(pf_df, "All Lines")
    log(f"\nEngine base requirements shape: {engine_req.shape}")
    saved = {}
    today_engine_vin = {}
    if dpt_all is not None:
        dpt = dpt_all.copy()
        dpt["ENGINE"] = dpt["ENGINE"].astype(str).str.strip()
        dpt_agg = dpt[dpt["ENGINE"].str.len() >= 5].groupby("ENGINE", as_index=False)["DPT_VIN"].sum()
        today_engine_vin = dict(zip(dpt_agg["ENGINE"], dpt_agg["DPT_VIN"]))

    engine_req["Clearance After 6:30AM"] = 100
    engine_req["Today VIN"] = engine_req["Engine Part No"].map(lambda e: today_engine_vin.get(e, 0))
    engine_req["Bal"] = engine_req["Clearance After 6:30AM"] - engine_req["Today VIN"]
    engine_req["PBS FLOAT"] = engine_req["PBS_FLOAT"]
    engine_req["Float UPTO SEALANT"] = engine_req["UPTO_SEALANT"]
    engine_req["TOTAL FLOAT"] = engine_req["TOTAL_FLOAT"]
    engine_req["With respect to PBS FLOAT"] = engine_req["Bal"] - engine_req["PBS FLOAT"]
    engine_req["With respect to Sealant FLOAT"] = engine_req["Bal"] - engine_req["Float UPTO SEALANT"]
    engine_req["With respect to Total FLOAT"] = engine_req["Bal"] - engine_req["TOTAL FLOAT"]
    engine_sum = app.add_engine_subtotals(engine_req)
    log(f"Engine Summary shape with subtotals: {engine_sum.shape}")
    log("Engine Summary sample:")
    log(engine_sum[['Engine Part No', 'Model', 'Clearance After 6:30AM', 'Today VIN', 'Bal', 'TOTAL FLOAT', 'With respect to Total FLOAT']].head(10).to_string())

    vin_vs_float = app.compute_vin_vs_float(pf_df, dpt_all, "All Lines")
    log(f"\nVIN vs Float shape: {vin_vs_float.shape}")
    log("VIN vs Float sample (worst 5 shortages):")
    log(vin_vs_float.head(5).to_string())

    log("\nAll tests ran and written to report successfully!")

print("Done! Test report written to test_report.txt")
