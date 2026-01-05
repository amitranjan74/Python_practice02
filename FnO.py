
# Improved, vectorized, and safer implementation.
import pandas as pd
import numpy as np


def _require_columns(df, cols):
    """Raise KeyError if any column in cols is missing from df."""
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required columns: {missing}")


def main(input_csv="FO_Data_cleaned.csv", output_csv=None):
    fo = pd.read_csv(input_csv)

    # Expected columns used by the rules
    expected_cols = [
        "CO_CODE",
        "[Total_Shareholders_Funds_(Latest)]",
        "Mcap",
        "F&O",
        "PET_Check",
        "F-Score",
        "Promoter_Pledge%",
        "Revenue-_Latest_FY",
        "Impact_Cost",
        "Median6M_ADTO_in_Rs_crs",
        "Aug_2025_ADTO_in_Rs_crs",
        "Sep_2025_ADTO_in_Rs_crs",
        "Oct_2025_ADTO_in_Rs_cr",
        "BSE_Series",
        "NSE_Series",
    ]

    try:
        _require_columns(fo, expected_cols)
    except KeyError as e:
        # Surface a readable error and re-raise for caller to handle.
        print(f"ERROR: {e}")
        raise

    # Start with NaN ratings
    fo["Rating_Amit"] = np.nan

    # Build boolean masks (vectorized) for the two rule-sets.
    # Use Series methods (.ge, .gt, .lt) to handle NaNs gracefully.
    allowed_bse = ~fo["BSE_Series"].isin(["Z", "P"])  # True when series is not Z or P
    allowed_nse = ~fo["NSE_Series"].isin(["BZ"])        # True when series is not BZ

    # Coerce numeric fields to numeric types so comparisons work correctly
    numeric_cols = [
        "[Total_Shareholders_Funds_(Latest)]",
        "Mcap",
        "Promoter_Pledge%",
        "Revenue-_Latest_FY",
        "Impact_Cost",
        "Median6M_ADTO_in_Rs_crs",
        "Aug_2025_ADTO_in_Rs_crs",
        "Sep_2025_ADTO_in_Rs_crs",
        "Oct_2025_ADTO_in_Rs_cr",
    ]
    for c in numeric_cols:
        if c in fo.columns:
            fo[c] = pd.to_numeric(fo[c], errors="coerce")

    # Bluechip conditions
    blue_conds = {
        "Total_Shareholders_Funds>=2000": fo["[Total_Shareholders_Funds_(Latest)]"].ge(2000),
        "Mcap>=5000": fo["Mcap"].ge(5000),
        "F&O==F&O": fo["F&O"] == "F&O",
        "PET_Check==PET Good": fo["PET_Check"] == "PET Good",
        "Impact_Cost<1": fo["Impact_Cost"].lt(1),
        "Revenue>5000": fo["Revenue-_Latest_FY"].gt(5000),
        "Promoter_Pledge%<25": fo["Promoter_Pledge%"].lt(25),
        "Median6M_ADTO>20": fo["Median6M_ADTO_in_Rs_crs"].gt(20),
        "Aug_2025_ADTO>20": fo["Aug_2025_ADTO_in_Rs_crs"].gt(20),
        "Sep_2025_ADTO>20": fo["Sep_2025_ADTO_in_Rs_crs"].gt(20),
        "Oct_2025_ADTO>20": fo["Oct_2025_ADTO_in_Rs_cr"].gt(20),
        "BSE_series_allowed": allowed_bse,
        "NSE_series_allowed": allowed_nse,
    }

    # Good conditions (less strict than Bluechip)
    good_conds = {
        "Total_Shareholders_Funds>=500": fo["[Total_Shareholders_Funds_(Latest)]"].ge(500),
        "Mcap>=500": fo["Mcap"].ge(500),
        "F&O==F&O": fo["F&O"] == "F&O",
        "PET_or_FScore": (fo["PET_Check"] == "PET Good") | (fo["F-Score"] == "F-Check Passed"),
        "Impact_Cost<1": fo["Impact_Cost"].lt(1),
        "Revenue>1000": fo["Revenue-_Latest_FY"].gt(1000),
        "Promoter_Pledge%<40": fo["Promoter_Pledge%"].lt(40),
        "Median6M_ADTO>15": fo["Median6M_ADTO_in_Rs_crs"].gt(15),
        "Aug_2025_ADTO>15": fo["Aug_2025_ADTO_in_Rs_crs"].gt(15),
        "Sep_2025_ADTO>15": fo["Sep_2025_ADTO_in_Rs_crs"].gt(15),
        "Oct_2025_ADTO>15": fo["Oct_2025_ADTO_in_Rs_cr"].gt(15),
        "BSE_series_allowed": allowed_bse,
        "NSE_series_allowed": allowed_nse,
    }

    # Average conditions (less strict than Good)
    average_conds = {
        "Total_Shareholders_Funds>=100": fo["[Total_Shareholders_Funds_(Latest)]"].ge(100),
        "Mcap>=100": fo["Mcap"].ge(100),
        "F&O==F&O": fo["F&O"] == "F&O",
        "PET_or_FScore": (fo["PET_Check"] == "PET Good") | (fo["F-Score"] == "F-Check Passed"),
        "Impact_Cost<1": fo["Impact_Cost"].lt(1),
        "Revenue>200": fo["Revenue-_Latest_FY"].gt(200),
        "Promoter_Pledge%<50": fo["Promoter_Pledge%"].lt(50),
        "Median6M_ADTO>2": fo["Median6M_ADTO_in_Rs_crs"].gt(2),
        "Aug_2025_ADTO>2": fo["Aug_2025_ADTO_in_Rs_crs"].gt(2),
        "Sep_2025_ADTO>2": fo["Sep_2025_ADTO_in_Rs_crs"].gt(2),
        "Oct_2025_ADTO>2": fo["Oct_2025_ADTO_in_Rs_cr"].gt(2),
        "BSE_series_allowed": allowed_bse,
        "NSE_series_allowed": allowed_nse,
    }

    # Poor conditions
    poor_conds = {
        "Total_Shareholders_Funds<100": fo["[Total_Shareholders_Funds_(Latest)]"].lt(100),
        "Mcap<100": fo["Mcap"].lt(100),
        "F&O_blank": fo["F&O"] == "",
        "PET_blank": fo["PET_Check"] == "",
        "F-Score_blank": fo["F-Score"] == "",
        "Impact_Cost<1": fo["Impact_Cost"].gt(1),
        "Revenue<=200": fo["Revenue-_Latest_FY"].le(200),
        "Promoter_Pledge%>=50": fo["Promoter_Pledge%"].ge(50),
        "Median6M_ADTO<1": fo["Median6M_ADTO_in_Rs_crs"].lt(1),
        "Aug_2025_ADTO<1": fo["Aug_2025_ADTO_in_Rs_crs"].lt(1),
        "Sep_2025_ADTO<1": fo["Sep_2025_ADTO_in_Rs_crs"].lt(1),
        "Oct_2025_ADTO<1": fo["Oct_2025_ADTO_in_Rs_cr"].lt(1),
        "BSE_series_allowed": allowed_bse,
        "NSE_series_allowed": allowed_nse,
    }

    # Combine conditions to form masks
    Bluechip_mask = np.logical_and.reduce([s.fillna(False) for s in blue_conds.values()])
    Good_mask = np.logical_and.reduce([s.fillna(False) for s in good_conds.values()])
    Average_mask = np.logical_and.reduce([s.fillna(False) for s in average_conds.values()])
    Poor_mask = np.logical_and.reduce([s.fillna(False) for s in poor_conds.values()])

    fo.loc[Bluechip_mask, "Rating_Amit"] = "Bluechip"
    fo.loc[Good_mask & fo["Rating_Amit"].isna(), "Rating_Amit"] = "Good"
    fo.loc[Average_mask & fo["Rating_Amit"].isna(), "Rating_Amit"] = "Average"
    fo.loc[Poor_mask & fo["Rating_Amit"].isna(), "Rating_Amit"] = "Poor"
    fo.loc[fo["Rating_Amit"].isna(), "Rating_Amit"] = "Left-Blank Hence Poor"

    # Build a human-readable reason column describing which stricter-tier conditions failed.
    fo["Reason_Amit"] = ""

    # For rows rated Good, list which Bluechip conditions they failed
    for idx in fo.index[fo["Rating_Amit"] == "Good"]:
        failed = [name for name, series in blue_conds.items() if not bool(series.loc[idx])]
        fo.at[idx, "Reason_Amit"] = ", ".join(failed) if failed else "met_all_Bluechip"

    # For rows rated Average, list which Good conditions they failed
    for idx in fo.index[fo["Rating_Amit"] == "Average"]:
        failed = [name for name, series in good_conds.items() if not bool(series.loc[idx])]
        fo.at[idx, "Reason_Amit"] = ", ".join(failed) if failed else "met_all_Good"

    # For rows rated Poor, list which Average conditions they failed
    for idx in fo.index[fo["Rating_Amit"] == "Poor"]:
        failed = [name for name, series in average_conds.items() if not bool(series.loc[idx])]
        fo.at[idx, "Reason_Amit"] = ", ".join(failed) if failed else "met_all_Average"

    # For rows rated Left-Blank Hence Poor, list which Blank and not meeting conditions
    for idx in fo.index[fo["Rating_Amit"] == "Left-Blank Hence Poor"]:
        failed = [name for name, series in poor_conds.items() if bool(series.loc[idx])]
        fo.at[idx, "Reason_Amit"] = ", ".join(failed) if failed else "met_all_Poor"




    # If needed, collect the CO_CODE

    # Optionally save the result to a new CSV if output_csv was provided
    if output_csv:
        fo.to_csv(output_csv, index=False)

    return fo


if __name__ == "__main__":
    # Example: will write file with ratings appended
    try:
        main(input_csv="FO_Data_cleaned.csv", output_csv="FO_Data_cleaned_with_ratings.csv")
    except Exception:
        # Reraise after printing so a caller/test harness can fail the run if needed
        raise