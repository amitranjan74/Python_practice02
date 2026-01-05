

# Improved, vectorized, and safer implementation.
import pandas as pd
import numpy as np
from datetime import date, datetime


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
        "PET_Check",
        "F-Score",
        "Promoter_Pledge%",
        "Revenue-_Latest_FY",
        "Impact_Cost",
        "Median6M_ADTO_in_Rs_crs",
        "Aug_2025_ADTO_in_Rs_crs",
        "Sep_2025_ADTO_in_Rs_crs",
        "Oct_2025_ADTO_in_Rs_crs",
        "BSE_Series",
        "NSE_Series",
        "Last_Trading_Date?",
        "Poor_Alz_Manual",
        "Poor_Alz_Manual_Comments",
        "Size_Exception_for_Average_to_Good",
        "Size_Exception_for_Poor_to_Average",
        "Not_traded?"
    ]

    try:
        _require_columns(fo, expected_cols)
    except KeyError as e:
        # Surface a readable error and re-raise for caller to handle.
        print(f"ERROR: {e}")
        raise


    # Convert Last_Trading_Date? to datetime.date for safe comparison
    if "Last_Trading_Date?" in fo.columns:
        fo["Last_Trading_Date?"] = pd.to_datetime(fo["Last_Trading_Date?"], errors="coerce").dt.date

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
        "Oct_2025_ADTO_in_Rs_crs",
    ]
    for c in numeric_cols:
        if c in fo.columns:
            fo[c] = pd.to_numeric(fo[c], errors="coerce")



    # Good conditions (less strict than Bluechip)
    good_conds = {
        "Total_Shareholders_Funds>=500": fo["[Total_Shareholders_Funds_(Latest)]"].ge(500),
        "Mcap>=500": fo["Mcap"].ge(500),
        "PET_or_FScore": fo["PET_Check"] == "PET Good",
        "Impact_Cost<1": fo["Impact_Cost"].lt(1),
        "Revenue>1000": fo["Revenue-_Latest_FY"].gt(1000),
        "Promoter_Pledge%<40": fo["Promoter_Pledge%"].lt(25),
        "Median6M_ADTO>15": fo["Median6M_ADTO_in_Rs_crs"].gt(2),
        "Aug_2025_ADTO>15": fo["Aug_2025_ADTO_in_Rs_crs"].gt(2),
        "Sep_2025_ADTO>15": fo["Sep_2025_ADTO_in_Rs_crs"].gt(2),
        "Oct_2025_ADTO>15": fo["Oct_2025_ADTO_in_Rs_crs"].gt(2),
        "BSE_series_allowed": allowed_bse,
        "NSE_series_allowed": allowed_nse,
        "Last_Traded_Date": fo["Last_Trading_Date?"].apply(lambda d: not pd.isna(d) and d > date(2025, 10, 25))
    }

    good_exception_conds = {
        "Size_Exception" : fo["Size_Exception_for_Average_to_Good"] == "Upgrade",
        "PET_or_FScore": fo["PET_Check"] == "PET Good",
        "Impact_Cost<1": fo["Impact_Cost"].lt(1),
        "Promoter_Pledge%<40": fo["Promoter_Pledge%"].lt(25),
        "Median6M_ADTO>15": fo["Median6M_ADTO_in_Rs_crs"].gt(2),
        "Aug_2025_ADTO>15": fo["Aug_2025_ADTO_in_Rs_crs"].gt(2),
        "Sep_2025_ADTO>15": fo["Sep_2025_ADTO_in_Rs_crs"].gt(2),
        "Oct_2025_ADTO>15": fo["Oct_2025_ADTO_in_Rs_crs"].gt(2),
        "BSE_series_allowed": allowed_bse,
        "NSE_series_allowed": allowed_nse,
        "Last_Traded_Date": fo["Last_Trading_Date?"].apply(lambda d: not pd.isna(d) and d > date(2025, 10, 25))
    }


    # Average conditions (less strict than Good)
    average_conds = {
        "Total_Shareholders_Funds>=100": fo["[Total_Shareholders_Funds_(Latest)]"].ge(100),
        "Mcap>=100": fo["Mcap"].ge(100),
        "PET_or_FScore": (fo["PET_Check"] == "PET Good") | (fo["F-Score"] == "F-Check Passed"),
        "Impact_Cost<1": fo["Impact_Cost"].lt(1),
        "Revenue>200": fo["Revenue-_Latest_FY"].gt(200),
        "Promoter_Pledge%<50": fo["Promoter_Pledge%"].lt(50),
        "Median6M_ADTO>2": fo["Median6M_ADTO_in_Rs_crs"].gt(1),
        "Aug_2025_ADTO>2": fo["Aug_2025_ADTO_in_Rs_crs"].gt(1),
        "Sep_2025_ADTO>2": fo["Sep_2025_ADTO_in_Rs_crs"].gt(1),
        "Oct_2025_ADTO>2": fo["Oct_2025_ADTO_in_Rs_crs"].gt(1),
        "BSE_series_allowed": allowed_bse,
        "NSE_series_allowed": allowed_nse,
        "Last_Traded_Date": fo["Last_Trading_Date?"].apply(lambda d: not pd.isna(d) and d > date(2025, 10, 25))
    }

    average_exception_conds = {
        "Size_Exception" : fo["Size_Exception_for_Poor_to_Average"] == "Upgrade",
        "Manual_Downgrade" : fo["Poor_Alz_Manual"] != "Downgrade",
        "PET_or_FScore": (fo["PET_Check"] == "PET Good") | (fo["F-Score"] == "F-Check Passed"),
        "Impact_Cost<1": fo["Impact_Cost"].lt(1),
        "Promoter_Pledge%<50": fo["Promoter_Pledge%"].lt(50),
        "Median6M_ADTO>2": fo["Median6M_ADTO_in_Rs_crs"].gt(1),
        "Aug_2025_ADTO>2": fo["Aug_2025_ADTO_in_Rs_crs"].gt(1),
        "Sep_2025_ADTO>2": fo["Sep_2025_ADTO_in_Rs_crs"].gt(1),
        "Oct_2025_ADTO>2": fo["Oct_2025_ADTO_in_Rs_crs"].gt(1),
        "BSE_series_allowed": allowed_bse,
        "NSE_series_allowed": allowed_nse,
        "Last_Traded_Date": fo["Last_Trading_Date?"].apply(lambda d: not pd.isna(d) and d > date(2025, 10, 25))
    }

    average_adto_exception_conds = {
        "Total_Shareholders_Funds>=100": fo["[Total_Shareholders_Funds_(Latest)]"].ge(100),
        "Mcap>=100": fo["Mcap"].ge(100),
        "PET_or_FScore": (fo["PET_Check"] == "PET Good") | (fo["F-Score"] == "F-Check Passed"),
        "Impact_Cost<1": fo["Impact_Cost"].lt(1),
        "Revenue>200": fo["Revenue-_Latest_FY"].gt(200),
        "Promoter_Pledge%<50": fo["Promoter_Pledge%"].lt(50),
        "ADTO_Excep" : fo["ADTO_Upgrade_(Poor_to_Average)"] == "Upgrade",
        "BSE_series_allowed": allowed_bse,
        "NSE_series_allowed": allowed_nse,
        "Last_Traded_Date": fo["Last_Trading_Date?"].apply(lambda d: not pd.isna(d) and d > date(2025, 10, 25))
    }

    # Poor conditions
    poor_open_conds_1 = {
        "Total_Shareholders_Funds<100": fo["[Total_Shareholders_Funds_(Latest)]"].lt(100),
        "Mcap<100": fo["Mcap"].lt(100),
        "PET_or_FScore": (fo["PET_Check"] == "PET Good") | (fo["F-Score"] == "F-Check Passed"),
        "Impact_Cost<1": fo["Impact_Cost"].gt(1),
        "Revenue<=200": fo["Revenue-_Latest_FY"].le(200),
        "Promoter_Pledge%>=50": fo["Promoter_Pledge%"].ge(50),
        "Median6M_ADTO<1": fo["Median6M_ADTO_in_Rs_crs"].lt(1),
        "Aug_2025_ADTO<1": fo["Aug_2025_ADTO_in_Rs_crs"].lt(1),
        "Sep_2025_ADTO<1": fo["Sep_2025_ADTO_in_Rs_crs"].lt(1),
        "Oct_2025_ADTO<1": fo["Oct_2025_ADTO_in_Rs_crs"].lt(1),
        "BSE_series_allowed": allowed_bse,
        "NSE_series_allowed": allowed_nse,
    }

    poor_open_conds_2 = {
        "Basic_Funda": (fo["[Total_Shareholders_Funds_(Latest)]"].lt(100)) | (fo["Mcap"].lt(100)) | (fo["Revenue-_Latest_FY"].le(200))
    }

    # Treat any non-empty (after stripping) manual entry as a manual rating.
    # Use fillna/astype/str to avoid NaN/None issues and strip whitespace.
    poor_manual = {
        "Manual_Rating": fo["Poor_Alz_Manual"].fillna("").astype(str).str.strip() != "",
    }

    # Combine conditions to form masks
    Good_mask = np.logical_and.reduce([s.fillna(False) for s in good_conds.values()])
    Good_exception_mask = np.logical_and.reduce([s.fillna(False) for s in good_exception_conds.values()])
    Average_mask = np.logical_and.reduce([s.fillna(False) for s in average_conds.values()])
    Average_exception_mask = np.logical_and.reduce([s.fillna(False) for s in average_exception_conds.values()])
    Average_adto_exception_mask = np.logical_and.reduce([s.fillna(False) for s in average_adto_exception_conds.values()])
    Poor_mask_1 = np.logical_and.reduce([s.fillna(False) for s in poor_open_conds_1.values()])
    Poor_mask_2 = np.logical_and.reduce([s.fillna(False) for s in poor_open_conds_2.values()])
    Poor_mask_manual= np.logical_and.reduce([s.fillna(False) for s in poor_manual.values()])

    fo.loc[Good_mask & fo["Rating_Amit"].isna(), "Rating_Amit"] = "Good"
    fo.loc[Good_exception_mask & fo["Rating_Amit"].isna(), "Rating_Amit"] = "Good_in_Size_Exception"
    fo.loc[Average_mask & fo["Rating_Amit"].isna(), "Rating_Amit"] = "Average"
    fo.loc[Average_exception_mask & fo["Rating_Amit"].isna(), "Rating_Amit"] = "Average_in_Size_Exception"
    fo.loc[Average_adto_exception_mask & fo["Rating_Amit"].isna(), "Rating_Amit"] = "Average_in_ADTO_Exception"
    fo.loc[Poor_mask_1 & fo["Rating_Amit"].isna(), "Rating_Amit"] = "Poor"
    fo.loc[Poor_mask_2 & fo["Rating_Amit"].isna(), "Rating_Amit"] = "Poor"
    fo.loc[Poor_mask_manual, "Rating_Amit"] = "Poor_Manual"
    fo.loc[fo["Rating_Amit"].isna(), "Rating_Amit"] = "Left-Blank Hence Poor"
    

    # Build a human-readable reason column describing which stricter-tier conditions failed.
    fo["Reason_Amit"] = ""


    # For rows rated Average, list which Good conditions they failed
    for idx in fo.index[fo["Rating_Amit"].isin(["Average", "Average_in_Size_Exception", "Average_in_ADTO_Exception"])]:
        failed = [name for name, series in good_conds.items() if not bool(series.loc[idx])]
        fo.at[idx, "Reason_Amit"] = ", ".join(failed) if failed else "met_all_Good"

    # For rows rated Poor, list which Average conditions they failed
    for idx in fo.index[fo["Rating_Amit"] == "Poor"]:
        failed = [name for name, series in average_conds.items() if not bool(series.loc[idx])]
        fo.at[idx, "Reason_Amit"] = ", ".join(failed) if failed else "met_all_Average"

    # For rows rated Left-Blank Hence Poor, list which Blank and not meeting conditions
    for idx in fo.index[fo["Rating_Amit"] == "Left-Blank Hence Poor"]:
        failed = [name for name, series in poor_open_conds_1.items() if bool(series.loc[idx])]
        fo.at[idx, "Reason_Amit"] = ", ".join(failed) if failed else "met_all_Poor"




    # If needed, collect the CO_CODE

    # Optionally save the result to a new CSV if output_csv was provided
    if output_csv:
        fo.to_csv(output_csv, index=False)

    return fo


if __name__ == "__main__":
    # Example: will write file with ratings appended
    try:
        main(input_csv="NonFnO_Data_cleaned.csv", output_csv="NonFnO_Data_cleaned_with_ratings.csv")
    except Exception:
        # Reraise after printing so a caller/test harness can fail the run if needed
        raise