
import pandas as pd


def load_csv(path: str) -> pd.DataFrame:
	"""Load a CSV file with a safe encoding fallback strategy.

	Tries utf-8 first, then cp1252 and latin1. If all fail, opens the file
	with utf-8 and errors='replace' to avoid UnicodeDecodeError and
	lets pandas parse the (possibly lossy) text.

	Returns the DataFrame loaded and prints which method succeeded.
	"""
	encodings_to_try = ("utf-8", "cp1252", "latin1")
	for enc in encodings_to_try:
		try:
			df = pd.read_csv(path, encoding=enc)
			print(f"Loaded '{path}' with encoding={enc}")
			return df
		except UnicodeDecodeError:
			# try next encoding
			continue
		except Exception as e:
			# other parsing errors should be surfaced but we print for debug
			print(f"pd.read_csv failed with encoding={enc}: {e}")

	# Last resort: open with replacement for undecodable bytes and let pandas
	# read from the file-like object. This prevents UnicodeDecodeError but
	# may replace problematic characters.
	print(
		"All encodings failed; opening file with errors='replace' using utf-8."
	)
	with open(path, "r", encoding="utf-8", errors="replace") as f:
		df = pd.read_csv(f)
	print(f"Loaded '{path}' with utf-8 and errors='replace'")
	return df


# Load data
fo_data = load_csv("non_fno.csv")

# Normalize column names: strip leading/trailing whitespace, replace any
# run of whitespace with a single underscore, and keep other characters same.
# This makes column names safe to use as identifiers (e.g., for attribute access).
fo_data.columns = (
	fo_data.columns
	.astype(str)
	.str.strip()
	.str.replace(r"\s+", "_", regex=True)
)

# Show the renamed columns
print(fo_data.columns)

# Optional: to overwrite the CSV with cleaned column names
fo_data.to_csv("NonFnO_Data_cleaned.csv", index=False)

# Writing the logic fot rating

# if [Total_Shareholders_Funds_(Latest)] 
# Mcap
# F&O
# PET_Check
# F-Score
# Impact_Cost
# Revenue-_Latest_FY
# Promoter_Pledge%
# 6M_Median_ADTO_in_Rs_crs

