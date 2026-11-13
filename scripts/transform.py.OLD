import pandas as pd

# Read the source Excel file
df = pd.read_excel("raw/product-lifecycle.xlsx")

result = pd.DataFrame()

# Vendor and Product/Edition/Release/Azure Feature/Version/LifecyclePolicy mapping
result["Vendor"] = "Microsoft"
result["Product"] = df["Product Listing Name"]

# Helper to pick a column name from possible variants (case/spacing differences)
def pick_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return df[c]
    return None

# Edition, Release, Azure Feature and Version
edition_col = pick_col(df, ["Edition", "Edition ", "Product Edition"]) 
release_col = pick_col(df, ["Release", "Release Number", "Release number"]) 
azure_feature_col = pick_col(df, ["Azure Feature", "AzureFeature", "Azure feature"]) 

# Lifecycle policy
lifecycle_col = pick_col(df, ["Support Policy", "Lifecycle Policy", "Support policy"])

# Assign columns in the desired order
result["Version"] = edition_col
result["Release"] = release_col
result["Azure Feature"] = azure_feature_col
result["LifecyclePolicy"] = lifecycle_col

# StartDate: prefer "Release Start Date", fallback to "Start Date"
start_src = pick_col(df, ["Release Start Date", "Release start date", "ReleaseStartDate"]) 
start_fallback = pick_col(df, ["Start Date", "Start date", "StartDate"]) 
if start_src is None:
    start = start_fallback
else:
    # use release start date where present, otherwise fallback
    if start_fallback is not None:
        start = start_src.fillna(start_fallback)
    else:
        start = start_src

start_dt = pd.to_datetime(start, errors='coerce')
# Format to Short Date (Dutch(Belgium)) -> dd/mm/YYYY
result["StartDate"] = start_dt.dt.strftime("%d/%m/%Y")

# EndOfSupport: prefer Release End Date, then Extended End Date, then Retirement
end_release = pick_col(df, ["Release End Date", "Release end date", "ReleaseEndDate"]) 
end_extended = pick_col(df, ["Extended End Date", "Extended end date", "ExtendedEndDate"]) 
end_retire = pick_col(df, ["Retirement", "Retirement Date"]) 

# Build end candidate by chaining fillna
end_combined = None
if end_release is not None:
    end_combined = end_release

if end_combined is None and end_extended is not None:
    end_combined = end_extended
elif end_combined is not None and end_extended is not None:
    end_combined = end_combined.fillna(end_extended)

if end_retire is not None:
    if end_combined is None:
        end_combined = end_retire
    else:
        end_combined = end_combined.fillna(end_retire)

# If still None create a Series of NaT so conversion below works
if end_combined is None:
    end_combined = pd.Series([pd.NaT] * len(df))

end_dt = pd.to_datetime(end_combined, errors='coerce')
# where there's no date at all, set to 31/12/2099
end_dt = end_dt.fillna(pd.Timestamp(year=2099, month=12, day=31))
# store formatted date
result["EndOfSupport"] = end_dt.dt.strftime("%d/%m/%Y")

# DaysToEOS: number of days between today and EndOfSupport
today = pd.Timestamp.now().normalize()
result["DaysToEOS"] = (end_dt - today).dt.days

# LifecycleStatus: if original "End of Support" equals "Out of Support" for that product, copy that value
# otherwise derive status from DaysToEOS using the thresholds given.
# Find the original End of Support column in the source sheet (several possible spellings)
eos_original = pick_col(df, ["End of Support", "End Of Support", "End of support", "EndOfSupport"]) 

# Also capture the original "Product Listing Name" to compare
original_product = df["Product Listing Name"]

# Function to compute lifecycle status per requirements
def compute_status(i, days):
    # check original End of Support explicit Out of Support for the same product
    if eos_original is not None:
        val = eos_original.iloc[i]
        if isinstance(val, str) and val.strip().lower() == "out of support":
            return val.strip()
    # if days is missing
    if pd.isna(days):
        return "UNKNOWN"
    try:
        d = int(days)
    except Exception:
        return "UNKNOWN"
    # follow requested thresholds
    if d <= 0:
        return "END_OF_SUPPORT"
    elif 1 <= d <= 90:
        return "EOS_LT_3_MONTHS"
    elif 91 <= d <= 180:
        return "EOS_LT_6_MONTHS"
    elif 181 <= d <= 365:
        return "EOS_LT_12_MONTHS"
    else:
        return "SUPPORTED"

# Apply compute_status
result["LifecycleStatus"] = [compute_status(i, days) for i, days in enumerate(result["DaysToEOS"]) ]

# Export CSV and JSON
# Ensure column order is consistent/explicit
cols = [
    "Vendor",
    "Product",
    "Version",
    "Release",
    "Azure Feature",
    "LifecyclePolicy",
    "StartDate",
    "EndOfSupport",
    "DaysToEOS",
    "LifecycleStatus",
]

# Export to CSV and JSON with all columns included
result.to_csv("data/microsoft-lifecycle.csv", index=False, columns=cols)
result.to_json("data/microsoft-lifecycle.json", orient="records", indent=2)

# Optionally produce a summary (kept from previous script)
summary = (
    result
    .groupby("LifecycleStatus")
    .size()
    .reset_index(name="Count")
)

print(summary)
