import pandas as pd

df = pd.read_excel(
    "raw/product-lifecycle.xlsx"
)

result = pd.DataFrame()

result["Vendor"] = "Microsoft"

result["Product"] = (
    df["Product Listing Name"]
)

result["Version"] = (
    df["Edition"]
)

result["LifecyclePolicy"] = (
    df["Support Policy"]
)

result["StartDate"] = (
    df["Start Date"]
)

result["EndOfSupport"] = (
    df["End of Support"]
)

today = pd.Timestamp.today()

result["EndOfSupport"] = (
    pd.to_datetime(
        result["EndOfSupport"],
        errors='coerce'
    )
)

result["DaysToEOS"] = (

    result["EndOfSupport"]

    - today

).dt.days

def lifecycle_status(days):

    if pd.isna(days):
        return "UNKNOWN"

    elif days < 0:
        return "END_OF_SUPPORT"

    elif days < 90:
        return "EOS_LT_3_MONTHS"

    elif days < 180:
        return "EOS_LT_6_MONTHS"

    elif days < 365:
        return "EOS_LT_12_MONTHS"

    return "SUPPORTED"

result["LifecycleStatus"] = (

    result["DaysToEOS"]

    .apply(
        lifecycle_status
    )

)

result.to_csv(

    "data/microsoft-lifecycle.csv",

    index=False

)

result.to_json(

    "data/microsoft-lifecycle.json",

    orient="records",

    indent=2

)

summary = (

    result
    .groupby(
        "LifecycleStatus"
    )
    .size()
    .reset_index(
        name="Count"
    )

)
