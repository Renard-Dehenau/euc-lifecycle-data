#!/usr/bin/env python3
"""Microsoft EUC lifecycle -> Power BI CSV with expanded lifecycle columns.

The 20 shared release fields are preserved; the 12 previously JSON-encoded
Description keys are published as source-specific, typed-friendly CSV columns.
"""
from __future__ import annotations

import argparse
import json
import os
import tempfile
from datetime import date
from pathlib import Path
from urllib.parse import urlsplit

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from release_contract import (
    OUTPUT_COLUMNS as BASE_OUTPUT_COLUMNS, business_content_hash,
    clean, create_record, digest, get, iso_utc, output_paths, run_pipeline,
)

ROOT = Path(__file__).resolve().parent.parent
SOURCE = "EUC_LIFECYCLE"
# Stable, additive lifecycle extension. Existing 20 base fields retain their order.
LIFECYCLE_DETAIL_COLUMNS = [
    "Vendor", "Edition", "Release", "AzureFeature", "LifecyclePolicy",
    "StartDate", "MainstreamEndDate", "ExtendedEndDate", "ReleaseEndDate",
    "RetirementDate", "EndOfSupport", "OriginalSupportStatus",
]
OUTPUT_COLUMNS = BASE_OUTPUT_COLUMNS + LIFECYCLE_DETAIL_COLUMNS
MICROSOFT_XLSX_URL = (
    "https://download.microsoft.com/download/"
    "a0c60d25-38e1-4d44-91e3-76c6d4c71275/product-lifecycle-data-new.xlsx"
)
PATHS = output_paths(ROOT, "microsoft-lifecycle")
NORMALIZED_FILE = PATHS.NORMALIZED_FILE
HISTORY_FILE = PATHS.HISTORY_FILE
REJECTED_FILE = PATHS.REJECTED_FILE
STATISTICS_FILE = PATHS.STATISTICS_FILE
STATISTICS_HISTORY_FILE = PATHS.STATISTICS_HISTORY_FILE
WARNINGS_FILE = PATHS.WARNINGS_FILE
LOG_FILE = PATHS.LOG_FILE


def download_workbook(destination: Path) -> Path:
    url = os.getenv("LIFECYCLE_SOURCE_URL", MICROSOFT_XLSX_URL)
    parsed = urlsplit(url)
    if parsed.scheme != "https" or parsed.hostname != "download.microsoft.com" or parsed.username or parsed.password:
        raise ValueError("Lifecycle source must be a Microsoft HTTPS download")
    session = requests.Session()
    retries = Retry(total=4, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504], respect_retry_after_header=True)
    session.mount("https://", HTTPAdapter(max_retries=retries))
    response = session.get(url, timeout=(15, 120), allow_redirects=True)
    response.raise_for_status()
    if urlsplit(response.url).scheme != "https":
        raise ValueError("Microsoft lifecycle redirected to a non-HTTPS address")
    if len(response.content) > 30 * 1024 * 1024:
        raise ValueError("Lifecycle workbook exceeded the 30 MB size limit")
    if not response.content.startswith(b"PK\x03\x04"):
        raise ValueError("Microsoft lifecycle response is not an XLSX workbook")
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile("wb", dir=destination.parent, suffix=".xlsx.tmp", delete=False) as file:
        file.write(response.content)
        temp_path = Path(file.name)
    try:
        # Verify the file is parseable *before* replacing the last valid raw copy.
        pd.read_excel(temp_path, engine="openpyxl", nrows=1)
        os.replace(temp_path, destination)
    finally:
        temp_path.unlink(missing_ok=True)
    return destination


def parse_workbook(path: Path) -> list[dict]:
    df = pd.read_excel(path, engine="openpyxl")
    if df.empty or "Product Listing Name" not in df.columns:
        raise RuntimeError("Unexpected lifecycle workbook: missing Product Listing Name or no records")
    df = df.astype(object).where(pd.notna(df), None)
    return df.to_dict(orient="records")


def date_only(value) -> str | None:
    val = iso_utc(value)
    return val[:10] if val else None


def normalize_lifecycle_status(source_status: str | None, effective_eos: str | None, snapshot: str) -> str:
    raw = (source_status or "").casefold().replace("_", " ").strip()
    if raw in {"out of support", "end of support", "retired", "no longer supported"}:
        return "END_OF_SUPPORT"
    if not effective_eos:
        return "SUPPORTED" if raw in {"supported", "in support"} else "UNKNOWN"
    days = (date.fromisoformat(effective_eos) - date.fromisoformat(snapshot)).days
    if days <= 0:
        return "END_OF_SUPPORT"
    if days <= 90:
        return "EOS_LT_3_MONTHS"
    if days <= 180:
        return "EOS_LT_6_MONTHS"
    if days <= 365:
        return "EOS_LT_12_MONTHS"
    return "SUPPORTED"


def map_lifecycle(item: dict, snapshot: str, ingested: str) -> dict:
    product = clean(get(item, "Product Listing Name", "Product"))
    edition = clean(get(item, "Edition", "Product Edition", "Version"))
    release = clean(get(item, "Release", "Release Number"))
    feature = clean(get(item, "Azure Feature", "AzureFeature"))
    policy = clean(get(item, "Support Policy", "LifecyclePolicy", "Lifecycle Policy"))
    original_status = clean(get(item, "End of Support", "LifecycleStatus"))
    start = date_only(get(item, "Release Start Date", "Start Date", "StartDate"))
    mainstream_end = date_only(get(item, "Mainstream End Date", "MainstreamEndDate"))
    extended_end = date_only(get(item, "Extended End Date", "ExtendedEndDate"))
    release_end = date_only(get(item, "Release End Date", "ReleaseEndDate"))
    retirement = date_only(get(item, "Retirement Date", "Retirement"))
    # Release end date is release-specific; other policy dates are fallbacks.
    # Do not substitute 2099 for missing dates.
    effective_eos = release_end or extended_end or retirement
    status = normalize_lifecycle_status(original_status, effective_eos, snapshot)
    identity = [product or "", edition or "", release or "", feature or ""]
    release_id = "lifecycle:" + digest(identity)[:32] if product else None
    details = {
        "Vendor": "Microsoft", "Edition": edition, "Release": release,
        "AzureFeature": feature, "LifecyclePolicy": policy,
        "StartDate": start, "MainstreamEndDate": mainstream_end,
        "ExtendedEndDate": extended_end, "ReleaseEndDate": release_end,
        "RetirementDate": retirement, "EndOfSupport": effective_eos,
        "OriginalSupportStatus": original_status,
    }
    row = create_record(
        source=SOURCE, release_id=release_id,
        workload="EUC", product=product, service="MICROSOFT_LIFECYCLE",
        category=policy or "PRODUCT_LIFECYCLE", release_type="LIFECYCLE",
        release_status=status,
        title=" — ".join(x for x in (product, edition, release) if x) or None,
        description=None,  # Description is now plain text, not serialized JSON.
        # The source has no independent announcement timestamp; preserve the
        # release/start-date proxy in PublishedDate and exact value in StartDate.
        published=start,
        impact=None, action_required=status in {"END_OF_SUPPORT", "EOS_LT_3_MONTHS"},
        url=get(item, "DocsUrl", "DocumentationUrl", "url"),
        snapshot=snapshot, ingested=ingested,
    )
    # Never stringify lifecycle details into Description: each key has its
    # own stable column with null represented as an empty CSV cell.
    row.update(details)
    row["ContentHash"] = business_content_hash(row)
    return row


def migrate_lifecycle_history(
    old_rows: list[dict[str, str]], old_columns: list[str], new_columns: list[str],
) -> list[dict] | None:
    """Losslessly expand the previous 20-column, JSON-in-Description history.

    A strict archive is made by release_contract *before* this function runs.
    Unknown JSON keys refuse migration rather than silently dropping data.
    """
    if old_columns != BASE_OUTPUT_COLUMNS or new_columns != OUTPUT_COLUMNS:
        return None
    converted = []
    for old in old_rows:
        raw_description = old.get("Description", "")
        details = {}
        if raw_description and raw_description.strip():
            try:
                payload = json.loads(raw_description)
            except json.JSONDecodeError:
                payload = None  # Preserve genuine free-text descriptions.
            if isinstance(payload, dict):
                unknown = set(payload) - set(LIFECYCLE_DETAIL_COLUMNS)
                if unknown:
                    raise ValueError(f"Unmapped legacy JSON keys: {sorted(unknown)}")
                details = payload
        row = dict(old)
        if details:
            row["Description"] = None
        # Restore the in-memory types used by create_record, so the migration
        # does not generate a spurious content-hash change on an identical row.
        row["ActionRequired"] = str(row.get("ActionRequired", "")).lower() in {"true", "1"}
        for key in BASE_OUTPUT_COLUMNS:
            if key not in {"ActionRequired"} and row.get(key) == "":
                row[key] = None
        for key in LIFECYCLE_DETAIL_COLUMNS:
            row[key] = details.get(key)
        row["ContentHash"] = business_content_hash(row)
        converted.append({key: row.get(key, "") for key in OUTPUT_COLUMNS})
    return converted


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, help="Use a local Microsoft XLSX workbook (offline testing)")
    parser.add_argument("--fixture-json", type=Path, help="Use JSON array of source workbook rows (offline testing)")
    args = parser.parse_args(argv)
    if args.input and args.fixture_json:
        parser.error("Use either --input or --fixture-json")
    def extract():
        if args.fixture_json:
            content = json.loads(args.fixture_json.read_text(encoding="utf-8"))
            if not isinstance(content, list):
                raise ValueError("Expected JSON array of lifecycle source records")
            return content
        workbook = args.input or download_workbook(ROOT / "raw" / "product-lifecycle.xlsx")
        return parse_workbook(workbook)
    return run_pipeline(
        source=SOURCE, paths=PATHS, extractor=extract, mapper=map_lifecycle,
        output_columns=OUTPUT_COLUMNS, history_migrator=migrate_lifecycle_history,
    )


if __name__ == "__main__":
    main()
