import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import release_contract as rc
import transform


def test_exact_schema():
    assert len(rc.OUTPUT_COLUMNS) == 20
    assert rc.OUTPUT_COLUMNS[0:3] == ["RecordId", "ReleaseId", "Source"]
    assert rc.OUTPUT_COLUMNS[-2:] == ["ContentHash", "SchemaVersion"]


def test_lifecycle_status_rules():
    assert transform.normalize_lifecycle_status("Out of Support", None, "2026-11-13") == "END_OF_SUPPORT"
    assert transform.normalize_lifecycle_status(None, None, "2026-11-13") == "UNKNOWN"
    assert transform.normalize_lifecycle_status(None, "2026-12-01", "2026-11-13") == "EOS_LT_3_MONTHS"
    assert transform.normalize_lifecycle_status(None, "2027-11-14", "2026-11-13") == "SUPPORTED"


def test_missing_eos_never_fabricated():
    item = {"Product Listing Name": "Azure A", "Edition": "Standard", "Release": "1.0",
            "Support Policy": "Modern", "Start Date": "2025-01-01", "End of Support": "Out of Support"}
    row = transform.map_lifecycle(item, "2026-11-13", "2026-11-13T12:00:00Z")
    assert row["ReleaseStatus"] == "END_OF_SUPPORT"
    assert '"EndOfSupport": null' in row["Description"]
    assert "2099" not in row["Description"]
    assert row["PublishedDate"] == "2025-01-01T00:00:00Z"


def test_stable_identity_across_updated_support_dates():
    item = {"Product Listing Name": "Windows 11", "Edition": "Enterprise", "Release": "24H2", "Release End Date": "2027-01-01"}
    a = transform.map_lifecycle(item, "2026-11-13", "2026-11-13T12:00:00Z")
    item["Release End Date"] = "2028-01-01"
    b = transform.map_lifecycle(item, "2026-11-14", "2026-11-14T12:00:00Z")
    assert a["RecordId"] == b["RecordId"]
    assert a["ContentHash"] != b["ContentHash"]
    assert list(a) == rc.OUTPUT_COLUMNS


def test_missing_product_rejected(tmp_path):
    paths = rc.output_paths(tmp_path, "euc")
    data = [{"Product Listing Name": "Windows A", "Release End Date": "2028-01-01"}, {"Product Listing Name": None}]
    rc.run_pipeline(source="EUC_LIFECYCLE", paths=paths, extractor=lambda: data, mapper=transform.map_lifecycle, rejection_limit=0.60)
    assert len(rc.read_csv(paths.NORMALIZED_FILE, rc.OUTPUT_COLUMNS)) == 1
    assert len(rc.read_csv(paths.REJECTED_FILE, rc.REJECTED_COLUMNS)) == 1
    assert list(rc.read_csv(paths.NORMALIZED_FILE, rc.OUTPUT_COLUMNS)[0]) == rc.OUTPUT_COLUMNS


def test_source_missing_preserves_last_good(tmp_path):
    paths = rc.output_paths(tmp_path, "euc")
    rc.run_pipeline(source="EUC_LIFECYCLE", paths=paths,
                    extractor=lambda: [{"Product Listing Name": "Windows A"}], mapper=transform.map_lifecycle)
    good = paths.NORMALIZED_FILE.read_bytes()
    with pytest.raises(RuntimeError, match="download failed"):
        rc.run_pipeline(source="EUC_LIFECYCLE", paths=paths,
                        extractor=lambda: (_ for _ in ()).throw(RuntimeError("download failed")), mapper=transform.map_lifecycle)
    assert paths.NORMALIZED_FILE.read_bytes() == good
    assert rc.read_csv(paths.STATISTICS_FILE, rc.STATISTICS_COLUMNS)[0]["Status"] == "FAILED"


def test_workbook_fixture_is_real_and_parseable():
    src = Path(__file__).resolve().parents[1] / "raw" / "product-lifecycle.xlsx"
    assert src.is_file()
    values = transform.parse_workbook(src)
    assert len(values) > 100
    assert "Product Listing Name" in values[0]
