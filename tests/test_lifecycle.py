"""Regression tests for the independent lifecycle exporter and expanded CSV."""
import csv
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))

import release_contract as rc
import transform


def pipeline(tmp_path, rows, **kwargs):
    paths = rc.output_paths(tmp_path, 'microsoft-lifecycle')
    outcome = rc.run_pipeline(
        source=transform.SOURCE,
        paths=paths,
        extractor=lambda: rows,
        mapper=transform.map_lifecycle,
        output_columns=transform.OUTPUT_COLUMNS,
        history_migrator=transform.migrate_lifecycle_history,
        **kwargs,
    )
    return paths, outcome


def test_exact_schema_and_json_keys_are_columns():
    assert len(rc.OUTPUT_COLUMNS) == 20
    assert len(transform.LIFECYCLE_DETAIL_COLUMNS) == 12
    assert len(transform.OUTPUT_COLUMNS) == 32
    assert transform.OUTPUT_COLUMNS[:20] == rc.OUTPUT_COLUMNS
    assert transform.OUTPUT_COLUMNS[20:] == transform.LIFECYCLE_DETAIL_COLUMNS
    assert len(set(transform.OUTPUT_COLUMNS)) == len(transform.OUTPUT_COLUMNS)


def test_lifecycle_status_rules():
    assert transform.normalize_lifecycle_status('Out of Support', None, '2026-11-13') == 'END_OF_SUPPORT'
    assert transform.normalize_lifecycle_status(None, None, '2026-11-13') == 'UNKNOWN'
    assert transform.normalize_lifecycle_status(None, '2026-12-01', '2026-11-13') == 'EOS_LT_3_MONTHS'
    assert transform.normalize_lifecycle_status(None, '2027-11-14', '2026-11-13') == 'SUPPORTED'


def test_description_is_not_json_and_missing_eos_never_fabricated():
    item = {'Product Listing Name': 'Azure A', 'Edition': 'Standard', 'Release': '1.0',
            'Support Policy': 'Modern', 'Start Date': '2025-01-01', 'End of Support': 'Out of Support'}
    row = transform.map_lifecycle(item, '2026-11-13', '2026-11-13T12:00:00Z')
    assert row['ReleaseStatus'] == 'END_OF_SUPPORT'
    assert row['Description'] is None
    assert row['EndOfSupport'] is None
    assert row['OriginalSupportStatus'] == 'Out of Support'
    assert row['Edition'] == 'Standard'
    assert row['Vendor'] == 'Microsoft'
    assert row['PublishedDate'] == '2025-01-01T00:00:00Z'
    assert row['StartDate'] == '2025-01-01'
    assert list(row) == transform.OUTPUT_COLUMNS


def test_stable_identity_and_hash_for_detail_changes():
    item = {'Product Listing Name': 'Windows 11', 'Edition': 'Enterprise', 'Release': '24H2',
            'Release End Date': '2027-01-01'}
    a = transform.map_lifecycle(item, '2026-11-13', '2026-11-13T12:00:00Z')
    item['Release End Date'] = '2028-01-01'
    b = transform.map_lifecycle(item, '2026-11-14', '2026-11-14T12:00:00Z')
    assert a['RecordId'] == b['RecordId']
    assert a['ContentHash'] != b['ContentHash']
    item['Release End Date'] = '2027-01-01'
    c = transform.map_lifecycle(item, '2026-11-15', '2026-11-15T12:00:00Z')
    assert c['ContentHash'] == a['ContentHash']


def test_csv_has_extended_columns_and_invalid_rows_are_rejected(tmp_path):
    data = [{'Product Listing Name': 'Windows A', 'Release End Date': '2028-01-01',
             'Edition': 'Enterprise'}, {'Product Listing Name': None}]
    paths, stats = pipeline(tmp_path, data, rejection_limit=0.60)
    result = rc.read_csv(paths.NORMALIZED_FILE, transform.OUTPUT_COLUMNS)
    assert stats['AcceptedRecords'] == 1
    assert stats['RejectedRecords'] == 1
    assert result[0]['Edition'] == 'Enterprise'
    assert result[0]['EndOfSupport'] == '2028-01-01'
    assert result[0]['Description'] == ''
    assert len(rc.read_csv(paths.REJECTED_FILE, rc.REJECTED_COLUMNS)) == 1
    assert list(result[0]) == transform.OUTPUT_COLUMNS
    assert paths.NORMALIZED_FILE.read_bytes().startswith(b'\xef\xbb\xbf')


def test_migration_preserves_history_and_archives_legacy_json(tmp_path):
    paths = rc.output_paths(tmp_path, 'microsoft-lifecycle')
    item = {'Product Listing Name': 'Windows A', 'Edition': 'Enterprise',
            'Release': '24H2', 'Release End Date': '2028-01-01',
            'Start Date': '2024-10-01', 'Support Policy': 'Modern'}
    mapped = transform.map_lifecycle(item, '2026-11-13', '2026-11-13T12:00:00Z')
    original = {k: mapped[k] for k in rc.OUTPUT_COLUMNS}
    original['Description'] = json.dumps({k: mapped[k] for k in transform.LIFECYCLE_DETAIL_COLUMNS})
    original['ContentHash'] = 'legacy-hash'
    rc.atomic_text(paths.HISTORY_FILE, rc.serialize_csv([original], rc.OUTPUT_COLUMNS))
    backup_bytes = paths.HISTORY_FILE.read_bytes()
    pipeline(tmp_path, [item])
    backups = list((tmp_path/'data'/'legacy').glob('microsoft-lifecycle-history.schema-*.csv'))
    assert len(backups) == 1
    assert backups[0].read_bytes() == backup_bytes
    history = rc.read_csv(paths.HISTORY_FILE, transform.OUTPUT_COLUMNS)
    assert len(history) == 1, 'An unchanged legacy record should not create a false new revision'
    assert history[0]['Edition'] == 'Enterprise'
    assert history[0]['Description'] == ''
    warnings = rc.read_csv(paths.WARNINGS_FILE, rc.WARNING_COLUMNS)
    assert any(w['WarningCode'] == 'LEGACY_CSV_EXPANDED' for w in warnings)


def test_second_run_does_not_duplicate_unchanged_history(tmp_path):
    item = {'Product Listing Name': 'Windows A', 'Release End Date': '2028-01-01'}
    paths, _ = pipeline(tmp_path, [item])
    pipeline(tmp_path, [item])
    assert len(rc.read_csv(paths.HISTORY_FILE, transform.OUTPUT_COLUMNS)) == 1


def test_missing_source_preserves_last_good_csv(tmp_path):
    paths, _ = pipeline(tmp_path, [{'Product Listing Name': 'Windows A'}])
    good = paths.NORMALIZED_FILE.read_bytes()
    with pytest.raises(RuntimeError, match='download failed'):
        rc.run_pipeline(
            source=transform.SOURCE, paths=paths,
            extractor=lambda: (_ for _ in ()).throw(RuntimeError('download failed')),
            mapper=transform.map_lifecycle,
            output_columns=transform.OUTPUT_COLUMNS,
            history_migrator=transform.migrate_lifecycle_history,
        )
    assert paths.NORMALIZED_FILE.read_bytes() == good
    assert rc.read_csv(paths.STATISTICS_FILE, rc.STATISTICS_COLUMNS)[0]['Status'] == 'FAILED'


def test_csv_formula_injection_is_escaped_in_edition(tmp_path):
    paths, _ = pipeline(tmp_path, [{'Product Listing Name': 'Windows A', 'Edition': '=HYPERLINK("https://example.org")'}])
    row = rc.read_csv(paths.NORMALIZED_FILE, transform.OUTPUT_COLUMNS)[0]
    assert row['Edition'].startswith("'=")


def test_workbook_fixture_is_real_and_parseable():
    src = Path(__file__).resolve().parents[1] / 'raw' / 'product-lifecycle.xlsx'
    assert src.is_file()
    rows = transform.parse_workbook(src)
    assert len(rows) > 100
    assert 'Product Listing Name' in rows[0]
