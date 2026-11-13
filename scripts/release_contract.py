"""Shared Power BI base contract with optional source-specific extension columns.

This module contains no source retrieval or source-specific business mapping.
Every exporter remains a separate script and publishes its own dataset.
"""
from __future__ import annotations

import csv
import hashlib
import html
import io
import json
import logging
import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from datetime import date, datetime, time, timezone
from pathlib import Path
from typing import Any, Callable, Iterable
from urllib.parse import urlsplit

SCHEMA_VERSION = "2.1"
OUTPUT_COLUMNS = [
    "RecordId", "ReleaseId", "Source", "Workload", "Product", "Service",
    "Category", "ReleaseType", "ReleaseStatus", "Title", "Description",
    "PublishedDate", "LastModifiedDate", "ImpactLevel", "ActionRequired",
    "Url", "SnapshotDate", "IngestedAtUtc", "ContentHash", "SchemaVersion",
]
REJECTED_COLUMNS = ["RejectedAtUtc", "Source", "SourceIndex", "ErrorCodes", "ErrorMessage", "RawRecordJson"]
WARNING_COLUMNS = ["WarningAtUtc", "Source", "WarningCode", "WarningMessage"]
STATISTICS_COLUMNS = [
    "RunAtUtc", "SnapshotDate", "Source", "Status", "RawRecords", "AcceptedRecords",
    "RejectedRecords", "WarningCount", "RejectionRate", "SchemaVersion",
]
CSV_SEPARATOR = ";"
LOG = logging.getLogger("release_export")


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def utc_string(value: datetime | None = None) -> str:
    return (value or utc_now()).astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def iso_utc(value: Any) -> str | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None  # ambiguous numeric dates require source-specific handling
    if isinstance(value, datetime):
        d = value
    elif isinstance(value, date):
        d = datetime.combine(value, time.min)
    else:
        text = str(value).strip()
        if not text or text.lower() in {"nat", "nan", "null", "none"}:
            return None
        if text in {"31/12/2099", "2099-12-31"}:
            return None  # fabricated historical placeholder
        try:
            d = datetime.fromisoformat(text.replace("Z", "+00:00"))
        except ValueError:
            return None
    if d.tzinfo is None:
        d = d.replace(tzinfo=timezone.utc)
    return utc_string(d)


def clean(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
    text = str(value).strip()
    if not text or text.lower() in {"nat", "nan", "none", "null"}:
        return None
    text = html.unescape(re.sub(r"<[^>]*>", " ", text))
    return re.sub(r"\s+", " ", text).strip() or None


def get(item: dict[str, Any], *aliases: str) -> Any:
    """Find the first nonempty alias, including nested dotted paths and casing drift."""
    for alias in aliases:
        current: Any = item
        for segment in alias.split("."):
            if not isinstance(current, dict):
                current = None
                break
            if segment in current:
                current = current[segment]
            else:
                normalized = re.sub(r"[^a-z0-9]", "", segment.lower())
                actual = next((k for k in current if re.sub(r"[^a-z0-9]", "", str(k).lower()) == normalized), None)
                current = current.get(actual) if actual is not None else None
        if current is None or (isinstance(current, str) and (not current.strip() or current.strip().lower() in {"none", "null", "nan", "nat"})):
            continue
        if isinstance(current, float) and current != current:  # NaN
            continue
        if isinstance(current, list) and not current:
            continue
        if str(current).strip() in {"NaT", "<NA>"}:
            continue
        return current
    return None


def join_values(value: Any) -> str | None:
    if isinstance(value, list):
        vals = []
        for item in value:
            if isinstance(item, dict):
                item = get(item, "name", "displayName", "value") or item
            v = clean(item)
            if v and v not in vals:
                vals.append(v)
        return "; ".join(vals) or None
    return clean(value)


def https_url(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    parsed = urlsplit(value)
    if parsed.scheme == "https" and parsed.hostname and not parsed.username and not parsed.password:
        return value
    return None


def digest(value: Any) -> str:
    data = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(data.encode("utf-8")).hexdigest()


def business_content_hash(row: dict[str, Any]) -> str:
    """Hash source business attributes, excluding ingestion metadata and timestamps."""
    return digest({k: v for k, v in row.items() if k not in {
        "RecordId", "ContentHash", "SnapshotDate", "IngestedAtUtc", "SchemaVersion",
    }})


def create_record(
    *, source: str, release_id: str | None, title: str | None,
    workload: Any = None, product: Any = None, service: Any = None,
    category: Any = None, release_type: Any = None, release_status: Any = None,
    description: Any = None, published: Any = None, modified: Any = None,
    impact: Any = None, action_required: bool = False, url: Any = None,
    snapshot: str | None = None, ingested: str | None = None,
) -> dict[str, Any]:
    identity = clean(release_id) or ""
    row = {
        "RecordId": digest([source, identity]) if identity else "",
        "ReleaseId": identity, "Source": source, "Workload": join_values(workload),
        "Product": join_values(product), "Service": join_values(service),
        "Category": join_values(category), "ReleaseType": clean(release_type) or "OTHER",
        "ReleaseStatus": clean(release_status) or "UNKNOWN",
        "Title": clean(title), "Description": clean(description),
        "PublishedDate": iso_utc(published), "LastModifiedDate": iso_utc(modified),
        "ImpactLevel": clean(impact), "ActionRequired": bool(action_required),
        "Url": https_url(url), "SnapshotDate": snapshot or utc_now().date().isoformat(),
        "IngestedAtUtc": ingested or utc_string(), "ContentHash": "", "SchemaVersion": SCHEMA_VERSION,
    }
    # SnapshotDate and IngestedAtUtc must NOT cause false business changes.
    row["ContentHash"] = business_content_hash(row)
    return row


def validate_record(row: dict[str, Any], columns: list[str] | None = None) -> list[str]:
    columns = columns or OUTPUT_COLUMNS
    errors = []
    if list(row) != columns:
        errors.append("OUTPUT_SCHEMA_MISMATCH")
    for col in ("RecordId", "ReleaseId", "Source", "Title", "ContentHash"):
        if not row.get(col):
            errors.append("MISSING_" + col.upper())
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", row.get("SnapshotDate", "")):
        errors.append("INVALID_SNAPSHOT_DATE")
    for col in ("PublishedDate", "LastModifiedDate", "IngestedAtUtc"):
        val = row.get(col)
        if val is not None and not re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z", str(val)):
            errors.append("INVALID_" + col.upper())
    return errors


@dataclass(frozen=True)
class Paths:
    NORMALIZED_FILE: Path
    HISTORY_FILE: Path
    REJECTED_FILE: Path
    STATISTICS_FILE: Path
    STATISTICS_HISTORY_FILE: Path
    WARNINGS_FILE: Path
    LOG_FILE: Path


def output_paths(root: Path, slug: str) -> Paths:
    return Paths(
        NORMALIZED_FILE=root / "data" / "normalized" / f"{slug}.csv",
        HISTORY_FILE=root / "data" / "history" / f"{slug}-history.csv",
        REJECTED_FILE=root / "data" / "rejected" / f"{slug}-rejected.csv",
        STATISTICS_FILE=root / "data" / "metadata" / f"{slug}-statistics.csv",
        STATISTICS_HISTORY_FILE=root / "data" / "metadata" / f"{slug}-statistics-history.csv",
        WARNINGS_FILE=root / "data" / "metadata" / f"{slug}-warnings.csv",
        LOG_FILE=root / "data" / "logs" / f"{slug}.log",
    )


def csv_safe(col: str, value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value).lower()
    s = str(value)
    if col not in {"RecordId", "ReleaseId", "ContentHash", "SchemaVersion", "PublishedDate", "LastModifiedDate", "SnapshotDate", "IngestedAtUtc", "EndOfSupport", "StartDate", "MainstreamEndDate", "ExtendedEndDate", "ReleaseEndDate", "RetirementDate"} and s.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + s
    return s


def serialize_csv(records: list[dict[str, Any]], columns: list[str]) -> str:
    buf = io.StringIO(newline="")
    writer = csv.DictWriter(buf, fieldnames=columns, delimiter=CSV_SEPARATOR, lineterminator="\n", extrasaction="raise")
    writer.writeheader()
    for row in records:
        writer.writerow({key: csv_safe(key, row.get(key)) for key in columns})
    return "\ufeff" + buf.getvalue()


class CsvSchemaMismatch(RuntimeError):
    """A stored CSV has a header incompatible with the active data contract."""


def read_csv(path: Path, columns: list[str]) -> list[dict[str, str]]:
    """Strict reader for the current normalized CSV contract.

    This intentionally does NOT migrate arbitrary CSV files when validating
    freshly published output. Only historical sidecars may be migrated.
    """
    if not path.exists():
        return []
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream, delimiter=CSV_SEPARATOR)
        if reader.fieldnames != columns:
            raise CsvSchemaMismatch(
                f"CSV schema mismatch at {path}; expected {len(columns)} columns, "
                f"found {len(reader.fieldnames or [])}"
            )
        rows = list(reader)
        if any(None in row for row in rows):
            raise RuntimeError(f"Malformed CSV rows in {path}; refusing overwrite")
        return rows


def read_historical_csv(
    path: Path,
    columns: list[str],
    *,
    warn: Callable[[str, str], None],
    migrate: Callable[[list[dict[str, str]], list[str], list[str]], list[dict[str, Any]] | None] | None = None,
) -> list[dict[str, str]]:
    """Migrate incompatible *historical* sidecars without losing old data.

    A legacy CSV is copied to data/legacy/ before any replacement. If it has
    precisely the expected column names in a different order or uses comma
    delimiters, its records can be preserved safely. Otherwise, the new-format
    history starts afresh, while the entire old history remains in the backup.
    """
    try:
        return read_csv(path, columns)
    except CsvSchemaMismatch:
        pass

    # Preserve source bytes; a digest-based file name makes retries idempotent.
    raw = path.read_bytes()
    digest_suffix = hashlib.sha256(raw).hexdigest()[:16]
    backup = path.parents[1] / "legacy" / f"{path.stem}.schema-{digest_suffix}{path.suffix}"
    backup.parent.mkdir(parents=True, exist_ok=True)
    if backup.exists():
        if backup.read_bytes() != raw:
            raise RuntimeError(f"Legacy backup collision at {backup}; refusing overwrite")
    else:
        shutil.copy2(path, backup)
    if backup.read_bytes() != raw:
        raise RuntimeError(f"Could not verify legacy history backup: {backup}")

    # Only migrate with a verified backup; map known legacy schemas explicitly.
    decoded = raw.decode("utf-8-sig")
    for delimiter in (CSV_SEPARATOR, ","):
        reader = csv.DictReader(io.StringIO(decoded), delimiter=delimiter)
        header = reader.fieldnames
        if header is None:
            continue
        rows = list(reader)
        if any(None in row for row in rows):
            raise RuntimeError(f"Malformed legacy history at {path}; backup at {backup}")
        if len(header) == len(columns) and set(header) == set(columns):
            warn(
                "LEGACY_CSV_REORDERED",
                f"Converted history column order/delimiter for {path}; original backed up to {backup}",
            )
            return [{column: row.get(column, "") for column in columns} for row in rows]
        if migrate is not None and delimiter == CSV_SEPARATOR:
            try:
                upgraded = migrate(rows, header, columns)
                if upgraded is not None:
                    if any(list(row) != columns for row in upgraded):
                        raise ValueError("Migrated history has unexpected columns")
                    warn("LEGACY_CSV_EXPANDED", f"Migrated {len(upgraded)} historical rows into expanded lifecycle columns; backup: {backup}")
                    return upgraded
            except (ValueError, TypeError, KeyError, json.JSONDecodeError) as e:
                warn("LEGACY_CSV_MIGRATION_SKIPPED", f"Could not safely migrate {path}: {e}; preserving backup at {backup}")

    warn(
        "LEGACY_CSV_ARCHIVED",
        f"Incompatible historical schema at {path}; old file preserved at {backup}; "
        "starting a new history with the exact current schema",
    )
    return []


def atomic_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", newline="", dir=path.parent, prefix=path.name + ".", suffix=".tmp", delete=False) as out:
        name = out.name
        out.write(content)
    os.replace(name, path)


def log_setup(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    LOG.setLevel(logging.INFO)
    LOG.handlers.clear()
    formatter = logging.Formatter("%(asctime)sZ %(levelname)s %(message)s")
    file_handler = logging.FileHandler(path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    LOG.addHandler(file_handler)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    LOG.addHandler(stream_handler)


def run_pipeline(
    *, source: str, paths: Paths,
    extractor: Callable[[], Iterable[dict[str, Any]]],
    mapper: Callable[[dict[str, Any], str, str], dict[str, Any]],
    rejection_limit: float = 0.05,
    output_columns: list[str] | None = None,
    history_migrator: Callable[[list[dict[str, str]], list[str], list[str]], list[dict[str, Any]] | None] | None = None,
) -> dict[str, Any]:
    """Fetch, normalize, validate; publish the current CSV *last* if all gates pass."""
    output_columns = output_columns or OUTPUT_COLUMNS
    if len(output_columns) != len(set(output_columns)) or output_columns[:len(OUTPUT_COLUMNS)] != OUTPUT_COLUMNS:
        raise ValueError("Source output columns must extend the ordered base contract without duplicates")
    log_setup(paths.LOG_FILE)
    snapshot, ingested = utc_now().date().isoformat(), utc_string()
    warnings: list[dict[str, str]] = []
    rejected: list[dict[str, str]] = []
    accepted: list[dict[str, Any]] = []
    raw_count = 0
    status = "FAILED"
    error_message = None

    def warn(code: str, message: str):
        warnings.append({"WarningAtUtc": utc_string(), "Source": source, "WarningCode": code, "WarningMessage": message})
        LOG.warning("%s %s", code, message)

    try:
        records = extractor()  # failure must NOT fall back silently to stale data
        seen: set[str] = set()
        for index, item in enumerate(records):
            # Explicit source diagnostics do not become rejected business rows.
            if isinstance(item, dict) and "__pipeline_warning__" in item:
                detail = item["__pipeline_warning__"]
                warn(str(detail.get("code", "SOURCE_WARNING")), str(detail.get("message", "Source data warning")))
                continue
            raw_count += 1
            try:
                if not isinstance(item, dict):
                    raise ValueError("Record is not a JSON object")
                row = mapper(item, snapshot, ingested)
                errors = validate_record(row, output_columns)
                if errors:
                    raise ValueError(",".join(errors))
                if row["RecordId"] in seen:
                    raise ValueError("DUPLICATE_RECORD_ID")
                seen.add(row["RecordId"])
                accepted.append(row)
            except Exception as e:
                rejected.append({
                    "RejectedAtUtc": utc_string(), "Source": source, "SourceIndex": str(index),
                    "ErrorCodes": str(e).split(":", 1)[0][:100], "ErrorMessage": str(e)[:500],
                    "RawRecordJson": json.dumps(item, ensure_ascii=False, sort_keys=True, default=str),
                })
        if not accepted:
            raise RuntimeError("NO_VALID_RECORDS")
        if len(rejected) / max(raw_count, 1) > rejection_limit:
            raise RuntimeError(f"REJECTION_THRESHOLD_EXCEEDED: {len(rejected)}/{raw_count}")
        if rejected:
            warn("REJECTED_RECORDS", f"{len(rejected)} of {raw_count} source records rejected")
        accepted.sort(key=lambda r: (r["Source"], r["ReleaseId"]))
        history = read_historical_csv(paths.HISTORY_FILE, output_columns, warn=warn, migrate=history_migrator)
        # Change-only history prevents unchanged daily snapshots inflating Git history.
        # Statistics-history still records every daily run. A reverted value is a change.
        history_rows = list(history)
        latest: dict[str, dict[str, Any]] = {}
        for old in sorted(history, key=lambda x: (x["SnapshotDate"], x["IngestedAtUtc"])):
            latest[old["RecordId"]] = old
        for row in accepted:
            previous = latest.get(row["RecordId"])
            if previous is None or previous["ContentHash"] != row["ContentHash"]:
                history_rows.append(row)
        # Save history first, current file last. A source failure never publishes stale data.
        atomic_text(paths.HISTORY_FILE, serialize_csv(history_rows, output_columns))
        status = "SUCCESS"
    except Exception as e:
        error_message = str(e)
        LOG.exception("Pipeline failed: %s", e)
        warn("PIPELINE_FAILED", error_message[:400])
        rejected.append({
            "RejectedAtUtc": utc_string(), "Source": source, "SourceIndex": "-1",
            "ErrorCodes": "PIPELINE_FAILED", "ErrorMessage": error_message[:500], "RawRecordJson": "",
        })
    finally:
        # Read/migrate statistics history separately from the primary export.
        # A second schema mismatch must never hide the initial failure.
        try:
            previous_stats = read_historical_csv(
                paths.STATISTICS_HISTORY_FILE, STATISTICS_COLUMNS, warn=warn
            )
        except Exception as diagnostic_error:
            LOG.exception("Statistics history recovery failed: %s", diagnostic_error)
            warn("DIAGNOSTIC_RECOVERY_FAILED", str(diagnostic_error)[:400])
            previous_stats = []
            # Do not overwrite a history that we could not safely preserve.
            preserve_stats_history = True
            if status == "SUCCESS":
                status = "FAILED"
                error_message = f"Statistics history recovery failed: {diagnostic_error}"
        else:
            preserve_stats_history = False

        stats = {
            "RunAtUtc": utc_string(), "SnapshotDate": snapshot, "Source": source,
            "Status": status, "RawRecords": raw_count, "AcceptedRecords": len(accepted),
            "RejectedRecords": len(rejected), "WarningCount": len(warnings),
            "RejectionRate": round(len([r for r in rejected if r["SourceIndex"] != "-1"]) / max(raw_count, 1), 6),
            "SchemaVersion": SCHEMA_VERSION,
        }
        # Diagnostics are available even on a failing execution. Write each
        # sidecar independently so errors in one do not mask a primary failure.
        for path, records, columns in (
            (paths.REJECTED_FILE, rejected, REJECTED_COLUMNS),
            (paths.WARNINGS_FILE, warnings, WARNING_COLUMNS),
            (paths.STATISTICS_FILE, [stats], STATISTICS_COLUMNS),
        ):
            try:
                atomic_text(path, serialize_csv(records, columns))
            except Exception:
                LOG.exception("Could not write diagnostic file: %s", path)
                if status == "SUCCESS":
                    status = "FAILED"
                    error_message = f"Could not write diagnostic file: {path}"
        if not preserve_stats_history:
            try:
                atomic_text(
                    paths.STATISTICS_HISTORY_FILE,
                    serialize_csv(previous_stats + [stats], STATISTICS_COLUMNS),
                )
            except Exception:
                LOG.exception("Could not write statistics history")
                if status == "SUCCESS":
                    status = "FAILED"
                    error_message = "Could not write statistics history"
    if status != "SUCCESS":
        raise RuntimeError(error_message or "Pipeline failed")
    atomic_text(paths.NORMALIZED_FILE, serialize_csv(accepted, output_columns))
    LOG.info("Published %s records to %s", len(accepted), paths.NORMALIZED_FILE)
    return stats
