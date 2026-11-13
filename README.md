# Microsoft EUC lifecycle — schema-tolerant Power BI export

This repository retains its **independent** Python script (`scripts/transform.py`) and publishes its own 20-column Power BI CSV. No consolidation with Windows Update Catalog, Windows Release Health or Microsoft Release Communications takes place.

Microsoft's lifecycle source is an XLSX workbook, **not a JSON endpoint**. The script downloads the workbook over HTTPS, validates it, parses rows into JSON-compatible records and normalizes them using the same record contract as Microsoft Release Communications.

## Exact 20-column contract

`RecordId; ReleaseId; Source; Workload; Product; Service; Category; ReleaseType; ReleaseStatus; Title; Description; PublishedDate; LastModifiedDate; ImpactLevel; ActionRequired; Url; SnapshotDate; IngestedAtUtc; ContentHash; SchemaVersion`

CSV uses **UTF-8 BOM and semicolon delimiter**; all dates exposed as timestamps are normalized to ISO 8601 UTC. Source-specific details (`Edition`, `Release`, `AzureFeature`, `LifecyclePolicy`, `StartDate`, `MainstreamEndDate`, `ExtendedEndDate`, `ReleaseEndDate`, `RetirementDate`, `EndOfSupport`, `OriginalSupportStatus`) remain available as JSON in `Description`. The `PublishedDate` field contains the source's release/start date as a proxy because the XLSX has no separate publication timestamp; refer to `Description.StartDate` for its original meaning. `LastModifiedDate` stays blank unless actually supplied.

## Independent output files

| Constant | Path |
| --- | --- |
| `NORMALIZED_FILE` | `data/normalized/microsoft-lifecycle.csv` |
| `HISTORY_FILE` | `data/history/microsoft-lifecycle-history.csv` |
| `REJECTED_FILE` | `data/rejected/microsoft-lifecycle-rejected.csv` |
| `STATISTICS_FILE` | `data/metadata/microsoft-lifecycle-statistics.csv` |
| `STATISTICS_HISTORY_FILE` | `data/metadata/microsoft-lifecycle-statistics-history.csv` |
| `WARNINGS_FILE` | `data/metadata/microsoft-lifecycle-warnings.csv` |
| `LOG_FILE` | `data/logs/microsoft-lifecycle.log` |

`HISTORY_FILE` stores the initial business record and subsequent **content changes only**. Statistics record daily run outcomes, including failed runs. Rejected source records are isolated with error codes and original JSON. The normalized CSV is published only after minimum valid record checks and a configurable rejection-rate quality gate (default 5%).

## Lifecycle correctness

- Missing end-of-support dates remain **null**, never a synthetic `2099-12-31`.
- Explicit `Out of Support` from Microsoft overrides dates for lifecycle status.
- Release-specific end dates are preferred where present, with extended-support and retirement dates retained separately in `Description` for correct policy interpretation.
- Stable `RecordId` derives from Product + Edition + Release + Azure Feature. Source date or status changes update `ContentHash`, not `RecordId`.
- An invalid source row without a product is rejected, not allowed into Power BI.

## Use

```bash
python -m pip install -r requirements.txt
python -m pytest -q tests
python scripts/transform.py                # Fresh Microsoft workbook
python scripts/transform.py --input raw/product-lifecycle.xlsx # Local fixture
```

The scheduled `.github/workflows/lifecycle-import.yml` workflow publishes the normalized CSV and diagnostics after validation and uploads execution logs as an artifact. If the Microsoft workbook download fails, the pipeline **does not** silently republish stale data.

For a public repository, Power BI can import the normalized dataset from:

`https://raw.githubusercontent.com/Renard-Dehenau/euc-lifecycle-data/main/data/normalized/microsoft-lifecycle.csv`

Select semicolon delimiter and UTF-8 encoding, promote headers, then assign appropriate types to `SnapshotDate`, `PublishedDate`, `LastModifiedDate` and `IngestedAtUtc`. The older `data/microsoft-lifecycle.csv` and `.json` remain untouched for migration purposes and must not be confused with the new 20-column output.
