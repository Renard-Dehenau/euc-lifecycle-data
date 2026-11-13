# EUC lifecycle — expanded Power BI columns

The existing **independent** `scripts/transform.py` exporter remains in this repository; nothing is merged into Microsoft Release Communications or Windows Update Catalog.

The Microsoft EUC Lifecycle data comes from a Microsoft XLSX workbook, not from a JSON API. The script continues to normalize it into a CSV for Power BI.

## Expanded CSV contract

`data/normalized/microsoft-lifecycle.csv` retains the original **20 Microsoft Release Communications columns** in their original order, followed by the **12 previously JSON-encoded `Description` keys**:

`Vendor; Edition; Release; AzureFeature; LifecyclePolicy; StartDate; MainstreamEndDate; ExtendedEndDate; ReleaseEndDate; RetirementDate; EndOfSupport; OriginalSupportStatus`

That is **32 stable columns** in a single EUC-lifecycle dataset. `Description` is retained for future human-readable text and is no longer populated with a JSON object. Blank fields are output as empty cells, not the string `null`. Support dates use ISO `YYYY-MM-DD` and all existing timestamp fields retain ISO 8601 UTC.

The field `PublishedDate` continues to use the product/release start date as a proxy, because the source workbook does not include a separate publication date; `StartDate` preserves the source meaning precisely.

The dataset continues to publish seven independent artifacts:

| Constant | File |
| --- | --- |
| `NORMALIZED_FILE` | `data/normalized/microsoft-lifecycle.csv` |
| `HISTORY_FILE` | `data/history/microsoft-lifecycle-history.csv` |
| `REJECTED_FILE` | `data/rejected/microsoft-lifecycle-rejected.csv` |
| `STATISTICS_FILE` | `data/metadata/microsoft-lifecycle-statistics.csv` |
| `STATISTICS_HISTORY_FILE` | `data/metadata/microsoft-lifecycle-statistics-history.csv` |
| `WARNINGS_FILE` | `data/metadata/microsoft-lifecycle-warnings.csv` |
| `LOG_FILE` | `data/logs/microsoft-lifecycle.log` |

The output is **UTF-8 BOM, semicolon-delimited**, suitable for Power BI's CSV connector.

## Safeguarded history migration

On the first run after installation, a legacy 20-column history CSV with JSON-encoded `Description` will be copied byte-for-byte into `data/legacy/` before its history is converted to the expanded 32-column format. The migration preserves previous business values and recalculates compatible content hashes. Missing dates are never replaced by `2099-12-31`.

If the legacy history contains unknown/unmappable fields, the original history remains in the `data/legacy` backup, and a new history starts rather than silently discarding values. The workflow commits the legacy backup along with the new normalized dataset.

## Install

Extract the update ZIP into the `euc-lifecycle-data` repository root, preserving directory paths, then run:

```bash
python -m pip install -r requirements.txt
python -m pytest -q tests
python scripts/transform.py --input raw/product-lifecycle.xlsx
```

For a fresh download (rather than the local workbook), omit `--input`. The GitHub Actions workflow runs the tests and export, validates all 32 columns, uploads diagnostics, and publishes from a **separate clean checkout**. This avoids the dirty-working-tree `git pull --rebase` issue previously seen in the Windows catalog repository.

Power BI source URL after GitHub publication:

`https://raw.githubusercontent.com/Renard-Dehenau/euc-lifecycle-data/main/data/normalized/microsoft-lifecycle.csv`

**Power BI schema change:** After updating, refresh your Power Query navigation/schema step to recognize 12 new columns. The first 20 are unchanged, but this lifecycle CSV is now 32 columns; this intentional extension does not alter any other repository dataset.
