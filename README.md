# idek

Personal projects and tools.

## Tools

### [CSV / Excel Duplicate Column Remover](csv_dedupe_tool/)

Removes duplicate columns from a CSV or Excel workbook, keeping the one that
contains the regulator name (`IESBA`) in the regulator row.

```bash
cd csv_dedupe_tool
python install_deps.py
python dedupe_regulator_cols.py
```

Supports huge workbooks (streaming), auto-detects format and encoding, creates
missing output folders, and can convert between Excel and CSV.

See [`csv_dedupe_tool/README.md`](csv_dedupe_tool/README.md) for full details.
