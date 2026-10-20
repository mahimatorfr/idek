# CSV / Excel Duplicate Column Remover (Regulator-Aware)

Removes duplicate columns from a CSV or Excel workbook, keeping the one that
contains the regulator name (default `IESBA`) in the regulator row.

**Duplicate column** = a column whose values are identical to another column
across *every* row.

---

## Install

```bash
python install_deps.py
```

Installs `openpyxl` (only required for Excel files). CSV files need no
dependencies at all.

## Usage

### Interactive (Windows friendly — drag & drop)

```bash
python dedupe_regulator_cols.py
```

Prompts for the input path, output path, regulator pattern, row index,
delimiter and sheet name, then shows a summary and asks `Proceed? [Y/n]`.

### One-shot

```bash
python dedupe_regulator_cols.py input.xlsx output.xlsx
python dedupe_regulator_cols.py input.csv  output.csv
python dedupe_regulator_cols.py input.xlsx output.csv    # convert to CSV
```

### Options

| Flag | Short | Default | Description |
|------|-------|---------|-------------|
| `--regulator` | `-r` | `IESBA` | Pattern that marks the keeper column |
| `--regulator-row` | `-R` | auto-detect | 0-based index of the regulator row |
| `--delimiter` | `-d` | auto-detect | CSV delimiter (`,` `;` `\t` `\|`) |
| `--sheet` | `-s` | first sheet | Excel sheet name |

---

## Features

- **Format auto-detection** by magic bytes, not just file extension
  - `PK\x03\x04` → Excel `.xlsx` / `.xlsm` / `.xltx`
  - BOMs → UTF-8-SIG, UTF-16 LE/BE
  - Otherwise → CSV with encoding fallback:
    `utf-8-sig → utf-8 → utf-16 → cp1252 → latin-1`
- **Huge files supported** — streams with `read_only` / `write_only` and
  incremental per-column hashing, so memory stays flat regardless of file size
- **Output folder is created automatically** if it doesn't exist
- **Output format follows the output extension** (`.xlsx` → Excel, else CSV)
- **Cross-format conversion** (Excel → CSV and back)
- Handles `utf-8-sig` BOMs (files saved from Excel)
- Clear error messages instead of tracebacks

## How it decides which duplicate to keep

1. Group columns by a hash of their values across all rows
2. For each group of 2+ identical columns, look at the regulator row
3. Keep the column whose regulator-row cell contains `IESBA`; drop the rest
4. If no column in the group matches, keep the leftmost and drop the rest

## Example

**Input**
```
Name,Regulator,Value,Regulator,Value
EntityA,IESBA,100,IESBA,100
EntityB,IESBA,200,IESBA,200
```

**Output**
```
Name,Regulator,Value
EntityA,IESBA,100
EntityB,IESBA,200
```

## Files

| File | Purpose |
|------|---------|
| `dedupe_regulator_cols.py` | Main tool |
| `install_deps.py` | Dependency installer (openpyxl) |
| `run_dedupe.py` | Launcher with environment checks |
| `requirements.txt` | Dependency manifest |

## Requirements

- Python 3.7+
- `openpyxl` (only for Excel files)
