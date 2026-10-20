#!/usr/bin/env python3
"""
Remove duplicate columns from a CSV or Excel workbook, keeping the column that contains
the regulator name (e.g., 'IESBA') in the regulator row.

Duplicate columns = columns with IDENTICAL values across ALL rows.

Handles:
  - .xlsx / .xlsm / .xltx  (single sheet, streamed for large files)
  - .csv / .tsv / .txt      (any encoding: utf-8, utf-8-sig, utf-16, cp1252, latin-1)
  - Output folder created automatically
  - Output format follows the output file extension

Zero external deps for CSV. openpyxl required only for Excel files.
"""

import csv
import sys
import os
import argparse
import hashlib
from collections import defaultdict
from pathlib import Path


# ---------------------------------------------------------------------------
# Format detection
# ---------------------------------------------------------------------------

EXCEL_EXTS = ('.xlsx', '.xlsm', '.xltx', '.xltm')
CSV_EXTS = ('.csv', '.tsv', '.txt', '.dat')

ENCODINGS = ['utf-8-sig', 'utf-8', 'utf-16', 'cp1252', 'latin-1']


def detect_format(path):
    """Detect file format by magic bytes first, then extension."""
    with open(path, 'rb') as f:
        magic = f.read(8)

    # Excel = ZIP container (PK\x03\x04)
    if magic[:4] == b'PK\x03\x04':
        return 'xlsx'

    # Old binary .xls = OLE2 compound file (D0 CF 11 E0)
    if magic[:4] == b'\xd0\xcf\x11\xe0':
        return 'xls'

    # BOMs
    if magic[:3] == b'\xef\xbb\xbf':
        return 'csv', 'utf-8-sig'
    if magic[:2] == b'\xff\xfe':
        return 'csv', 'utf-16'
    if magic[:2] == b'\xfe\xff':
        return 'csv', 'utf-16'

    ext = os.path.splitext(path)[1].lower()
    if ext in EXCEL_EXTS:
        return 'xlsx'
    if ext in CSV_EXTS:
        return 'csv'
    return 'csv'  # default: try as CSV


def detect_encoding(path):
    """Try decoding with a fallback chain; return the first that works."""
    for enc in ENCODINGS:
        try:
            with open(path, 'r', encoding=enc) as f:
                f.read(65536)
            return enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    return 'latin-1'  # never fails


def sniff_delimiter(path, encoding):
    """Guess the delimiter from the first few lines."""
    try:
        with open(path, 'r', encoding=encoding, newline='') as f:
            sample = f.read(16384)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=',;\t|')
            return dialect.delimiter
        except Exception:
            return ','
    except Exception:
        return ','


# ---------------------------------------------------------------------------
# Row iteration (streaming, memory-safe for huge files)
# ---------------------------------------------------------------------------

def iter_xlsx_rows(path, sheet_name=None):
    """Stream rows from a single sheet of an xlsx workbook."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        print("❌ Error: openpyxl is required for Excel files. Install with:", file=sys.stderr)
        print("     pip install openpyxl", file=sys.stderr)
        raise

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                print(f"❌ Error: Sheet '{sheet_name}' not found. Available: {wb.sheetnames}", file=sys.stderr)
                return
            ws = wb[sheet_name]
        else:
            ws = wb[wb.sheetnames[0]]
            if len(wb.sheetnames) > 1:
                print(f"  → Workbook has {len(wb.sheetnames)} sheets; using first: '{ws.title}'")
        for row in ws.iter_rows(values_only=True):
            yield list(row) if row is not None else []
    finally:
        wb.close()


def iter_csv_rows(path, encoding, delimiter):
    """Stream rows from a CSV file."""
    with open(path, 'r', encoding=encoding, newline='') as f:
        reader = csv.reader(f, delimiter=delimiter)
        for row in reader:
            yield row


# ---------------------------------------------------------------------------
# Core algorithm
# ---------------------------------------------------------------------------

def cell_to_str(v):
    """Normalize a cell value to a string for comparison/output."""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def analyze_and_dedupe(rows_iter, regulator_pattern, regulator_row_idx):
    """
    Pass 1: stream all rows, hash each column, capture regulator row.
    Returns (keep_cols, num_cols, total_rows, regulator_row_idx, col_hashes)
    """
    col_hashes = [hashlib.md5(b'').hexdigest() for _ in range(0)]  # dynamic list
    row_count = 0
    max_cols = 0
    found_regulator_row = None
    regulator_values = None

    for row in rows_iter:
        row = [cell_to_str(c) for c in row]
        n = len(row)
        if n > max_cols:
            for _ in range(n - max_cols):
                col_hashes.append(hashlib.md5(b'').hexdigest())
            max_cols = n

        if found_regulator_row is None:
            for cell in row:
                if regulator_pattern.lower() in cell.lower():
                    found_regulator_row = row_count
                    regulator_values = row[:]
                    break

        for i, cell in enumerate(row):
            col_hashes[i] = hashlib.md5(
                (col_hashes[i] + '\x00' + cell).encode('utf-8', 'replace')
            ).hexdigest()

        row_count += 1

    if row_count == 0:
        return None, 0, 0, None, None

    if regulator_row_idx is None:
        regulator_row_idx = found_regulator_row
    elif regulator_row_idx >= row_count:
        return None, max_cols, row_count, regulator_row_idx, None

    # Group columns by hash
    groups = defaultdict(list)
    for i, h in enumerate(col_hashes):
        groups[h].append(i)

    keep = set()
    removed = 0
    dup_groups = 0

    for h, idxs in groups.items():
        if len(idxs) == 1:
            keep.add(idxs[0])
            continue
        dup_groups += 1
        keeper = None
        if regulator_values is not None and regulator_row_idx is not None:
            for i in idxs:
                val = regulator_values[i] if i < len(regulator_values) else ''
                if regulator_pattern.lower() in val.lower():
                    keeper = i
                    break
        if keeper is None:
            keeper = idxs[0]
            print(f"  → Duplicate group (cols {idxs}): no '{regulator_pattern}' found, keeping col {keeper}, removing {len(idxs)-1}")
        else:
            print(f"  → Duplicate group (cols {idxs}): keeping col {keeper} (has '{regulator_pattern}'), removing {len(idxs)-1}")
        keep.add(keeper)
        removed += len(idxs) - 1

    return keep, max_cols, row_count, regulator_row_idx, groups


# ---------------------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------------------

def write_csv(path, rows_iter, keep_cols, delimiter):
    parent = os.path.dirname(path)
    if parent and not os.path.exists(parent):
        os.makedirs(parent, exist_ok=True)
        print(f"  → Created output folder: {parent}")
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f, delimiter=delimiter)
        for row in rows_iter:
            w.writerow([row[i] if i < len(row) else '' for i in keep_cols])


def write_xlsx(path, rows_iter, keep_cols):
    try:
        from openpyxl import Workbook
    except ImportError:
        print("❌ Error: openpyxl is required for Excel output. pip install openpyxl", file=sys.stderr)
        return False

    parent = os.path.dirname(path)
    if parent and not os.path.exists(parent):
        os.makedirs(parent, exist_ok=True)
        print(f"  → Created output folder: {parent}")

    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="Sheet1")
    for row in rows_iter:
        ws.append([row[i] if i < len(row) else '' for i in keep_cols])
    wb.save(path)
    wb.close()
    return True


# ---------------------------------------------------------------------------
# High-level processing
# ---------------------------------------------------------------------------

def process(input_path, output_path, regulator_pattern, regulator_row_idx, delimiter, sheet_name, out_format):
    fmt_info = detect_format(input_path)
    if isinstance(fmt_info, tuple):
        src_format, forced_encoding = fmt_info
    else:
        src_format, forced_encoding = fmt_info, None

    print(f"  Detected format: {src_format.upper()}")
    print(f"  Output format  : {out_format.upper()}")

    if src_format == 'xls':
        print("❌ Error: Legacy .xls (binary) format is not supported.", file=sys.stderr)
        print("   Open it in Excel and 'Save As' .xlsx first.", file=sys.stderr)
        return False

    # Resolve encoding & delimiter for CSV
    encoding = forced_encoding
    if src_format == 'csv' and encoding is None:
        encoding = detect_encoding(input_path)
        print(f"  Detected encoding: {encoding}")
        if delimiter is None:
            delimiter = sniff_delimiter(input_path, encoding)
            print(f"  Detected delimiter: {delimiter!r}")
    elif delimiter is None:
        delimiter = ','

    # Build a fresh row iterator factory (we need to stream twice)
    def make_iter():
        if src_format == 'xlsx':
            return iter_xlsx_rows(input_path, sheet_name)
        return iter_csv_rows(input_path, encoding, delimiter)

    # Pass 1 - analyze
    keep, num_cols, total_rows, reg_row, _ = analyze_and_dedupe(
        make_iter(), regulator_pattern, regulator_row_idx)

    if keep is None:
        print("❌ Error: File is empty or regulator row index is out of range.", file=sys.stderr)
        return False

    if reg_row is None:
        print(f"❌ Error: Could not find regulator pattern '{regulator_pattern}' in any row.", file=sys.stderr)
        print("   Use --regulator-row N to specify the row manually (0-based).", file=sys.stderr)
        return False

    print(f"  Rows: {total_rows}  |  Columns: {num_cols}  |  Regulator row: {reg_row}")

    # Pass 2 - write
    kept_sorted = sorted(keep)
    print()
    try:
        if out_format == 'xlsx':
            ok = write_xlsx(output_path, make_iter(), kept_sorted)
            if not ok:
                return False
        else:
            write_csv(output_path, make_iter(), kept_sorted, delimiter)
    except PermissionError:
        print(f"❌ Error: Permission denied writing to {output_path}", file=sys.stderr)
        print("   Close the file in Excel and try again.", file=sys.stderr)
        return False
    except Exception as e:
        print(f"❌ Error writing output: {type(e).__name__}: {e}", file=sys.stderr)
        return False

    if not os.path.exists(output_path):
        print(f"❌ Error: Output file was not created: {output_path}", file=sys.stderr)
        return False

    size = os.path.getsize(output_path)
    removed = num_cols - len(kept_sorted)
    print()
    print(f"✓ Done! Kept {len(kept_sorted)} columns, removed {removed} duplicate column(s).")
    print(f"✓ Output file created: {output_path}  ({size:,} bytes)")
    return True


# ---------------------------------------------------------------------------
# Interactive mode
# ---------------------------------------------------------------------------

def interactive_mode():
    print("=" * 64)
    print("  CSV / Excel Duplicate Column Remover  (Regulator-Aware)")
    print("=" * 64)
    print()
    print("Drag & drop your file (CSV or Excel) into this window,")
    print("or paste the full path. Spaces are fine - no quotes needed.")
    print()

    while True:
        inp = input("1. Input file path: ").strip()
        if not inp:
            print("   → Error: Input file is required.")
            continue
        input_path = str(Path(inp.strip().strip('"\'')).expanduser().resolve())
        if not os.path.exists(input_path):
            print(f"   → Error: File not found: {input_path}")
            continue
        if not os.path.isfile(input_path):
            print(f"   → Error: Not a file: {input_path}")
            continue
        break
    print(f"   ✓ Input : {input_path}")

    # Default output name
    src_fmt = detect_format(input_path)
    src_fmt = src_fmt[0] if isinstance(src_fmt, tuple) else src_fmt
    default_ext = '.xlsx' if src_fmt == 'xlsx' else '.csv'
    stem = Path(input_path).stem
    default_output = str(Path(input_path).with_name(stem + '_deduped' + default_ext))

    while True:
        out = input(f"\n2. Output file path [{default_output}]: ").strip()
        output_path = str(Path(out.strip().strip('"\'')).expanduser().resolve()) if out else default_output
        if os.path.isdir(output_path):
            output_path = str(Path(output_path) / (stem + '_deduped' + default_ext))
            print(f"   → Folder selected, will write: {output_path}")
        if output_path == input_path:
            print("   → Error: Output must differ from input. Try another name.")
            continue
        break
    print(f"   ✓ Output: {output_path}")

    reg = input("\n3. Regulator pattern to keep [IESBA]: ").strip()
    regulator_pattern = reg if reg else "IESBA"
    print(f"   ✓ Regulator: {regulator_pattern}")

    row_inp = input("\n4. Regulator row index, 0-based (Enter = auto-detect): ").strip()
    regulator_row_idx = int(row_inp) if row_inp.lstrip('-').isdigit() else None
    print(f"   ✓ Regulator row: {'auto-detect' if regulator_row_idx is None else regulator_row_idx}")

    delim = input("\n5. Delimiter for CSV (Enter = auto-detect): ").strip()
    delimiter = delim if delim else None
    if delimiter == '\\t':
        delimiter = '\t'
    print(f"   ✓ Delimiter: {'auto-detect' if delimiter is None else repr(delimiter)}")

    sheet = input("\n6. Excel sheet name (Enter = first sheet): ").strip()
    sheet_name = sheet if sheet else None
    print(f"   ✓ Sheet: {sheet_name or 'first sheet'}")

    print()
    print("-" * 64)
    print("  SUMMARY - confirm to proceed")
    print("-" * 64)
    print(f"  Input    : {input_path}")
    print(f"  Output   : {output_path}")
    print(f"  Regulator: {regulator_pattern}")
    print(f"  Row      : {'auto-detect' if regulator_row_idx is None else regulator_row_idx}")
    print(f"  Delimiter: {'auto-detect' if delimiter is None else repr(delimiter)}")
    print(f"  Sheet    : {sheet_name or 'first sheet'}")
    print("-" * 64)

    while True:
        confirm = input("\nProceed? [Y/n]: ").strip().lower()
        if confirm in ('', 'y', 'yes'):
            break
        if confirm in ('n', 'no', 'q', 'quit'):
            print("Cancelled.")
            return False
        print("   → Please type 'y' or 'n'.")

    print()
    print("Processing...")
    print(f"  Input : {input_path}")
    print(f"  Output: {output_path}")
    print()

    out_fmt = 'xlsx' if output_path.lower().endswith(EXCEL_EXTS) else 'csv'
    return process(input_path, output_path, regulator_pattern, regulator_row_idx,
                   delimiter, sheet_name, out_fmt)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) == 1:
        try:
            sys.exit(0 if interactive_mode() else 1)
        except (KeyboardInterrupt, EOFError):
            print("\n\nCancelled.")
            sys.exit(1)

    p = argparse.ArgumentParser(
        description='Remove duplicate CSV/Excel columns, keeping the one with the regulator name',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python dedupe_regulator_cols.py                       (interactive, drag-and-drop)
  python dedupe_regulator_cols.py in.csv out.csv
  python dedupe_regulator_cols.py in.xlsx out.xlsx
  python dedupe_regulator_cols.py in.xlsx out.csv        (convert to CSV)
  python dedupe_regulator_cols.py in.csv out.csv -r IESBA -R 5
  python dedupe_regulator_cols.py in.csv out.csv -d ";"
        """)
    p.add_argument('input', nargs='?', help='Input file (.csv/.xlsx)')
    p.add_argument('output', nargs='?', help='Output file (format from extension)')
    p.add_argument('--regulator', '-r', default='IESBA', help='Regulator pattern (default: IESBA)')
    p.add_argument('--regulator-row', '-R', type=int, help='Regulator row index, 0-based (default: auto-detect)')
    p.add_argument('--delimiter', '-d', default=None, help='CSV delimiter (default: auto-detect)')
    p.add_argument('--sheet', '-s', default=None, help='Excel sheet name (default: first sheet)')
    args = p.parse_args()

    if not args.input:
        p.error("Input file required (or run with no args for interactive mode)")

    input_path = str(Path(args.input.strip().strip('"\'')).expanduser().resolve())
    if not os.path.exists(input_path):
        p.error(f"Input file not found: {input_path}")

    if args.output:
        output_path = str(Path(args.output.strip().strip('"\'')).expanduser().resolve())
    else:
        src_fmt = detect_format(input_path)
        src_fmt = src_fmt[0] if isinstance(src_fmt, tuple) else src_fmt
        ext = '.xlsx' if src_fmt == 'xlsx' else '.csv'
        output_path = str(Path(input_path).with_name(Path(input_path).stem + '_deduped' + ext))

    if output_path == input_path:
        p.error("Output must differ from input")

    out_fmt = 'xlsx' if output_path.lower().endswith(EXCEL_EXTS) else 'csv'

    print(f"Input:    {input_path}")
    print(f"Output:   {output_path}")
    print(f"Regulator: {args.regulator}")
    print(f"Row:       {'auto-detect' if args.regulator_row is None else args.regulator_row}")
    print()

    sys.exit(0 if process(input_path, output_path, args.regulator, args.regulator_row,
                         args.delimiter, args.sheet, out_fmt) else 1)


if __name__ == '__main__':
    main()