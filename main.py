#!/usr/bin/env python3
"""
Remove duplicate columns from a CSV, keeping the one that contains the regulator name (e.g., 'IESBA')
in the regulator row. Duplicate columns are defined as having identical values across all rows.

Windows-friendly CLI with drag-and-drop / paste-path support.
"""

import csv
import sys
import argparse
import os
from collections import defaultdict
from pathlib import Path


def find_regulator_row(rows, regulator_pattern):
    """Find the row index that contains the regulator pattern."""
    for i, row in enumerate(rows):
        for cell in row:
            if regulator_pattern.lower() in str(cell).lower():
                return i
    return None


def get_column_values(rows, col_idx):
    """Get all values for a specific column as a tuple (hashable)."""
    return tuple(row[col_idx] if col_idx < len(row) else '' for row in rows)


def normalize_path(path_str):
    """Normalize a file path: strip quotes, expand user, resolve."""
    path_str = path_str.strip().strip('"\'')
    return str(Path(path_str).expanduser().resolve())


def process_csv(input_path, output_path, regulator_pattern, regulator_row_idx, delimiter):
    """Core processing logic."""
    # Read all rows
    with open(input_path, 'r', newline='', encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter=delimiter)
        rows = list(reader)

    if not rows:
        print("Error: Empty file", file=sys.stderr)
        return False

    num_cols = max(len(r) for r in rows)
    # Pad rows to same length
    for r in rows:
        if len(r) < num_cols:
            r.extend([''] * (num_cols - len(r)))

    # Find regulator row
    if regulator_row_idx is None:
        regulator_row_idx = find_regulator_row(rows, regulator_pattern)
        if regulator_row_idx is None:
            print(f"Error: Could not find regulator pattern '{regulator_pattern}' in any row", file=sys.stderr)
            return False
        print(f"  → Found regulator pattern in row {regulator_row_idx} (0-based)")
    else:
        if regulator_row_idx >= len(rows):
            print(f"Error: Regulator row index {regulator_row_idx} exceeds file rows ({len(rows)})", file=sys.stderr)
            return False
        print(f"  → Using explicit regulator row: {regulator_row_idx}")

    # Group columns by their values across all rows
    col_groups = defaultdict(list)
    for col_idx in range(num_cols):
        col_key = get_column_values(rows, col_idx)
        col_groups[col_key].append(col_idx)

    # Determine which columns to keep
    keep_cols = set()
    remove_count = 0

    for col_key, col_indices in col_groups.items():
        if len(col_indices) == 1:
            keep_cols.add(col_indices[0])
        else:
            # Duplicate columns - find which one has the regulator name in regulator row
            regulator_col = None
            for idx in col_indices:
                cell_value = rows[regulator_row_idx][idx] if idx < len(rows[regulator_row_idx]) else ''
                if regulator_pattern.lower() in str(cell_value).lower():
                    regulator_col = idx
                    break

            if regulator_col is not None:
                keep_cols.add(regulator_col)
                remove_count += len(col_indices) - 1
                print(f"  → Duplicate group {col_indices}: keeping column {regulator_col} (has '{regulator_pattern}'), removing {len(col_indices)-1} duplicates")
            else:
                keep_cols.add(col_indices[0])
                remove_count += len(col_indices) - 1
                print(f"  → Duplicate group {col_indices}: no regulator found, keeping first column {col_indices[0]}, removing {len(col_indices)-1} duplicates")

    # Build output rows with only kept columns (preserve original order)
    kept_sorted = sorted(keep_cols)
    output_rows = []
    for row in rows:
        output_rows.append([row[i] for i in kept_sorted])

    # Write output
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=delimiter)
        writer.writerows(output_rows)

    print(f"\n✓ Done! Kept {len(kept_sorted)} columns, removed {remove_count} duplicate columns.")
    print(f"  Output: {output_path}")
    return True


def interactive_mode():
    """Interactive mode for Windows drag-and-drop / paste paths."""
    print("=" * 60)
    print("  CSV Duplicate Column Remover (Regulator-Aware)")
    print("  " + "=" * 58)
    print()
    print("Instructions:")
    print("  • Drag and drop your CSV file here, or paste the full path")
    print("  • Paths with spaces work fine (no need to add quotes)")
    print("  • Press Enter with empty input to use defaults")
    print()

    # Input file
    while True:
        inp = input("Input CSV file path: ").strip()
        if not inp:
            print("  Error: Input file is required")
            continue
        input_path = normalize_path(inp)
        if not os.path.exists(input_path):
            print(f"  Error: File not found: {input_path}")
            continue
        if not os.path.isfile(input_path):
            print(f"  Error: Not a file: {input_path}")
            continue
        break

    # Output file (default: same folder with _deduped suffix)
    default_output = str(Path(input_path).with_stem(Path(input_path).stem + '_deduped'))
    out = input(f"Output CSV file path [{default_output}]: ").strip()
    output_path = normalize_path(out) if out else default_output

    # Regulator pattern
    reg = input("Regulator pattern to keep [IESBA]: ").strip()
    regulator_pattern = reg if reg else "IESBA"

    # Regulator row (optional)
    row_inp = input("Regulator row index 0-based (empty = auto-detect): ").strip()
    regulator_row_idx = int(row_inp) if row_inp.isdigit() else None

    # Delimiter
    delim = input("Delimiter [,]: ").strip()
    delimiter = delim if delim else ","

    print()
    print("Processing...")
    print(f"  Input:    {input_path}")
    print(f"  Output:   {output_path}")
    print(f"  Regulator: {regulator_pattern}")
    print(f"  Row:       {'auto-detect' if regulator_row_idx is None else regulator_row_idx}")
    print(f"  Delimiter: '{delimiter}'")
    print()

    return process_csv(input_path, output_path, regulator_pattern, regulator_row_idx, delimiter)


def main():
    # If no args provided, run interactive mode (great for Windows drag-and-drop)
    if len(sys.argv) == 1:
        try:
            success = interactive_mode()
            sys.exit(0 if success else 1)
        except KeyboardInterrupt:
            print("\n\nCancelled.")
            sys.exit(1)
        except EOFError:
            print("\n\nCancelled.")
            sys.exit(1)

    # Otherwise, CLI mode
    parser = argparse.ArgumentParser(
        description='Remove duplicate CSV columns, keeping the one with the regulator name',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive mode (drag-and-drop friendly)
  python dedupe_regulator_cols.py

  # CLI mode
  python dedupe_regulator_cols.py input.csv output.csv
  python dedupe_regulator_cols.py "C:\\Data\\file.csv" "C:\\Output\\clean.csv" --regulator IESBA
  python dedupe_regulator_cols.py data.csv clean.csv --regulator-row 5 --delimiter ";"
        """
    )
    parser.add_argument('input', nargs='?', help='Input CSV file path')
    parser.add_argument('output', nargs='?', help='Output CSV file path (default: input_deduped.csv)')
    parser.add_argument('--regulator', '-r', default='IESBA', help='Regulator pattern to identify keeper column (default: IESBA)')
    parser.add_argument('--regulator-row', '-R', type=int, help='Explicit 0-based row index for regulator row (default: auto-detect)')
    parser.add_argument('--delimiter', '-d', default=',', help='CSV delimiter (default: comma)')

    args = parser.parse_args()

    if not args.input:
        parser.error("Input file required (or run without args for interactive mode)")

    input_path = normalize_path(args.input)

    if not os.path.exists(input_path):
        parser.error(f"Input file not found: {input_path}")

    if args.output:
        output_path = normalize_path(args.output)
    else:
        output_path = str(Path(input_path).with_stem(Path(input_path).stem + '_deduped'))

    print(f"Input:    {input_path}")
    print(f"Output:   {output_path}")
    print(f"Regulator: {args.regulator}")
    print(f"Row:       {'auto-detect' if args.regulator_row is None else args.regulator_row}")
    print(f"Delimiter: '{args.delimiter}'")
    print()

    success = process_csv(input_path, output_path, args.regulator, args.regulator_row, args.delimiter)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
