#!/usr/bin/env python3
"""
Remove duplicate columns from a CSV, keeping the one that contains the regulator name (e.g., 'IESBA')
in the regulator row. Duplicate columns are defined as having identical values across all rows.
"""

import csv
import sys
import argparse
from collections import defaultdict


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


def main():
    parser = argparse.ArgumentParser(description='Remove duplicate CSV columns, keeping the regulator column')
    parser.add_argument('input', help='Input CSV file')
    parser.add_argument('output', help='Output CSV file')
    parser.add_argument('--regulator', default='IESBA', help='Regulator pattern to identify the keeper column (default: IESBA)')
    parser.add_argument('--regulator-row', type=int, help='Explicit row index (0-based) for regulator row. If not set, auto-detect.')
    parser.add_argument('--delimiter', default=',', help='CSV delimiter (default: comma)')
    args = parser.parse_args()

    # Read all rows
    with open(args.input, 'r', newline='', encoding='utf-8') as f:
        reader = csv.reader(f, delimiter=args.delimiter)
        rows = list(reader)

    if not rows:
        print("Empty file", file=sys.stderr)
        sys.exit(1)

    num_cols = max(len(r) for r in rows)
    # Pad rows to same length
    for r in rows:
        if len(r) < num_cols:
            r.extend([''] * (num_cols - len(r)))

    # Find regulator row
    regulator_row_idx = args.regulator_row
    if regulator_row_idx is None:
        regulator_row_idx = find_regulator_row(rows, args.regulator)
        if regulator_row_idx is None:
            print(f"Could not find regulator pattern '{args.regulator}' in any row", file=sys.stderr)
            sys.exit(1)
        print(f"Found regulator pattern in row {regulator_row_idx} (0-based)")

    # Group columns by their values across all rows
    col_groups = defaultdict(list)
    for col_idx in range(num_cols):
        col_key = get_column_values(rows, col_idx)
        col_groups[col_key].append(col_idx)

    # Determine which columns to keep
    keep_cols = set()
    remove_cols = set()

    for col_key, col_indices in col_groups.items():
        if len(col_indices) == 1:
            # Unique column, keep it
            keep_cols.add(col_indices[0])
        else:
            # Duplicate columns - find which one has the regulator name in regulator row
            regulator_col = None
            for idx in col_indices:
                cell_value = rows[regulator_row_idx][idx] if regulator_row_idx < len(rows) and idx < len(rows[regulator_row_idx]) else ''
                if args.regulator.lower() in str(cell_value).lower():
                    regulator_col = idx
                    break

            if regulator_col is not None:
                keep_cols.add(regulator_col)
                remove_cols.update(idx for idx in col_indices if idx != regulator_col)
                print(f"Duplicate group {col_indices}: keeping column {regulator_col} (has regulator), removing {remove_cols}")
            else:
                # No regulator found in this group - keep first, remove rest
                keep_cols.add(col_indices[0])
                remove_cols.update(col_indices[1:])
                print(f"Duplicate group {col_indices}: no regulator found, keeping first column {col_indices[0]}")

    # Build output rows with only kept columns (preserve original order)
    kept_sorted = sorted(keep_cols)
    output_rows = []
    for row in rows:
        output_rows.append([row[i] for i in kept_sorted])

    # Write output
    with open(args.output, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=args.delimiter)
        writer.writerows(output_rows)

    print(f"Done. Kept {len(kept_sorted)} columns, removed {len(remove_cols)} duplicate columns.")
    print(f"Output written to {args.output}")


if __name__ == '__main__':
    main()
