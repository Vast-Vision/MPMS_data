"""
Convert Quantum Design MPMS3 .dat files to Apple Numbers compatible format (.xlsx)
Apple Numbers can directly open .xlsx files.

Usage:
    python dat_to_numbers.py <input_path> <output_dir>

Example:
    python dat_to_numbers.py ../MPMS_data/raw\ data/ ./raw_data_numbers_files/
"""

import argparse
from pathlib import Path
from io import StringIO
import pandas as pd


def parse_mpms_dat(filepath: Path) -> tuple[pd.DataFrame, dict]:
    """
    Parse an MPMS3 .dat file into a DataFrame and header metadata.

    Returns:
        (df, header_info): DataFrame of the [Data] section, dict of header key-value pairs.
    """
    with open(filepath, encoding="utf-8", errors="replace") as f:
        lines = f.readlines()

    header_info = {}
    in_data = False
    data_start = None

    for i, line in enumerate(lines):
        line = line.rstrip("\n\r")
        if line == "[Data]":
            in_data = True
            data_start = i + 1
            break
        if line.startswith("INFO,"):
            parts = line.split(",", 2)
            if len(parts) >= 3:
                key = parts[1].strip()
                val = parts[2].strip()
                header_info[key] = val
        elif not line.startswith(";") and "," in line and "=" not in line and not line.startswith("["):
            key_val = line.split(",", 1)
            if len(key_val) == 2:
                header_info[key_val[0].strip()] = key_val[1].strip()

    if not in_data or data_start is None:
        raise ValueError(f"No [Data] section found in {filepath}")

    # Parse data block as CSV
    data_block = "\n".join(l.rstrip("\n\r") for l in lines[data_start:])
    df = pd.read_csv(StringIO(data_block))

    # Coerce numeric columns
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, header_info


def convert_dat_to_numbers(dat_path: Path, output_dir: Path) -> Path:
    """
    Convert a single .dat file to .xlsx format (Apple Numbers compatible).

    Args:
        dat_path: Path to input .dat file
        output_dir: Directory for output file

    Returns:
        Path to created .xlsx file
    """
    # Parse the .dat file
    df, header_info = parse_mpms_dat(dat_path)

    # Create output filename (.numbers extension, but actually xlsx format)
    # Apple Numbers can open .xlsx files directly
    output_name = dat_path.stem + ".xlsx"
    output_path = output_dir / output_name

    # Write to Excel format with openpyxl
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Write data to 'Data' sheet
        df.to_excel(writer, sheet_name='Data', index=False)

        # Write metadata to 'Metadata' sheet
        if header_info:
            meta_df = pd.DataFrame(
                list(header_info.items()),
                columns=['Property', 'Value']
            )
            meta_df.to_excel(writer, sheet_name='Metadata', index=False)

    return output_path


def find_dat_files(path: Path) -> list[Path]:
    """Find all .dat files in a path (file or directory)."""
    if path.is_file() and path.suffix.lower() == ".dat":
        return [path]
    elif path.is_dir():
        return sorted(path.rglob("*.dat"))
    return []


def main():
    parser = argparse.ArgumentParser(
        description="Convert MPMS .dat files to Apple Numbers format (.xlsx)"
    )
    parser.add_argument(
        "input_path",
        type=Path,
        help="Input .dat file or directory containing .dat files"
    )
    parser.add_argument(
        "output_dir",
        type=Path,
        help="Output directory for .xlsx files"
    )
    parser.add_argument(
        "--skip-rw",
        action="store_true",
        help="Skip .rw.dat files (raw waveform data, typically very large)"
    )

    args = parser.parse_args()

    # Create output directory if needed
    args.output_dir.mkdir(parents=True, exist_ok=True)

    # Find all .dat files
    dat_files = find_dat_files(args.input_path)

    if not dat_files:
        print(f"No .dat files found in {args.input_path}")
        return

    print(f"Found {len(dat_files)} .dat files")

    # Convert each file
    converted = 0
    skipped = 0
    errors = 0

    for dat_file in dat_files:
        # Skip .rw.dat files if requested (they're very large raw data)
        if args.skip_rw and ".rw.dat" in dat_file.name:
            print(f"  Skipping (raw waveform): {dat_file.name}")
            skipped += 1
            continue

        try:
            output_path = convert_dat_to_numbers(dat_file, args.output_dir)
            print(f"  Converted: {dat_file.name} -> {output_path.name}")
            converted += 1
        except Exception as e:
            print(f"  Error converting {dat_file.name}: {e}")
            errors += 1

    print(f"\nDone! Converted: {converted}, Skipped: {skipped}, Errors: {errors}")
    print(f"Output directory: {args.output_dir.absolute()}")


if __name__ == "__main__":
    main()
