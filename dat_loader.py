"""
Load Quantum Design MPMS3 .dat files and graph magnetization data.
Handles [Header] / [Data] structure and DC Moment columns.
"""

from pathlib import Path
from typing import Optional, Tuple, Union

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


def parse_mpms_dat(filepath: Union[str, Path]) -> Tuple[pd.DataFrame, dict]:
    """
    Parse an MPMS3 .dat file into a DataFrame and header metadata.

    Returns:
        (df, header_info): DataFrame of the [Data] section, dict of header key-value pairs.
    """
    filepath = Path(filepath)
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

    # Parse data block as CSV (header row then data rows)
    from io import StringIO
    data_block = "\n".join(l.rstrip("\n\r") for l in lines[data_start:])
    df = pd.read_csv(StringIO(data_block))

    # Coerce numeric columns
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, header_info


def get_moment_column(df: pd.DataFrame) -> Optional[str]:
    """Choose the best moment column: DC Moment Free Ctr, then DC Moment Fixed Ctr, then Moment (emu)."""
    candidates = [
        "DC Moment Free Ctr (emu)",
        "DC Moment Fixed Ctr (emu)",
        "Moment (emu)",
    ]
    for c in candidates:
        if c in df.columns and df[c].notna().any():
            return c
    return None


def load_dat(filepath: Union[str, Path]) -> pd.DataFrame:
    """
    Load a single .dat file and return a DataFrame with standard columns
    and a 'Moment' column set to the best available moment data.
    """
    df, _ = parse_mpms_dat(filepath)
    moment_col = get_moment_column(df)
    if moment_col:
        df = df.assign(Moment=df[moment_col].values)
    return df


def plot_moment_vs_field(
    df: pd.DataFrame,
    *,
    moment_col: Optional[str] = None,
    field_col: str = "Magnetic Field (Oe)",
    title: Optional[str] = None,
    ax: Optional[plt.Axes] = None,
    label: Optional[str] = None,
) -> plt.Figure:
    """Plot moment (emu) vs magnetic field (Oe)."""
    if moment_col is None:
        moment_col = get_moment_column(df)
    if moment_col is None:
        moment_col = "Moment"
    if moment_col not in df.columns and "Moment" in df.columns:
        moment_col = "Moment"
    if field_col not in df.columns:
        raise ValueError(f"Column '{field_col}' not in DataFrame. Available: {list(df.columns)}")

    x = df[field_col].values
    y = df[moment_col].values
    mask = np.isfinite(x) & np.isfinite(y)
    x, y = x[mask], y[mask]
    if x.size == 0:
        raise ValueError("No finite (field, moment) pairs to plot.")

    fig = None
    if ax is None:
        fig, ax = plt.subplots(1, 1, figsize=(8, 5))
    ax.plot(x, y, "o-", ms=4, label=label or moment_col)
    ax.set_xlabel("Magnetic Field (Oe)")
    ax.set_ylabel("Moment (emu)")
    if title:
        ax.set_title(title)
    ax.axhline(0, color="gray", ls="--", alpha=0.7)
    ax.axvline(0, color="gray", ls="--", alpha=0.7)
    ax.grid(True, alpha=0.3)
    if label:
        ax.legend()
    if fig is None:
        fig = plt.gcf()
    return fig


def plot_moment_vs_temperature(
    df: pd.DataFrame,
    *,
    moment_col: Optional[str] = None,
    temp_col: str = "Temperature (K)",
    title: Optional[str] = None,
    ax: Optional[plt.Axes] = None,
) -> plt.Figure:
    """Plot moment vs temperature."""
    if moment_col is None:
        moment_col = get_moment_column(df) or "Moment"
    if temp_col not in df.columns:
        raise ValueError(f"Column '{temp_col}' not in DataFrame.")

    x = df[temp_col].values
    y = df[moment_col].values
    mask = np.isfinite(x) & np.isfinite(y)
    x, y = x[mask], y[mask]
    if x.size == 0:
        raise ValueError("No finite (temperature, moment) pairs to plot.")

    fig = None
    if ax is None:
        fig, ax = plt.subplots(1, 1, figsize=(8, 5))
    ax.plot(x, y, "o-", ms=4)
    ax.set_xlabel("Temperature (K)")
    ax.set_ylabel("Moment (emu)")
    if title:
        ax.set_title(title)
    ax.grid(True, alpha=0.3)
    if fig is None:
        fig = plt.gcf()
    return fig


def find_dat_files(directory: Union[str, Path]) -> list:
    """Return list of .dat files under the given directory."""
    directory = Path(directory) if not isinstance(directory, Path) else directory
    if directory.is_file() and directory.suffix.lower() == ".dat":
        return [directory]
    return sorted(directory.rglob("*.dat"))


def main():
    """Example: load .dat files from Jan_22_2025 and plot M vs H."""
    import argparse
    parser = argparse.ArgumentParser(description="Load MPMS .dat files and plot data.")
    parser.add_argument(
        "path",
        nargs="?",
        default=Path(__file__).resolve().parent / "Jan_22_2025",
        help="File or directory containing .dat files",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="Only list .dat files and columns, do not plot",
    )
    parser.add_argument(
        "--temp",
        action="store_true",
        help="Plot moment vs temperature instead of vs field",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Plot all .dat files in directory on one figure (M vs H only)",
    )
    parser.add_argument(
        "--save",
        metavar="FILE",
        nargs="?",
        const="",
        help="Save figure to plots/<name>.png (or FILE if given) instead of showing",
    )
    args = parser.parse_args()

    plots_dir = Path("plots")
    path = Path(args.path)
    dat_files = find_dat_files(path)
    if not dat_files:
        print(f"No .dat files found under {path}")
        return

    if args.list:
        for f in dat_files:
            df, info = parse_mpms_dat(f)
            moment_col = get_moment_column(df)
            print(f"\n{f.name}")
            print(f"  Rows: {len(df)}, Moment column: {moment_col}")
            if "Magnetic Field (Oe)" in df.columns:
                h = df["Magnetic Field (Oe)"]
                print(f"  Field range: {h.min():.2f} .. {h.max():.2f} Oe")
        return

    if args.all and len(dat_files) > 1:
        fig, ax = plt.subplots(1, 1, figsize=(9, 6))
        for f in dat_files:
            try:
                df = load_dat(f)
                moment_col = get_moment_column(df)
                if moment_col and "Magnetic Field (Oe)" in df.columns:
                    plot_moment_vs_field(
                        df,
                        moment_col=moment_col,
                        ax=ax,
                        label=f.stem,
                    )
            except Exception as e:
                print(f"Skip {f.name}: {e}")
        ax.legend()
        ax.set_title("Moment vs Field")
        plt.tight_layout()
        if args.save is not None:
            out = Path(args.save) if args.save else plots_dir / "moment_vs_field.png"
            out.parent.mkdir(parents=True, exist_ok=True)
            plt.savefig(out, dpi=150)
            print(f"Saved {out}")
        else:
            plt.show()
        return

    # Single file or first file
    for f in dat_files[:1] if not args.all else dat_files:
        df = load_dat(f)
        moment_col = get_moment_column(df)
        if not moment_col:
            print(f"No moment column found in {f.name}. Columns: {list(df.columns)}")
            continue
        title = f.stem
        if args.temp:
            plot_moment_vs_temperature(df, moment_col=moment_col, title=title)
        else:
            plot_moment_vs_field(df, moment_col=moment_col, title=title)
        plt.tight_layout()
        if args.save is not None:
            out = Path(args.save) if args.save else plots_dir / f"{f.stem}.png"
            out.parent.mkdir(parents=True, exist_ok=True)
            plt.savefig(out, dpi=150)
            print(f"Saved {out}")
        else:
            plt.show()


if __name__ == "__main__":
    main()
