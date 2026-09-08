#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
plot_coalition_women.py

1. Summarizes the number of women per coalition per year, from a CSV with
   (at least) the columns: year, party, number of women
   (produced earlier as coalition_by_knesset.csv).
2. Adds the expected number of women in the NEXT coalition, twice - once
   for a "קואליצית ימין-חרדים" (right-wing/Haredi coalition) scenario and
   once for a "קואליצית גוש שינוי" (change-bloc coalition) scenario. These
   two numbers are supplied by the user, either as command-line arguments
   or interactively when the script runs.
3. Produces two bar plots (one per scenario): the historical per-year
   totals in blue, plus the single added/predicted data point in purple.
4. The Y axis always extends up to 61 seats.
5. Both plots are titled "מספר נשים בקואליציה לאורך השנים".
6. Every bar gets a data label with its value, above the bar.
7. The added (purple) data point additionally gets its scenario name
   written above it.

Usage:
    python plot_coalition_women.py
        (prompts interactively for the two predicted values)

    python plot_coalition_women.py --right-haredi 22 --change-bloc 35
        (supplies the two predicted values directly, no prompts)

    python plot_coalition_women.py --csv path/to/coalition_by_knesset.csv --outdir out/
"""

import argparse
import csv
import os
import platform
import sys
import warnings

# Noto Sans Hebrew (like several Hebrew-only fonts) has no Latin glyphs, and
# matplotlib's text-metrics code probes fonts with reference Latin letters
# ("l", "p", ...) to measure ascent/descent - harmless here since none of our
# actual on-chart text is Latin, but it prints a confusing warning, so it's
# silenced.
warnings.filterwarnings("ignore", message="Glyph .* missing from font")


def _ensure_utf8_console():
    """Make printing/typing Hebrew text at the console safe.

    Windows terminals often default to a legacy codepage (e.g. cp1252) that
    can't encode Hebrew characters at all, which crashes plain print()/input()
    calls with a UnicodeEncodeError. This switches the console (on Windows)
    and stdout/stderr to UTF-8 up front. Best-effort: never raises.
    """
    if platform.system() == "Windows":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
            ctypes.windll.kernel32.SetConsoleCP(65001)
        except Exception:
            pass
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


_ensure_utf8_console()

import matplotlib
matplotlib.use("Agg")  # no display needed - we just save PNG files
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

try:
    from bidi.algorithm import get_display
except ImportError:
    sys.exit(
        "Missing dependency 'python-bidi', needed to display Hebrew text "
        "correctly. Install it with:\n\n    pip install python-bidi\n"
    )

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

DEFAULT_CSV_PATH = "coalition_by_knesset.csv"
DEFAULT_OUTDIR = "."

Y_AXIS_MAX = 61  # requirement 4

TITLE = "מספר נשים בקואליציה לאורך השנים"  # requirement 5
Y_LABEL = "מספר נשים"

BLUE = "#3d7dbf"    # historical data points
PURPLE = "#7b2d8e"  # the single added/predicted data point

SCENARIOS = [
    {
        "arg": "right_haredi",
        "name": "קואליצית ימין-חרדים",
        "out_file": "coalition_women_right_haredi.png",
    },
    {
        "arg": "change_bloc",
        "name": "קואליצית גוש שינוי",
        "out_file": "coalition_women_change_bloc.png",
    },
]


# --------------------------------------------------------------------------
# Hebrew / RTL helpers
# --------------------------------------------------------------------------

def find_hebrew_font():
    """Return a FontProperties for a font that can render Hebrew glyphs."""
    candidates = ["Noto Sans Hebrew", "Arial", "David Libre", "Alef", "Rubik"]
    for name in candidates:
        matches = [f for f in fm.fontManager.ttflist if name.lower() in f.name.lower()]
        if matches:
            return fm.FontProperties(fname=matches[0].fname)
    print(
        "Warning: no Hebrew-capable font was found on this system. "
        "Hebrew text may render as boxes. Install e.g. 'fonts-noto-core'.",
        file=sys.stderr,
    )
    return fm.FontProperties()


def rtl(text):
    """Reorder Hebrew (right-to-left) text so matplotlib draws it correctly."""
    return get_display(text)


# --------------------------------------------------------------------------
# Step 1: summarize the number of women per coalition per year
# --------------------------------------------------------------------------

def load_women_per_year(csv_path):
    """Read the per-party CSV and sum 'number of women' for each year.

    Returns an ordered dict: {year: total_women}
    """
    totals = {}
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        required = {"year", "number of women"}
        if not required.issubset(set(reader.fieldnames or [])):
            sys.exit(
                f"'{csv_path}' must have at least the columns {sorted(required)}; "
                f"found {reader.fieldnames}"
            )
        for row in reader:
            if not row.get("year", "").strip():
                continue
            year = int(row["year"])
            women = int(row["number of women"])
            totals[year] = totals.get(year, 0) + women
    if not totals:
        sys.exit(f"No data rows found in '{csv_path}'.")
    return dict(sorted(totals.items()))


# --------------------------------------------------------------------------
# Step 2: get the two predicted values from the user
# --------------------------------------------------------------------------

def prompt_for_number(scenario_name):
    while True:
        raw = input(
            f"Expected number of women in the next coalition under "
            f"'{scenario_name}': "
        ).strip()
        try:
            value = int(raw)
        except ValueError:
            print("Please enter a whole number.")
            continue
        if value < 0:
            print("Please enter a non-negative number.")
            continue
        return value


def resolve_predicted_values(args):
    """Fill in scenario['predicted_women'] from CLI args, or prompt for it."""
    cli_values = {"right_haredi": args.right_haredi, "change_bloc": args.change_bloc}
    for scenario in SCENARIOS:
        value = cli_values[scenario["arg"]]
        if value is None:
            value = prompt_for_number(scenario["name"])
        scenario["predicted_women"] = value


# --------------------------------------------------------------------------
# Steps 3-7: build one bar plot for a given scenario
# --------------------------------------------------------------------------

def build_plot(year_totals, scenario, hebrew_font, out_path):
    years = list(year_totals.keys())
    women = list(year_totals.values())

    # x labels: historical years as plain numbers, the new point as "הבאה" (next)
    x_labels = [str(y) for y in years] + [rtl("הבאה")]
    all_women = women + [scenario["predicted_women"]]
    colors = [BLUE] * len(women) + [PURPLE]  # requirement 3

    fig, ax = plt.subplots(figsize=(12, 7))
    bars = ax.bar(x_labels, all_women, color=colors, width=0.6, zorder=3)

    # requirement 4: y axis up to 61
    ax.set_ylim(0, Y_AXIS_MAX)

    # requirement 5: title
    title_obj = ax.set_title(rtl(TITLE), fontproperties=hebrew_font, fontweight="bold", pad=45)
    title_obj.set_fontsize(40)
    ax.set_ylabel(rtl(Y_LABEL), fontsize=13, fontproperties=hebrew_font)
    ax.set_xlabel(rtl("שנה"), fontsize=13, fontproperties=hebrew_font)

    # Noto Sans Hebrew (and some other Hebrew fonts) don't include Latin/digit
    # glyphs, so only apply it to labels that actually contain Hebrew text -
    # plain numeric year labels keep the default font, which does have digits.
    for tick in ax.get_xticklabels():
        text = tick.get_text()
        if any(0x0590 <= ord(ch) <= 0x05FF for ch in text):
            tick.set_fontproperties(hebrew_font)
        tick.set_fontsize(11)

    ax.grid(axis="y", linestyle="--", alpha=0.3, zorder=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # requirement 6: a data label above every bar
    for bar, value in zip(bars, all_women):
        ax.annotate(
            str(value),
            xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
            xytext=(0, 4),
            textcoords="offset points",
            ha="center", va="bottom",
            fontsize=11, fontweight="bold",
        )

    # requirement 7: name the added (purple) data point above it
    added_bar = bars[-1]
    annotate_obj=ax.annotate(
        rtl(scenario["name"]),
        xy=(added_bar.get_x() + added_bar.get_width() / 2, added_bar.get_height()),
        xytext=(0, 24),
        textcoords="offset points",
        ha="center", va="bottom",
        fontsize=12, fontweight="bold", color=PURPLE,
        fontproperties=hebrew_font,
    )
    annotate_obj.set_fontsize(20)

    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    print(f"Saved: {out_path}")


# --------------------------------------------------------------------------

def parse_args():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--csv", default=DEFAULT_CSV_PATH,
                    help=f"Path to the per-party CSV (default: {DEFAULT_CSV_PATH})")
    p.add_argument("--outdir", default=DEFAULT_OUTDIR,
                    help="Directory to save the two plots into (default: current directory)")
    p.add_argument("--right-haredi", type=int, default=None,
                    help="Expected number of women under 'קואליצית ימין-חרדים' "
                         "(skips the interactive prompt if given)")
    p.add_argument("--change-bloc", type=int, default=None,
                    help="Expected number of women under 'קואליצית גוש שינוי' "
                         "(skips the interactive prompt if given)")
    return p.parse_args()


def main():
    args = parse_args()

    if not os.path.exists(args.csv):
        sys.exit(f"Could not find '{args.csv}'. Pass --csv to point at the right file.")
    os.makedirs(args.outdir, exist_ok=True)

    hebrew_font = find_hebrew_font()

    # 1. summarize the number of women per coalition per year
    year_totals = load_women_per_year(args.csv)

    # 2. get the two predicted values from the user
    resolve_predicted_values(args)

    # 3-7. build the two bar plots
    for scenario in SCENARIOS:
        out_path = os.path.join(args.outdir, scenario["out_file"])
        build_plot(year_totals, scenario, hebrew_font, out_path)


if __name__ == "__main__":
    main()
