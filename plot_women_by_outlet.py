#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
plot_women_by_outlet.py

For each poll (outlet + date) in the "חישוב לפי ערוץ" sheet of the most recent
party-comparison-updated-vXX.xlsx workbook, draws a stacked bar chart of the
expected number of women (bottom, purple) vs. men (top, gray) MKs per party,
and saves it as a separate .jpg file.

Requirements:
    pip install pandas matplotlib openpyxl python-bidi

Usage:
    python plot_women_by_outlet.py [--input-dir DIR] [--output-dir DIR]

    --input-dir   Folder to search for party-comparison-updated-vXX.xlsx
                  files (defaults to the current directory). The file with
                  the highest vXX number is used.
    --output-dir  Folder to write the .jpg files to (defaults to the
                  input directory).
"""

import argparse
import re
import sys
from pathlib import Path

# On Windows, the console's default codepage (e.g. cp1252) can't encode
# Hebrew, which crashes plain print() calls. Force UTF-8 output (with
# fallback to '?'-style replacement) so the script never dies on a print.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass  # older Python without reconfigure(); printing may still fail

import matplotlib
matplotlib.use("Agg")
import matplotlib.colors as mcolors
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Patch
import pandas as pd

try:
    from bidi.algorithm import get_display
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "Missing dependency 'python-bidi'. Install it with:\n"
        "    pip install python-bidi"
    ) from exc


SHEET_NAME = "חישוב לפי ערוץ"

COL_OUTLET = "כלי תקשורת"
COL_DATE = "תאריך הסקר"
COL_PARTY = "מפלגה"
COL_SEATS = "מנדטים"
COL_WOMEN = "כמות נשים צפויה"
COL_MEN = "כמות גברים צפויה"

COLOR_WOMEN = "#7B2D8E"   # purple
COLOR_MEN = "#BFBFBF"     # light gray

VERSION_RE = re.compile(r"^party-comparison-updated-v(\d+)\.xlsx$")

# How far the pseudo-3D extrusion reaches, in x-axis units (bar-index
# spacing) and y-axis units (data units) respectively.
DEPTH_X = 0.16
DEPTH_Y_FRAC = 0.035  # fraction of tick_max


def _shade(color, amount):
    """Blend `color` toward white (amount > 0) or black (amount < 0)."""
    r, g, b = mcolors.to_rgb(color)
    if amount >= 0:
        r, g, b = (r + (1 - r) * amount, g + (1 - g) * amount, b + (1 - b) * amount)
    else:
        r, g, b = (r * (1 + amount), g * (1 + amount), b * (1 + amount))
    return (r, g, b)


def draw_3d_bar_segment(ax, center, y0, y1, width, color, draw_top, depth_y, zorder_base=3):
    """Draw one stacked segment (from y0 to y1) of a pseudo-3D bar: a flat
    front face, a darker side face extruded to the upper-right, and
    (only for the topmost segment) a lighter top cap."""
    if y1 <= y0:
        return
    left = center - width / 2
    right = center + width / 2
    dx, dy = DEPTH_X, depth_y

    front_color = color
    side_color = _shade(color, -0.30)
    top_color = _shade(color, 0.35)

    # Front face
    ax.add_patch(Polygon(
        [(left, y0), (right, y0), (right, y1), (left, y1)],
        closed=True, facecolor=front_color, edgecolor="none",
        zorder=zorder_base,
    ))
    # Side (right) face
    ax.add_patch(Polygon(
        [(right, y0), (right + dx, y0 + dy), (right + dx, y1 + dy), (right, y1)],
        closed=True, facecolor=side_color, edgecolor="none",
        zorder=zorder_base - 0.5,
    ))
    # Top cap (only for the segment that forms the visible top of the stack)
    if draw_top:
        ax.add_patch(Polygon(
            [(left, y1), (right, y1), (right + dx, y1 + dy), (left + dx, y1 + dy)],
            closed=True, facecolor=top_color, edgecolor="none",
            zorder=zorder_base + 0.5,
        ))


def rtl(text: str) -> str:
    """Reorder Hebrew (and mixed Hebrew/number) text for correct display
    in matplotlib, which does not apply the Unicode bidi algorithm itself."""
    return get_display(str(text))


def find_latest_workbook(input_dir: Path) -> Path:
    candidates = []
    for f in input_dir.glob("party-comparison-updated-v*.xlsx"):
        if f.name.startswith("~$"):
            continue
        m = VERSION_RE.match(f.name)
        if m:
            candidates.append((int(m.group(1)), f))
    if not candidates:
        raise FileNotFoundError(
            f"No party-comparison-updated-vXX.xlsx file found in {input_dir}"
        )
    candidates.sort(key=lambda t: t[0])
    return candidates[-1][1]


def format_poll_date(raw) -> str:
    """Format a poll date as D.M (day.month, no leading zeros), Israeli style."""
    ts = pd.to_datetime(raw)
    return f"{ts.day}.{ts.month}"


def sanitize_filename(text: str) -> str:
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    text = text.strip().replace(" ", "_")
    return text


def plot_poll(df_poll: pd.DataFrame, outlet: str, date_raw, out_path: Path) -> None:
    # Match the reference chart's party ordering: sorted by total seats
    # (מנדטים) descending, ties keep the workbook's original row order.
    df_sorted = df_poll.sort_values(COL_SEATS, ascending=False, kind="stable")

    parties = df_sorted[COL_PARTY].tolist()
    women = df_sorted[COL_WOMEN].tolist()
    men = df_sorted[COL_MEN].tolist()
    totals = [w + m for w, m in zip(women, men)]

    total_women = int(df_poll[COL_WOMEN].sum())
    date_str = format_poll_date(date_raw)

    x_labels = [rtl(p) for p in parties]
    x = range(len(parties))

    fig, ax = plt.subplots(figsize=(13, 8.5))

    # Standard y-axis range (just a little headroom above the tallest bar),
    # computed up front since the 3D extrusion depth is scaled from it.
    max_total = max(totals) if totals else 0
    tick_max = max(5, (max_total // 5 + 1) * 5)
    y_max = tick_max * 1.08
    depth_y = tick_max * DEPTH_Y_FRAC

    width = 0.6
    for i, (w, m) in enumerate(zip(women, men)):
        draw_3d_bar_segment(ax, i, 0, w, width, COLOR_WOMEN, draw_top=(m == 0 and w > 0), depth_y=depth_y)
        draw_3d_bar_segment(ax, i, w, w + m, width, COLOR_MEN, draw_top=(m > 0), depth_y=depth_y)

    # Legend proxies (the bars themselves are drawn as custom polygons, not
    # ax.bar patches, so they need explicit legend handles).
    legend_handles = [
        Patch(facecolor=COLOR_WOMEN, label=rtl("כמות נשים צפויה")),
        Patch(facecolor=COLOR_MEN, label=rtl("כמות גברים צפויה")),
    ]

    # Data labels inside each non-zero segment
    for i, (w, m) in enumerate(zip(women, men)):
        if w > 0:
            ax.text(i, w / 2, str(w), ha="center", va="center",
                    color="white", fontsize=23, fontweight="bold", zorder=4)
        if m > 0:
            ax.text(i, w + m / 2, str(m), ha="center", va="center",
                    color="black", fontsize=23, fontweight="bold", zorder=4)
        if w == 0 and m == 0:
            ax.text(i, 0.4, "0", ha="center", va="bottom",
                    color="black", fontsize=19, fontweight="bold", zorder=4)

    ax.set_xlim(-0.5, len(parties) - 1 + 0.5 + DEPTH_X)
    ax.set_ylim(0, y_max)
    ax.set_yticks(range(0, tick_max + 1, 5))

    ax.set_xticks(list(x))
    ax.set_xticklabels(x_labels, rotation=35, ha="right", fontsize=17)

    for spine in ("top", "right", "left"):
        ax.spines[spine].set_visible(False)
    ax.tick_params(axis="y", length=0, labelsize=16)
    ax.tick_params(axis="x", length=0)

    title = "\n".join([
        rtl('מספר ח"כים וח"כיות צפויה'),
        rtl(f"לפי סקר {outlet} מה- {date_str}"),
        rtl(f"רק {total_women} נשים בכנסת הבאה"),
    ])
    # Floated inside the plot area, over the open space above the shorter
    # middle bars (like a text box dragged onto the chart in Excel) rather
    # than anchored to the top edge or pushing the graph down.
    ax.text(0.5, 0.75, title, transform=ax.transAxes, ha="center", va="center",
            fontsize=27, linespacing=1.5, zorder=5)

    ax.legend(
        handles=legend_handles,
        loc="upper center", bbox_to_anchor=(0.5, -0.32),
        ncol=2, frameon=False, fontsize=18,
    )

    fig.subplots_adjust(top=0.95, bottom=0.36, left=0.06, right=0.97)
    fig.savefig(out_path, format="jpg", dpi=200)
    plt.close(fig)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", default=".",
                         help="Folder containing party-comparison-updated-vXX.xlsx")
    parser.add_argument("--output-dir", default=None,
                         help="Folder to write the .jpg files to "
                              "(defaults to --input-dir)")
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir) if args.output_dir else input_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    workbook_path = find_latest_workbook(input_dir)
    print(f"Using workbook: {workbook_path.name}")

    df = pd.read_excel(workbook_path, sheet_name=SHEET_NAME)

    # Preserve the order in which outlet/date polls first appear in the sheet
    polls = df[[COL_OUTLET, COL_DATE]].drop_duplicates().itertuples(index=False)

    for outlet, date_raw in polls:
        df_poll = df[(df[COL_OUTLET] == outlet) & (df[COL_DATE] == date_raw)]
        date_str = format_poll_date(date_raw)
        fname = f"women_seats_{sanitize_filename(outlet)}_{date_str.replace('.', '-')}.jpg"
        out_path = output_dir / fname
        plot_poll(df_poll, outlet, date_raw, out_path)
        print(f"  wrote {out_path.name}")

    print("Done.")


if __name__ == "__main__":
    main()
