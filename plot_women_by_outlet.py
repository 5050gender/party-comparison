#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
plot_women_by_outlet.py

For each poll (outlet + date) in the "חישוב לפי ערוץ" sheet of the most recent
party-comparison-updated-vXX.xlsx workbook, draws a stacked bar chart of the
expected number of women (bottom, purple) vs. men (top, gray) MKs per party,
and saves it as a separate .jpg file.

Optionally (with --with-pie-charts), also draws a pie chart per poll showing
the expected number of women split across party "blocs" (groups), using the
party -> bloc mapping in mapping.csv (columns: מפלגה, גוש).

Requirements:
    pip install pandas matplotlib openpyxl python-bidi

Usage:
    python plot_women_by_outlet.py [--input-dir DIR] [--output-dir DIR]
                                    [--with-pie-charts] [--mapping-csv FILE]

    --input-dir         Folder to search for party-comparison-updated-vXX.xlsx
                         files (defaults to the current directory). The file
                         with the highest vXX number is used.
    --output-dir        Folder to write the .jpg files to (defaults to the
                         input directory).
    --with-pie-charts   Also generate a per-poll pie chart of expected women
                         by bloc (optional; off by default).
    --mapping-csv       Path to the party -> bloc mapping CSV (defaults to
                         mapping.csv inside --input-dir). Only used with
                         --with-pie-charts.
"""

import argparse
import math
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

MAPPING_COL_PARTY = "מפלגה"
MAPPING_COL_GROUP = "גוש"

# The pie chart reuses the bar chart's purple: the "אופוזיציה" (opposition)
# bloc gets the same full-strength purple as the "women" bars, and every
# other bloc gets a progressively lighter tint of that same purple (never an
# unrelated hue), so the two charts read as one visual system.
OPPOSITION_GROUP_NAME = "אופוזיציה"
OTHER_GROUP_TINTS = [0.55, 0.80, 0.92]  # lighten amounts for non-opposition blocs

# Pseudo-3D pie geometry: vertical squash (perspective tilt) and extrusion
# depth, both as fractions of the pie radius.
PIE_SQUASH = 0.58
PIE_DEPTH = 0.16

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


def _pie_point(theta_deg, r, squash=PIE_SQUASH):
    """A point on the pie's rim/interior at angle `theta_deg` (standard math
    convention: 0=east, counterclockwise-positive) and radius `r`, with the
    y-coordinate squashed to fake the tilted-disk 3D perspective."""
    t = math.radians(theta_deg)
    return r * math.cos(t), r * squash * math.sin(t)


def draw_3d_pie(ax, values, colors, start_angle=90, radius=1.0,
                 squash=PIE_SQUASH, depth=PIE_DEPTH):
    """Draw a pseudo-3D pie: a darker extruded side wall along the front
    (bottom) rim, topped with the flat elliptical pie itself. `values`
    entries of 0 simply produce a degenerate (invisible) wedge but still
    reserve their angular position, which callers can use to anchor a
    callout. Returns the wedge boundary angles (len(values) + 1)."""
    total = sum(values)
    boundaries = [start_angle]
    theta = start_angle
    for v in values:
        theta -= (v / total) * 360.0 if total else 0.0
        boundaries.append(theta)

    side_colors = [_shade(c, -0.35) for c in colors]

    def wedge_index(theta):
        for i in range(len(values)):
            hi, lo = boundaries[i], boundaries[i + 1]
            if hi >= theta >= lo - 1e-6:
                return i
        return len(values) - 1

    # Side wall: sample the rim finely and keep only the front-facing
    # (bottom) half, i.e. where the un-squashed sine is negative.
    step = 1.0
    n_steps = int(360 / step)
    rim_thetas = [boundaries[0] - i * step for i in range(n_steps + 1)]
    for i in range(len(rim_thetas) - 1):
        thA, thB = rim_thetas[i], rim_thetas[i + 1]
        mid = (thA + thB) / 2
        if math.sin(math.radians(mid)) < 0:
            widx = wedge_index(mid)
            if values[widx] <= 0:
                continue
            xA, yA = _pie_point(thA, radius, squash)
            xB, yB = _pie_point(thB, radius, squash)
            ax.add_patch(Polygon(
                [(xA, yA), (xB, yB), (xB, yB - depth), (xA, yA - depth)],
                closed=True, facecolor=side_colors[widx], edgecolor="none",
                zorder=1,
            ))

    # Top face: the flat (elliptical) pie wedges, drawn over the side walls.
    for i, v in enumerate(values):
        if v <= 0:
            continue
        t0, t1 = boundaries[i], boundaries[i + 1]
        n = max(2, int(abs(t0 - t1)) + 1)
        pts = [(0.0, 0.0)]
        for k in range(n):
            t = t0 + (t1 - t0) * k / (n - 1)
            pts.append(_pie_point(t, radius, squash))
        pts.append((0.0, 0.0))
        ax.add_patch(Polygon(pts, closed=True, facecolor=colors[i],
                              edgecolor="white", linewidth=2, zorder=3))

    return boundaries


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


def build_title(outlet: str, date_str: str, total_women: int) -> str:
    return "\n".join([
        rtl('מספר ח"כים וח"כיות צפויה'),
        rtl(f"לפי סקר {outlet} מה- {date_str}"),
        rtl(f"רק {total_women} נשים בכנסת הבאה"),
    ])


def load_mapping(mapping_csv: Path) -> dict:
    """Read the party -> bloc mapping CSV (columns: מפלגה, גוש) into a dict."""
    map_df = pd.read_csv(mapping_csv, encoding="utf-8-sig")
    return dict(zip(map_df[MAPPING_COL_PARTY], map_df[MAPPING_COL_GROUP]))


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

    title = build_title(outlet, date_str, total_women)
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


def bloc_colors(group_order: list) -> dict:
    """Fixed, non-cycled color per bloc: the opposition bloc gets the same
    purple as the bar chart's "women" color; every other bloc gets a
    progressively lighter tint of that same purple, assigned in the order
    blocs first appear in mapping.csv (so a bloc's color/tint never changes
    from poll to poll)."""
    colors = {}
    tint_i = 0
    for g in group_order:
        if g == OPPOSITION_GROUP_NAME:
            colors[g] = COLOR_WOMEN
        else:
            amount = OTHER_GROUP_TINTS[tint_i % len(OTHER_GROUP_TINTS)]
            colors[g] = _shade(COLOR_WOMEN, amount)
            tint_i += 1
    return colors


def plot_pie_poll(df_poll: pd.DataFrame, outlet: str, date_raw, mapping: dict,
                   out_path: Path) -> None:
    """Pseudo-3D pie chart of expected women split across party blocs (per
    mapping.csv). Blocs with 0 expected women still get a callout showing 0
    rather than silently disappearing."""
    df = df_poll.copy()
    df["__group"] = df[COL_PARTY].map(mapping)

    unmapped = sorted(df.loc[df["__group"].isna(), COL_PARTY].unique())
    if unmapped:
        print(f"  warning: no bloc mapping for: {', '.join(unmapped)} "
              f"(grouped under \"{unmapped[0]}\" as-is)")
        df["__group"] = df["__group"].fillna(df[COL_PARTY])

    group_order = list(dict.fromkeys(mapping.values()))
    group_color = bloc_colors(group_order)

    # Keep every bloc (including ones with 0 women) in a fixed order, so
    # an empty bloc still reserves an angular slot for its callout.
    by_group = df.groupby("__group", sort=False)[COL_WOMEN].sum()
    by_group = by_group.reindex(group_order, fill_value=0)

    total_women = int(by_group.sum())
    date_str = format_poll_date(date_raw)

    if total_women == 0:
        print(f"  skipping pie chart for {outlet} ({date_str}): 0 expected women")
        return

    values = [int(v) for v in by_group.values]
    colors = [group_color.get(g, "#8C8C8C") for g in group_order]

    fig, ax = plt.subplots(figsize=(9, 9))

    radius = 1.0
    boundaries = draw_3d_pie(ax, values, colors, start_angle=90, radius=radius)

    # Inside count/% labels + outside bloc-name labels, for non-empty wedges.
    for i, (g, v) in enumerate(zip(group_order, values)):
        if v <= 0:
            continue
        mid = (boundaries[i] + boundaries[i + 1]) / 2
        pct = v / total_women * 100

        lx, ly = _pie_point(mid, radius * 0.62)
        label_color = "white" if g == OPPOSITION_GROUP_NAME else "black"
        ax.text(lx, ly, f"{v}\n({pct:.0f}%)", ha="center", va="center",
                fontsize=17, fontweight="bold", color=label_color, zorder=4)

        # Push labels further out when they fall on the front (bottom) half
        # of the pie, so they clear the extruded 3D side wall down there
        # instead of overlapping it.
        label_dist = radius * (1.5 if math.sin(math.radians(mid)) < 0 else 1.2)
        ox, oy = _pie_point(mid, label_dist)
        ha = "left" if math.cos(math.radians(mid)) >= 0 else "right"
        ax.text(ox, oy, rtl(g), ha=ha, va="center",
                fontsize=17, fontweight="bold", zorder=4)

    # Callout for any bloc with 0 expected women (e.g. "חדש תע"ל"): it has
    # no wedge to label directly, so point a leader line at its reserved
    # (degenerate) boundary angle instead of letting it vanish silently.
    for i, (g, v) in enumerate(zip(group_order, values)):
        if v > 0:
            continue
        anchor_theta = boundaries[i]
        ax_x, ax_y = _pie_point(anchor_theta, radius)
        tx, ty = _pie_point(anchor_theta, radius * 1.55)
        ax.annotate(
            f"{rtl(g)}\n0 (0%)",
            xy=(ax_x, ax_y), xytext=(tx, ty),
            ha="center", va="center", fontsize=15, fontweight="bold",
            arrowprops=dict(arrowstyle="-", color="#999999", lw=1.3),
            zorder=5,
        )

    pad = radius * 0.75
    ax.set_xlim(-radius - pad, radius + pad)
    ax.set_ylim(-radius * PIE_SQUASH - PIE_DEPTH - pad * 0.6, radius * PIE_SQUASH + pad * 0.6)
    ax.set_aspect("equal", adjustable="box")
    ax.axis("off")

    title = build_title(outlet, date_str, total_women)
    ax.set_title(title, fontsize=22, linespacing=1.5, pad=20)

    # The pie's data aspect (wide, short ellipse) is much wider than the
    # square figure, which would otherwise letterbox a lot of blank space
    # above/below it; crop to the actual content instead.
    fig.savefig(out_path, format="jpg", dpi=200, bbox_inches="tight", pad_inches=0.35)
    plt.close(fig)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", default=".",
                         help="Folder containing party-comparison-updated-vXX.xlsx")
    parser.add_argument("--output-dir", default=None,
                         help="Folder to write the .jpg files to "
                              "(defaults to --input-dir)")
    parser.add_argument("--with-pie-charts", action="store_true",
                         help="Also generate a per-poll pie chart of expected "
                              "women by bloc (needs mapping.csv)")
    parser.add_argument("--mapping-csv", default=None,
                         help="Path to the party -> bloc mapping CSV "
                              "(defaults to mapping.csv inside --input-dir)")
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir) if args.output_dir else input_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    workbook_path = find_latest_workbook(input_dir)
    print(f"Using workbook: {workbook_path.name}")

    df = pd.read_excel(workbook_path, sheet_name=SHEET_NAME)

    mapping = None
    if args.with_pie_charts:
        mapping_csv = Path(args.mapping_csv) if args.mapping_csv else input_dir / "mapping.csv"
        if not mapping_csv.exists():
            raise FileNotFoundError(
                f"--with-pie-charts was given but {mapping_csv} was not found."
            )
        mapping = load_mapping(mapping_csv)
        print(f"Using bloc mapping: {mapping_csv.name} "
              f"({len(set(mapping.values()))} blocs, {len(mapping)} parties)")

    # Preserve the order in which outlet/date polls first appear in the sheet
    polls = df[[COL_OUTLET, COL_DATE]].drop_duplicates().itertuples(index=False)

    for outlet, date_raw in polls:
        df_poll = df[(df[COL_OUTLET] == outlet) & (df[COL_DATE] == date_raw)]
        date_str = format_poll_date(date_raw)

        fname = f"women_seats_{sanitize_filename(outlet)}_{date_str.replace('.', '-')}.jpg"
        out_path = output_dir / fname
        plot_poll(df_poll, outlet, date_raw, out_path)
        print(f"  wrote {out_path.name}")

        if mapping is not None:
            pie_fname = f"women_by_bloc_{sanitize_filename(outlet)}_{date_str.replace('.', '-')}.jpg"
            pie_out_path = output_dir / pie_fname
            plot_pie_poll(df_poll, outlet, date_raw, mapping, pie_out_path)
            print(f"  wrote {pie_out_path.name}")

    print("Done.")


if __name__ == "__main__":
    main()
