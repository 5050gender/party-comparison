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

Optionally (with --create-email-drafts), also creates a Gmail DRAFT (never
sent automatically) per outlet, summarizing that poll and attaching its
charts.

Requirements:
    pip install pandas matplotlib openpyxl python-bidi

Usage:
    python plot_women_by_outlet.py [--input-dir DIR] [--output-dir DIR]
                                    [--with-pie-charts] [--mapping-csv FILE]
                                    [--create-email-drafts]
                                    [--gmail-app-password PASSWORD]

    --input-dir            Folder to search for
                            party-comparison-updated-vXX.xlsx files (defaults
                            to the current directory). The file with the
                            highest vXX number is used.
    --output-dir            Folder to write the .jpg files to (defaults to
                            the input directory).
    --with-pie-charts       Also generate a per-poll pie chart of expected
                            women by bloc (optional; off by default).
    --mapping-csv           Path to the party -> bloc mapping CSV (defaults
                            to mapping.csv inside --input-dir). Used with
                            --with-pie-charts and/or --create-email-drafts.
    --create-email-drafts   Also create a Gmail draft per outlet (not sent)
                            summarizing the poll, with the charts attached.
                            Requires IMAP enabled on the Gmail account and a
                            Gmail App Password (see --gmail-app-password).
    --gmail-app-password    Gmail App Password for the sending account
                            (Google Account -> Security -> 2-Step
                            Verification -> App Passwords -- a normal login
                            password will NOT work). Defaults to the
                            GMAIL_APP_PASSWORD environment variable, which is
                            safer than passing it on the command line.
"""

import argparse
import imaplib
import math
import os
import re
import sys
import time
from email.message import EmailMessage
from email.utils import formataddr
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
from matplotlib.patches import Wedge
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

# "Mean poll" sheet: an aggregate (average-of-polls) estimate per party, not
# tied to any single outlet/date. Shares the same women/men column names as
# the per-outlet sheet, but its seats column is named differently, and it has
# a trailing totals row that must be excluded before plotting.
MEAN_SHEET_NAME = "חישוב 2026"
MEAN_COL_SEATS = "מנדטים צפויים"
TOTAL_ROW_LABEL = 'סה"כ'

COLOR_WOMEN = "#7B2D8E"   # purple
COLOR_MEN = "#BFBFBF"     # light gray

MAPPING_COL_PARTY = "מפלגה"
MAPPING_COL_GROUP = "גוש"

# The pie chart reuses the bar chart's purple: the "אופוזיציה" (opposition)
# bloc gets the same full-strength purple as the "women" bars, and every
# other bloc gets a progressively lighter tint of that same purple (never an
# unrelated hue), so the two charts read as one visual system.
OPPOSITION_GROUP_NAME = "אופוזיציה"
COALITION_GROUP_NAME = "קואליציה"
OTHER_GROUP_TINTS = [0.55, 0.80, 0.92]  # lighten amounts for non-opposition blocs

# --- Email drafts (--create-email-drafts) ---------------------------------
# Creates a Gmail DRAFT (never sends) per outlet, summarizing that poll and
# attaching its bar + bloc charts. Uses IMAP APPEND to Gmail's Drafts folder,
# which requires IMAP enabled on the account and a 16-character Gmail "App
# Password" (Google Account -> Security -> 2-Step Verification -> App
# Passwords) -- a normal Gmail login password will NOT work here. Pass the
# app password via --gmail-app-password, or (safer, keeps it out of shell
# history) set the GMAIL_APP_PASSWORD environment variable.
EMAIL_SENDER = "einatact50@gmail.com"
EMAIL_TO_NAME = "info.5050@merkazim.org"
EMAIL_TO_ADDR = "info@5050il.co.il"
IMAP_HOST = "imap.gmail.com"
IMAP_DRAFTS_FOLDER = "[Gmail]/Drafts"

# Half-donut ("half bagel") geometry: inner/outer ring radius and the
# angular gap left between adjacent segments, in degrees.
DONUT_OUTER_R = 1.0
DONUT_INNER_R = 0.55
DONUT_GAP_DEG = 1.5

VERSION_RE = re.compile(r"^party-comparison-updated-v(\d+)\.xlsx$")


def _shade(color, amount):
    """Blend `color` toward white (amount > 0) or black (amount < 0)."""
    r, g, b = mcolors.to_rgb(color)
    if amount >= 0:
        r, g, b = (r + (1 - r) * amount, g + (1 - g) * amount, b + (1 - b) * amount)
    else:
        r, g, b = (r * (1 + amount), g * (1 + amount), b * (1 + amount))
    return (r, g, b)


def _donut_point(theta_deg, r):
    """A point at angle `theta_deg` (standard math convention: 0=east,
    counterclockwise-positive) and radius `r`, centered on the origin."""
    t = math.radians(theta_deg)
    return r * math.cos(t), r * math.sin(t)


def draw_half_donut(ax, values, colors, start_angle=180, end_angle=0,
                     outer_r=DONUT_OUTER_R, inner_r=DONUT_INNER_R,
                     gap_deg=DONUT_GAP_DEG):
    """Draw a flat half-donut ("half bagel"): a semicircular ring split into
    segments proportional to `values`, opening downward (start_angle=180 on
    the left sweeping clockwise to end_angle=0 on the right, through the top).
    A `values` entry of 0 draws no ring segment but still reserves its
    angular position, which callers can use to place a callout label.
    Returns the segment boundary angles (len(values) + 1)."""
    total = sum(values)
    span = start_angle - end_angle
    boundaries = [start_angle]
    theta = start_angle
    for v in values:
        theta -= (v / total) * span if total else 0.0
        boundaries.append(theta)

    for i, v in enumerate(values):
        if v <= 0:
            continue
        lo, hi = sorted((boundaries[i], boundaries[i + 1]))
        g = min(gap_deg, (hi - lo) * 0.15)
        lo, hi = lo + g, hi - g
        ax.add_patch(Wedge(
            (0.0, 0.0), outer_r, lo, hi, width=outer_r - inner_r,
            facecolor=colors[i], edgecolor="white", linewidth=2, zorder=3,
        ))

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


def build_detail_lines(outlet: str, date_str: str) -> str:
    """Single-line subtitle (smaller font, right-aligned when drawn)."""
    return rtl(f'מספר ח"כים וח"כיות בכנסת הבאה לפי סקר {outlet} מה- {date_str}')


def build_leading_line(parties: list, women: list) -> str:
    """Top, larger-font headline: the party with the most expected women,
    e.g. 'הליכוד מובילה עם 5 נשים'. Ties keep the first occurrence in the
    given (seat-sorted) order."""
    if not parties:
        return ""
    max_women = max(women)
    idx = women.index(max_women)
    leading_party = parties[idx]
    return rtl(f"{leading_party} מובילה עם {max_women} נשים")


def build_women_count_line(total_women: int) -> str:
    """Prominent 'N women' line, shared by the bloc (half-donut) chart's
    per-outlet and mean-poll variants."""
    return rtl(f"רק {total_women} נשים בכנסת הבאה")


def load_mapping(mapping_csv: Path) -> dict:
    """Read the party -> bloc mapping CSV (columns: מפלגה, גוש) into a dict."""
    map_df = pd.read_csv(mapping_csv, encoding="utf-8-sig")
    return dict(zip(map_df[MAPPING_COL_PARTY], map_df[MAPPING_COL_GROUP]))


def build_mean_detail_lines() -> str:
    """Single-line subtitle for the mean-poll chart (smaller font,
    right-aligned when drawn)."""
    return rtl('מספר ח"כים וח"כיות בכנסת הבאה לפי ממוצע הסקרים')


def render_bar_chart(parties: list, women: list, men: list, detail_lines: str,
                      out_path: Path) -> None:
    """Shared flat-horizontal-bar renderer, used for both the per-outlet
    poll charts and the mean-poll (average-of-polls) chart."""
    totals = [w + m for w, m in zip(women, men)]

    y_labels = [rtl(p) for p in parties]
    y = list(range(len(parties)))

    fig, ax = plt.subplots(figsize=(12, 9))

    # Women form the base of each bar (starting at the left, x=0), with men
    # stacked after them (further right).
    bar_height = 0.62
    ax.barh(y, women, height=bar_height, color=COLOR_WOMEN,
            label=rtl("כמות נשים צפויה"), zorder=3)
    ax.barh(y, men, left=women, height=bar_height, color=COLOR_MEN,
            label=rtl("כמות גברים צפויה"), zorder=3)

    # Data labels inside each non-zero segment
    for i, (w, m) in enumerate(zip(women, men)):
        if w > 0:
            ax.text(w / 2, i, str(w), ha="center", va="center",
                    color="white", fontsize=18, fontweight="bold", zorder=4)
        if m > 0:
            ax.text(w + m / 2, i, str(m), ha="center", va="center",
                    color="black", fontsize=18, fontweight="bold", zorder=4)
        if w == 0 and m == 0:
            ax.text(0.4, i, "0", ha="left", va="center",
                    color="black", fontsize=15, fontweight="bold", zorder=4)

    # Standard x-axis range (just a little headroom past the longest bar).
    # The axis itself is hidden (data labels are printed on the bars), so
    # this range only controls layout, not any visible ticks/labels.
    max_total = max(totals) if totals else 0
    tick_max = max(5, (max_total // 5 + 1) * 5)
    x_max = tick_max * 1.08
    ax.set_xlim(0, x_max)
    ax.set_xticks([])

    ax.set_yticks(y)
    ax.set_yticklabels(y_labels, fontsize=15)
    # Largest party first in `parties` -> drawn at the top of the chart.
    ax.invert_yaxis()

    for spine in ("top", "right", "left", "bottom"):
        ax.spines[spine].set_visible(False)
    ax.tick_params(axis="x", length=0, labelbottom=False)
    ax.tick_params(axis="y", length=0)

    # Title block at the very top of the figure (not overlaid on the bars).
    # A larger-font headline (leading party by expected women) sits on top,
    # with the previous two-line detail block underneath it.
    leading_line = build_leading_line(parties, women)
    fig.text(0.95, 0.985, leading_line, ha="right", va="top",
              fontsize=29, fontweight="bold", linespacing=1.3, zorder=5)
    fig.text(0.95, 0.895, detail_lines, ha="right", va="top",
              fontsize=17, zorder=5)

    ax.legend(
        loc="upper center", bbox_to_anchor=(0.5, -0.09),
        ncol=2, frameon=False, fontsize=16,
    )

    fig.subplots_adjust(top=0.76, bottom=0.16, left=0.16, right=0.97)
    fig.savefig(out_path, format="jpg", dpi=200, bbox_inches="tight", pad_inches=0.25)
    plt.close(fig)


def plot_poll(df_poll: pd.DataFrame, outlet: str, date_raw, out_path: Path) -> None:
    # Ordered by expected number of women descending (most women at the
    # top), ties keep the workbook's original row order.
    df_sorted = df_poll.sort_values(COL_WOMEN, ascending=False, kind="stable")

    parties = df_sorted[COL_PARTY].tolist()
    women = df_sorted[COL_WOMEN].tolist()
    men = df_sorted[COL_MEN].tolist()

    date_str = format_poll_date(date_raw)
    detail_lines = build_detail_lines(outlet, date_str)

    render_bar_chart(parties, women, men, detail_lines, out_path)


def plot_mean_poll(df_mean: pd.DataFrame, out_path: Path) -> None:
    """Chart for the aggregate 'mean poll' (average-of-polls) estimate in
    the חישוב 2026 sheet -- same visual style as a per-outlet poll chart,
    but with no outlet/date in the title and the trailing totals row
    (סה"כ) excluded before sorting/plotting."""
    df = df_mean[df_mean[COL_PARTY] != TOTAL_ROW_LABEL]
    df_sorted = df.sort_values(COL_WOMEN, ascending=False, kind="stable")

    parties = df_sorted[COL_PARTY].tolist()
    women = df_sorted[COL_WOMEN].tolist()
    men = df_sorted[COL_MEN].tolist()

    detail_lines = build_mean_detail_lines()

    render_bar_chart(parties, women, men, detail_lines, out_path)


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


def build_mean_pie_title(total_women: int) -> str:
    """Title for the mean-poll bloc chart -- same three-line format as
    build_pie_title, but with no outlet/date (matches build_mean_detail_lines'
    'ממוצע הסקרים' framing)."""
    return "\n".join([
        rtl('מספר ח"כים וח"כיות צפויה'),
        rtl("לפי ממוצע הסקרים"),
        rtl(f"רק {total_women} נשים בכנסת הבאה"),
    ])


def render_bloc_chart(df: pd.DataFrame, mapping: dict, count_line: str,
                       detail_line: str, out_path: Path, skip_label: str) -> bool:
    """Shared half-donut ("half bagel") bloc-chart renderer, used for both
    per-outlet polls and the mean-poll aggregate. Blocs with 0 expected
    women still get a callout showing 0 rather than silently disappearing.
    Returns False (writing nothing) if there are 0 expected women overall."""
    df = df.copy()
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

    if total_women == 0:
        print(f"  skipping pie chart for {skip_label}: 0 expected women")
        return False

    values = [int(v) for v in by_group.values]
    colors = [group_color.get(g, "#8C8C8C") for g in group_order]

    fig, ax = plt.subplots(figsize=(9, 6.5))

    outer_r, inner_r = DONUT_OUTER_R, DONUT_INNER_R
    boundaries = draw_half_donut(ax, values, colors, start_angle=180, end_angle=0,
                                  outer_r=outer_r, inner_r=inner_r)

    # Non-empty blocs get their value + name/% printed directly on their own
    # ring segment (stacked along the segment's own radial direction, so it
    # reads correctly at any angle), colored for contrast against that
    # segment's fill.
    r_mid = (inner_r + outer_r) / 2
    r_offset = (outer_r - inner_r) * 0.24

    for i, (g, v) in enumerate(zip(group_order, values)):
        if v <= 0:
            continue
        mid = (boundaries[i] + boundaries[i + 1]) / 2
        pct = v / total_women * 100
        on_segment_color = "white" if g == OPPOSITION_GROUP_NAME else "#262626"

        vx, vy = _donut_point(mid, r_mid + r_offset)
        ax.text(vx, vy, str(v), ha="center", va="center",
                fontsize=24, fontweight="bold", color=on_segment_color, zorder=4)

        nx, ny = _donut_point(mid, r_mid - r_offset)
        ax.text(nx, ny, f"{rtl(g)}\n({pct:.0f}%)", ha="center", va="center",
                fontsize=13, fontweight="bold", color=on_segment_color, zorder=4)

    # Blocs with 0 expected women (e.g. "חדש תע"ל") have no colored segment
    # to print on, so they still get a muted callout in the hole below the
    # ring rather than silently disappearing. Spread multiple empty blocs
    # into an evenly-spaced row so they never overlap each other.
    empty_groups = [g for g, v in zip(group_order, values) if v <= 0]
    if empty_groups:
        y_value, y_name = inner_r * 0.55, inner_r * 0.22
        slot_max = min(0.35, inner_r * 0.6)
        if len(empty_groups) > 1:
            slot_x = [-slot_max + i * (2 * slot_max) / (len(empty_groups) - 1)
                      for i in range(len(empty_groups))]
        else:
            slot_x = [0.0]
        for x, g in zip(slot_x, empty_groups):
            ax.text(x, y_value, "0", ha="center", va="center",
                    fontsize=24, fontweight="bold", color="#9a9a9a", zorder=4)
            ax.text(x, y_name, f"{rtl(g)}\n(0%)", ha="center", va="center",
                    fontsize=13, fontweight="bold", color="#9a9a9a", zorder=4)

    pad = outer_r * 0.35
    ax.set_xlim(-outer_r - pad, outer_r + pad)
    ax.set_ylim(-outer_r * 0.15, outer_r + pad)
    ax.set_aspect("equal", adjustable="box")
    ax.axis("off")

    # Prominent "N women" line on top (mirrors the bar chart's leading-party
    # headline), with the smaller, right-aligned subtitle beneath it.
    ax.text(0.98, 1.14, count_line, transform=ax.transAxes, ha="right", va="bottom",
            fontsize=22, fontweight="bold", zorder=5)
    ax.text(0.98, 1.04, detail_line, transform=ax.transAxes, ha="right", va="bottom",
            fontsize=16, zorder=5)

    fig.savefig(out_path, format="jpg", dpi=200, bbox_inches="tight", pad_inches=0.35)
    plt.close(fig)
    return True


def plot_pie_poll(df_poll: pd.DataFrame, outlet: str, date_raw, mapping: dict,
                   out_path: Path) -> bool:
    """Half-donut bloc chart for a single outlet's poll. Returns False
    (writing nothing) if there are 0 expected women overall."""
    date_str = format_poll_date(date_raw)
    total_women = int(df_poll[COL_WOMEN].sum())
    count_line = build_women_count_line(total_women)
    detail_line = build_detail_lines(outlet, date_str)
    return render_bloc_chart(df_poll, mapping, count_line, detail_line, out_path,
                              skip_label=f"{outlet} ({date_str})")


def plot_mean_pie_poll(df_mean: pd.DataFrame, mapping: dict, out_path: Path) -> bool:
    """Half-donut bloc chart for the aggregate 'mean poll' (average-of-polls)
    estimate in the חישוב 2026 sheet, with the trailing totals row (סה"כ)
    excluded first. Returns False (writing nothing) if there are 0 expected
    women overall."""
    df = df_mean[df_mean[COL_PARTY] != TOTAL_ROW_LABEL]
    total_women = int(df[COL_WOMEN].sum())
    count_line = build_women_count_line(total_women)
    detail_line = build_mean_detail_lines()
    return render_bloc_chart(df, mapping, count_line, detail_line, out_path,
                              skip_label="ממוצע הסקרים")


def compute_poll_email_stats(df_poll: pd.DataFrame, mapping: dict) -> dict:
    """Leading party (by expected women) + opposition/coalition/total women
    totals for a single poll, used to fill in the email template. Note:
    unlike rtl(), these party/bloc names are left in normal (logical) Hebrew
    order -- email clients do their own bidi rendering, so reordering here
    would show up backwards."""
    women_by_party = df_poll.set_index(COL_PARTY)[COL_WOMEN]
    leading_party = women_by_party.idxmax()
    leading_women = int(women_by_party.max())

    groups = df_poll[COL_PARTY].map(mapping)
    opposition_women = int(df_poll.loc[groups == OPPOSITION_GROUP_NAME, COL_WOMEN].sum())
    coalition_women = int(df_poll.loc[groups == COALITION_GROUP_NAME, COL_WOMEN].sum())
    total_women = int(df_poll[COL_WOMEN].sum())

    return {
        "leading_party": leading_party,
        "leading_women": leading_women,
        "opposition_women": opposition_women,
        "coalition_women": coalition_women,
        "total_women": total_women,
    }


def build_email_subject(outlet: str, date_str: str) -> str:
    return (f'מספר ח"כיות וחכ"ים צפויה בכל מפלגה ובכל גוש, '
            f'לפי סקר {outlet} מתאריך {date_str}')


def build_email_body(date_str: str, stats: dict) -> str:
    return (
        "שלום רב,\n"
        f"על פי הסקר שלכם מתאריך {date_str} מסתמן שבמפלגה המובילה מבחינת "
        f"מספר נשים, היא מפלגת {stats['leading_party']} עם "
        f"{stats['leading_women']} נשים.\n"
        f"בחלוקה עפ\"י גושים, בגוש האופוזיציה יש {stats['opposition_women']} "
        f"נשים, ובגוש הקואליציה יש {stats['coalition_women']} נשים.\n"
        f"סה\"כ צפויות להיות בכנסת הבאה {stats['total_women']} נשים.\n"
        "אתם מוזמנים להשתמש בגרפים המצורפים.\n"
        "\n"
        "בברכה,\n"
        "צוות 5050"
    )


def build_email_message(subject: str, body: str, attachments: list) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = EMAIL_SENDER
    msg["To"] = formataddr((EMAIL_TO_NAME, EMAIL_TO_ADDR))
    msg["Subject"] = subject
    msg.set_content(body)
    for path in attachments:
        data = Path(path).read_bytes()
        msg.add_attachment(data, maintype="image", subtype="jpeg",
                            filename=Path(path).name)
    return msg


def save_gmail_draft(msg: EmailMessage, app_password: str) -> None:
    """Append `msg` to the Gmail Drafts folder via IMAP (never sends it)."""
    with imaplib.IMAP4_SSL(IMAP_HOST) as imap:
        status, _ = imap.login(EMAIL_SENDER, app_password)
        if status != "OK":
            raise RuntimeError(f"IMAP login failed for {EMAIL_SENDER}")
        status, _ = imap.append(
            IMAP_DRAFTS_FOLDER, "\\Draft",
            imaplib.Time2Internaldate(time.time()), msg.as_bytes(),
        )
        if status != "OK":
            raise RuntimeError(f"Failed to save draft to {IMAP_DRAFTS_FOLDER}")


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
    parser.add_argument("--create-email-drafts", action="store_true",
                         help=f"Also create a Gmail DRAFT (not sent) per "
                              f"outlet in {EMAIL_SENDER}'s Drafts folder, "
                              f"summarizing that poll and attaching its "
                              f"charts. Needs mapping.csv (loaded "
                              f"automatically even without --with-pie-charts) "
                              f"and a Gmail App Password.")
    parser.add_argument("--gmail-app-password", default=None,
                         help="Gmail App Password for %s (needs IMAP enabled "
                              "+ a 16-char App Password from Google Account "
                              "-> Security -> App Passwords). Defaults to the "
                              "GMAIL_APP_PASSWORD environment variable, which "
                              "is safer than passing it on the command line."
                              % EMAIL_SENDER)
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir) if args.output_dir else input_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    workbook_path = find_latest_workbook(input_dir)
    print(f"Using workbook: {workbook_path.name}")

    df = pd.read_excel(workbook_path, sheet_name=SHEET_NAME)

    # Skip incomplete rows (missing outlet or date). This is usually not
    # actually missing data -- it typically happens when a poll's row uses
    # formulas (=INDEX/MATCH...) that pull from another sheet, and the
    # workbook was last saved by a tool that writes cells directly (e.g. via
    # openpyxl) without Excel/LibreOffice recalculating first. Python can
    # only read the last-cached formula result, not evaluate the formula
    # itself, so it sees blanks even though Excel would show real values.
    # Without this filter, a missing date also silently matches nothing when
    # filtered by equality (NaN != NaN), which produced bogus "nan-nan"
    # charts and crashed the e-mail stats step.
    incomplete_mask = df[COL_OUTLET].isna() | df[COL_DATE].isna()
    if incomplete_mask.any():
        # Name the affected outlet(s) so this is obvious at a glance, not a
        # silent skip. A row missing even its outlet name is reported as
        # "(row N)" using its 1-based position in the sheet (+2 for the
        # header row and 0-based index).
        affected = (
            df.loc[incomplete_mask, COL_OUTLET]
            .fillna("")
            .replace("", pd.NA)
        )
        labels = []
        for idx, outlet_name in affected.items():
            if pd.isna(outlet_name):
                labels.append(f"(row {idx + 2})")
            else:
                labels.append(str(outlet_name))
        counts = pd.Series(labels).value_counts()
        summary = ", ".join(f"{name} ({n} row{'s' if n != 1 else ''})"
                             for name, n in counts.items())
        print(f"  warning: skipping {int(incomplete_mask.sum())} row(s) with "
              f"a missing outlet/date, excluded from all charts/emails: {summary}")
        print(f"    (if the data looks complete when you open the workbook, "
              f"this is likely stale formula cells -- open the file in "
              f"Excel/LibreOffice, let it recalculate, and save before "
              f"rerunning)")
        df = df[~incomplete_mask]

    app_password = None
    if args.create_email_drafts:
        app_password = args.gmail_app_password or os.environ.get("GMAIL_APP_PASSWORD")
        if not app_password:
            raise SystemExit(
                "--create-email-drafts needs a Gmail App Password: pass "
                "--gmail-app-password or set the GMAIL_APP_PASSWORD "
                "environment variable."
            )

    mapping = None
    if args.with_pie_charts or args.create_email_drafts:
        mapping_csv = Path(args.mapping_csv) if args.mapping_csv else input_dir / "mapping.csv"
        if not mapping_csv.exists():
            raise FileNotFoundError(
                f"--with-pie-charts/--create-email-drafts needs {mapping_csv}, "
                f"which was not found."
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

        pie_out_path = None
        if mapping is not None:
            pie_fname = f"women_by_bloc_{sanitize_filename(outlet)}_{date_str.replace('.', '-')}.jpg"
            pie_out_path = output_dir / pie_fname
            plot_pie_poll(df_poll, outlet, date_raw, mapping, pie_out_path)
            if pie_out_path.exists():
                print(f"  wrote {pie_out_path.name}")
            else:
                pie_out_path = None  # plot_pie_poll skips 0-women polls

        if args.create_email_drafts:
            if df_poll.empty or df_poll[COL_WOMEN].sum() == 0:
                print(f"  skipping email draft for {outlet} ({date_str}): no data")
                continue
            stats = compute_poll_email_stats(df_poll, mapping)
            subject = build_email_subject(outlet, date_str)
            body = build_email_body(date_str, stats)
            attachments = [out_path] + ([pie_out_path] if pie_out_path else [])
            msg = build_email_message(subject, body, attachments)
            save_gmail_draft(msg, app_password)
            print(f"  created email draft: {subject}")

    # Also plot the aggregate "mean poll" (average-of-polls) estimate.
    df_mean = pd.read_excel(workbook_path, sheet_name=MEAN_SHEET_NAME)
    mean_out_path = output_dir / "women_seats_ממוצע_סקרים.jpg"
    plot_mean_poll(df_mean, mean_out_path)
    print(f"  wrote {mean_out_path.name}")

    if mapping is not None:
        mean_pie_out_path = output_dir / "women_by_bloc_ממוצע_סקרים.jpg"
        if plot_mean_pie_poll(df_mean, mapping, mean_pie_out_path):
            print(f"  wrote {mean_pie_out_path.name}")

    print("Done.")


if __name__ == "__main__":
    main()
