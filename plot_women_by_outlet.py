#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
plot_women_by_outlet.py

For each poll (outlet + date) in the "חישוב לפי ערוץ" sheet of the most recent
party-comparison-updated-vXX.xlsx workbook, draws a bar chart of the expected
number of women (purple) vs. men (gray) MKs per party, and saves it as a
separate .jpg file.

Optionally (with --with-pie-charts), also draws a half-donut "arc" chart per
poll showing the expected number of women/men split across the opposition
and coalition blocs, using the party -> bloc mapping in mapping.csv (columns:
מפלגה, גוש).

Optionally (with --create-email-drafts), also creates a Gmail DRAFT (never
sent automatically) per outlet, summarizing that poll and attaching its
charts.

Charts are drawn as HTML/CSS/SVG (see the "templates" folder next to this
script) and rendered to JPG with a headless browser (Playwright), not with
matplotlib -- this keeps the exact chart design in editable HTML/CSS/SVG
files, so a visual tweak (colors, spacing, fonts) only ever needs a template
edit, never a change to this script.

Requirements:
    pip install pandas openpyxl jinja2 playwright pillow
    playwright install chromium

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
import base64
import imaplib
import io
import math
import os
import re
import sys
import time
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

from PIL import Image

# On Windows, the console's default codepage (e.g. cp1252) can't encode
# Hebrew, which crashes plain print() calls. Force UTF-8 output (with
# fallback to '?'-style replacement) so the script never dies on a print.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass  # older Python without reconfigure(); printing may still fail

import pandas as pd

try:
    from jinja2 import Environment, FileSystemLoader
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "Missing dependency 'jinja2'. Install it with:\n"
        "    pip install jinja2"
    ) from exc

try:
    from playwright.sync_api import sync_playwright
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "Missing dependency 'playwright'. Install it with:\n"
        "    pip install playwright\n"
        "    playwright install chromium"
    ) from exc

# Charts are HTML/CSS/SVG templates (see the "templates" folder next to this
# script) rendered to JPG via a headless browser -- edit the .html.j2 files
# to change fonts, colors, spacing, or layout; this script only ever
# supplies the data.
TEMPLATES_DIR = Path(__file__).resolve().parent / "templates"
FONTS_DIR = Path(__file__).resolve().parent / "fonts"
LOGO_PATH = TEMPLATES_DIR / "assets" / "logo_5050.jpg"
EMAIL_TEMPLATE_PATH = TEMPLATES_DIR / "assets" / "text_for_email.txt"


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

MAPPING_COL_PARTY = "מפלגה"
MAPPING_COL_GROUP = "גוש"

OPPOSITION_GROUP_NAME = "אופוזיציה"
COALITION_GROUP_NAME = "קואליציה"

# Shortened party names for chart display only -- the workbook/mapping.csv
# keep using the full official name everywhere else (candidate lists, bloc
# lookups, etc.); only what's actually drawn on a chart uses the shorter
# form, so this must never be applied before a mapping.get(party) lookup.
CHART_DISPLAY_NAME_OVERRIDES = {
    "ביחד (בנט-לפיד)": "ביחד",
    "המילואימניקים-הכלכלית": "המילואימניקים",
}


def _chart_display_name(party: str) -> str:
    return CHART_DISPLAY_NAME_OVERRIDES.get(party, party)

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

VERSION_RE = re.compile(r"^party-comparison-updated-v(\d+)\.xlsx$")


# --- Arc (half-donut) chart geometry ---------------------------------------
# Ports the math from the approved knesset_arc_v3 SVG design into Python, so
# the arc_chart.html.j2 template only ever drops in already-computed
# numbers/paths -- there is no runtime trig in the template itself.
ARC_CX, ARC_CY = 310, 318
ARC_RO, ARC_RI = 252, 148  # outer / inner radius
ARC_TOTAL_SEATS = 120

ARC_COLOR_GRAY = "#74829e"        # unmapped/other parties
ARC_COLOR_OPP_MEN = "#7ecaff"
ARC_COLOR_OPP_WOMEN = "#00c7f2"
ARC_COLOR_COAL_WOMEN = "#a71d2a"
ARC_COLOR_COAL_MEN = "#ff6384"
ARC_CONTOUR_OPP_WOMEN = "#008ba8"
ARC_CONTOUR_COAL_WOMEN = "#5e0b14"

ARC_BLOC_CHANGE_LABEL = "שינוי+משותפת"  # opposition + unmapped/gray seats
ARC_BLOC_COALITION_LABEL = "ימין+חרדים"  # coalition seats
ARC_BLOC_RECT_WIDTH = 110
ARC_BLOC_CHANGE_LABEL_FONT_SIZE = 16
ARC_BLOC_COALITION_LABEL_FONT_SIZE = 17

# Canvas: the title now lives in an HTML header above the <svg> (matching
# the bar chart's header), not inside the SVG itself, so the viewBox only
# needs to frame the arc + bloc boxes + top callout -- no top title margin.
# Tightened horizontally to hug the actual content (the bloc-total boxes are
# the widest thing, at cx +/- mid_r +/- half the box width) plus a small
# breathing-room margin, rather than the original design's much wider canvas
# -- that wide canvas left a lot of dead white space on both sides once the
# chart is square-padded into the final social-post JPG.
_ARC_MID_R = (ARC_RO + ARC_RI) / 2
ARC_CONTENT_MARGIN_X = 25
ARC_VIEWBOX_MIN_X = ARC_CX - _ARC_MID_R - ARC_BLOC_RECT_WIDTH / 2 - ARC_CONTENT_MARGIN_X
ARC_VIEWBOX_WIDTH = 2 * (_ARC_MID_R + ARC_BLOC_RECT_WIDTH / 2 + ARC_CONTENT_MARGIN_X)
ARC_VIEWBOX_MIN_Y = 0
ARC_VIEWBOX_HEIGHT = 420

ARC_LOGO_WIDTH = 110
ARC_LOGO_HEIGHT = 65
ARC_LOGO_POS = (ARC_CX - ARC_LOGO_WIDTH / 2, ARC_CY + 10)

# Gap (px) between the outlet logo and the title/subtitle text column in the
# header's .header-text-row -- must match that rule's CSS `gap` in
# arc_chart.html.j2 (and bar_chart.html.j2's, same value), since
# render_arc_chart_html() also uses it to widen the render viewport so the
# title doesn't get squeezed into wrapping (see there).
ARC_HEADER_LOGO_GAP = 15

# Geometry for the top arrows + "X חברות כנסת" callout. Each arrow starts
# at its own women-segment's label position (opp_women / coal_women -- same
# angle the "X נשים" number is drawn at, see build_arc_chart_data), so it
# always points from the actual colored segment it's labeling, not a fixed
# angle -- the reference design's own hardcoded angles only happened to line
# up for its particular sample data and land inside the wrong segment for a
# real poll's data (verified: with the current workbook's seat split, the
# reference's fixed "right arrow" angle of 93 falls inside the cyan
# opp_women segment, not the dark-red coal_women segment it's meant to
# start from). Only the arrows' *target* end (pointing at the total-women
# callout) is fixed/data-independent, matching the approved design.
ARC_ARROW_LEFT_TARGET_X_OFFSET = -15
ARC_ARROW_RIGHT_TARGET_X_OFFSET = 5
ARC_ARROW_TARGET_Y_OFFSET = -97
ARC_TOTAL_LABEL_Y_OFFSET = -78
ARC_TOTAL_LABEL_FONT_SIZE = 17


def _arc_rad(d):
    return d * math.pi / 180


def _arc_pt(r, alpha):
    t = _arc_rad(180 + alpha)
    return (round(ARC_CX + r * math.cos(t), 2), round(ARC_CY + r * math.sin(t), 2))


def _arc_clean_seg_path(a1, a2):
    ox1, oy1 = _arc_pt(ARC_RO, a1)
    ox2, oy2 = _arc_pt(ARC_RO, a2)
    ix1, iy1 = _arc_pt(ARC_RI, a1)
    ix2, iy2 = _arc_pt(ARC_RI, a2)
    lg = 1 if (a2 - a1) > 180 else 0
    return (f"M{ox1},{oy1} A{ARC_RO},{ARC_RO} 0 {lg} 1 {ox2},{oy2} "
            f"L{ix2},{iy2} A{ARC_RI},{ARC_RI} 0 {lg} 0 {ix1},{iy1} Z")


def _arc_split_label_lines(name: str) -> list:
    """Best-effort split of a party/label name into up to 2 short lines for
    the small white label inside the gray/unmapped arc segment (mirrors the
    design's 'הרשימה' / 'המשותפת' two-line split)."""
    words = name.split()
    if len(words) <= 1:
        return [name]
    if len(words) == 2:
        return words
    mid = (len(words) + 1) // 2
    return [" ".join(words[:mid]), " ".join(words[mid:])]


def build_arc_chart_data(opp_women: int, opp_men: int, coal_women: int,
                          coal_men: int, gray_men: int, gray_women: int,
                          gray_label: str,
                          total_women_label: str, title_line1: str,
                          title_line2: str, logo_data_uri: str,
                          outlet_name: str = None,
                          poll_date_str: str = None) -> dict:
    """Compute every geometric value the arc_chart.html.j2 template needs:
    segment paths/colors/labels, separator lines, the two bloc totals at the
    base, the gray/unmapped segments' labels, and the top arrows + total-
    women callout. All angles are derived from the seat counts as a share of
    the total seats actually accounted for across the segments -- not a
    fixed 120, since some polls' seat counts sum to less than 120 (parties
    below the electoral threshold that this workbook doesn't reallocate),
    which would otherwise leave the arc visibly short of a full half-circle.
    Pass real per-bloc, per-gender seat totals and this reproduces the
    approved design for any dataset.

    Unmapped/"gray" parties (missing from mapping.csv, or mapped to
    something other than the opposition/coalition group) are gender-split
    the same way the opposition and coalition blocs are: gray_men draws as
    one peripheral segment at the very start of the arc (bottom-left, next
    to the "שינוי+משותפת" total), and gray_women draws as a second, separate
    gray segment inserted between opp_women and coal_women -- right at the
    seam between the two blocs -- instead of every unmapped party's seats
    (men and women together) being folded into one single gray blob."""
    segs = [
        {"key": "gray_men", "n": gray_men, "fill": ARC_COLOR_GRAY, "bloc": "gray",
         "label": gray_label},
        {"key": "opp_men", "n": opp_men, "fill": ARC_COLOR_OPP_MEN, "bloc": "opposition"},
        {"key": "opp_women", "n": opp_women, "fill": ARC_COLOR_OPP_WOMEN, "bloc": "opposition",
         "num": str(opp_women), "sub": "נשים", "contour": ARC_CONTOUR_OPP_WOMEN},
        {"key": "gray_women", "n": gray_women, "fill": ARC_COLOR_GRAY,
         "bloc": "boundary"},
        {"key": "coal_women", "n": coal_women, "fill": ARC_COLOR_COAL_WOMEN, "bloc": "coalition",
         "num": str(coal_women), "sub": "נשים", "contour": ARC_CONTOUR_COAL_WOMEN,
         "label_shift": 2.5},
        {"key": "coal_men", "n": coal_men, "fill": ARC_COLOR_COAL_MEN, "bloc": "coalition"},
    ]
    segs_by_key = {s["key"]: s for s in segs}

    total_seats = sum(s["n"] for s in segs) or ARC_TOTAL_SEATS  # guard div-by-0

    cum = 0.0
    for s in segs:
        s["a1"] = cum
        s["a2"] = cum + (s["n"] / total_seats) * 180
        cum = s["a2"]

    mid_r = (ARC_RO + ARC_RI) / 2

    segments_out = []
    for s in segs:
        if s["n"] <= 0:
            continue
        entry = {"d": _arc_clean_seg_path(s["a1"], s["a2"]), "fill": s["fill"]}
        if s.get("num"):
            lx, ly = _arc_pt(mid_r, (s["a1"] + s["a2"]) / 2 + s.get("label_shift", 0))
            entry["num"] = s["num"]
            entry["sub"] = s["sub"]
            entry["num_pos"] = (lx, ly + 4)
            entry["sub_pos"] = (lx, ly + 28)
            entry["contour"] = s["contour"]
        segments_out.append(entry)

    # Separators: solid between different blocs, dashed between the
    # men/women split within the same bloc.
    separators = []
    present = [s for s in segs if s["n"] > 0]
    for i in range(len(present) - 1):
        a, b = present[i], present[i + 1]
        x1, y1 = _arc_pt(ARC_RO + 2, a["a2"])
        x2, y2 = _arc_pt(ARC_RI - 2, a["a2"])
        separators.append({
            "x1": x1, "y1": y1, "x2": x2, "y2": y2,
            "dashed": a["bloc"] == b["bloc"],
        })

    # Gray/unmapped segments' labels (small 1-2 line white text), each
    # centered in its own arc slice -- every segment with a non-empty
    # "label" gets one (currently just gray_main; gray_boundary is
    # deliberately left unlabeled since it's usually too narrow a sliver
    # for even one word of text to fit legibly).
    gray_labels = []
    for s in segs:
        if s["n"] > 0 and s.get("label"):
            glx, gly = _arc_pt(mid_r, (s["a1"] + s["a2"]) / 2 + s.get("label_shift", 0))
            lines = _arc_split_label_lines(s["label"])
            if len(lines) > 1:
                line_entries = [{"text": t, "y": gly - 4 + i * 13} for i, t in enumerate(lines)]
            else:
                line_entries = [{"text": lines[0], "y": gly + 4}]
            gray_labels.append({"x": glx, "lines": line_entries})

    # The unmapped/"gray" parties (both the main gray segment and the
    # between-the-blocs boundary segment) are shown as a visually distinct
    # color in the arc, but folded into the change/opposition bloc's total
    # below -- this matches the approved design exactly (verified against
    # real data).
    change_total = opp_women + opp_men + gray_men + gray_women
    coalition_total = coal_women + coal_men

    bloc_rects = [
        {"cx": ARC_CX - mid_r, "label": ARC_BLOC_CHANGE_LABEL, "total": change_total,
         "width": ARC_BLOC_RECT_WIDTH, "label_font_size": ARC_BLOC_CHANGE_LABEL_FONT_SIZE},
        {"cx": ARC_CX + mid_r, "label": ARC_BLOC_COALITION_LABEL, "total": coalition_total,
         "width": ARC_BLOC_RECT_WIDTH, "label_font_size": ARC_BLOC_COALITION_LABEL_FONT_SIZE},
    ]

    # Top arrows + total-women callout. Each arrow starts at its own
    # women-segment's actual label angle (opp_women/cyan, coal_women/dark-
    # red -- the same angle their "X נשים" number is drawn at), so it always
    # points from the segment it's labeling, whatever that segment's real
    # position turns out to be for this dataset. Only the target end (the
    # total-women callout above the arc's apex) is fixed/data-independent,
    # matching the approved design.
    opp_women_seg, coal_women_seg = segs_by_key["opp_women"], segs_by_key["coal_women"]
    opp_women_angle = ((opp_women_seg["a1"] + opp_women_seg["a2"]) / 2
                        + opp_women_seg.get("label_shift", 0))
    coal_women_angle = ((coal_women_seg["a1"] + coal_women_seg["a2"]) / 2
                         + coal_women_seg.get("label_shift", 0))
    w1x, w1y = _arc_pt(ARC_RI - 2, opp_women_angle)
    w2x, w2y = _arc_pt(ARC_RI - 2, coal_women_angle)
    target_y = ARC_CY + ARC_ARROW_TARGET_Y_OFFSET
    arrows = [
        {"x1": w1x, "y1": w1y, "x2": ARC_CX + ARC_ARROW_LEFT_TARGET_X_OFFSET, "y2": target_y},
        {"x1": w2x, "y1": w2y, "x2": ARC_CX + ARC_ARROW_RIGHT_TARGET_X_OFFSET, "y2": target_y},
    ]
    total_label_pos = (ARC_CX, ARC_CY + ARC_TOTAL_LABEL_Y_OFFSET)

    return {
        "cx": ARC_CX, "cy": ARC_CY,
        "segments": segments_out,
        "separators": separators,
        "gray_labels": gray_labels,
        "bloc_rects": bloc_rects,
        "arrows": arrows,
        "total_label_text": total_women_label,
        "total_label_pos": total_label_pos,
        "total_label_font_size": ARC_TOTAL_LABEL_FONT_SIZE,
        "title_line1": title_line1,
        "title_line2": title_line2,
        "outlet_name": outlet_name,
        "poll_date_str": poll_date_str,
        "outlet_logo_data_uri": _outlet_logo_data_uri(outlet_name) if outlet_name else None,
        "outlet_logo_width": ARC_LOGO_WIDTH,
        "outlet_logo_height": ARC_LOGO_HEIGHT,
        "logo_data_uri": logo_data_uri,
        "logo_pos": ARC_LOGO_POS,
        "logo_width": ARC_LOGO_WIDTH,
        "logo_height": ARC_LOGO_HEIGHT,
        "view_box": f"{ARC_VIEWBOX_MIN_X} {ARC_VIEWBOX_MIN_Y} {ARC_VIEWBOX_WIDTH} {ARC_VIEWBOX_HEIGHT}",
        "svg_width": ARC_VIEWBOX_WIDTH,
        "svg_height": ARC_VIEWBOX_HEIGHT,
    }


# --- HTML/SVG rendering (Jinja2 + headless Chromium) -----------------------

_JINJA_ENV = Environment(loader=FileSystemLoader(str(TEMPLATES_DIR)))


def _data_uri(path: Path, mime: str) -> str:
    b64 = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{b64}"


def _heebo_data_uris() -> dict:
    return {
        "heebo_regular_data_uri": _data_uri(FONTS_DIR / "Heebo-Regular.ttf", "font/ttf"),
        "heebo_bold_data_uri": _data_uri(FONTS_DIR / "Heebo-Bold.ttf", "font/ttf"),
    }


def _logo_data_uri():
    if LOGO_PATH.exists():
        return _data_uri(LOGO_PATH, "image/jpeg")
    print(f"  warning: logo file not found at {LOGO_PATH}, charts will be "
          f"generated without it")
    return None


OUTLET_LOGOS_DIR = TEMPLATES_DIR / "assets" / "outlet_logos"

# Maps a poll's outlet name (exactly as it appears in COL_OUTLET) to a small
# logo image file inside templates/assets/outlet_logos/, shown inline next
# to the outlet's name in each chart's header subtitle. An outlet with no
# entry here (or whose file is missing) just shows its name as plain text --
# same as every chart looked before this feature existed -- so add entries
# here incrementally as logo files are supplied, no other outlet is affected.
OUTLET_LOGO_FILES = {
    "ערוץ 14": "channel_14.png",
    "חדשות 12": "channel_12.png",
    "חדשות 13": "channel_13.png",
    "כאן חדשות": "kan_11.png",
    "i24 news": "i24_news.png",
    "מעריב": "maariv.png",
}


def _outlet_logo_data_uri(outlet: str):
    filename = OUTLET_LOGO_FILES.get(outlet)
    if not filename:
        return None
    path = OUTLET_LOGOS_DIR / filename
    if not path.exists():
        print(f"  warning: outlet logo file not found for \"{outlet}\": "
              f"{path} -- showing the outlet name without it")
        return None
    mime = "image/png" if path.suffix.lower() == ".png" else "image/jpeg"
    return _data_uri(path, mime)


def _render_html_to_square_jpg(html: str, out_path: Path, css_width: int,
                                css_height: int = None, size: int = 945,
                                anchor: str = "center") -> None:
    """Render an HTML string with headless Chromium and save it as an exact
    `size`x`size` pixel square JPEG. Renders at `css_width`x`css_height`
    (scaled up for crisp text), pads the shorter side with white to make it
    square, then resizes down to the exact target size. `anchor` controls
    where the extra square-up padding goes: "top" keeps content flush
    against the top edge (all padding added below); "center" splits it
    evenly above/below.

    css_height=None auto-fits the viewport to the page's *actual* rendered
    content height instead of using a fixed guess. This matters because a
    fixed height has to be tall enough for the longest possible party list,
    but for any shorter list, Playwright's full_page screenshot still
    captures the whole (mostly empty) viewport -- that leftover blank
    space then balloons into extra padding on every side once the image is
    squared up. Auto-fitting means the screenshot is exactly as tall as the
    real content, so the only padding left is what's actually needed to
    make a portrait card square.
    """
    html_path = out_path.with_suffix(".render.html")
    html_path.write_text(html, encoding="utf-8")
    try:
        scale = size / css_width
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(
                viewport={"width": css_width, "height": css_height or 100},
                device_scale_factor=scale,
            )
            # .resolve().as_uri() (not an f-string) so this works with a
            # relative --output-dir (e.g. "graphs/") and on Windows, where
            # a bare "file://<path>" is not a valid URL.
            page.goto(html_path.resolve().as_uri())
            # Wait for the embedded @font-face (Heebo) to actually finish
            # loading before measuring/screenshotting -- a fixed sleep here
            # was a race: on a slow load, text metrics/wrapping would still
            # briefly reflect the fallback system font (which can be
            # noticeably wider for bold Hebrew), occasionally causing a
            # headline to wrap onto an extra line only sometimes, for the
            # exact same input.
            page.evaluate("document.fonts.ready")
            page.wait_for_timeout(30)  # settle any final layout/paint
            if css_height is None:
                measured = page.evaluate("document.documentElement.scrollHeight")
                page.set_viewport_size({"width": css_width, "height": measured})
            png_bytes = page.screenshot(full_page=True)
            browser.close()
    finally:
        html_path.unlink(missing_ok=True)

    img = Image.open(io.BytesIO(png_bytes)).convert("RGB")
    w, h = img.size
    side = max(w, h)
    canvas = Image.new("RGB", (side, side), "white")
    y_offset = 0 if anchor == "top" else (side - h) // 2
    canvas.paste(img, ((side - w) // 2, y_offset))
    canvas = canvas.resize((size, size), Image.LANCZOS)
    canvas.save(out_path, format="JPEG", quality=92)


BAR_CHART_CSS_WIDTH = 830  # widened (from the reference mockup's 695) so the
                            # chart's own width:height ratio is closer to 1:1
                            # -- less dead white space once _render_html_to_
                            # square_jpg pads it out to a square social post
BAR_NAME_COL_WIDTH = 200   # matches templates/bar_chart.html.j2's .chart-body
                           # grid-template-columns first track (party name)
BAR_AREA_COL_WIDTH = 590   # ...and the second track (the bar itself) -- the
                           # pixel budget each row's bar is scaled to fit
BAR_HEADROOM_MULTIPLIER = 1.3  # empty space reserved past the longest bar,
                                # as a multiple of the rounded-up tick max --
                                # smaller = longer bars (closer to filling
                                # BAR_AREA_COL_WIDTH); 1.0 would let the
                                # longest bar touch the column's edge
BAR_LOGO_SIZE = 110  # matches templates/bar_chart.html.j2's .logo-floating
                      # size and the reference mockup's .logo-floating

# A women-bar this narrow can't fit its own number left-aligned with padding
# at the normal font size -- two tiers, matching the reference mockup: a
# moderately-narrow bar (e.g. the הליכוד row) just centers the number with
# no padding at the normal size; a very-narrow bar (e.g. the רע"מ row) also
# shrinks the font.
BAR_NARROW_CENTER_PX = 40
BAR_NARROW_FONT_PX = 20


def render_bar_chart_html(title_line1: str, title_line2: str, rows: list,
                           total_women: int, out_path: Path,
                           outlet_name: str = None,
                           poll_date_str: str = None) -> None:
    """Renders bar_chart.html.j2: a fixed "סה"כ X נשים" badge under the
    title, a two-column (party name | bar) grid of per-party rows, and the
    5050 logo centered below everything -- matches the reference mockup's
    layout exactly (see templates/bar_chart.html.j2's header comment).

    outlet_name/poll_date_str (per-outlet charts only -- the mean-poll chart
    passes neither) let the template rebuild the "לפי סקר X | תאריך: Y"
    subtitle itself instead of taking title_line2 as opaque text, so it can
    drop the outlet's logo (see OUTLET_LOGO_FILES) inline right next to the
    outlet's name when one is on file; with no entry the subtitle renders
    identically to plain title_line2 text."""
    template = _JINJA_ENV.get_template("bar_chart.html.j2")
    html = template.render(
        title_line1=title_line1, title_line2=title_line2, rows=rows,
        total_women=total_women,
        logo_data_uri=_logo_data_uri(),
        logo_size=BAR_LOGO_SIZE,
        name_col_width=BAR_NAME_COL_WIDTH,
        bar_col_width=BAR_AREA_COL_WIDTH,
        outlet_name=outlet_name,
        poll_date_str=poll_date_str,
        outlet_logo_data_uri=_outlet_logo_data_uri(outlet_name) if outlet_name else None,
        outlet_logo_size=BAR_LOGO_SIZE,
        **_heebo_data_uris(),
    )
    _render_html_to_square_jpg(html, out_path, css_width=BAR_CHART_CSS_WIDTH,
                                css_height=None, anchor="center")


def render_arc_chart_html(arc_data: dict, out_path: Path) -> None:
    template = _JINJA_ENV.get_template("arc_chart.html.j2")
    html = template.render(**arc_data, **_heebo_data_uris())
    # The header (title/subtitle) now sits above the <svg> as ordinary HTML,
    # not inside the SVG's own viewBox, so the page's total rendered height
    # is taller than svg_height alone -- auto-fit to actual content height
    # (css_height=None) instead of assuming it equals the SVG's height.
    #
    # The page renders at a fixed viewport width (css_width below); normally
    # that's just the SVG's own width (svg_width), which is exactly wide
    # enough for the title text alone. When there's an outlet logo, the
    # header is now a row of [logo, text] (see .header-text-row), so the
    # text column loses (outlet_logo_width + its gap) of width to the logo
    # -- without extra room here, that's enough to force the title onto a
    # second line. Widen the render viewport by that same amount so the
    # text column keeps the same width it always had; the (now narrower,
    # relatively) SVG still ends up centered under it via chart-wrapper's
    # align-items: center.
    css_width = arc_data["svg_width"]
    if arc_data.get("outlet_logo_data_uri"):
        css_width += arc_data["outlet_logo_width"] + ARC_HEADER_LOGO_GAP
    _render_html_to_square_jpg(html, out_path, css_width=css_width,
                                css_height=None, anchor="center")


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


def format_date_short(raw) -> str:
    """Format a date as DD.MM.YY (zero-padded day/month, 2-digit year),
    Israeli style -- used in the bar chart's header subtitle."""
    ts = pd.to_datetime(raw)
    return ts.strftime("%d.%m.%y")


def sanitize_filename(text: str) -> str:
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    text = text.strip().replace(" ", "_")
    return text


# NOTE: unlike the old matplotlib charts, none of this text is pre-reordered
# with a bidi helper -- the HTML templates set dir="rtl"/direction:rtl and
# the browser applies the Unicode bidi algorithm itself. Pre-reordering here
# would double-flip and garble the text.

def build_bar_headline() -> str:
    """Top, larger-font headline for the bar chart -- fixed text, the same
    on every bar chart (per-outlet and mean-poll alike)."""
    return "כמה נשים תהיינה בכנסת הבאה?"


def build_bar_subheadline(outlet: str, date_raw) -> str:
    """Second (smaller) line of the bar chart's title block, naming the
    poll it's based on and the poll's own date."""
    return f"לפי סקר {outlet} | תאריך: {format_date_short(date_raw)}"


def build_bar_subheadline_mean() -> str:
    """Second (smaller) line of the bar chart's title block for the
    mean-poll (average-of-polls) chart -- there's no single poll date to
    show (it aggregates every outlet's latest poll), so this uses today's
    date instead, i.e. when this snapshot was generated."""
    return f"לפי ממוצע הסקרים | תאריך: {format_date_short(datetime.now())}"


def load_mapping(mapping_csv: Path) -> dict:
    """Read the party -> bloc mapping CSV (columns: מפלגה, גוש) into a dict."""
    map_df = pd.read_csv(mapping_csv, encoding="utf-8-sig")
    return dict(zip(map_df[MAPPING_COL_PARTY], map_df[MAPPING_COL_GROUP]))


def render_bar_chart(parties: list, women: list, men: list, title_line1: str,
                      title_line2: str, out_path: Path,
                      outlet_name: str = None, poll_date_str: str = None) -> None:
    """Shared bar-chart renderer (bar_chart.html.j2), used for both the
    per-outlet poll charts and the mean-poll (average-of-polls) chart.
    Each row's bar is drawn at a fixed pixels-per-mandate scale (matching
    the reference mockup's literal "each mandate = 15px" bars) rather than
    stretched to fill a shared row width -- a party's bar-container is only
    as wide as its own seat total. The pixels-per-mandate value itself is
    NOT hardcoded to the reference's 15px, though: it's derived from
    BAR_AREA_COL_WIDTH so the single longest bar in THIS chart's data always
    fits inside the template's fixed bar column, however many seats that
    turns out to be for a given poll."""
    totals = [w + m for w, m in zip(women, men)]

    # Headroom past the longest bar (BAR_HEADROOM_MULTIPLIER), so the
    # tallest bar doesn't touch the far edge of the fixed bar column.
    max_total = max(totals) if totals else 0
    tick_max = max(5, (max_total // 5 + 1) * 5)
    x_max = tick_max * BAR_HEADROOM_MULTIPLIER
    px_per_seat = BAR_AREA_COL_WIDTH / x_max if x_max else 0

    rows = []
    for p, w, m, t in zip(parties, women, men, totals):
        total_px = round(t * px_per_seat)
        women_px = round(w * px_per_seat)
        men_px = total_px - women_px  # avoid a rounding-gap between the two segments
        rows.append({
            "party": p,
            "women": w,
            "men": m,
            "total": t,
            "total_px": total_px,
            "women_px": women_px,
            "men_px": men_px,
            "narrow": women_px < BAR_NARROW_CENTER_PX,
            "narrow_font": women_px < BAR_NARROW_FONT_PX,
        })

    total_women = sum(women)
    render_bar_chart_html(title_line1, title_line2, rows, total_women, out_path,
                          outlet_name=outlet_name, poll_date_str=poll_date_str)


def plot_poll(df_poll: pd.DataFrame, outlet: str, date_raw, out_path: Path) -> None:
    # Ordered by expected number of women descending (most women at the
    # top); ties broken by expected number of men descending; remaining
    # ties keep the workbook's original row order.
    df_sorted = df_poll.sort_values(
        [COL_WOMEN, COL_MEN], ascending=[False, False], kind="stable"
    )

    parties = [_chart_display_name(p) for p in df_sorted[COL_PARTY].tolist()]
    women = df_sorted[COL_WOMEN].tolist()
    men = df_sorted[COL_MEN].tolist()

    title_line1 = build_bar_headline()
    title_line2 = build_bar_subheadline(outlet, date_raw)

    render_bar_chart(parties, women, men, title_line1, title_line2, out_path,
                      outlet_name=outlet, poll_date_str=format_date_short(date_raw))


def plot_mean_poll(df_mean: pd.DataFrame, out_path: Path) -> None:
    """Chart for the aggregate 'mean poll' (average-of-polls) estimate in
    the חישוב 2026 sheet -- same visual style as a per-outlet poll chart,
    but with no outlet/date in the title and the trailing totals row
    (סה"כ) excluded before sorting/plotting."""
    df = df_mean[df_mean[COL_PARTY] != TOTAL_ROW_LABEL]
    df_sorted = df.sort_values(
        [COL_WOMEN, COL_MEN], ascending=[False, False], kind="stable"
    )

    parties = [_chart_display_name(p) for p in df_sorted[COL_PARTY].tolist()]
    women = df_sorted[COL_WOMEN].tolist()
    men = df_sorted[COL_MEN].tolist()

    title_line1 = build_bar_headline()
    title_line2 = build_bar_subheadline_mean()

    render_bar_chart(parties, women, men, title_line1, title_line2, out_path)


def render_bloc_chart(df: pd.DataFrame, mapping: dict, title_line1: str,
                       title_line2: str, out_path: Path, skip_label: str,
                       outlet_name: str = None, poll_date_str: str = None) -> bool:
    """Shared arc (half-donut) bloc-chart renderer (arc_chart.html.j2), used
    for both per-outlet polls and the mean-poll aggregate. Splits seats into
    opposition/coalition men+women; any party whose mapped group is not
    exactly OPPOSITION_GROUP_NAME/COALITION_GROUP_NAME (missing from
    mapping.csv, or mapped to some other/self-referential group) is gender-
    split into two gray "unmapped" segments the same way the opposition and
    coalition blocs are (see build_arc_chart_data), and its seats are folded
    into the change bloc's ("גוש השינוי") total either way. Returns False
    (writing nothing) if there are 0 expected women overall."""
    total_women = int(df[COL_WOMEN].sum())

    if total_women == 0:
        print(f"  skipping pie chart for {skip_label}: 0 expected women")
        return False

    opp_women = opp_men = coal_women = coal_men = 0
    gray_men = gray_women = 0
    gray_parties = []
    for _, row in df.iterrows():
        party = row[COL_PARTY]
        w, m = int(row[COL_WOMEN]), int(row[COL_MEN])
        group = mapping.get(party)
        if group == OPPOSITION_GROUP_NAME:
            opp_women += w
            opp_men += m
        elif group == COALITION_GROUP_NAME:
            coal_women += w
            coal_men += m
        else:
            gray_men += m
            gray_women += w
            if w + m > 0:
                gray_parties.append(_chart_display_name(party))

    if gray_parties:
        print(f"  note: unmapped, gender-split into two gray segments (men "
              f"at the start of the arc, women between the two blocs), "
              f"folded into \"{ARC_BLOC_CHANGE_LABEL}\": {', '.join(gray_parties)}")

    gray_label = " / ".join(gray_parties) if gray_parties else ""
    total_label_text = f"{total_women} חברות כנסת"

    arc_data = build_arc_chart_data(
        opp_women=opp_women, opp_men=opp_men,
        coal_women=coal_women, coal_men=coal_men,
        gray_men=gray_men, gray_women=gray_women, gray_label=gray_label,
        total_women_label=total_label_text,
        title_line1=title_line1, title_line2=title_line2,
        logo_data_uri=_logo_data_uri(),
        outlet_name=outlet_name, poll_date_str=poll_date_str,
    )
    render_arc_chart_html(arc_data, out_path)
    return True


def plot_pie_poll(df_poll: pd.DataFrame, outlet: str, date_raw, mapping: dict,
                   out_path: Path) -> bool:
    """Arc/bloc chart for a single outlet's poll. Returns False (writing
    nothing) if there are 0 expected women overall. Title block is the same
    headline + poll/date subhead as the matching bar chart (build_bar_headline
    / build_bar_subheadline), so both charts read as one consistent pair."""
    date_str = format_poll_date(date_raw)
    title_line1 = build_bar_headline()
    title_line2 = build_bar_subheadline(outlet, date_raw)
    return render_bloc_chart(df_poll, mapping, title_line1, title_line2, out_path,
                              skip_label=f"{outlet} ({date_str})",
                              outlet_name=outlet, poll_date_str=format_date_short(date_raw))


def plot_mean_pie_poll(df_mean: pd.DataFrame, mapping: dict, out_path: Path) -> bool:
    """Arc/bloc chart for the aggregate 'mean poll' (average-of-polls)
    estimate in the חישוב 2026 sheet, with the trailing totals row (סה"כ)
    excluded first. Returns False (writing nothing) if there are 0 expected
    women overall. Same title block as plot_mean_poll's bar chart."""
    df = df_mean[df_mean[COL_PARTY] != TOTAL_ROW_LABEL]
    title_line1 = build_bar_headline()
    title_line2 = build_bar_subheadline_mean()
    return render_bloc_chart(df, mapping, title_line1, title_line2, out_path,
                              skip_label="ממוצע הסקרים")


def compute_poll_email_stats(df_poll: pd.DataFrame, mapping: dict) -> dict:
    """Leading party (by expected women) + per-bloc/total women totals for a
    single poll, used to fill in the email template. Note: unlike rtl(),
    these party/bloc names are left in normal (logical) Hebrew order --
    email clients do their own bidi rendering, so reordering here would
    show up backwards.

    "change_bloc_women" folds in unmapped/"gray" parties' women (parties
    not mapped to exactly OPPOSITION_GROUP_NAME or COALITION_GROUP_NAME),
    matching the arc chart's "גוש השינוי" total (see render_bloc_chart)."""
    women_by_party = df_poll.set_index(COL_PARTY)[COL_WOMEN]
    leading_party = women_by_party.idxmax()
    leading_women = int(women_by_party.max())

    groups = df_poll[COL_PARTY].map(mapping)
    opposition_women = int(df_poll.loc[groups == OPPOSITION_GROUP_NAME, COL_WOMEN].sum())
    coalition_women = int(df_poll.loc[groups == COALITION_GROUP_NAME, COL_WOMEN].sum())
    gray_women = int(df_poll.loc[
        ~groups.isin([OPPOSITION_GROUP_NAME, COALITION_GROUP_NAME]), COL_WOMEN
    ].sum())
    total_women = int(df_poll[COL_WOMEN].sum())

    return {
        "leading_party": leading_party,
        "leading_women": leading_women,
        "opposition_women": opposition_women,
        "coalition_women": coalition_women,
        "gray_women": gray_women,
        "change_bloc_women": opposition_women + gray_women,
        "coalition_bloc_women": coalition_women,
        "total_women": total_women,
    }


def _read_text_asset(path: Path) -> str:
    """Read a user-edited text asset, tolerating whichever encoding it was
    saved with -- Windows Notepad defaults to UTF-16 (with a BOM), but a
    plain UTF-8 or UTF-8-with-BOM save should also just work."""
    raw = path.read_bytes()
    for encoding in ("utf-16", "utf-8-sig", "utf-8"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise SystemExit(f"Could not decode {path} as UTF-16 or UTF-8.")


EMAIL_SUBJECT_PREFIX = "נושא:"


def _load_email_template() -> tuple:
    """Split templates/assets/text_for_email.txt into (subject_template,
    body_template): the first line is the subject (prefixed with
    "נושא: "), and everything after the following blank line is the body.
    Placeholders like [שם הערוץ/עיתון] and [מספר] are filled in by
    build_email_subject/build_email_body -- edit the wording in the .txt
    file freely, just keep the bracketed placeholder names intact."""
    if not EMAIL_TEMPLATE_PATH.exists():
        raise SystemExit(
            f"--create-email-drafts needs {EMAIL_TEMPLATE_PATH}, "
            f"which was not found."
        )
    text = _read_text_asset(EMAIL_TEMPLATE_PATH).replace("\r\n", "\n").replace("\r", "\n")
    subject_line, _, rest = text.partition("\n\n")
    subject_template = subject_line.strip()
    if subject_template.startswith(EMAIL_SUBJECT_PREFIX):
        subject_template = subject_template[len(EMAIL_SUBJECT_PREFIX):].strip()
    body_template = rest.strip("\n")
    return subject_template, body_template


def _fill_sequential(text: str, placeholder: str, values: list) -> str:
    """Replace successive occurrences of `placeholder` with each value from
    `values`, in order -- used for the email template's repeated [מספר]
    placeholder, which stands for a different number each time it appears.
    A mismatched count (template edited to add/remove a placeholder) is
    reported rather than silently filling in the wrong number; unmatched
    placeholders beyond len(values) are left as literal text."""
    count = text.count(placeholder)
    if count != len(values):
        print(f'  warning: email template has {count} "{placeholder}" '
              f'placeholder(s) but {len(values)} value(s) were expected; '
              f'filling in order, any extra placeholders are left as-is')
    values_iter = iter(values)

    def _replace(_match):
        try:
            return str(next(values_iter))
        except StopIteration:
            return _match.group(0)

    return re.sub(re.escape(placeholder), _replace, text)


def build_email_subject(outlet: str, date_str: str) -> str:
    subject_template, _ = _load_email_template()
    subject = subject_template.replace("[שם הערוץ/עיתון]", outlet)
    subject = subject.replace("[תאריך]", date_str)
    return subject


def build_email_body(outlet: str, date_str: str, stats: dict) -> str:
    _, body_template = _load_email_template()
    body = body_template.replace("[ערוץ תקשורת]", outlet)
    body = body.replace("[שם הערוץ]", outlet)
    body = body.replace("[תאריך]", date_str)
    body = body.replace("[שם המפלגה]", stats["leading_party"])
    body = _fill_sequential(body, "[מספר]", [
        stats["leading_women"],
        stats["change_bloc_women"],
        stats["coalition_bloc_women"],
        stats["total_women"],
    ])
    return body


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
            body = build_email_body(outlet, date_str, stats)
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
