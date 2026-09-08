"""Exports the tunable design parameters used by plot_women_by_outlet.py and
templates/*.j2 (colors, fonts, sizes, spacing, positions, fixed label text)
to a CSV -- not the poll data itself. Companion to apply_chart_parameters.py,
which reads a CSV of this same shape back and patches the code/templates.

Usage:
    python3 extract_chart_parameters.py [OUTPUT_CSV]

    OUTPUT_CSV   Where to write the CSV (default: chart_template_parameters.csv
                 next to this script).

This is a curated list, not something parsed live out of the files -- if you
add or restructure a design knob in the templates or script, add a matching
row here (and a matching rule in apply_chart_parameters.py's REGISTRY) so the
two stay in sync.
"""
import csv
import sys
from pathlib import Path

BAR = "templates/bar_chart.html.j2"
ARC = "templates/arc_chart.html.j2"
PY = "plot_women_by_outlet.py"

ROWS = [
    # ---------- bar_chart.html.j2 ----------
    (BAR, "@font-face Heebo", "font-weight (regular)", "400",
     "Embedded Heebo regular weight, loaded from fonts/Heebo-Regular.ttf as a data: URI"),
    (BAR, "@font-face Heebo", "font-weight (bold)", "700",
     "Embedded Heebo bold weight, loaded from fonts/Heebo-Bold.ttf as a data: URI"),
    (BAR, "body", "font-family",
     "'Heebo', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif",
     "Primary chart font"),
    (BAR, "body", "background-color", "#ffffff", "Card background"),
    (BAR, "body", "color", "#000000", "Default text color"),
    (BAR, "body", "padding", "12px 14px 14px", "Outer card padding (top / sides / bottom)"),
    (BAR, ".chart-wrapper", "width", "400px", "Fixed card width (BAR_CHART_CSS_WIDTH in the script)"),
    (BAR, ".header-area", "margin-bottom", "10px", "Gap between title block and first bar row"),
    (BAR, ".header-title", "font-size", "21px", "Headline text size (\"כמה נשים תהיינה בכנסת הבאה?\")"),
    (BAR, ".header-title", "font-weight", "900", "Headline weight"),
    (BAR, ".header-title", "letter-spacing", "-0.4px", "Headline letter spacing"),
    (BAR, ".header-title", "line-height", "1.15", "Headline line height"),
    (BAR, ".header-subtitle", "font-size", "12px", "Subtitle text size (poll/date line)"),
    (BAR, ".header-subtitle", "color", "#333", "Subtitle text color"),
    (BAR, ".header-subtitle", "font-weight", "500", "Subtitle weight"),
    (BAR, ".header-subtitle", "line-height", "1.25", "Subtitle line height"),
    (BAR, ".chart-body", "gap", "5.5px", "Vertical gap between party rows"),
    (BAR, ".logo-overlay-box", "bottom", "85px", "Logo+label block's fixed offset from the bottom of .chart-body"),
    (BAR, ".logo-overlay-box", "right", "12px", "Logo+label block's offset from the right edge"),
    (BAR, ".logo-overlay-text", "font-size", "21px", "\"סה\"כ X נשים\" callout text size"),
    (BAR, ".logo-overlay-text", "font-weight", "900", "Callout text weight"),
    (BAR, ".logo-overlay-text", "color", "#702283", "Callout text color (brand purple)"),
    (BAR, ".logo-overlay-text", "letter-spacing", "-0.4px", "Callout letter spacing"),
    (BAR, ".logo-overlay-text", "margin-bottom", "4px", "Gap between callout text and logo image below it"),
    (BAR, ".logo-overlay-text", "background", "rgba(255, 255, 255, 0.85)",
     "Translucent pill background behind the callout text"),
    (BAR, ".logo-overlay-text", "padding", "2px 6px", "Callout pill padding"),
    (BAR, ".logo-overlay-text", "border-radius", "4px", "Callout pill corner radius"),
    (BAR, ".logo-footer-text / .logo-overlay-text", "font styling", "same as above",
     "Footer fallback (no zero-seat rows) uses identical text styling, no background pill"),
    (BAR, ".logo-footer img", "width / height", "70px", "Footer-fallback logo size"),
    (BAR, ".party-name", "font-size", "11.5px", "Party name label size"),
    (BAR, ".party-name", "font-weight", "700", "Party name weight"),
    (BAR, ".party-name", "color", "#111", "Party name color"),
    (BAR, ".party-name", "margin-bottom", "1.5px", "Gap between party name and its bar row"),
    (BAR, ".party-name", "line-height", "1.1", "Party name line height"),
    (BAR, ".bar-row", "height", "17px", "Bar height"),
    (BAR, ".bar-women", "background-color", "#702283", "Women-seats bar color (brand purple)"),
    (BAR, ".bar-women", "color", "#ffffff", "Women count label color (inside the bar)"),
    (BAR, ".bar-women", "font-size", "11px", "Women count label size"),
    (BAR, ".bar-women", "border-radius", "2.5px 0 0 2.5px", "Rounded left end of the women segment"),
    (BAR, ".bar-men", "background-color", "#bdbdbd", "Men-seats bar color (gray)"),
    (BAR, ".bar-men", "border-radius", "0 2.5px 2.5px 0", "Rounded right end of the men segment"),
    (BAR, ".bar-men.full", "border-radius", "2.5px",
     "All four corners rounded when a party has 0 women (bar is men-only)"),
    (BAR, ".total-label", "font-size", "12.5px", "Seat-total number printed after each bar"),
    (BAR, ".total-label", "font-weight", "900", "Seat-total weight"),
    (BAR, ".total-label", "margin-left", "6px", "Gap between bar end and the total number"),
    (BAR, ".zero-label", "font-size", "12.5px", "\"0\" label for parties with 0 seats"),
    (BAR, ".zero-label", "font-weight", "900", "Zero-label weight"),
    (BAR, ".legend", "gap", "16px", "Gap between the two legend items"),
    (BAR, ".legend", "margin-top", "12px", "Gap between last party row and the legend"),
    (BAR, ".legend", "font-size", "11.5px", "Legend text size"),
    (BAR, ".legend", "font-weight", "600", "Legend text weight"),
    (BAR, ".legend-color", "width x height", "16px x 10px", "Legend color swatch size"),
    (BAR, ".legend-color.women", "background-color", "#702283", "Women swatch color"),
    (BAR, ".legend-color.men", "background-color", "#bdbdbd", "Men swatch color"),
    (BAR, "fixed text", "legend labels", "כמות נשים צפויה / כמות גברים צפויה", "Legend caption text"),
    (BAR, "fixed text", "logo callout prefix", "סה\"כ {total_women} נשים", "Total-women callout text template"),

    # ---------- arc_chart.html.j2 ----------
    (ARC, "title_line1 <text>", "font-size", "21", "Arc headline size (same wording as the bar chart's headline)"),
    (ARC, "title_line1 <text>", "fill", "#000000", "Headline color"),
    (ARC, "title_line1 <text>", "font-weight", "900", "Headline weight"),
    (ARC, "title_line2 <text>", "font-size", "12", "Arc subtitle size (poll/date line)"),
    (ARC, "title_line2 <text>", "fill", "#333333", "Subtitle color"),
    (ARC, "title_line2 <text>", "font-weight", "500", "Subtitle weight"),
    (ARC, "segment number <text> (seg.num)", "font-size", "30", "Big number inside the women segments (e.g. \"28\")"),
    (ARC, "segment number <text>", "fill", "white", "Number color"),
    (ARC, "segment number <text>", "stroke-width", "5", "Contour outline width around the number"),
    (ARC, "segment sub-label <text> (seg.sub)", "font-size", "24", "\"נשים\" label under the number"),
    (ARC, "segment sub-label <text>", "stroke-width", "4", "Contour outline width around the sub-label"),
    (ARC, "bloc rect", "width x height", "104 x 76", "White box behind each bloc's label/total at the base of the arc"),
    (ARC, "bloc rect", "fill", "#fff", "Box background color"),
    (ARC, "bloc rect label <text>", "font-size", "24", "Bloc name text size (e.g. \"גוש השינוי + המשותפת\")"),
    (ARC, "bloc rect total <text>", "font-size", "30", "Bloc seat-total text size"),
    (ARC, "gray/unmapped label <text>", "font-size", "11", "Label inside the gray (unmapped-party) arc segment"),
    (ARC, "gray/unmapped label <text>", "fill", "#fff", "Label color"),
    (ARC, "separators (solid, between blocs)", "stroke-width", "3", "Solid white divider width"),
    (ARC, "separators (dashed, within a bloc)", "stroke-width / stroke-dasharray", "1.5 / \"5,4\"",
     "Dashed divider between a bloc's men/women segments"),
    (ARC, "arrows", "stroke / stroke-width", "#000000 / 2",
     "Arrows pointing to the total-women callout, with an arrowhead marker (6x6)"),
    (ARC, "total_label_text <text> (e.g. \"30 חברות כנסת\")", "font-size", "17", "Top callout text size"),
    (ARC, "total_label_text <text>", "font-weight", "bold", "Top callout weight"),
    (ARC, "colors (fill)", "gray / unmapped", "#8890a8", "ARC_COLOR_GRAY"),
    (ARC, "colors (fill)", "opposition men", "#0A85ED", "ARC_COLOR_OPP_MEN"),
    (ARC, "colors (fill)", "opposition women", "#08C8F9", "ARC_COLOR_OPP_WOMEN"),
    (ARC, "colors (fill)", "coalition women", "#0061BF", "ARC_COLOR_COAL_WOMEN"),
    (ARC, "colors (fill)", "coalition men", "#003F88", "ARC_COLOR_COAL_MEN"),
    (ARC, "colors (contour)", "opposition-women number outline", "#06A0C7", "ARC_CONTOUR_OPP_WOMEN"),
    (ARC, "colors (contour)", "coalition-women number outline", "#002244", "ARC_CONTOUR_COAL_WOMEN"),

    # ---------- plot_women_by_outlet.py -- arc geometry/labels ----------
    (PY, "ARC_CX, ARC_CY", "arc center point", "310, 318", "SVG coordinates of the arc's center"),
    (PY, "ARC_RO, ARC_RI", "outer / inner radius", "252, 148", "Arc thickness (ring between these two radii)"),
    (PY, "ARC_TOTAL_SEATS", "fallback seat total", "120",
     "Used only if a poll's segments sum to 0 (div-by-zero guard)"),
    (PY, "ARC_BLOC_CHANGE_LABEL", "left bloc label text", "גוש השינוי + המשותפת",
     "Opposition + unmapped/gray seats"),
    (PY, "ARC_BLOC_COALITION_LABEL", "right bloc label text", "גוש ימין-חרדים", "Coalition seats"),
    (PY, "ARC_VIEWBOX_MIN_X, ARC_VIEWBOX_MIN_Y", "viewBox origin", "-60, -55", "Top-left of the visible SVG crop"),
    (PY, "ARC_VIEWBOX_WIDTH", "viewBox width", "720", ""),
    (PY, "ARC_TOP_MARGIN", "extra top margin", "55", "Headroom reserved above the arc for the two title lines"),
    (PY, "ARC_BOTTOM_MARGIN", "extra bottom margin", "45", "Headroom reserved below the arc for the logo"),
    (PY, "ARC_VIEWBOX_HEIGHT", "viewBox height (computed)", "385 + ARC_TOP_MARGIN + ARC_BOTTOM_MARGIN = 485",
     "Derived -- edit ARC_TOP_MARGIN/ARC_BOTTOM_MARGIN instead"),
    (PY, "ARC_TITLE_LINE1_POS", "headline position", "(ARC_CX, 14)", ""),
    (PY, "ARC_TITLE_LINE2_POS", "subtitle position", "(ARC_CX, 36)", ""),
    (PY, "ARC_LOGO_SIZE", "logo size", "90", "5050 logo width/height in the arc chart"),
    (PY, "ARC_LOGO_POS", "logo position (computed)", "(ARC_CX - ARC_LOGO_SIZE/2, ARC_CY)",
     "Derived -- edit ARC_CX/ARC_CY/ARC_LOGO_SIZE instead"),
    (PY, "build_arc_chart_data(): angle_offset", "arrow spread angle", "16.5",
     "Degrees each top arrow is offset from the arc's apex (90°)"),
    (PY, "build_arc_chart_data(): target_offset", "arrow target x-offset", "20",
     "How far apart the two arrows' tips land, left/right of center"),
    (PY, "build_arc_chart_data(): target_y", "arrow target y", "ARC_CY - 97",
     "Vertical position the arrows point to (the total-women callout)"),

    # ---------- plot_women_by_outlet.py -- bar chart sizing ----------
    (PY, "BAR_CHART_CSS_WIDTH", "card width", "400", "Must match templates/bar_chart.html.j2's .chart-wrapper width"),
    (PY, "BAR_LOGO_SIZE", "logo size", "95",
     "5050 logo width/height in the bar chart overlay (matches .logo-overlay-img)"),
    (PY, "render_bar_chart(): bar headroom multiplier", "x_max = tick_max * 1.8", "1.8",
     "Extra headroom past the longest bar, so bars don't stretch edge-to-edge"),
    (PY, "render_bar_chart(): tick rounding step", "tick_max = max(5, (max_total // 5 + 1) * 5)", "5",
     "Rounds the longest bar's total up to the nearest multiple of 5"),

    # ---------- plot_women_by_outlet.py -- fixed chart text ----------
    (PY, "build_bar_headline()", "headline text", "כמה נשים תהיינה בכנסת הבאה?",
     "Same on every bar chart and arc chart"),
    (PY, "build_bar_subheadline(outlet, date)", "subtitle template", "לפי סקר {outlet} | תאריך: {DD.MM.YY}",
     "Per-outlet poll charts -- mixes fixed text with live expressions, edit by hand"),
    (PY, "build_bar_subheadline_mean()", "subtitle template (mean poll)", "לפי ממוצע הסקרים | תאריך: {DD.MM.YY}",
     "Mean-poll chart; date is today's date -- mixes fixed text with a live expression, edit by hand"),
    (PY, "format_date_short()", "date format", "DD.MM.YY",
     "Zero-padded day/month, 2-digit year -- a strftime pattern, edit by hand"),
    (PY, "format_poll_date()", "date format (filenames)", "D.M",
     "No leading zeros; used in output filenames, not the chart text -- edit by hand"),

    # ---------- plot_women_by_outlet.py -- workbook/schema references (not chart data itself) ----------
    (PY, "SHEET_NAME", "per-outlet sheet name", "חישוב לפי ערוץ", "Workbook sheet the per-outlet charts read from"),
    (PY, "MEAN_SHEET_NAME", "mean-poll sheet name", "חישוב 2026", "Workbook sheet the mean-poll chart reads from"),
    (PY, "COL_OUTLET / COL_DATE / COL_PARTY / COL_SEATS / COL_WOMEN / COL_MEN", "column headers read",
     "כלי תקשורת / תאריך הסקר / מפלגה / מנדטים / כמות נשים צפויה / כמות גברים צפויה",
     "Column names expected in SHEET_NAME"),
    (PY, "TOTAL_ROW_LABEL", "totals-row marker", "סה\"כ", "Row excluded from the mean-poll chart before plotting"),
    (PY, "MAPPING_COL_PARTY / MAPPING_COL_GROUP", "mapping.csv column headers", "מפלגה / גוש", ""),
    (PY, "OPPOSITION_GROUP_NAME / COALITION_GROUP_NAME", "bloc names in mapping.csv", "אופוזיציה / קואליציה",
     "Any other/missing value is drawn as the gray \"unmapped\" segment"),
]


def main():
    out_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent / "chart_template_parameters.csv"
    with out_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["source_file", "section", "parameter", "value", "description"])
        writer.writerows(ROWS)
    print(f"Wrote {len(ROWS)} rows to {out_path}")


if __name__ == "__main__":
    main()
