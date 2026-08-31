"""Live-preview helper for the two chart templates (templates/bar_chart.html.j2
and templates/arc_chart.html.j2) -- for tweaking CSS/layout without waiting on
a full run.

Unlike plot_women_by_outlet.py, this does NOT read any workbook and does NOT
render any JPGs. It fills the templates with realistic sample data and writes
plain, already-rendered .html files that open directly in a browser.

Usage:
    1. Edit templates/bar_chart.html.j2 or templates/arc_chart.html.j2
       (colors, fonts, spacing, wording, ...).
    2. Run:  python preview_charts.py
    3. It writes preview/bar_chart_preview.html and
       preview/arc_chart_preview.html and opens both in your default
       browser. After further edits, just re-run this script and refresh
       the browser tab (or run it once and refresh after each edit -- the
       files are rewritten in place, only the content changes).

The bar chart's logo overlay is positioned using the same measurement step
the real script uses, so its placement here matches a real run. Everything
else (colors, fonts, spacing, arc geometry, text) is exactly what a real
run would produce for this sample data -- this is the *same* rendering
code, just pointed at made-up numbers instead of a workbook, and stopped
short of the final screenshot-to-JPG step.
"""
import importlib.util
import webbrowser
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
MAIN_SCRIPT = SCRIPT_DIR / "plot_women_by_outlet.py"
PREVIEW_DIR = SCRIPT_DIR / "preview"

# Import plot_women_by_outlet.py as a module so this stays in lockstep with
# the real renderer -- no separate copy of the geometry/rendering logic to
# keep in sync.
_spec = importlib.util.spec_from_file_location("plot_women_by_outlet", MAIN_SCRIPT)
m = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(m)

# --- Sample data (no workbook needed) ---------------------------------------
# A realistic 12-party spread, including a few zero-seat parties at the tail
# (like a real poll) so the bar chart's logo-placement logic has something
# to react to.
SAMPLE_PARTIES = [
    # (party name,            women, men)
    ("ישר!",               12, 12),
    ("ביחד (בנט-לפיד)",      7,  7),
    ("הדמוקרטים",            5,  5),
    ("ישראל ביתנו",          4,  5),
    ("הליכוד",               2, 21),
    ("עוצמה יהודית",         1,  8),
    ("הציונות הדתית",        1,  4),
    ('רע"מ',                 1,  4),
    ("יהדות התורה",          0,  8),
    ('ש"ס',                  0,  7),
    ("הרשימה המשותפת",       0,  6),
    ("כחול לבן",             0,  0),
]

SAMPLE_ARC = dict(
    opp_women=28, opp_men=33, coal_women=5, coal_men=47,
    gray_seats=7, gray_label="הרשימה המשותפת",
)


def _preview_bar_chart_html() -> str:
    parties = [p for p, _w, _m in SAMPLE_PARTIES]
    women = [w for _p, w, _m in SAMPLE_PARTIES]
    men = [mm for _p, _w, mm in SAMPLE_PARTIES]
    totals = [w + mm for w, mm in zip(women, men)]

    max_total = max(totals)
    tick_max = max(5, (max_total // 5 + 1) * 5)
    x_max = tick_max * 1.8

    rows = [{
        "party": p, "women": w, "men": mm, "total": t,
        "women_pct": round(w / x_max * 100, 2),
        "men_pct": round(mm / x_max * 100, 2),
    } for p, w, mm, t in zip(parties, women, men, totals)]

    title_line1 = m.build_bar_headline()
    title_line2 = m.build_bar_subheadline_mean()
    total_women = sum(women)
    has_zero_row = any(r["total"] == 0 for r in rows)
    logo_uri = m._logo_data_uri()

    template = m._JINJA_ENV.get_template("bar_chart.html.j2")
    # Same fixed-offset placement render_bar_chart_html uses -- see its
    # docstring -- so this preview matches a real run.
    logo_overlay_uri = logo_uri if (logo_uri and has_zero_row) else None
    logo_footer_uri = logo_uri if (logo_uri and not has_zero_row) else None

    return template.render(
        title_line1=title_line1, title_line2=title_line2, rows=rows,
        has_zero_row=has_zero_row, total_women=total_women,
        logo_data_uri=logo_overlay_uri,
        logo_footer_uri=logo_footer_uri,
        logo_size=m.BAR_LOGO_SIZE,
        **m._heebo_data_uris(),
    )


def _preview_arc_chart_html() -> str:
    total_women = SAMPLE_ARC["opp_women"] + SAMPLE_ARC["coal_women"]
    title_line1 = m.build_bar_headline()
    title_line2 = m.build_bar_subheadline_mean()
    arc_data = m.build_arc_chart_data(
        **SAMPLE_ARC,
        total_women_label=f"{total_women} חברות כנסת",
        title_line1=title_line1, title_line2=title_line2,
        logo_data_uri=m._logo_data_uri(),
    )
    template = m._JINJA_ENV.get_template("arc_chart.html.j2")
    return template.render(**arc_data, **m._heebo_data_uris())


def main():
    PREVIEW_DIR.mkdir(exist_ok=True)
    bar_path = PREVIEW_DIR / "bar_chart_preview.html"
    arc_path = PREVIEW_DIR / "arc_chart_preview.html"

    bar_path.write_text(_preview_bar_chart_html(), encoding="utf-8")
    arc_path.write_text(_preview_arc_chart_html(), encoding="utf-8")

    print(f"Wrote {bar_path}")
    print(f"Wrote {arc_path}")
    print("Opening in your default browser (re-run this script and refresh "
          "the tab after further template edits)...")
    webbrowser.open(bar_path.resolve().as_uri())
    webbrowser.open(arc_path.resolve().as_uri())


if __name__ == "__main__":
    main()
