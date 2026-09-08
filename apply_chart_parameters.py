"""Applies a "chart parameters" CSV (same shape as chart_template_parameters.csv,
which extract_chart_parameters.py produces) back onto plot_women_by_outlet.py
and templates/*.j2 -- the reverse direction: CSV -> code, instead of code -> CSV.

This lets you tweak colors, sizes, spacing, positions and fixed label text in
a spreadsheet, then apply the whole set back in one pass, instead of hand-
editing CSS/SVG attributes/Python constants directly.

Usage:
    python3 apply_chart_parameters.py PARAMS_CSV [--project-dir DIR] [--dry-run] [--no-backup]

    PARAMS_CSV      A CSV with columns: source_file, section, parameter, value
                     (a "description" column, if present, is ignored). Start
                     from chart_template_parameters.csv, edit the "value"
                     column for whatever you want to change, and pass that
                     file in -- rows you didn't touch are simply re-applied
                     as no-ops.
    --project-dir    Folder containing plot_women_by_outlet.py and templates/
                     (default: this script's own folder).
    --dry-run        Print what would change (unified diffs) without writing
                     anything.
    --no-backup      Skip writing .bak copies of files before overwriting them
                     (backups are written by default).

Value column conventions
-------------------------
* CSS / SVG attribute rows: the value is used verbatim as the property's new
  value text (e.g. "21px", "#702283", "12px 14px 14px").
* Python constants holding a *string* (e.g. SHEET_NAME, ARC_BLOC_CHANGE_LABEL):
  give the plain text, with or without surrounding quotes -- either is
  accepted, and the script writes it back as a properly quoted, escaped
  Python string literal.
* Python constants holding a number/tuple/expression (e.g. ARC_CX, ARC_CY or
  ARC_TITLE_LINE1_POS): give the exact Python source text that belongs after
  the "=" (e.g. "310, 318" or "(ARC_CX, 14)").
* Combined rows (one CSV row covering several attributes/constants, e.g.
  "width x height" or the COL_* column-name bundle): split on the same
  separator shown in the original export ("x", " / ").

Not every row from the export is safely auto-patchable -- a handful describe
*derived* values (e.g. ARC_VIEWBOX_HEIGHT, which is computed from
ARC_TOP_MARGIN + ARC_BOTTOM_MARGIN) or text templates that interleave fixed
text with live Python expressions (e.g. build_bar_subheadline's poll-name/
date interpolation, or the @font-face weight declarations). Those rows are
reported as "skipped (manual edit needed)" with a pointer to where to edit
by hand, rather than risking a wrong guess at the surrounding code.

After patching, both templates are checked with Jinja2's parser and
plot_women_by_outlet.py is checked with Python's own compiler before
anything is written, so a bad edit fails loudly instead of producing a
broken file.
"""
import argparse
import ast
import csv
import difflib
import re
import sys
from pathlib import Path

try:
    from jinja2 import Environment
except ImportError:
    sys.exit("Missing dependency 'jinja2'. Install it with:\n    pip install jinja2")


# --- generic patchers -------------------------------------------------------

class PatchError(Exception):
    """Raised when a row's anchor text can't be found in the target file --
    almost always means the file has drifted from what the CSV assumes."""


def _py_str_literal(value: str, preferred_quote: str = '"') -> str:
    """Turn plain text into a properly-quoted Python string literal, stripping
    one layer of surrounding "..." if the value already has it (so both a
    quoted and unquoted CSV cell work). Prefers `preferred_quote` (normally
    whatever quote character the existing assignment already used, so an
    unchanged value round-trips back to identical source text) but switches
    to the other quote character when the text itself contains the preferred
    one and not the other -- e.g. TOTAL_ROW_LABEL = 'סה"כ' (contains a
    literal double quote, so single quotes are used)."""
    v = value.strip()
    if len(v) >= 2 and v[0] == '"' and v[-1] == '"':
        v = v[1:-1]
    other_quote = "'" if preferred_quote == '"' else '"'
    q = other_quote if (preferred_quote in v and other_quote not in v) else preferred_quote
    escaped = v.replace("\\", "\\\\").replace(q, "\\" + q)
    return f"{q}{escaped}{q}"


def py_const_patch(text: str, lhs: str, new_value: str) -> str:
    """Patches a `LHS = <value>  [# comment]` line -- module-level or
    indented inside a function -- identified by its exact left-hand side
    text (a single NAME or a comma-separated tuple of NAMEs). Preserves
    indentation and any trailing comment. Auto-detects string- vs
    expression-valued constants from the *current* right-hand side (if it
    starts with a quote, the new value is written back as a proper Python
    string literal; otherwise the CSV value is used verbatim as raw source)."""
    # The value alternative tries a quoted string first (so a "#" *inside*
    # the string, e.g. a hex color, isn't mistaken for the start of a
    # trailing "# comment") before falling back to a bare expression.
    pattern = re.compile(
        r"(?m)^([ \t]*" + re.escape(lhs) + r"\s*=\s*)"
        r"(\"(?:[^\"\\]|\\.)*\"|'(?:[^'\\]|\\.)*'|[^#\n]+?)"
        r"(\s*(?:#.*)?)$"
    )
    m = pattern.search(text)
    if not m:
        raise PatchError(f"couldn't find assignment `{lhs} = ...`")
    old_rhs = m.group(2).strip()
    if old_rhs.startswith('"') or old_rhs.startswith("'"):
        new_rhs = _py_str_literal(new_value, preferred_quote=old_rhs[0])
    else:
        new_rhs = new_value.strip()
    return text[: m.start()] + m.group(1) + new_rhs + m.group(3) + text[m.end():]


def css_patch(text: str, selector: str, prop: str, new_value: str) -> str:
    """Patches `prop: <value>;` inside the named CSS rule (`selector { ... }`)
    in a <style> block. `selector` must match the rule's opening line
    exactly (as it's written in the template, e.g. ".bar-men.full")."""
    block_re = re.compile(
        r"(?m)(^[ \t]*" + re.escape(selector) + r"\s*\{)([^}]*)(\})"
    )
    bm = block_re.search(text)
    if not bm:
        raise PatchError(f"couldn't find CSS rule `{selector} {{ ... }}`")
    block = bm.group(2)
    # Negative lookbehind so a short property name (e.g. "color") can't
    # match as the tail of a longer hyphenated one (e.g. "background-color").
    prop_re = re.compile(r"(?<![-\w])(" + re.escape(prop) + r"\s*:\s*)([^;]+)(;)")
    pm = prop_re.search(block)
    if not pm:
        raise PatchError(f"couldn't find `{prop}:` inside `{selector} {{ ... }}`")
    new_block = block[: pm.start()] + pm.group(1) + new_value.strip() + pm.group(3) + block[pm.end():]
    return text[: bm.start()] + bm.group(1) + new_block + bm.group(3) + text[bm.end():]


def css_patch_multi(text: str, selector: str, props: list, new_value: str) -> str:
    """Like css_patch, but writes the same value into several properties at
    once (e.g. a combined "width x height" row where both are equal)."""
    for prop in props:
        text = css_patch(text, selector, prop, new_value)
    return text


def svg_attr_patch(text: str, tag: str, anchor: str, attr: str, new_value: str) -> str:
    """Patches `attr="..."` inside the opening `<tag ...>` that contains the
    given anchor substring (a distinctive bit of the tag's *other* text --
    a Jinja expression like "title_line1_pos[0]" -- unique in the source)."""
    idx = text.find(anchor)
    if idx == -1:
        raise PatchError(f"couldn't find anchor `{anchor}`")
    tag_start = text.rfind(f"<{tag}", 0, idx)
    if tag_start == -1:
        raise PatchError(f"couldn't find enclosing <{tag}> for anchor `{anchor}`")
    tag_end = text.find(">", tag_start)
    if tag_end == -1:
        raise PatchError(f"unterminated <{tag}> tag near anchor `{anchor}`")
    opening = text[tag_start:tag_end]
    attr_re = re.compile(re.escape(attr) + r'="([^"]*)"')
    am = attr_re.search(opening)
    if not am:
        raise PatchError(f"couldn't find `{attr}=\"...\"` on the <{tag}> containing `{anchor}`")
    new_opening = opening[: am.start(1)] + new_value.strip() + opening[am.end(1):]
    return text[:tag_start] + new_opening + text[tag_end:]


def literal_replace(text: str, old_literal: str, new_literal: str) -> str:
    if text.count(old_literal) == 0:
        raise PatchError(f"couldn't find literal text `{old_literal}`")
    return text.replace(old_literal, new_literal, 1)


# --- bespoke handlers for the not-plainly-mechanical rows -------------------

def _bundle(delim):
    def handler(names: list):
        def apply(text: str, new_value: str) -> str:
            pieces = [p.strip() for p in new_value.split(delim)]
            if len(pieces) != len(names):
                raise PatchError(
                    f"expected {len(names)} values separated by \"{delim}\", got {len(pieces)}"
                )
            for name, piece in zip(names, pieces):
                text = py_const_patch(text, name, piece)
            return text
        return apply
    return handler


def _bar_1p8_multiplier(text: str, new_value: str) -> str:
    return literal_replace(text, "x_max = tick_max * 1.8", f"x_max = tick_max * {new_value.strip()}")


def _bar_tick_step(text: str, new_value: str) -> str:
    v = new_value.strip()
    old = "tick_max = max(5, (max_total // 5 + 1) * 5)"
    new = f"tick_max = max({v}, (max_total // {v} + 1) * {v})"
    return literal_replace(text, old, new)


def _bar_headline_literal(text: str, new_value: str) -> str:
    m = re.search(r"(?ms)^def build_bar_headline\(\).*?return\s+(\"[^\"]*\")", text)
    if not m:
        raise PatchError("couldn't find build_bar_headline()'s return statement")
    new_literal = _py_str_literal(new_value)
    return text[: m.start(1)] + new_literal + text[m.end(1):]


def _bar_logo_footer_img_wh(text: str, new_value: str) -> str:
    return css_patch_multi(text, ".logo-footer img", ["width", "height"], new_value)


def _legend_color_wh(text: str, new_value: str) -> str:
    parts = [p.strip() for p in re.split(r"\s+x\s+", new_value)]
    if len(parts) != 2:
        raise PatchError('expected "WIDTH x HEIGHT", e.g. "16px x 10px"')
    text = css_patch(text, ".legend-color", "width", parts[0])
    text = css_patch(text, ".legend-color", "height", parts[1])
    return text


def _legend_labels(text: str, new_value: str) -> str:
    parts = [p.strip() for p in new_value.split(" / ")]
    if len(parts) != 2:
        raise PatchError('expected "WOMEN LABEL / MEN LABEL"')
    women_label, men_label = parts
    m = re.search(r"(<span>)([^<]*)(</span>\s*</div>\s*<div class=\"legend-item\">)", text)
    if not m:
        raise PatchError("couldn't find the legend's women-label <span>")
    text = text[: m.start(2)] + women_label + text[m.end(2):]
    m2 = re.search(r"(legend-color men.*?<span>)([^<]*)(</span>)", text, re.S)
    if not m2:
        raise PatchError("couldn't find the legend's men-label <span>")
    text = text[: m2.start(2)] + men_label + text[m2.end(2):]
    return text


def _logo_callout_text(text: str, new_value: str) -> str:
    # new_value looks like: סה"כ {total_women} נשים -- keep the
    # {total_women} slot as a real Jinja expression either way.
    templated = new_value.replace("{total_women}", "{{ total_women }}").strip()
    count = text.count('סה"כ {{ total_women }} נשים')
    if count == 0:
        raise PatchError("couldn't find the logo callout text")
    return text.replace('סה"כ {{ total_women }} נשים', templated)


def _bloc_rect_wh(text: str, new_value: str) -> str:
    parts = [p.strip() for p in re.split(r"\s+x\s+", new_value)]
    if len(parts) != 2:
        raise PatchError('expected "WIDTH x HEIGHT", e.g. "104 x 76"')
    text = svg_attr_patch(text, "rect", "rect.cx - 52", "width", parts[0])
    text = svg_attr_patch(text, "rect", "rect.cx - 52", "height", parts[1])
    return text


def _sep_solid_width(text: str, new_value: str) -> str:
    return literal_replace(text, 'stroke-width="3"', f'stroke-width="{new_value.strip()}"')


def _sep_dashed(text: str, new_value: str) -> str:
    parts = [p.strip() for p in new_value.split(" / ")]
    if len(parts) != 2:
        raise PatchError('expected "STROKE-WIDTH / DASHARRAY", e.g. "1.5 / \\"5,4\\""')
    width, dash = parts
    dash = dash.strip('"')
    old = 'stroke-width="1.5" stroke-dasharray="5,4"'
    new = f'stroke-width="{width}" stroke-dasharray="{dash}"'
    return literal_replace(text, old, new)


def _arrow_stroke(text: str, new_value: str) -> str:
    parts = [p.strip() for p in new_value.split(" / ")]
    if len(parts) != 2:
        raise PatchError('expected "STROKE COLOR / STROKE-WIDTH"')
    color, width = parts
    text = svg_attr_patch(text, "line", "marker-end", "stroke", color)
    text = svg_attr_patch(text, "line", "marker-end", "stroke-width", width)
    return text


def _skip(reason: str):
    def handler(_text: str, _value: str):
        raise Skip(reason)
    return handler


class Skip(Exception):
    def __init__(self, reason):
        super().__init__(reason)
        self.reason = reason


# --- the registry: (source_file, section, parameter) -> (target_file, handler) ---

BAR = "templates/bar_chart.html.j2"
ARC = "templates/arc_chart.html.j2"
PY = "plot_women_by_outlet.py"

# ARC_COLOR_*/ARC_CONTOUR_* are declared in the .py file, even though the
# CSV lists them under the arc template (that's where they're *used*).
_ARC_COLOR_CONST = {
    "gray / unmapped": "ARC_COLOR_GRAY",
    "opposition men": "ARC_COLOR_OPP_MEN",
    "opposition women": "ARC_COLOR_OPP_WOMEN",
    "coalition women": "ARC_COLOR_COAL_WOMEN",
    "coalition men": "ARC_COLOR_COAL_MEN",
    "opposition-women number outline": "ARC_CONTOUR_OPP_WOMEN",
    "coalition-women number outline": "ARC_CONTOUR_COAL_WOMEN",
}

REGISTRY = {}


def reg(source, section, parameter, target, handler):
    REGISTRY[(source, section, parameter)] = (target, handler)


def reg_css(section, parameter):
    reg(BAR, section, parameter, BAR, lambda t, v, s=section, p=parameter: css_patch(t, s, p, v))


def reg_py(section, parameter):
    reg(PY, section, parameter, PY, lambda t, v, s=section: py_const_patch(t, s, v))


def reg_svg(section, parameter, tag, anchor, attr):
    reg(ARC, section, parameter, ARC,
        lambda t, v, tag=tag, anchor=anchor, attr=attr: svg_attr_patch(t, tag, anchor, attr, v))


# -- bar_chart.html.j2 : plain CSS properties --
for sel, prop in [
    ("body", "font-family"), ("body", "background-color"), ("body", "color"), ("body", "padding"),
    (".chart-wrapper", "width"),
    (".header-area", "margin-bottom"),
    (".header-title", "font-size"), (".header-title", "font-weight"),
    (".header-title", "letter-spacing"), (".header-title", "line-height"),
    (".header-subtitle", "font-size"), (".header-subtitle", "color"),
    (".header-subtitle", "font-weight"), (".header-subtitle", "line-height"),
    (".chart-body", "gap"),
    (".logo-overlay-box", "bottom"), (".logo-overlay-box", "right"),
    (".logo-overlay-text", "font-size"), (".logo-overlay-text", "font-weight"),
    (".logo-overlay-text", "color"), (".logo-overlay-text", "letter-spacing"),
    (".logo-overlay-text", "margin-bottom"), (".logo-overlay-text", "background"),
    (".logo-overlay-text", "padding"), (".logo-overlay-text", "border-radius"),
    (".party-name", "font-size"), (".party-name", "font-weight"), (".party-name", "color"),
    (".party-name", "margin-bottom"), (".party-name", "line-height"),
    (".bar-row", "height"),
    (".bar-women", "background-color"), (".bar-women", "color"), (".bar-women", "font-size"),
    (".bar-women", "border-radius"),
    (".bar-men", "background-color"), (".bar-men", "border-radius"),
    (".bar-men.full", "border-radius"),
    (".total-label", "font-size"), (".total-label", "font-weight"), (".total-label", "margin-left"),
    (".zero-label", "font-size"), (".zero-label", "font-weight"),
    (".legend", "gap"), (".legend", "margin-top"), (".legend", "font-size"), (".legend", "font-weight"),
    (".legend-color.women", "background-color"), (".legend-color.men", "background-color"),
]:
    reg_css(sel, prop)

reg(BAR, ".logo-footer img", "width / height", BAR, _bar_logo_footer_img_wh)
reg(BAR, ".legend-color", "width x height", BAR, _legend_color_wh)
reg(BAR, "fixed text", "legend labels", BAR, _legend_labels)
reg(BAR, "fixed text", "logo callout prefix", BAR, _logo_callout_text)
reg(BAR, "@font-face Heebo", "font-weight (regular)", None,
    _skip("structural -- pairs a font file with its weight, not a tunable value"))
reg(BAR, "@font-face Heebo", "font-weight (bold)", None,
    _skip("structural -- pairs a font file with its weight, not a tunable value"))
reg(BAR, ".logo-footer-text / .logo-overlay-text", "font styling", None,
    _skip("informational note, not a single value"))

# -- arc_chart.html.j2 : SVG <text>/<rect>/<line> attributes --
reg_svg("title_line1 <text>", "font-size", "text", "title_line1_pos[0]", "font-size")
reg_svg("title_line1 <text>", "fill", "text", "title_line1_pos[0]", "fill")
reg_svg("title_line1 <text>", "font-weight", "text", "title_line1_pos[0]", "font-weight")
reg_svg("title_line2 <text>", "font-size", "text", "title_line2_pos[0]", "font-size")
reg_svg("title_line2 <text>", "fill", "text", "title_line2_pos[0]", "fill")
reg_svg("title_line2 <text>", "font-weight", "text", "title_line2_pos[0]", "font-weight")
reg_svg("segment number <text> (seg.num)", "font-size", "text", "seg.num_pos[0]", "font-size")
reg_svg("segment number <text>", "fill", "text", "seg.num_pos[0]", "fill")
reg_svg("segment number <text>", "stroke-width", "text", "seg.num_pos[0]", "stroke-width")
reg_svg("segment sub-label <text> (seg.sub)", "font-size", "text", "seg.sub_pos[0]", "font-size")
reg_svg("segment sub-label <text>", "stroke-width", "text", "seg.sub_pos[0]", "stroke-width")
reg(ARC, "bloc rect", "width x height", ARC, _bloc_rect_wh)
reg_svg("bloc rect", "fill", "rect", "rect.cx - 52", "fill")
reg_svg("bloc rect label <text>", "font-size", "text", "cy + 28", "font-size")
reg_svg("bloc rect total <text>", "font-size", "text", "cy + 60", "font-size")
reg_svg("gray/unmapped label <text>", "font-size", "text", "gray_label_pos[0]", "font-size")
reg_svg("gray/unmapped label <text>", "fill", "text", "gray_label_pos[0]", "fill")
reg(ARC, "separators (solid, between blocs)", "stroke-width", ARC, _sep_solid_width)
reg(ARC, "separators (dashed, within a bloc)", "stroke-width / stroke-dasharray", ARC, _sep_dashed)
reg(ARC, "arrows", "stroke / stroke-width", ARC, _arrow_stroke)
reg_svg("total_label_text <text> (e.g. \"30 חברות כנסת\")", "font-size", "text", "total_label_pos[0]", "font-size")
reg_svg("total_label_text <text>", "font-weight", "text", "total_label_pos[0]", "font-weight")
for param, const in _ARC_COLOR_CONST.items():
    kind = "colors (contour)" if "outline" in param else "colors (fill)"
    reg(ARC, kind, param, PY, (lambda t, v, c=const: py_const_patch(t, c, v)))

# ARC_VIEWBOX_MIN_X / ARC_VIEWBOX_MIN_Y are two SEPARATE assignment lines in
# the source (not a combined tuple line like "ARC_CX, ARC_CY = 310, 318"), so
# this needs the bundle-split handler rather than the generic reg_py() loop.
reg(PY, "ARC_VIEWBOX_MIN_X, ARC_VIEWBOX_MIN_Y", "viewBox origin", PY,
    _bundle(", ")(["ARC_VIEWBOX_MIN_X", "ARC_VIEWBOX_MIN_Y"]))

# -- plot_women_by_outlet.py : constants and small expressions --
for section, parameter in [
    ("ARC_CX, ARC_CY", "arc center point"),
    ("ARC_RO, ARC_RI", "outer / inner radius"),
    ("ARC_TOTAL_SEATS", "fallback seat total"),
    ("ARC_BLOC_CHANGE_LABEL", "left bloc label text"),
    ("ARC_BLOC_COALITION_LABEL", "right bloc label text"),
    ("ARC_VIEWBOX_WIDTH", "viewBox width"),
    ("ARC_TOP_MARGIN", "extra top margin"),
    ("ARC_BOTTOM_MARGIN", "extra bottom margin"),
    ("ARC_TITLE_LINE1_POS", "headline position"),
    ("ARC_TITLE_LINE2_POS", "subtitle position"),
    ("ARC_LOGO_SIZE", "logo size"),
    ("BAR_CHART_CSS_WIDTH", "card width"),
    ("BAR_LOGO_SIZE", "logo size"),
    ("SHEET_NAME", "per-outlet sheet name"),
    ("MEAN_SHEET_NAME", "mean-poll sheet name"),
    ("TOTAL_ROW_LABEL", "totals-row marker"),
]:
    reg_py(section, parameter)

reg(PY, "build_arc_chart_data(): angle_offset", "arrow spread angle", PY,
    lambda t, v: py_const_patch(t, "angle_offset", v))
reg(PY, "build_arc_chart_data(): target_offset", "arrow target x-offset", PY,
    lambda t, v: py_const_patch(t, "target_offset", v))
reg(PY, "build_arc_chart_data(): target_y", "arrow target y", PY,
    lambda t, v: py_const_patch(t, "target_y", v))
reg(PY, "render_bar_chart(): bar headroom multiplier", "x_max = tick_max * 1.8", PY, _bar_1p8_multiplier)
reg(PY, "render_bar_chart(): tick rounding step", "tick_max = max(5, (max_total // 5 + 1) * 5)", PY, _bar_tick_step)
reg(PY, "build_bar_headline()", "headline text", PY, _bar_headline_literal)
reg(PY, "COL_OUTLET / COL_DATE / COL_PARTY / COL_SEATS / COL_WOMEN / COL_MEN", "column headers read", PY,
    _bundle(" / ")(["COL_OUTLET", "COL_DATE", "COL_PARTY", "COL_SEATS", "COL_WOMEN", "COL_MEN"]))
reg(PY, "MAPPING_COL_PARTY / MAPPING_COL_GROUP", "mapping.csv column headers", PY,
    _bundle(" / ")(["MAPPING_COL_PARTY", "MAPPING_COL_GROUP"]))
reg(PY, "OPPOSITION_GROUP_NAME / COALITION_GROUP_NAME", "bloc names in mapping.csv", PY,
    _bundle(" / ")(["OPPOSITION_GROUP_NAME", "COALITION_GROUP_NAME"]))

# -- rows that are derived/computed, or mix fixed text with live expressions --
reg(PY, "ARC_VIEWBOX_HEIGHT", "viewBox height (computed)", None,
    _skip("derived from ARC_TOP_MARGIN + ARC_BOTTOM_MARGIN -- edit those instead"))
reg(PY, "ARC_LOGO_POS", "logo position (computed)", None,
    _skip("derived from ARC_CX/ARC_CY/ARC_LOGO_SIZE -- edit those instead"))
reg(PY, "build_bar_subheadline(outlet, date)", "subtitle template", None,
    _skip("mixes fixed text with the {outlet}/{date} expressions -- "
          "edit build_bar_subheadline()'s return line by hand"))
reg(PY, "build_bar_subheadline_mean()", "subtitle template (mean poll)", None,
    _skip("mixes fixed text with a date expression -- "
          "edit build_bar_subheadline_mean()'s return line by hand"))
reg(PY, "format_date_short()", "date format", None,
    _skip("a strftime() pattern, not free text -- edit format_date_short()'s "
          "strftime string by hand"))
reg(PY, "format_poll_date()", "date format (filenames)", None,
    _skip("built without strftime (no leading zeros) -- edit format_poll_date() by hand"))


# --- driver -------------------------------------------------------------

def load_csv_rows(path: Path) -> list:
    with path.open(newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        required = {"source_file", "section", "parameter", "value"}
        missing = required - set(reader.fieldnames or [])
        if missing:
            sys.exit(f"CSV is missing column(s): {', '.join(sorted(missing))}")
        return list(reader)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("params_csv", type=Path)
    ap.add_argument("--project-dir", type=Path, default=Path(__file__).resolve().parent)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()

    rows = load_csv_rows(args.params_csv)

    file_paths = {
        BAR: args.project_dir / "templates" / "bar_chart.html.j2",
        ARC: args.project_dir / "templates" / "arc_chart.html.j2",
        PY: args.project_dir / "plot_women_by_outlet.py",
    }
    for key, p in file_paths.items():
        if not p.exists():
            sys.exit(f"Expected {key} at {p}, not found. Pass --project-dir if the "
                      f"project lives somewhere else.")

    originals = {key: p.read_text(encoding="utf-8") for key, p in file_paths.items()}
    working = dict(originals)

    applied, unchanged, skipped, errors = [], [], [], []

    for row in rows:
        key = (row["source_file"].strip(), row["section"].strip(), row["parameter"].strip())
        value = (row.get("value") or "").strip()
        label = f'{key[0]} :: {key[1]} :: {key[2]}'

        entry = REGISTRY.get(key)
        if entry is None:
            skipped.append((label, "no rule registered for this row (unrecognized "
                                    "source_file/section/parameter -- check for a typo, "
                                    "or this row was hand-added to the CSV)"))
            continue

        target, handler = entry
        if target is None:
            try:
                handler(None, value)
            except Skip as s:
                skipped.append((label, s.reason))
            continue

        before = working[target]
        try:
            after = handler(before, value)
        except PatchError as e:
            errors.append((label, str(e)))
            continue
        except Skip as s:
            skipped.append((label, s.reason))
            continue

        if after == before:
            unchanged.append(label)
        else:
            working[target] = after
            applied.append(label)

    # Sanity-check every changed file before writing anything.
    jinja_env = Environment()
    for key in (BAR, ARC):
        if working[key] != originals[key]:
            try:
                jinja_env.parse(working[key])
            except Exception as e:
                sys.exit(f"Refusing to write {key}: it no longer parses as valid "
                          f"Jinja2 after patching ({e}). No files were changed.")
    if working[PY] != originals[PY]:
        try:
            ast.parse(working[PY])
        except SyntaxError as e:
            sys.exit(f"Refusing to write {PY}: it no longer parses as valid Python "
                      f"after patching ({e}). No files were changed.")

    changed_files = [k for k in file_paths if working[k] != originals[k]]

    print(f"{len(applied)} value(s) applied, {len(unchanged)} already matched, "
          f"{len(skipped)} skipped, {len(errors)} error(s).\n")

    if changed_files:
        print("Files that would change:" if args.dry_run else "Files changed:")
        for key in changed_files:
            print(f"  {key}")
            diff = difflib.unified_diff(
                originals[key].splitlines(keepends=True),
                working[key].splitlines(keepends=True),
                fromfile=f"{key} (before)", tofile=f"{key} (after)",
            )
            sys.stdout.writelines("    " + line if not line.endswith("\n") else "    " + line
                                   for line in diff)
        print()
    else:
        print("No files need changes.\n")

    if skipped:
        print(f"Skipped ({len(skipped)}) -- not auto-patchable, see note:")
        for label, reason in skipped:
            print(f"  - {label}\n      {reason}")
        print()

    if errors:
        print(f"Errors ({len(errors)}) -- anchor text not found (file may have drifted "
              f"from what this row assumes):")
        for label, reason in errors:
            print(f"  - {label}\n      {reason}")
        print()

    if args.dry_run:
        print("Dry run -- no files were written.")
        return

    if not changed_files:
        return

    for key in changed_files:
        path = file_paths[key]
        if not args.no_backup:
            backup = path.with_suffix(path.suffix + ".bak")
            backup.write_text(originals[key], encoding="utf-8")
            print(f"Backed up {key} -> {backup.name}")
        path.write_text(working[key], encoding="utf-8")
        print(f"Wrote {key}")


if __name__ == "__main__":
    main()
