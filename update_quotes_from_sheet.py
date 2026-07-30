#!/usr/bin/env python3
"""
update_quotes_from_sheet.py

Pulls per-party data from the public Google Sheet used by the
party-comparison project and writes it into two separate fields inside
party-comparison-concepts.html's PARTIES array:

  * `quote` / `quoteDate`  <- column P  ("פרטי התחייבות")
    A direct, attributed leader quote (e.g. 'בנט: "..."'). Rendered by
    quoteBlock() in the italic "cf-quote" style.

  * `platform`             <- columns Y + Z
    Y = "יש התייחסות לשיוויון?"  (does the party's platform address
        gender equality?)
    Z = "מה ההתייחסות?"          (what does it say?)
    This is the party's "עמדה במצע" (platform position). It's rendered
    separately from the quote, under the existing
    "האם יש התייחסות לשוויון מגדרי במצע?" heading (see renderQuoteSection
    in the HTML - look for the ".platform-q"/".platform-a" classes).

Sheet: https://docs.google.com/spreadsheets/d/1fLUd-R07JiAkldRabkRJds7eGOuPi78w6JRSqvh4G74/edit

Usage:
    python3 update_quotes_from_sheet.py INPUT_HTML OUTPUT_HTML

    INPUT_HTML  - a local copy of party-comparison-concepts.html
                  (e.g. downloaded from GitHub)
    OUTPUT_HTML - path to write the updated file to

Notes / assumptions baked in (edit SHEET_NAME_TO_ID / TRIVIAL_Y below if
the sheet layout or the HTML's party ids ever change):

  * Column P sometimes holds more than one quote (separated by blank
    lines), each usually formatted as 'Name: "quote text"'. Each line is
    parsed to strip the "Name: " prefix and the wrapping quote marks
    (quoteBlock() adds its own quote marks around the whole value), and
    multiple quotes for the same party are joined with " / ". If P is
    empty for a party, `quote` is set to null (no direct quote known)
    rather than left stale.
  * `platform` = Z, unless Y is a substantive sentence (not just a bare
    "אין"/"לא"/"מעט"/empty), in which case platform = "Y — Z". If both
    Y and Z are empty, `platform` is set to null.
  * `quoteDate` is cleared to '' for every party this script touches,
    since the sheet has no date column for this data.
  * Rows in the sheet with no matching HTML party id (e.g. "יש עתיד",
    which is folded into "ביחד" in the HTML) are skipped and reported.
  * This does NOT push anything to GitHub - it only writes a local
    output file. Review the diff, then commit it yourself (e.g. via
    the GitHub web UI or `git commit`).
"""

import csv
import io
import re
import sys
import urllib.request

# On Windows, the console's default codepage (cp1252 etc.) can't encode the
# Hebrew text in this script's docstring/output. Force stdout/stderr to
# UTF-8 so `python update_quotes_from_sheet.py` (no args) and any Hebrew in
# printed warnings don't crash with UnicodeEncodeError.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding='utf-8')
    except (AttributeError, ValueError):
        pass

SHEET_ID = "1fLUd-R07JiAkldRabkRJds7eGOuPi78w6JRSqvh4G74"
GVIZ_URL = (
    "https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
    "?tqx=out:csv&gid=0&tq=select+A,P,Y,Z"
).format(sheet_id=SHEET_ID)

# Sheet's "שם המפלגה" (column A) -> HTML PARTIES[].id
SHEET_NAME_TO_ID = {
    'עוצמה יהודית': 'bengvir',
    'ישר': 'eisenkot',
    'כחול לבן': 'gantz',
    'ש"ס': 'deri',
    'ביחד': 'bennett',
    'חד"ש תע"ל': 'odeh',
    'הליכוד': 'netanyahu',
    'יהדות התורה': 'goldknopf',
    'רע"ם': 'abbas',
    'הדמוקרטים': 'golan',
    'הציונות הדתית': 'smotrich',
    'ישראל ביתנו': 'liberman',
    # 'יש עתיד' intentionally has no HTML entry - it's folded into 'bennett'.
}

# Y answers this short/empty are treated as "no real content" and dropped
# from the combined platform text (only Z is used in that case).
TRIVIAL_Y = {'', 'אין', 'לא', 'מעט'}

# Matches a single "Name: "quoted text"" line (straight quotes only).
_QUOTE_LINE_RE = re.compile(r'^\s*([^":]{1,25}):\s*"(.*)"\s*$')
# Matches a bare "quoted text" line with no "Name:" prefix.
_BARE_QUOTE_LINE_RE = re.compile(r'^\s*"(.*)"\s*$')


def fetch_sheet_rows():
    """Fetch columns A/P/Y/Z from the public sheet as a list of dict rows."""
    with urllib.request.urlopen(GVIZ_URL) as resp:
        raw = resp.read().decode('utf-8')
    reader = csv.reader(io.StringIO(raw))
    rows = list(reader)
    header, data_rows = rows[0], rows[1:]
    expected = ['שם המפלגה', 'פרטי התחייבות', 'יש התייחסות לשיוויון?', 'מה ההתייחסות?']
    assert header == expected, (
        "Sheet header changed - update the gviz `tq` select or this check: %r" % header
    )
    return [
        {'name': r[0], 'p': r[1].strip(), 'y': r[2].strip(), 'z': r[3].strip()}
        for r in data_rows
    ]


def build_quote(p_text):
    """Turn column P into a clean quote string, or None if empty."""
    if not p_text:
        return None
    parts = []
    for line in p_text.split('\n'):
        line = line.strip()
        if not line:
            continue
        m = _QUOTE_LINE_RE.match(line)
        if m:
            parts.append(m.group(2))
            continue
        m = _BARE_QUOTE_LINE_RE.match(line)
        parts.append(m.group(1) if m else line)
    return ' / '.join(parts) if parts else None


def build_platform(y, z):
    """Turn columns Y+Z into the platform ("עמדה במצע") text, or None."""
    if not y and not z:
        return None
    if not z:
        return y
    if y in TRIVIAL_Y:
        return z
    return "{} — {}".format(y, z)


def js_escape(s):
    return s.replace('\\', '\\\\').replace("'", "\\'").replace('\n', ' ')


def js_value(s):
    return 'null' if s is None else "'{}'".format(js_escape(s))


def patch_html(html_text, data_by_id):
    """Replace quote:/quoteDate: and platform: for each id in data_by_id."""
    missing_quote = []
    missing_platform = []
    for party_id, fields in data_by_id.items():
        # --- quote / quoteDate ---
        quote_re = re.compile(
            r"(id: '" + re.escape(party_id) + r"'[\s\S]*?)"
            r"quote: (?:null|'(?:[^'\\]|\\.)*'), quoteDate: (?:null|'(?:[^'\\]|\\.)*')"
        )
        quote_replacement = "quote: {}, quoteDate: ''".format(js_value(fields['quote']))
        new_text, n = quote_re.subn(
            lambda m: m.group(1) + quote_replacement, html_text, count=1)
        if n == 0:
            missing_quote.append(party_id)
        else:
            html_text = new_text

        # --- platform ---
        platform_re = re.compile(
            r"(id: '" + re.escape(party_id) + r"'[\s\S]*?)"
            r"platform: (?:null|'(?:[^'\\]|\\.)*')"
        )
        platform_replacement = "platform: {}".format(js_value(fields['platform']))
        new_text, n = platform_re.subn(
            lambda m: m.group(1) + platform_replacement, html_text, count=1)
        if n == 0:
            missing_platform.append(party_id)
        else:
            html_text = new_text

    return html_text, missing_quote, missing_platform


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    input_path, output_path = sys.argv[1], sys.argv[2]

    rows = fetch_sheet_rows()

    data_by_id = {}
    unmatched_sheet_rows = []
    for row in rows:
        party_id = SHEET_NAME_TO_ID.get(row['name'])
        if not party_id:
            unmatched_sheet_rows.append(row['name'])
            continue
        data_by_id[party_id] = {
            'quote': build_quote(row['p']),
            'platform': build_platform(row['y'], row['z']),
        }

    with open(input_path, 'r', encoding='utf-8') as f:
        html_text = f.read()

    updated_html, missing_quote, missing_platform = patch_html(html_text, data_by_id)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(updated_html)

    print("Updated quote/platform for {} matched parties.".format(len(data_by_id)))
    if missing_quote:
        print("WARNING: 'quote' field not found for ids: {}".format(missing_quote))
    if missing_platform:
        print("WARNING: 'platform' field not found for ids: {}".format(missing_platform))
    if unmatched_sheet_rows:
        print("Sheet rows with no HTML party mapping (skipped): {}".format(
            unmatched_sheet_rows))
    print("Wrote: {}".format(output_path))


if __name__ == '__main__':
    main()
