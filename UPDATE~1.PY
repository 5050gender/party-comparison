#!/usr/bin/env python3
"""
update_quotes_from_sheet.py

Pulls the "gender equality reference" data (columns Y and Z) from the
public Google Sheet used by the party-comparison project, and writes the
combined text into each party's `quote` / `quoteDate` fields inside the
party-comparison-concepts.html PARTIES array.

Sheet: https://docs.google.com/spreadsheets/d/1fLUd-R07JiAkldRabkRJds7eGOuPi78w6JRSqvh4G74/edit
  Column Y = "יש התייחסות לשיוויון?"  (does the party address gender equality?)
  Column Z = "מה ההתייחסות?"          (what is that reference / detail?)

Usage:
    python3 update_quotes_from_sheet.py INPUT_HTML OUTPUT_HTML

    INPUT_HTML  - a local copy of party-comparison-concepts.html
                  (e.g. downloaded from GitHub, or the file open in your
                  browser's Downloads folder)
    OUTPUT_HTML - path to write the updated file to

Notes / assumptions baked in (edit SHEET_NAME_TO_ID and the "trivial Y"
list below if the sheet or the HTML's party ids ever change):

  * quote = Z, unless Y is a substantive sentence (not just a bare
    "אין"/"לא"/"מעט"/empty), in which case quote = "Y — Z".
  * quoteDate is cleared to '' for every party this script touches,
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

SHEET_ID = "1fLUd-R07JiAkldRabkRJds7eGOuPi78w6JRSqvh4G74"
GVIZ_URL = (
    "https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
    "?tqx=out:csv&gid=0&tq=select+A,Y,Z"
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
# from the combined quote (only Z is used in that case).
TRIVIAL_Y = {'', 'אין', 'לא', 'מעט'}


def fetch_sheet_rows():
    """Fetch column A/Y/Z from the public sheet as a list of dict rows."""
    with urllib.request.urlopen(GVIZ_URL) as resp:
        raw = resp.read().decode('utf-8')
    reader = csv.reader(io.StringIO(raw))
    rows = list(reader)
    header, data_rows = rows[0], rows[1:]
    assert header == ['שם המפלגה', 'יש התייחסות לשיוויון?', 'מה ההתייחסות?'], (
        "Sheet header changed - update the gviz `tq` select or this check: %r" % header
    )
    return [{'name': r[0], 'y': r[1].strip(), 'z': r[2].strip()} for r in data_rows]


def build_quote(y, z):
    if not z:
        return None  # nothing usable for this party
    if y in TRIVIAL_Y:
        return z
    return "{} — {}".format(y, z)


def js_escape(s):
    return s.replace('\\', '\\\\').replace("'", "\\'").replace('\n', ' ')


def patch_html(html_text, quotes_by_id):
    """Replace quote:/quoteDate: for each id in quotes_by_id inside PARTIES."""
    missing = []
    for party_id, quote_text in quotes_by_id.items():
        pattern = re.compile(
            r"(id: '" + re.escape(party_id) + r"'[\s\S]*?)"
            r"quote: (?:null|'(?:[^'\\]|\\.)*'), quoteDate: (?:null|'(?:[^'\\]|\\.)*')"
        )
        replacement = "quote: '{}', quoteDate: ''".format(js_escape(quote_text))
        new_text, n = pattern.subn(lambda m: m.group(1) + replacement, html_text, count=1)
        if n == 0:
            missing.append(party_id)
        else:
            html_text = new_text
    return html_text, missing


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    input_path, output_path = sys.argv[1], sys.argv[2]

    rows = fetch_sheet_rows()

    quotes_by_id = {}
    unmatched_sheet_rows = []
    for row in rows:
        party_id = SHEET_NAME_TO_ID.get(row['name'])
        if not party_id:
            unmatched_sheet_rows.append(row['name'])
            continue
        quote = build_quote(row['y'], row['z'])
        if quote:
            quotes_by_id[party_id] = quote

    with open(input_path, 'r', encoding='utf-8') as f:
        html_text = f.read()

    updated_html, missing_ids = patch_html(html_text, quotes_by_id)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(updated_html)

    print("Updated {} of {} matched parties.".format(
        len(quotes_by_id) - len(missing_ids), len(quotes_by_id)))
    if missing_ids:
        print("WARNING: these HTML party ids were not found in the input file "
              "(id renamed? file already up to date?): {}".format(missing_ids))
    if unmatched_sheet_rows:
        print("Sheet rows with no HTML party mapping (skipped): {}".format(
            unmatched_sheet_rows))
    print("Wrote: {}".format(output_path))


if __name__ == '__main__':
    main()
