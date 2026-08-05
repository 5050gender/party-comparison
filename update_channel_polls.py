#!/usr/bin/env python3
"""
update_channel_polls.py

Refreshes the "סקרים לפי ערוץ" tab of the party-comparison workbook from
themadad.com/polls26's own embedded poll history, and reports the numbers
that should also be patched into party-comparison-concepts.html's
CHANNEL_DATA/CHANNELS constants (see renderChannelTabs()/renderScoreRows()).

How it works
------------
themadad.com/polls26/ embeds its ENTIRE poll history (one record per
published poll, since late 2022) as a plain JSON array inside an inline
<script> tag on the page - no headless browser or clicking through the
site's UI is needed. Each record looks like:

    {"pollNumber": "661", "publisher": "מעריב", "pollster": "מנחם לזר",
     "date": "2026-07-31", "respondents": "501",
     "likud": "22", "utj": "8", ... }

This script:
  1. Downloads the page's HTML.
  2. Finds the `allPolls    = [...]` array literal and extracts it by
     bracket-matching (it's valid JSON - no eval needed).
  3. Groups records by `publisher` (the media channel) and keeps the
     most recent (`date`) record per channel.
  4. Optionally filters to channels whose latest poll is on/after a given
     cutoff date (default: no filter - pass --year to restrict, e.g.
     --year 2026, matching how this tab was first built).
  5. Maps the site's field slugs to the workbook's party names and writes
     them into the "סקרים לפי ערוץ" tab (values only - openpyxl formulas
     in "חישוב לפי ערוץ" and elsewhere are left untouched and will
     recalculate from the new inputs).

Usage
-----
    python3 update_channel_polls.py INPUT_XLSX OUTPUT_XLSX [--year 2026]

    INPUT_XLSX   - local copy of the workbook (e.g. downloaded from GitHub)
    OUTPUT_XLSX  - path to write the updated workbook to
    --year YYYY  - only include channels whose latest poll is in this
                   year or later (default: MIN_YEAR below - polls from
                   before that year are never added; --year can raise
                   this floor further but not lower it)

This does NOT push anything to GitHub, and does NOT touch
party-comparison-concepts.html. After running it:
  1. Open the diff / new "סקרים לפי ערוץ" tab and sanity-check it.
  2. Re-run this project's normal xlsx recalc step.
  3. Regenerate the HTML's CHANNEL_DATA/CHANNELS constants from the new
     "חישוב לפי ערוץ" tab (same COUNTIFS-based women/pct-per-channel
     numbers used when this feature was first built) and commit both
     files, the same way earlier data updates in this project were done.

Field mapping (themadad.com slug -> workbook party name), current as of
2026-08-05 - update here if themadad.com adds/renames fields or if a
tracked party's name changes.
"""

import argparse
import http.cookiejar
import json
import re
import sys
import urllib.request

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding='utf-8')
    except (AttributeError, ValueError):
        pass

POLLS_URL = "https://themadad.com/polls26/"

# This project only tracks polls for the current (2026) election cycle -
# a channel whose latest poll predates this year is dropped entirely rather
# than showing stale pre-cycle numbers. See main()'s --year handling below.
MIN_YEAR = 2026

# themadad.com field slug -> party name as used in this workbook's
# "מועמידים 2026" / "חישוב 2026" tabs. Only fields for parties we track are
# listed; other fields in each poll record (avoda/meretz overlap, yamina,
# economy, tikvahHadash, unifiedArabList, ballad) are ignored.
FIELD_TO_PARTY = [
    ('likud', 'הליכוד'),
    ('utj', 'יהדות התורה'),
    ('shas', 'ש"ס'),
    ('bw', 'כחול לבן'),
    ('hadashTal', 'חדש תע"ל'),
    ('israelBeitanu', 'ישראל ביתנו'),
    ('avoda', 'הדמוקרטים'),        # post-Meretz/Avoda merger, site kept the old slug
    ('smotrich', 'הציונות הדתית'),
    ('raam', 'רע"מ'),
    ('otzma', 'עוצמה יהודית'),
    ('bennett', 'ביחד (בנט-לפיד)'),
    ('eisenkot', 'ישר!'),
    ('miluimnikim', 'טרופר-הנדל'),  # "בית ציוני - המילואימניקים" (Tropper-Hendel), added 2026-08
]

SHEET_NAME = 'סקרים לפי ערוץ'


def fetch_all_polls():
    """Download polls26 and extract the embedded `allPolls` JSON array.

    themadad.com sends a "set a cookie, then redirect back to the same URL"
    response on the first request (consent/bot-check). urllib.request's
    default opener doesn't keep cookies across redirects, so it bounces
    between the same two URLs forever and Python raises
    "HTTPError: HTTP Error 302: ... infinite loop". Using a cookie-jar-backed
    opener lets it carry the cookie through the redirect like a real browser
    would, so it resolves normally.
    """
    cookie_jar = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookie_jar))
    headers = {
        'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                        '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'he-IL,he;q=0.9,en-US;q=0.8,en;q=0.7',
    }
    req = urllib.request.Request(POLLS_URL, headers=headers)
    with opener.open(req, timeout=30) as resp:
        html = resp.read().decode('utf-8', errors='replace')

    # The page contains several unrelated occurrences of the bare word
    # "allPolls" (e.g. references/usages elsewhere in its scripts) - only ONE
    # of them is the actual `const allPolls    = [ ... ]` declaration. Matching
    # on the bare word and grabbing the next '[' after it is fragile: it can
    # lock onto an unrelated, tiny array next to a false-positive match and
    # produce invalid/truncated JSON. Anchor on the assignment itself instead.
    decl_match = re.search(r'allPolls\s*=\s*\[', html)
    if decl_match is None:
        raise RuntimeError("Could not find an 'allPolls = [...]' declaration in the "
                            "page - themadad.com may have changed its page structure.")
    bracket_start = decl_match.end() - 1  # position of the '[' itself
    depth = 0
    started = False
    i = bracket_start
    for i in range(bracket_start, len(html)):
        ch = html[i]
        if ch == '[':
            depth += 1
            started = True
        elif ch == ']':
            depth -= 1
            if started and depth == 0:
                i += 1
                break
    arr_text = html[bracket_start:i]
    return json.loads(arr_text)


def latest_per_channel(all_polls, min_year=None):
    """Group poll records by publisher, keep the most recent per channel."""
    by_publisher = {}
    for p in all_polls:
        pub = p.get('publisher')
        if not pub:
            continue
        if pub not in by_publisher or by_publisher[pub]['date'] < p['date']:
            by_publisher[pub] = p
    if min_year is not None:
        cutoff = "{}-01-01".format(min_year)
        by_publisher = {k: v for k, v in by_publisher.items() if v['date'] >= cutoff}
    return by_publisher


def write_sheet(wb, channel_polls):
    """(Re)write the סקרים לפי ערוץ tab with the given {channel: poll_record} data."""
    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]
    ws = wb.create_sheet(SHEET_NAME)

    headers = ['כלי תקשורת', 'תאריך הסקר', 'סוקר', 'מספר משיבים'] + [p for _, p in FIELD_TO_PARTY]
    ws.append(headers)

    channel_order = sorted(channel_polls.keys(), key=lambda c: channel_polls[c]['date'], reverse=True)
    for chan in channel_order:
        d = channel_polls[chan]
        row = [chan, d.get('date'), d.get('pollster'), d.get('respondents')]
        for field, _ in FIELD_TO_PARTY:
            val = d.get(field)
            row.append(int(val) if val not in (None, '') else None)
        ws.append(row)
    return channel_order


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('input_xlsx')
    ap.add_argument('output_xlsx')
    ap.add_argument('--year', type=int, default=MIN_YEAR,
                     help='Only include channels whose latest poll is in this year or later '
                          '(default: {}, matching this project\'s cutoff - pre-{} polls are '
                          'never included; pass a later year to filter further, not an earlier '
                          'one)'.format(MIN_YEAR, MIN_YEAR))
    args = ap.parse_args()

    if args.year < MIN_YEAR:
        sys.exit("--year {} is before this project's cutoff ({}) - polls before {} are never "
                  "added, so --year can only be {} or later.".format(
                      args.year, MIN_YEAR, MIN_YEAR, MIN_YEAR))

    import openpyxl

    print("Fetching {} ...".format(POLLS_URL))
    all_polls = fetch_all_polls()
    print("Found {} total poll records.".format(len(all_polls)))

    channel_polls = latest_per_channel(all_polls, min_year=args.year)
    print("Latest poll per channel (>= {}):".format(args.year))
    for chan, d in sorted(channel_polls.items(), key=lambda kv: kv[1]['date'], reverse=True):
        print("  {:14s} {}  ({})".format(chan, d['date'], d.get('pollster', '')))

    wb = openpyxl.load_workbook(args.input_xlsx, data_only=False)
    channel_order = write_sheet(wb, channel_polls)
    wb.save(args.output_xlsx)

    print("\nWrote {} channels to '{}' tab in {}".format(len(channel_order), SHEET_NAME, args.output_xlsx))
    print("Next: run this project's recalc step, regenerate the HTML's "
          "CHANNEL_DATA/CHANNELS from the 'חישוב לפי ערוץ' tab, and commit both files.")


if __name__ == '__main__':
    main()
