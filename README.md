# party-comparison

An interactive HTML page comparing Israeli political parties ahead of the 2026 Knesset election, focused on gender representation - current vs. expected number and percentage of women in each party's list.

**Live page:** https://5050gender.github.io/party-comparison/party-comparison-concepts.html

## Files

### party-comparison-concepts.html
The main deliverable - a single self-contained HTML page (party, poll and channel data embedded inline as JS constants: `PARTIES`, `CHANNELS`, `CHANNEL_DATA`, `LOGOS`). Served via GitHub Pages at the live URL above; can also be embedded in an iframe.

### party-comparison-updated-vNN.xlsx
The data workbook backing the page (v16 is current; older versions are kept for history). Key tabs:
- `סקר 2026` - overall seat survey
- `מועמדים 2026` - per-party candidate lists with gender
- `חישוב 2026` - computed current/expected women counts and percentages
- `סקרים לפי ערוץ` - latest poll per media channel
- `חישוב לפי ערוץ` - per-channel computed women counts and percentages

## Scripts

### update_channel_polls.py
Refreshes the `סקרים לפי ערוץ` tab from themadad.com/polls26's embedded poll history.

```
python3 update_channel_polls.py INPUT_XLSX OUTPUT_XLSX [--year YYYY]
```
- `INPUT_XLSX` - local copy of the workbook (e.g. downloaded from GitHub)
- `OUTPUT_XLSX` - path to write the updated workbook to
- `--year YYYY` - only include channels whose latest poll is in this year or later (default: no filter)

Does not touch the HTML or push to GitHub. After running: sanity-check the new tab, regenerate the HTML's `CHANNEL_DATA`/`CHANNELS` from the updated `חישוב לפי ערוץ` tab, and commit both files.

### update_html_from_xlsx.py
Syncs `party-comparison-concepts.html`'s `PARTIES` numeric fields and `CHANNEL_DATA`/`CHANNELS` from the workbook.

```
python3 update_html_from_xlsx.py HTML_IN XLSX_IN HTML_OUT [--skip-channels]
```
- `HTML_IN` - local copy of the current HTML (e.g. downloaded from GitHub)
- `XLSX_IN` - local copy of the current workbook
- `HTML_OUT` - path to write the updated HTML to
- `--skip-channels` - only update the PARTIES numbers, skip regenerating CHANNEL_DATA/CHANNELS

### update_quotes_from_sheet.py
Pulls each party's quote and platform position from the project's [Google Sheet](https://docs.google.com/spreadsheets/d/1fLUd-R07JiAkldRabkRJds7eGOuPi78w6JRSqvh4G74/edit) into the HTML's `quote`, `quoteDate` and `platform` fields.

```
python3 update_quotes_from_sheet.py INPUT_HTML OUTPUT_HTML
```
- `INPUT_HTML` - local copy of the current HTML (e.g. downloaded from GitHub)
- `OUTPUT_HTML` - path to write the updated file to

Does not push to GitHub - only writes a local output file. Review the diff, then commit it yourself.
