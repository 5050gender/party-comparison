# party-comparison

An interactive HTML page comparing Israeli political parties ahead of the 2026 Knesset election, focused on gender representation - current vs. expected number and percentage of women in each party's list.

**Live page:** https://5050gender.github.io/party-comparison/party-comparison-concepts.html

## Files

### party-comparison-concepts.html
The main deliverable - a single self-contained HTML page (party, poll and channel data embedded inline as JS constants: `PARTIES`, `CHANNELS`, `CHANNEL_DATA`, `LOGOS`). Served via GitHub Pages at the live URL above; can also be embedded in an iframe.

### party-comparison-updated-vNN.xlsx
The data workbook backing the page (v40 is current; older versions are kept for history). Key tabs:
- `סקר 2026` - overall seat survey
- `מועמדים 2026` - per-party candidate lists with gender
- `חישוב 2026` - computed current/expected women counts and percentages
- `סקרים לפי ערוץ` - latest poll per media channel
- `חישוב לפי ערוץ` - per-channel computed women counts and percentages

### mapping.csv
Party -> bloc mapping (`אופוזיציה` / `קואליציה`) used by `plot_women_by_outlet.py`'s `--with-pie-charts` and `--create-email-drafts`. A party that's missing here, or mapped to anything other than those two exact values, is drawn as its own gray "unmapped" segment in the arc chart and folded into the גוש השינוי total.

### templates/
HTML/CSS/SVG chart templates used by `plot_women_by_outlet.py` and `preview_charts.py`:
- `bar_chart.html.j2` - the per-party women/men seats bar chart
- `arc_chart.html.j2` - the גוש השינוי vs. גוש ימין-חרדים arc (half-donut) chart
- `assets/logo_5050.jpg` - the 5050 logo, overlaid on both charts
- `assets/text_for_email.txt` - editable subject + body template for `--create-email-drafts` (see below)

Edit these directly to change fonts, colors, spacing, or layout - `plot_women_by_outlet.py` only ever supplies the data that fills them in.

### fonts/
`Heebo-Regular.ttf` / `Heebo-Bold.ttf`, embedded into the rendered charts so they don't depend on the network at render time.

### graphs/
Suggested output folder for the generated chart JPGs (pass `--output-dir graphs/` to `plot_women_by_outlet.py`).

## Scripts

### plot_women_by_outlet.py
Generates the "how many women are expected in the next Knesset" social graphics from the workbook: a bar chart of expected women/men seats per party for each poll in `סקרים לפי ערוץ`/`חישוב לפי ערוץ` plus the aggregate mean-poll estimate, and (optionally) an arc chart of the גוש השינוי / גוש ימין-חרדים bloc split. Can also draft (never send) a per-outlet summary email in Gmail.

```
python plot_women_by_outlet.py [--input-dir DIR] [--output-dir DIR] [--with-pie-charts] [--mapping-csv PATH] [--create-email-drafts] [--gmail-app-password APP_PASSWORD]
```
- `--input-dir` - folder containing `party-comparison-updated-vNN.xlsx` (defaults to the current folder; auto-picks the highest version number found there)
- `--output-dir` - folder to write the `.jpg` files to (defaults to `--input-dir`; try `graphs/`)
- `--with-pie-charts` - also generate the bloc arc chart for each poll (needs `mapping.csv`)
- `--mapping-csv` - path to the party -> bloc mapping CSV (defaults to `mapping.csv` inside `--input-dir`)
- `--create-email-drafts` - also create a Gmail draft (never sent) per outlet, summarizing that poll and attaching its charts, built from `templates/assets/text_for_email.txt`. Needs `mapping.csv` (loaded automatically even without `--with-pie-charts`) and a Gmail App Password (`--gmail-app-password`, or the `GMAIL_APP_PASSWORD` environment variable - safer than passing it on the command line)

Charts are drawn as HTML/CSS/SVG (see `templates/` above) and rendered to JPG with a headless browser (Playwright), not matplotlib - design tweaks belong in the templates, not this script.

Requirements: `pip install pandas openpyxl jinja2 playwright pillow` and `playwright install chromium`.

### preview_charts.py
A fast way to see template edits without a full run. Fills `templates/bar_chart.html.j2` and `templates/arc_chart.html.j2` with realistic sample data (no workbook needed) and writes plain, already-rendered `.html` files - skipping the JPG-rendering step entirely.

```
python preview_charts.py
```
Writes `preview/bar_chart_preview.html` and `preview/arc_chart_preview.html` and opens both in your default browser. Edit a template, re-run, refresh the tab.

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
