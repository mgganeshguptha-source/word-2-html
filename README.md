# Word → Knowledge Hub converter

## One-time setup
    pip install python-docx

## Regenerate the page
    python docx2hub.py BCBSM-Knowledge_Management-Input_Template.docx BCBSM-Knowledge-Hub.html

## Validate without writing a file
    python docx2hub.py BCBSM-Knowledge_Management-Input_Template.docx out.html --check

`hub.css` must sit in the same folder as `docx2hub.py`. All styling lives there —
edit it to change colours or layout without touching the converter.

## The contract

| In Word | Renders as |
|---|---|
| Heading 1 | New section + sidebar nav entry |
| Heading 2 `Group: <name>` | Nav group for the sections that follow |
| First paragraphs after Heading 1 | Grey lead text |
| Table `Field \| Value` | Page header (title, tagline, pills, footer) |
| Table `Value \| Label` | Stat boxes |
| Table `Title \| Description` | Card grid |
| Table `Title \| Description \| Accent` | Cards with accent border |
| Table `Title \| Description \| URL \| Type` | Document link cards |
| Table `Term \| Definition` | Glossary |
| Table `Narration` | Spoken script for the Listen button |
| Numbered list | Auto-numbered step cards |
| Style `Intense Quote` | Orange warning callout |
| Style `Subtle Emphasis` | Small italic footnote |

Accent values: `primary`, `teal`, `orange`, `deep` (case-insensitive).
Type values: `sp` (SharePoint), `ext` (external), `apx` (Apexon internal).

## Images

Paste or insert a picture anywhere in the document — in a paragraph, or inside a
table cell — and it renders in place. Images are embedded directly in the HTML as
base64, so the page stays a single self-contained file.

- **Alt text**: right-click the image > Edit Alt Text. Used as the `alt` attribute.
  A warning is printed for any image without it.
- **Caption**: the alt-text pane's *Title* field, if set, renders as a caption.
- Images wider than 1600px are automatically downscaled to control page weight.
- EMF/WMF images (pasted from PowerPoint or Visio as metafiles) cannot render in a
  browser. The converter warns; re-paste using Paste Special > Picture (PNG).
- Base64 adds roughly 33% to image size. The converter reports the finished page
  size and warns above 3 MB, which matters if you distribute by email.

## News feed

Add a table with headers `Feed | URL | Max` to any section. The converter fetches
each RSS or Atom feed **at build time** and bakes the headlines into the page, so
the output stays a single static file with no runtime dependency.

| Feed | URL | Max |
|---|---|---|
| MI Blue Daily | https://www.bcbsm.mibluedaily.com/categories/our-news/feed | 5 |

- `Max` is optional (defaults to 5 headlines per feed).
- Only the headline, date, source and link are stored — never article body text.
- Headlines from all feeds are merged, de-duplicated and sorted newest first.
- The page shows the retrieval date. News is only as fresh as your last run.
- Results are cached in `news-cache.json` beside the output file. If a feed is
  unreachable (proxy block, site down), the previous headlines are shown and a
  warning printed — the section is never left blank.
- `--offline` skips the network entirely and uses the cache.

## Org charts

Add a table with headers `Name | Title | Reports To`. Each table renders as one
chart, in the position it occupies in the document — so several charts is simply
several tables. Give each its own Heading 1 so it gets a nav entry.

| Name | Title | Reports To |
|---|---|---|
| John Smith | CIO | Sarah Chen |
| Mary Jones | VP, Digital Experience | John Smith |
| David Brown | Director, MPA | Mary Jones |

- **Root**: leave `Reports To` blank, or use a placeholder — `-`, `—`, `N/A`,
  `None`, `.` are all treated as "no manager". The chart starts at that person.
- Naming a manager who is NOT listed as a row creates a dashed placeholder box
  above them. Use a placeholder instead if you don't want that.
- Levels are calculated from `Reports To` — there is no Level column, and box
  colours follow depth automatically.
- Duplicate names and reporting loops are reported as warnings; the chart still
  renders with the bad rows dropped rather than failing.
- Rendered as HTML/CSS, not an image: text wraps, stays searchable with Ctrl+F,
  and is readable by screen readers. Wide charts scroll horizontally; on narrow
  screens the tree stacks vertically.

## Controlling how many boxes per row

By default the converter picks a column count that avoids leaving one item alone
on the last row — 6 stats render 3+3, 4 cards render 2+2.

To force a specific number, put it in brackets after the FIRST header cell:

| Value (4) | Label |
|---|---|
| 1939 | Founded |

That renders 4 across, so six items become 4+2. Works on any grid table:
`Value (4) | Label`, `Title (2) | Description`, `Title (1) | Description | URL | Type`.
Allowed values are 1 to 6; anything else is ignored and auto-balancing applies.

On narrow screens the grid collapses regardless of the authored count.

## Notes

- Hyperlinks made with Ctrl+K are preserved (text + destination). Bare URLs are auto-linked.
- Curly quotes and non-breaking spaces from Word autocorrect are normalised.
- Author instruction paragraphs (the orange boxes) are filtered out by text match.
  If you add your own instruction text, add a pattern to `INSTRUCTION_PATTERNS`.
- A `Narration` block merged into the bottom of another table is detected and
  recovered, with a warning. Prefer a separate one-column table.
- Link rows without a valid `http` URL are skipped with a warning.
